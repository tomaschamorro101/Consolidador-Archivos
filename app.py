#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse, re, warnings, sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd
warnings.filterwarnings("ignore", message="Cannot parse header or footer", category=UserWarning, module="openpyxl")

DEFAULT_HEADERS = {"personas": 6, "cartera": 7, "ventas": 7}
DAYFIRST_DEFAULT = True

def log(msg, logfile):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def _make_unique_columns(columns):
    seen = {}; out = []
    for c in columns:
        c = str(c)
        if c in seen:
            seen[c] += 1; out.append(f"{c}__{seen[c]}")
        else:
            seen[c] = 0; out.append(c)
    return out

def detect_header_row(df, lookahead=30):
    pat = re.compile(r"(ruc|c[eé]dula|identif|cliente|raz[oó]n|factura|fecha|documento|subtotal|total|vendedor|saldo|por vencer|30|60|90|120)", re.I)
    best, best_score = 0, -1
    n = min(len(df), max(lookahead, 5))
    for i in range(n):
        row = df.iloc[i]; score = 0
        for v in row.values:
            if isinstance(v, str) and pat.search(v or ""): score += 2
        score += row.notna().sum() * 0.1
        if score > best_score: best, best_score = i, score
    return best

def read_table_auto(path, auto_header=True, forced_header_1based=None):
    raw = pd.read_excel(path, header=None, dtype=str)
    if not auto_header and forced_header_1based:
        header_idx = max(0, int(forced_header_1based)-1)
    else:
        header_idx = detect_header_row(raw)
    cols = (raw.iloc[header_idx].astype(str).str.strip()
              .str.replace(r"\s+", " ", regex=True)
              .str.replace(r"[^\w\s/áéíóúÁÉÍÓÚñÑ#\.]", "", regex=True))
    df = raw.iloc[header_idx+1:].copy()
    df.columns = cols
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed", case=False)]
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    df.columns = (df.columns.astype(str).str.strip()
                    .str.replace(r"\s+", " ", regex=True)
                    .str.replace(r"[^\w\s/áéíóúÁÉÍÓÚñÑ#\.]", "", regex=True))
    df.columns = _make_unique_columns(list(df.columns))
    return df, header_idx+1

def digits_only(x):
    if pd.isna(x): return None
    import re as _re
    s = _re.sub(r"\D", "", str(x))
    return s if s else None

def build_id_unico(df):
    import re as _re
    id_cols = [c for c in df.columns if _re.search(r"(ruc|c[eé]dula|identificaci[oó]n|ci/ruc|ci\b|nro\s*id|numero.*(doc|id))", str(c), re.I)]
    def to_id(row):
        for c in id_cols:
            d = digits_only(row.get(c))
            if d and len(d) >= 13: return d[:10]
        for c in id_cols:
            d = digits_only(row.get(c))
            if d and len(d) == 10: return d
        return None
    df = df.copy(); df["ID_UNICO"] = df.apply(to_id, axis=1); return df

def normalize_amounts(df):
    import re as _re
    df = df.copy()
    pat = _re.compile(r"(total|subtotal|\biva\b|monto|valor|saldo|por vencer|30|60|90|120|>120)", re.I)
    for col in list(df.columns):
        if pat.search(str(col)):
            s = df[col].astype(str)
            s = (s.str.replace(r"[^\d\-,.]", "", regex=True)
                   .str.replace(".", "", regex=False)
                   .str.replace(",", ".", regex=False))
            df[col] = pd.to_numeric(s, errors="coerce")
    return df

def parse_dates(df, dayfirst=True):
    import re as _re
    date_cols = [c for c in df.columns if _re.search(r"(fecha|emisi[oó]n|venc|mes|a[ñn]o)", str(c), re.I)]
    for c in date_cols:
        s = df[c]
        try:
            df[c] = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst)
        except Exception:
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
                try:
                    df[c] = pd.to_datetime(s, format=fmt, errors="coerce"); break
                except Exception: continue
    return df

def pick_cols(df, patterns):
    import re as _re
    cols = []
    for p in patterns: cols.extend([c for c in df.columns if _re.search(p, c, re.I)])
    seen, ordered = set(), []
    for c in cols:
        if c not in seen: seen.add(c); ordered.append(c)
    if "ID_UNICO" not in ordered: ordered = ["ID_UNICO"] + ordered
    ordered = [c for c in ordered if c in df.columns]
    return df[ordered].copy()

def add_prefix(df, prefix):
    df = df.copy()
    df.columns = [col if col == "ID_UNICO" else f"{prefix}_{col}" for col in df.columns]
    return df

def export_sqlite(sqlite_path, personas, cartera, ventas, consolidado):
    con = sqlite3.connect(sqlite_path)
    try:
        personas.to_sql("personas", con, if_exists="replace", index=False)
        cartera.to_sql("cartera", con, if_exists="replace", index=False)
        ventas.to_sql("ventas", con, if_exists="replace", index=False)
        consolidado.to_sql("consolidado", con, if_exists="replace", index=False)
    finally:
        con.close()

def list_tables(sqlite_path):
    con = sqlite3.connect(sqlite_path)
    try:
        return [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    finally:
        con.close()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", default="salida_app")
    ap.add_argument("--headers", nargs=3, type=int, help="Filas 1-based: personas cartera ventas")
    ap.add_argument("--no-auto-header", action="store_true")
    ap.add_argument("--dayfirst", action="store_true")
    ap.add_argument("--no-sqlite", action="store_true")
    args = ap.parse_args()

    script_dir = Path(__file__).resolve().parent
    out_dir = Path(args.output)
    if not out_dir.is_absolute(): out_dir = script_dir / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    logfile = str(out_dir / "etl_log.txt")

    log("===== INICIO APP v1.5 (patch FIX2) =====", logfile)
    dayfirst_effective = args.dayfirst or DAYFIRST_DEFAULT
    log(f"dayfirst={dayfirst_effective}", logfile)

    per_path = script_dir / "Personas.xlsx"
    car_path = script_dir / "CarteraPorCobrar.xlsx"
    ven_path = script_dir / "Ventas Personalizado.xlsx"
    for label, p in [("Personas", per_path), ("Cartera", car_path), ("Ventas", ven_path)]:
        if not p.exists():
            raise FileNotFoundError(f"No se encontró el archivo requerido para {label}: {p}")

    log(f"Personas: {per_path}", logfile)
    log(f"Cartera:  {car_path}", logfile)
    log(f"Ventas:   {ven_path}", logfile)

    headers = DEFAULT_HEADERS.copy()
    if args.headers:
        headers = {"personas": args.headers[0], "cartera": args.headers[1], "ventas": args.headers[2]}
        auto_header = False if args.no_auto_header else True
    else:
        auto_header = True if not args.no_auto_header else False

    personas, per_hdr = read_table_auto(per_path, auto_header=auto_header, forced_header_1based=headers["personas"] if not auto_header else None)
    cartera,  car_hdr = read_table_auto(car_path,  auto_header=auto_header, forced_header_1based=headers["cartera"] if not auto_header else None)
    ventas,   ven_hdr = read_table_auto(ven_path,   auto_header=auto_header, forced_header_1based=headers["ventas"] if not auto_header else None)
    log(f"Header (1-based): Personas={per_hdr} Cartera={car_hdr} Ventas={ven_hdr}", logfile)

    personas = build_id_unico(personas)
    cartera  = parse_dates(normalize_amounts(build_id_unico(cartera)), dayfirst=dayfirst_effective)
    ventas   = parse_dates(normalize_amounts(build_id_unico(ventas)),  dayfirst=dayfirst_effective)

    personas_l = pick_cols(personas, [r"\bruc\b", r"c[eé]dula", r"raz[oó]n", r"nombre comercial", r"categor", r"vendedor", r"cr[eé]dito|dias"])
    cartera_l  = pick_cols(cartera,  [r"(cliente|raz[oó]n|ruc|c[eé]dula)", r"(doc|factura)", r"(fecha|venc)", r"(por vencer|30|60|90|120|>120|saldo|total|valor documento|cobros|retencion)"])
    ventas_l   = pick_cols(ventas,   [r"(cliente|raz[oó]n|ruc|c[eé]dula)", r"(factura|doc|#)", r"(fecha)", r"\bsubtotal\b|sub\.", r"\biva\b", r"\btotal\b", r"(producto|cantidad|bodega|vendedor)"])

    p_pref = add_prefix(personas_l, "PER")
    c_pref = add_prefix(cartera_l,  "CAR")
    v_pref = add_prefix(ventas_l,   "VEN")
    cons = p_pref.merge(c_pref, on="ID_UNICO", how="outer").merge(v_pref, on="ID_UNICO", how="outer")

    (out_dir/"Personas_Limpio.csv").write_text(personas.to_csv(index=False, encoding="utf-8-sig"))
    (out_dir/"Cartera_Limpio.csv").write_text(cartera.to_csv(index=False, encoding="utf-8-sig"))
    (out_dir/"Ventas_Limpio.csv").write_text(ventas.to_csv(index=False, encoding="utf-8-sig"))
    (out_dir/"Consolidado.csv").write_text(cons.to_csv(index=False, encoding="utf-8-sig"))

    with pd.ExcelWriter(out_dir/"Consolidado_APP_v1_patch.xlsx", engine="openpyxl") as w:
        p_pref.to_excel(w, sheet_name="Personas", index=False)
        c_pref.to_excel(w, sheet_name="Cartera", index=False)
        v_pref.to_excel(w, sheet_name="Ventas", index=False)
        cons.to_excel(w, sheet_name="Consolidado", index=False)

    if not args.no_sqlite:
        sqlite_path = out_dir/"consolidado.db"
        export_sqlite(str(sqlite_path), personas, cartera, ventas, cons)
        log(f"SQLite exportado: {Path(sqlite_path).resolve()}", logfile)
        tabs = list_tables(str(sqlite_path))
        log(f"Tablas en DB: {tabs}", logfile)
        con = sqlite3.connect(sqlite_path)
        try:
            for t in ["personas","cartera","ventas","consolidado"]:
                try:
                    cols = [r[1] for r in con.execute(f"PRAGMA table_info({t})")]
                    log(f"Columnas {t}: {cols[:15]}{' ...' if len(cols)>15 else ''}", logfile)
                except Exception as e:
                    log(f"(Aviso) No se pudo leer columnas de {t}: {e}", logfile)
        finally:
            con.close()

    log("===== FIN APP v1.5 (patch FIX2) =====", logfile)

if __name__ == "__main__":
    main()
