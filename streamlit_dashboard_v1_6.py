import sqlite3
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Consolidador v1.6", layout="wide")
st.title("Consolidador v1.6 – Ventas • Cartera • Alertas")

@st.cache_data(ttl=120)
def read_table(db, name):
    con = sqlite3.connect(db)
    try:
        return pd.read_sql(f"SELECT * FROM {name}", con)
    finally:
        con.close()

cand = [Path("salida_app/consolidado.db"), Path("./consolidado.db")]
default = ""
for c in cand:
    if c.exists():
        default = str(c.resolve()); break
db_path = st.text_input("Ruta a SQLite (consolidado.db)", value=default)
if not db_path: st.stop()
dbp = Path(db_path)
if not dbp.exists():
    st.error(f"No encuentro la DB en: {dbp}"); st.stop()

tabs = st.tabs(["Ventas","Cartera","Alertas"])
ventas = read_table(str(dbp), "ventas")
cartera = read_table(str(dbp), "cartera")

with tabs[0]:
    st.subheader("Ventas (resumen rápido)")
    # detectar subtotal/total
    vcols = {c.lower(): c for c in ventas.columns}
    usar = vcols.get("total") or next((v for k,v in vcols.items() if "total" in k), None)
    if usar:
        ventas[usar] = pd.to_numeric(ventas[usar], errors="coerce")
        st.metric("Ventas totales", f"{ventas[usar].sum():,.2f}")
    # Formatear columnas numéricas
    ventas_fmt = ventas.copy()
    for col in ventas_fmt.select_dtypes(include=['float', 'int']).columns:
        ventas_fmt[col] = ventas_fmt[col].map(lambda x: f"{x:,.2f}" if pd.notnull(x) else "")
    st.dataframe(ventas_fmt.head(200))

with tabs[1]:
    st.subheader("Cartera (resumen rápido)")
    ccols = {c.lower(): c for c in cartera.columns}
    tot = next((v for k,v in ccols.items() if "total" in k or " saldo" in k or k=="saldo"), None)
    if tot:
        cartera[tot] = pd.to_numeric(cartera[tot], errors="coerce")
        st.metric("Cartera total", f"{cartera[tot].sum():,.2f}")
    cartera_fmt = cartera.copy()
    for col in cartera_fmt.select_dtypes(include=['float', 'int']).columns:
        cartera_fmt[col] = cartera_fmt[col].map(lambda x: f"{x:,.2f}" if pd.notnull(x) else "")
    st.dataframe(cartera_fmt.head(200))

with tabs[2]:
    st.subheader("Alertas (clientes con mora y con facturación)")
    # unión simple por nombre normalizado
    def norm(s):
        try:
            import unidecode; return s.astype(str).map(lambda x: unidecode.unidecode(x).strip().lower())
        except:
            return s.astype(str).str.strip().str.lower()
    v_name = next((c for c in ventas.columns if "raz" in c.lower() or "cliente" in c.lower() or "nombre" in c.lower()), None)
    c_name = next((c for c in cartera.columns if "raz" in c.lower() or "cliente" in c.lower() or "nombre" in c.lower()), None)
    v_total = next((c for c in ventas.columns if "total" in c.lower()), None)
    overdue_cols = [c for c in cartera.columns if any(k in c.lower() for k in ["30","60","90","120",">120"])]
    if v_name and c_name and v_total and overdue_cols:
        vv = ventas.copy(); vv["monto"] = pd.to_numeric(vv[v_total], errors="coerce").fillna(0.0)
        vv["_k"] = norm(vv[v_name])
        cc = cartera.copy(); cc["overdue"] = cc[overdue_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
        cc = cc.groupby(c_name, as_index=False)["overdue"].sum(); cc["_k"] = norm(cc[c_name])
        alert = vv.merge(cc[["_k","overdue"]], on="_k", how="inner")
        res = alert.groupby(v_name, as_index=False).agg(facturas=("monto","size"), monto=("monto","sum"), vencido=("overdue","max"))
        st.metric("Clientes con alerta", f"{len(res):,}")
        res_fmt = res.copy()
        for col in res_fmt.select_dtypes(include=['float', 'int']).columns:
            res_fmt[col] = res_fmt[col].map(lambda x: f"{x:,.2f}" if pd.notnull(x) else "")
        st.dataframe(res_fmt.sort_values("monto", ascending=False).head(100))
    else:
        st.info("No se detectaron columnas para generar alertas.")
