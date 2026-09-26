"""
Resumen de ventas por mes, siempre separado en SOLES y DÓLARES (nunca se mezclan),
y selector global de periodo (todo o un mes específico) en la barra lateral.
"""
import pandas as pd
import plotly.express as px
import streamlit as st

from core.utils import convert_to_excel, filtrar_empresa

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
TODO = "TODO"
SIMBOLO = {"PEN": "S/", "USD": "US$"}
NOMBRE_MONEDA = {"PEN": "Soles", "USD": "Dólares"}
MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def etiqueta_mes(clave: str) -> str:
    """'2026-09' -> 'Septiembre 2026'."""
    if clave == TODO:
        return "📅 Todo el periodo"
    try:
        anio, mes = clave.split("-")
        return f"{MESES[int(mes) - 1]} {anio}"
    except (ValueError, IndexError):
        return clave


def selector_periodo(empresa_activa: str) -> str:
    """Selector en la barra lateral. Guarda la elección en st.session_state['periodo']."""
    df = filtrar_empresa(empresa_activa, aplicar_periodo=False)
    meses = []
    if not df.empty and "Fecha" in df.columns:
        meses = sorted(df["Fecha"].dropna().dt.strftime("%Y-%m").unique(), reverse=True)
    opciones = [TODO] + list(meses)
    actual = st.session_state.get("periodo", TODO)
    if actual not in opciones:
        actual = TODO
    elegido = st.sidebar.selectbox("📅 PERIODO A REVISAR:", opciones, index=opciones.index(actual),
                                   format_func=etiqueta_mes,
                                   help="Se aplica a todos los módulos de ventas. "
                                        "Retención de clientes siempre usa todo el historial.")
    st.session_state["periodo"] = elegido
    return elegido


def _moneda(df: pd.DataFrame) -> pd.Series:
    return df["Moneda"].fillna("PEN").replace({"": "PEN"}) if "Moneda" in df.columns else pd.Series("PEN", index=df.index)


def tabla_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por mes con soles y dólares por separado (con y sin IGV) + fila TOTAL."""
    d = df.copy()
    d["_mon"] = _moneda(d)
    d["_mes"] = d["Fecha"].dt.strftime("%Y-%m").fillna("Sin fecha")
    filas = []
    for mes, g in d.groupby("_mes"):
        fila = {"Mes": etiqueta_mes(mes) if mes != "Sin fecha" else mes, "_orden": mes}
        for mon in ["PEN", "USD"]:
            gm = g[g["_mon"] == mon]
            fila[f"{NOMBRE_MONEDA[mon]} con IGV"] = gm["Total_Linea"].sum()
            fila[f"{NOMBRE_MONEDA[mon]} sin IGV"] = gm["Venta_Neta"].sum()
        fila["Comprobantes"] = g.loc[g["Venta_Neta"] > 0, "Comprobante"].nunique() if "Comprobante" in g else 0
        fila["Clientes"] = g.loc[g["Venta_Neta"] > 0, "Cliente_ID"].nunique()
        filas.append(fila)
    t = pd.DataFrame(filas).sort_values("_orden").drop(columns="_orden")
    total = {"Mes": "TOTAL"}
    for c in t.columns[1:]:
        total[c] = t[c].sum()
    ventas = d[d["Venta_Neta"] > 0]
    total["Comprobantes"] = ventas["Comprobante"].nunique() if "Comprobante" in ventas else 0
    total["Clientes"] = ventas["Cliente_ID"].nunique()
    t = pd.concat([t, pd.DataFrame([total])], ignore_index=True)
    if t[["Dólares con IGV", "Dólares sin IGV"]].abs().sum().sum() == 0:  # sin dólares: no mostrar columnas vacías
        t = t.drop(columns=["Dólares con IGV", "Dólares sin IGV"])
    return t


def mostrar_resumen(df: pd.DataFrame, clave: str = "res", con_grafico: bool = True):
    """Bloque principal: totales por moneda, tabla por mes y gráfico. Nunca suma soles con dólares."""
    if df.empty:
        return
    periodo = st.session_state.get("periodo", TODO)
    d = df.copy()
    d["_mon"] = _moneda(d)
    fechas = d["Fecha"].dropna()
    rango = f"del {fechas.min():%d/%m/%Y} al {fechas.max():%d/%m/%Y}" if not fechas.empty else ""
    st.markdown(f"#### 📌 Resumen de ventas · {etiqueta_mes(periodo).replace('📅 ', '')} {rango}")

    monedas = [m for m in ["PEN", "USD"] if d.loc[d["_mon"] == m, "Total_Linea"].abs().sum() > 0] or ["PEN"]
    cols = st.columns(len(monedas) * 2 + 2)
    i = 0
    for mon in monedas:
        g = d[d["_mon"] == mon]
        s = SIMBOLO[mon]
        cols[i].metric(f"Ventas en {NOMBRE_MONEDA[mon].lower()} (con IGV)", f"{s} {g['Total_Linea'].sum():,.2f}",
                       help="Ya descuenta las notas de crédito (devoluciones).")
        cols[i + 1].metric(f"{NOMBRE_MONEDA[mon]} sin IGV", f"{s} {g['Venta_Neta'].sum():,.2f}")
        i += 2
    ventas = d[d["Venta_Neta"] > 0]
    cols[i].metric("Comprobantes", f"{ventas['Comprobante'].nunique():,}" if "Comprobante" in ventas else "—")
    cols[i + 1].metric("Clientes", f"{ventas['Cliente_ID'].nunique():,}")

    tabla = tabla_mensual(d)
    vista = tabla.copy()
    for c in vista.columns:
        if c.startswith("Soles"):
            vista[c] = vista[c].map(lambda x: f"S/ {x:,.2f}")
        elif c.startswith("Dólares"):
            vista[c] = vista[c].map(lambda x: f"US$ {x:,.2f}")
    with st.expander("📅 Detalle por mes (soles y dólares por separado)", expanded=True):
        st.dataframe(vista, use_container_width=True, hide_index=True)
        c1, c2 = st.columns([1, 3])
        c1.download_button("📥 Descargar resumen (.xlsx)", data=convert_to_excel(tabla, sheet_name="Resumen mensual"),
                           file_name="resumen_ventas_por_mes.xlsx", mime=MIME_XLSX, key=f"dl_resumen_{clave}")
    if con_grafico and len(tabla) > 2:
        with st.expander("📊 Ver gráfico por mes", expanded=False):
            graf = tabla[tabla["Mes"] != "TOTAL"].melt(id_vars="Mes", value_vars=[c for c in tabla.columns if "con IGV" in c],
                                                        var_name="Moneda", value_name="Venta con IGV")
            fig = px.bar(graf, x="Mes", y="Venta con IGV", color="Moneda", barmode="group", text_auto=",.0f",
                         title="Ventas por mes (con IGV)")
            fig.update_layout(height=360, legend_title_text="", xaxis_title="", margin=dict(t=50, b=10))
            st.plotly_chart(fig, use_container_width=True, key=f"graf_resumen_{clave}")
    if "USD" in monedas:
        st.caption("Soles y dólares se muestran siempre por separado: no se suman entre sí.")
