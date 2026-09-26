import streamlit as st
import pandas as pd

from core.utils import filtrar_empresa, cobertura_costos


def render(empresa_activa):
    st.title("⚖️ Balance Financiero y Costos Estratégicos")
    df_v = filtrar_empresa(empresa_activa)  # copia con Venta_Neta sin IGV y notas de crédito restadas

    gastos = st.session_state.dfs
    if "Aquaz" in empresa_activa:
        df_g = gastos.get('Gastos_Aquaz', pd.DataFrame())
    elif "Quimaroma" in empresa_activa:
        df_g = gastos.get('Gastos_Quima', pd.DataFrame())
    else:
        df_g = pd.concat([gastos.get('Gastos_Aquaz', pd.DataFrame()), gastos.get('Gastos_Quima', pd.DataFrame())])

    # Las mermas de planta son de Aquaz: no se restan a Quimaroma
    df_p = pd.DataFrame() if "Quimaroma" in empresa_activa else gastos.get('Produccion', pd.DataFrame())

    hay_costos = not df_v.empty and df_v['Utilidad_Bruta'].notna().any()
    ut_bruta = df_v['Utilidad_Bruta'].sum() if hay_costos else 0
    cob = cobertura_costos(df_v)
    if not df_g.empty and 'Monto' not in df_g.columns:
        st.warning("El Excel de gastos no tiene una columna 'Monto': los gastos fijos aparecen en 0.")
    g_fijos = pd.to_numeric(df_g['Monto'], errors='coerce').sum() if not df_g.empty and 'Monto' in df_g.columns else 0
    mermas = pd.to_numeric(df_p['Merma_Soles'], errors='coerce').sum() if not df_p.empty and 'Merma_Soles' in df_p.columns else 0
    resultado = ut_bruta - g_fijos - mermas

    st.markdown("### Resumen Consolidado")
    st.caption("Utilidad comercial = venta sin IGV − costo de lo vendido (notas de crédito descontadas).")
    c1, c2, c3, c4 = st.columns(4)
    if not hay_costos:
        st.info("Todavía no hay costos cargados, así que la utilidad y el resultado neto quedan en blanco. "
                f"Venta neta registrada: S/ {df_v['Venta_Neta'].sum() if not df_v.empty else 0:,.2f} (sin IGV).")
    elif cob < 0.999:
        st.warning(f"Solo el {cob:.0%} de la venta tiene costo: la utilidad y el resultado son parciales.")
    c1.metric("1. Utilidad Comercial", f"S/ {ut_bruta:,.2f}" if hay_costos else "—")
    c2.metric("2. Pérdidas Planta", f"- S/ {mermas:,.2f}")
    c3.metric("3. Gastos Fijos", f"- S/ {g_fijos:,.2f}")
    c4.metric("💰 RESULTADO NETO", f"S/ {resultado:,.2f}" if hay_costos else "—")
