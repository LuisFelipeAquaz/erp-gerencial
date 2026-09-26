import streamlit as st
import pandas as pd


def render(empresa_activa):
    st.title("🏭 Eficiencia de Planta y MRP Predictivo")
    if empresa_activa == "Quimaroma (Tienda Fiori)":
        st.info("⚠️ El módulo de Producción es exclusivo de la fábrica.")
    else:
        df = st.session_state.dfs.get('Produccion', pd.DataFrame())
        if not df.empty:
            c1, c2 = st.columns(2)
            c1.subheader("Volumen Fabricado")
            if 'Cantidad_Producida' in df.columns:
                c1.bar_chart(df.groupby('Producto')['Cantidad_Producida'].sum().sort_values(ascending=False).head(10))
            c2.subheader("Pérdidas por Mermas (S/)")
            if 'Merma_Soles' in df.columns:
                c2.bar_chart(df.groupby('Producto')['Merma_Soles'].sum().sort_values(ascending=False).head(10), color="#ff4b4b")
