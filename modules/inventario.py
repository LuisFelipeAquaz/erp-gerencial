import streamlit as st
import pandas as pd
from core.utils import resaltar_stock_critico


def render():
    st.title("📦 Inventario en Tiempo Real")
    df = st.session_state.dfs.get('Inventario', pd.DataFrame())
    if df.empty:
        st.info("No hay datos de Inventario en la matriz.")
    else:
        st.dataframe(df.style.apply(resaltar_stock_critico, axis=1), use_container_width=True, height=500)
