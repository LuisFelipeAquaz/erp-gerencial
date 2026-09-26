import streamlit as st
import pandas as pd


def render(empresa_activa):
    st.title("📊 Panel de Control Principal")
    st.write(f"Resumen gerencial de **{empresa_activa}**.")
    if st.session_state.dfs.get('Ventas', pd.DataFrame()).empty:
        st.warning("Base de datos conectada. 👈 Ve a 'Carga de Datos' y procesa tus archivos por primera vez para subirlos a la nube.")
    else:
        st.success("✅ Base de datos en la nube perfectamente sincronizada.")
