import streamlit as st
import pandas as pd

from core.resumen import mostrar_resumen
from core.utils import filtrar_empresa


def render(empresa_activa):
    st.title("📊 Panel de Control Principal")
    st.write(f"Resumen gerencial de **{empresa_activa}**.")
    df = filtrar_empresa(empresa_activa)
    if df.empty:
        st.warning("No hay ventas para este entorno o periodo. 👈 Ve a 'Carga de Datos' y sube tu reporte, "
                   "o elige 'Todo el periodo' en la barra lateral.")
        return
    mostrar_resumen(df, clave="dash")
    if df["Empresa"].nunique() > 1:
        st.markdown("#### 🏢 Por empresa")
        t = (df.assign(Moneda=df["Moneda"].replace({"PEN": "Soles con IGV", "USD": "Dólares con IGV"}))
             .pivot_table(index="Empresa", columns="Moneda", values="Total_Linea", aggfunc="sum", fill_value=0))
        st.dataframe(t.style.format("{:,.2f}"), use_container_width=True)
