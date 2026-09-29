import streamlit as st
import pandas as pd
import plotly.express as px
from core.utils import convert_to_excel, filtrar_empresa, cobertura_costos, texto_utilidad
from core.resumen import mostrar_resumen
from modules import productos_vendidos


def render(empresa_activa):
    st.title("💰 Análisis de Ventas y Fidelización")
    # Copia preparada: Venta_Neta (sin IGV, con notas de crédito) y Utilidad_Bruta ya calculadas
    df = filtrar_empresa(empresa_activa)

    if df.empty:
        st.info(f"Sube datos de ventas para ver la analítica.")
    else:
        mostrar_resumen(df, clave="analitica")
        st.markdown("---")
        productos_vendidos.render(df)
        st.markdown("---")
        st.subheader("📈 Vendedores y clientes")
        # Los gráficos se hacen en una sola moneda para no mezclar soles con dólares
        monedas = [m for m in ["PEN", "USD"] if (df["Moneda"] == m).any()]
        if len(monedas) > 1:
            mon = st.radio("Moneda de los gráficos:", monedas, horizontal=True,
                           format_func=lambda m: "Soles" if m == "PEN" else "Dólares")
            df = df[df["Moneda"] == mon]
        simbolo = "US$" if (df["Moneda"] == "USD").all() else "S/"
        cob = cobertura_costos(df)
        v2, v3 = st.columns(2)
        v2.metric("Unidades", f"{df['Cantidad'].sum():,.0f}")
        v3.metric("Utilidad bruta", texto_utilidad(df).replace("S/", simbolo))
        if cob == 0:
            st.caption("Sin costos cargados: la utilidad queda en blanco y los gráficos se miden por venta.")
        elif cob < 0.999:
            st.caption(f"⚠️ Solo el {cob:.0%} de la venta tiene costo: la utilidad es parcial.")

        with st.expander("📊 Ver gráficos de vendedores y clientes", expanded=False):
            opciones = ["Venta neta (sin IGV)", "Venta con IGV"] + (["Utilidad bruta"] if cob > 0 else [])
            medida = st.radio("Medir por:", opciones, horizontal=True)
            m = {"Venta neta (sin IGV)": "Venta_Neta", "Venta con IGV": "Total_Linea"}.get(medida, "Utilidad_Bruta")

            # --- SECCIÓN DE GRÁFICOS PERSONALIZADOS ---
            st.markdown("### ⚙️ Personalización de Gráficos")
            tipo_grafico = st.radio("Elige el estilo visual:", ["📊 Gráfico de Barras", "🍩 Gráfico Circular (Porcentajes)", "📈 Gráfico de Líneas"], horizontal=True)

            c1, c2 = st.columns(2)

            # --- 1. PROCESAMIENTO VENDEDORES ---
            df_vend = df.groupby('Vendedor')[m].sum().reset_index()
            df_vend = df_vend.sort_values(by=m, ascending=False)
            total_vend = df_vend[m].sum()
            df_vend['Porcentaje (%)'] = (df_vend[m] / total_vend) * 100 if total_vend > 0 else 0

            c1.subheader(f"Rendimiento por Vendedor ({medida.lower()})")
            if df_vend["Vendedor"].nunique() == 1:
                c1.caption(f"Todas las ventas figuran como '{df_vend['Vendedor'].iloc[0]}': este reporte no trae el vendedor.")
            if "Barras" in tipo_grafico:
                fig1 = px.bar(df_vend, x='Vendedor', y=m, text=df_vend['Porcentaje (%)'].apply(lambda x: f'{x:.1f}%'))
            elif "Circular" in tipo_grafico:
                fig1 = px.pie(df_vend, names='Vendedor', values=m, hole=0.4)
            else:
                fig1 = px.line(df_vend, x='Vendedor', y=m, markers=True)

            fig1.update_layout(margin=dict(t=20, b=20, l=0, r=0))
            c1.plotly_chart(fig1, use_container_width=True)
            # Botón para descargar el Excel de los Vendedores
            c1.download_button("📥 Exportar Tabla Vendedores (Excel)", data=convert_to_excel(df_vend), file_name="Rendimiento_Vendedores.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            # --- 2. PROCESAMIENTO CLIENTES ---
            df_cli = df.groupby('Cliente_ID').agg(Cliente=('Cliente', 'first'), **{m: (m, 'sum')}).reset_index(drop=True)
            df_cli = df_cli[['Cliente', m]]
            df_cli = df_cli.sort_values(by=m, ascending=False).head(10)
            total_cli = df_cli[m].sum()
            df_cli['Porcentaje (%)'] = (df_cli[m] / total_cli) * 100 if total_cli > 0 else 0

            c2.subheader(f"Top 10 Clientes ({medida.lower()})")
            if "Barras" in tipo_grafico:
                fig2 = px.bar(df_cli, x='Cliente', y=m, text=df_cli['Porcentaje (%)'].apply(lambda x: f'{x:.1f}%'))
            elif "Circular" in tipo_grafico:
                fig2 = px.pie(df_cli, names='Cliente', values=m, hole=0.4)
            else:
                fig2 = px.line(df_cli, x='Cliente', y=m, markers=True)

            fig2.update_layout(margin=dict(t=20, b=20, l=0, r=0))
            c2.plotly_chart(fig2, use_container_width=True)
            # Botón para descargar el Excel del Top Clientes
            c2.download_button("📥 Exportar Tabla Clientes (Excel)", data=convert_to_excel(df_cli), file_name="Top_Clientes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            st.info("💡 **Tip Pro:** Si quieres guardar el gráfico como imagen, solo pasa el mouse por encima del dibujo y haz clic en el ícono de la **cámara de fotos** que aparece en la esquina superior derecha.")

        st.markdown("---")
        st.subheader("📱 Métricas Digitales (ROI/CAC)")
        col_m1, col_m2, col_m3 = st.columns(3)
        v_cerradas = col_m1.number_input("Ventas por Redes (S/)", value=5000.0)
        sueldo = col_m2.number_input("Sueldo Vendedora (S/)", value=1025.0)
        inv = col_m3.number_input("Inversión Ads (S/)", value=300.0)
        roi = (v_cerradas - (sueldo + inv)) / (sueldo + inv) * 100 if (sueldo + inv) > 0 else 0
        st.metric("ROI Digital", f"{roi:.1f}%")
