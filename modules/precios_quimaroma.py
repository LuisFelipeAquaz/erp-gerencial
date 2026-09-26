import streamlit as st
import pandas as pd
import io
from core.database import guardar_en_nube
from core.utils import norm_texto


def render(supabase):
    st.title("🏪 Control de Costos y Precios (Quimaroma - Fiori)")
    st.info("Sube aquí tu Excel quincenal con la plantilla oficial. El 'Costo Estándar' se usará para calcular la rentabilidad real de tus ventas.")

    archivo_precios = st.file_uploader("Sube tu Excel de Precios Actualizado", type=["xlsx", "xls"])
    if archivo_precios:
        if st.button("Cargar Lista Oficial de Precios", type="primary", use_container_width=True):
            try:
                df_precios = pd.read_excel(archivo_precios)
                df_precios.columns = df_precios.columns.astype(str).str.strip()

                # El molde trae 'Costo_Real' y 'Costo_Estandar_Soles'. El estándar en soles es el que manda:
                # se quita el otro para no terminar con dos columnas 'Costo_Real' (eso impedía guardar).
                if 'Costo_Estandar_Soles' in df_precios.columns and 'Costo_Real' in df_precios.columns:
                    df_precios = df_precios.drop(columns=['Costo_Real'])
                df_precios = df_precios.rename(columns={
                    'Nombre_Insumo': 'Producto',
                    'Costo_Estandar_Soles': 'Costo_Real'
                })
                if 'Producto' not in df_precios.columns or 'Costo_Real' not in df_precios.columns:
                    st.error("Faltan las columnas 'Nombre_Insumo' y/o 'Costo_Estandar_Soles'. Usa el molde de abajo.")
                    return
                df_precios = df_precios.dropna(subset=['Producto'])
                repetidos = df_precios['Producto'].map(norm_texto).duplicated(keep='last')
                if repetidos.any():
                    st.warning(f"{repetidos.sum()} productos estaban repetidos; se usó la última fila de cada uno: "
                               + ", ".join(df_precios.loc[repetidos, 'Producto'].astype(str).head(5)))
                    df_precios = df_precios[~repetidos]

                if guardar_en_nube(supabase, 'Maestro_Costos', df_precios):
                    st.session_state.dfs['Maestro_Costos'] = df_precios
                    st.success(f"✅ Lista de precios actualizada en la nube ({len(df_precios)} productos). "
                               "Los costos nuevos se aplican a los reportes de ventas que subas desde ahora.")
            except Exception as e:
                st.error(f"Error procesando el Excel. Asegúrate de tener las columnas correctas: {e}")

    st.markdown("---")
    df_actual = st.session_state.dfs.get('Maestro_Costos', pd.DataFrame())
    if not df_actual.empty:
        st.subheader("📋 Lista Vigente Oficial")
        df_mostrar = df_actual.copy()
        if 'Costo_Real' in df_mostrar.columns:
            df_mostrar = df_mostrar.rename(columns={'Costo_Real': 'Costo Estándar (S/)'})
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ No hay una lista de precios cargada. Descarga el molde y súbelo.")

    st.markdown("---")
    with st.expander("📥 Descargar Molde de Excel en Blanco"):
        df_molde = pd.DataFrame(columns=['Codigo_SKU', 'Nombre_Insumo', 'Moneda', 'Costo_Real', 'Costo_Estandar_Soles', 'Precio_Venta_Publico'])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_molde.to_excel(writer, index=False, sheet_name='Precios')
        st.download_button("Descargar Plantilla", data=output.getvalue(), file_name="Molde_Precios_Quimaroma.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
