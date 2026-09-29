import streamlit as st
import pandas as pd

from core.database import guardar_en_nube
from core.lectores import leer_reporte, combinar_por_comprobante
from core.utils import parse_fecha

TABLAS_POR_EMPRESA = {
    "Aquaz": ['Ventas', 'Produccion', 'Gastos_Aquaz'],
    "Quimaroma": ['Ventas_Quima', 'Gastos_Quima'],
    # El consolidado borra movimientos de ambas empresas, pero NO el Maestro de Costos ni el Inventario
    "Consolidado": ['Ventas', 'Produccion', 'Gastos_Aquaz', 'Ventas_Quima', 'Gastos_Quima', 'Gastos'],
}


def _clave_empresa(empresa_activa):
    return "Aquaz" if "Aquaz" in empresa_activa else ("Quimaroma" if "Quimaroma" in empresa_activa else "Consolidado")


def _guardar(supabase, tabla, df) -> bool:
    """Guarda en la nube y, solo si funcionó, actualiza la sesión."""
    if guardar_en_nube(supabase, tabla, df):
        st.session_state.dfs[tabla] = df
        return True
    return False


def render(empresa_activa, supabase):
    st.title("📥 Centro de Inyección de Datos (Nube)")

    tab_ventas, tab_bases, tab_planta, tab_borrar = st.tabs(
        ["🚀 Ventas", "🗄️ Bases y Gastos", "🏭 Producción", "🗑️ Borrar datos"])

    with tab_ventas:
        _cargar_ventas(supabase)
    with tab_bases:
        _cargar_bases(supabase)
    with tab_planta:
        _cargar_produccion(supabase)
    with tab_borrar:
        _borrar(empresa_activa, supabase)


# ==========================================
# VENTAS: UNA SOLA ZONA DE CARGA
# ==========================================
def _cargar_ventas(supabase):
    st.markdown("#### Sube tu reporte de ventas")
    st.caption("Reporte detallado de FACEL (Aquaz) o Reporte de Ventas Detallado (Quimaroma), tal cual lo descargas. "
               "De FACEL sirven el Informe de Ventas y el reporte detallado (VENTAS GENERAL o separado por pestañas). "
               "El ERP reconoce cuál es, agrega lo nuevo y nunca duplica ni borra lo que ya tenías.")

    archivo = st.file_uploader("Arrastra aquí el Excel", type=["xlsx", "xls"], key="ventas_unico")
    if not archivo or not st.button("Procesar reporte", type="primary", use_container_width=True):
        return

    try:
        r = leer_reporte(archivo, st.session_state.dfs.get('Maestro_Costos', pd.DataFrame()))
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    st.markdown(f"**{r['nombre']}**" + (f" · hojas: {', '.join(r['hojas'])}" if r['hojas'] else ""))
    for nivel, msg in r["avisos"]:
        getattr(st, nivel)(msg)

    nuevo = r["datos"]
    if nuevo.empty:
        return

    tabla = 'Ventas' if r["tipo"] in ("facel_detallado", "facel_resumido") else 'Ventas_Quima'
    empresa = 'Aquaz' if tabla == 'Ventas' else 'Quimaroma'
    actual = st.session_state.dfs.get(tabla, pd.DataFrame())
    if not actual.empty and 'Comprobante' not in actual.columns:
        st.error(f"Tienes ventas de {empresa} cargadas con la versión anterior del ERP (sin número de comprobante). "
                 f"Para no duplicarlas, elige {empresa} en la barra lateral, ve a '🗑️ Borrar datos', bórralas una sola "
                 "vez y vuelve a subir tus reportes.")
        return
    # Se SUMA a lo que ya está guardado: solo se actualizan los comprobantes que vienen repetidos
    resultado, n_nuevos, n_actualizados = combinar_por_comprobante(actual, nuevo)
    detalle = f"{empresa}: {n_nuevos} comprobantes nuevos"
    if n_actualizados:
        detalle += f" y {n_actualizados} que ya existían (se actualizaron, no se duplicaron)"
    detalle += f". Total guardado: {resultado['Comprobante'].nunique()} comprobantes"

    if _guardar(supabase, tabla, resultado):
        st.success(f"✅ {detalle}.")


# ==========================================
# MATRIZ, COSTOS Y GASTOS
# ==========================================
def _cargar_bases(supabase):
    colA, colB = st.columns(2)
    with colA:
        st.markdown("#### 📘 Matriz Principal y Costos")
        archivo_matriz = st.file_uploader("Arrastra tu archivo MATRIZ", type=["xlsx", "xls"], key="matriz")
        if archivo_matriz and st.button("Procesar Matriz", use_container_width=True):
            try:
                xls = pd.ExcelFile(archivo_matriz)
                cargadas = []
                for h in list(st.session_state.dfs.keys()):
                    if h in xls.sheet_names:
                        df_cargado = pd.read_excel(xls, h)
                        df_cargado.columns = df_cargado.columns.astype(str).str.strip()
                        if _guardar(supabase, h, df_cargado):
                            cargadas.append(h)
                if cargadas:
                    st.success("✅ Guardadas en la nube: " + ", ".join(cargadas))
                else:
                    st.warning(f"No encontré hojas reconocidas. Hojas del archivo: {xls.sheet_names}")
            except Exception as e:
                st.error(f"Error: {e}")

    with colB:
        st.markdown("#### 💸 Gastos Operativos")
        archivo_gastos = st.file_uploader("Arrastra tu Excel de GASTOS", type=["xlsx", "xls"], key="gastos")
        if archivo_gastos and st.button("Procesar Gastos", use_container_width=True):
            try:
                xls_g = pd.ExcelFile(archivo_gastos)
                ok = []
                if 'AQUAZ' in xls_g.sheet_names:
                    if _guardar(supabase, 'Gastos_Aquaz', pd.read_excel(xls_g, 'AQUAZ')):
                        ok.append('Gastos_Aquaz')
                if 'QUIMA' in xls_g.sheet_names:
                    if _guardar(supabase, 'Gastos_Quima', pd.read_excel(xls_g, 'QUIMA', header=6)):
                        ok.append('Gastos_Quima')
                if ok:
                    st.success("✅ Gastos guardados: " + ", ".join(ok))
                    for t in ok:
                        if 'Monto' not in st.session_state.dfs[t].columns:
                            st.warning(f"La hoja de {t} no tiene una columna llamada 'Monto'; "
                                       "Finanzas no podrá sumar esos gastos.")
                else:
                    st.warning(f"No encontré las hojas AQUAZ o QUIMA. Hojas del archivo: {xls_g.sheet_names}")
            except Exception as e:
                st.error(f"Error: {e}")


# ==========================================
# PRODUCCIÓN
# ==========================================
def _cargar_produccion(supabase):
    st.markdown("#### 🏭 Reporte de Planta")
    archivo_prod = st.file_uploader("Arrastra reporte de planta", type=["xlsx", "xls"], key="prod")
    if not archivo_prod or not st.button("Procesar Producción", use_container_width=True, type="primary"):
        return
    try:
        df_bruto = pd.read_excel(archivo_prod)
        df_bruto.columns = df_bruto.columns.astype(str).str.strip()
        col = lambda c, d: df_bruto[c] if c in df_bruto.columns else pd.Series([d] * len(df_bruto), index=df_bruto.index)
        df_p = pd.DataFrame(index=df_bruto.index)
        df_p['Fecha'] = parse_fecha(col('Fecha', None))
        df_p['Empresa'] = "Aquaz"
        df_p['Lote'] = col('Número de Lote (o de Orden)', "S/L").fillna("S/L").astype(str)
        df_p['Producto'] = col('Producto Fabricado', "SIN NOMBRE").fillna("SIN NOMBRE")
        df_p['Cantidad_Producida'] = pd.to_numeric(col('Cantidad Producida', 0), errors='coerce').fillna(0)
        df_p['Costo_Materia_Prima'] = pd.to_numeric(col('Costo de Materia Prima', 0), errors='coerce').fillna(0)
        df_p['Merma_Soles'] = pd.to_numeric(col('Merma', 0), errors='coerce').fillna(0)
        df_p['Operario'] = col('Operario / Responsable', "No especificado").fillna("No especificado")

        actual = st.session_state.dfs.get('Produccion', pd.DataFrame())
        # Evita duplicar: un lote que ya existía se reemplaza
        if not actual.empty and 'Lote' in actual.columns:
            lotes = set(df_p.loc[df_p['Lote'] != "S/L", 'Lote'])
            actual = actual[~actual['Lote'].astype(str).isin(lotes)]
        if _guardar(supabase, 'Produccion', pd.concat([actual, df_p], ignore_index=True)):
            st.success(f"✅ {len(df_p)} registros guardados en la nube.")
    except Exception as e:
        st.error(f"Error: {e}")


# ==========================================
# BORRADO CON CONFIRMACIÓN
# ==========================================
def _borrar(empresa_activa, supabase):
    tablas = TABLAS_POR_EMPRESA[_clave_empresa(empresa_activa)]
    st.markdown(f"#### Borrar datos de {empresa_activa}")
    st.warning("Esto elimina de la nube: " + ", ".join(tablas) + ". No se puede deshacer. "
               "El Maestro de Costos y el Inventario no se tocan.")

    confirmacion = st.text_input("Escribe BORRAR para confirmar", key="confirmar_borrado")
    if st.button("🗑️ Borrar definitivamente", type="primary",
                 disabled=confirmacion.strip().upper() != "BORRAR"):
        errores = [t for t in tablas if not _guardar(supabase, t, pd.DataFrame())]
        if errores:
            st.error("No se pudieron borrar: " + ", ".join(errores))
        else:
            st.success(f"Datos de {empresa_activa} borrados.")
