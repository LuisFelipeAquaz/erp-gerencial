import io

import pandas as pd
import streamlit as st
from supabase import create_client, Client

TABLAS = ['Ventas', 'Ventas_Quima', 'Produccion', 'Gastos', 'Inventario',
          'Maestro_Costos', 'Gastos_Aquaz', 'Gastos_Quima']


@st.cache_resource
def init_supabase():
    url = st.secrets["supabase"]["URL"]
    key = st.secrets["supabase"]["KEY"]
    return create_client(url, key)


def conectar_supabase():
    """Devuelve el cliente de Supabase o None si falla."""
    try:
        return init_supabase()
    except Exception as e:
        st.error(f"Error conectando a la base de datos: {e}")
        return None


def cargar_memoria_nube(supabase):
    """
    Descarga todas las tablas de Supabase a st.session_state.dfs.

    PROTECCIÓN: si la descarga falla, se marca la sesión como 'sin nube' y
    guardar_en_nube() se niega a escribir. Antes, un fallo silencioso dejaba
    las tablas vacías y la siguiente carga reemplazaba todo el historial.
    """
    if 'dfs' not in st.session_state:
        st.session_state.dfs = {t: pd.DataFrame() for t in TABLAS}

    if st.session_state.get('nube_ok'):
        return  # ya se descargó en esta sesión

    if not supabase:
        st.session_state.nube_ok = False
        _aviso_sin_nube("No hay conexión con Supabase.")
        return

    try:
        respuesta = supabase.table('base_datos_erp').select('*').execute()
        nuevas = {t: pd.DataFrame() for t in TABLAS}
        for fila in respuesta.data:
            nombre = fila['nombre_tabla']
            contenido = fila['contenido']
            if contenido and contenido != "[]":
                # dtype=False: no convertir textos como '000037' o '75770151' en números
                df = pd.read_json(io.StringIO(contenido), orient='records', dtype=False, convert_dates=False)
                if 'Fecha' in df.columns:
                    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
                nuevas[nombre] = df
        st.session_state.dfs = nuevas
        st.session_state.nube_ok = True
    except Exception as e:
        st.session_state.nube_ok = False
        _aviso_sin_nube(f"Falló la descarga de datos: {e}")


def _aviso_sin_nube(motivo: str):
    st.error(f"⚠️ {motivo} Para proteger tu historial, el ERP NO guardará cambios en la nube "
             "hasta que la conexión funcione.")
    if st.button("🔄 Reintentar conexión"):
        st.session_state.pop('nube_ok', None)
        st.rerun()


def guardar_en_nube(supabase, nombre_tabla, df) -> bool:
    """Guarda una tabla completa. Devuelve True solo si se guardó de verdad."""
    if not st.session_state.get('nube_ok'):
        st.error(f"No se guardó '{nombre_tabla}': la sesión no está sincronizada con la nube. "
                 "Recarga la página e inténtalo de nuevo.")
        return False
    if not supabase:
        return False
    try:
        json_str = df.to_json(orient='records', date_format='iso') if not df.empty else "[]"
        supabase.table('base_datos_erp').upsert(
            {'nombre_tabla': nombre_tabla, 'contenido': json_str}
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error al guardar '{nombre_tabla}' en la nube: {e}")
        return False
