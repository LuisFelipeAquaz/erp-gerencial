import streamlit as st
import pandas as pd
import numpy as np
import os
import warnings

warnings.filterwarnings('ignore')

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="ERP Holding Gerencial", layout="wide")

# ==========================================
# 1. SISTEMA DE SEGURIDAD
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔒 Acceso al ERP Gerencial")
    usuario = st.text_input("Usuario")
    clave = st.text_input("Contraseña", type="password")
    
    if st.button("Entrar"):
        # Verifica la caja fuerte de Streamlit
        if usuario in st.secrets["passwords"] and st.secrets["passwords"][usuario] == clave:
            st.session_state["autenticado"] = True
            st.rerun() 
        else:
            st.error("😕 Usuario o contraseña incorrectos")
    
    st.stop() # Detiene la carga del resto de la página si no hay login

# ==========================================
# 2. MEMORIA DEL SISTEMA (SESSION STATE)
# ==========================================
# Aquí le decimos a la app que recuerde los archivos cargados
if 'matriz_principal' not in st.session_state:
    st.session_state['matriz_principal'] = None
if 'ventas_facel' not in st.session_state:
    st.session_state['ventas_facel'] = None
if 'planta_produccion' not in st.session_state:
    st.session_state['planta_produccion'] = None


# ==========================================
# 3. BARRA LATERAL (MENÚ)
# ==========================================
with st.sidebar:
    st.title("⚙️ Panel de Control")
    st.selectbox("ENTORNO DE TRABAJO:", ["Aquaz (Planta/Mayorista)"])
    
    st.markdown("---")
    st.markdown("**Navegación Estratégica**")
    menu = st.radio("", [
        "🏠 Inicio (Dashboard)", 
        "📥 Carga de Datos", 
        "💰 1. Ventas & Analítica",
        "🏭 2. Producción & MRP",
        "📈 3. Finanzas & Costos",
        "📦 4. Inventario",
        "🏪 5. Tienda Fiori (Unit Economics)"
    ])


# ==========================================
# 4. RUTAS DE LAS PESTAÑAS
# ==========================================

# --- PESTAÑA: INICIO ---
if menu == "🏠 Inicio (Dashboard)":
    st.title("📊 Panel de Control Principal")
    st.markdown("Resumen gerencial de **Aquaz (Planta/Mayorista)**.")
    st.info("👈 Ve a la pestaña 'Carga de Datos' para arrancar el sistema.")

# --- PESTAÑA: CARGA DE DATOS ---
elif menu == "📥 Carga de Datos":
    st.title("📥 Inyección de Datos")
    st.markdown("Aquí actualizas la base de datos de Aquaz (Planta/Mayorista).")
    
    # Botón maestro para limpiar la memoria
    if st.button("🗑️ Borrar toda la información y reiniciar", type="primary"):
        st.session_state['matriz_principal'] = None
        st.session_state['ventas_facel'] = None
        st.session_state['planta_produccion'] = None
        st.success("¡Memoria borrada! El sistema está en cero.")
        st.rerun()

    st.markdown("---")
    
    # Subida Matriz
    st.subheader("1️⃣ Sube tu Matriz Principal")
    archivo_matriz = st.file_uploader("Arrastra tu archivo MATRIZ (Historial)", type=["xlsx", "xls"], key="up_matriz")
    if archivo_matriz is not None:
        # Lo guardamos directo en la memoria global
        st.session_state['matriz_principal'] = pd.read_excel(archivo_matriz)
        st.success("✅ Matriz cargada en memoria.")

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🚀 Ventas (FACEL)")
        archivo_ventas = st.file_uploader("Arrastra reporte de Facel", type=["xlsx", "xls"], key="up_ventas")
        if st.button("Procesar Ventas"):
            if archivo_ventas is not None:
                # Al procesar, lo guardamos en la memoria global
                df_temp_ventas = pd.read_excel(archivo_ventas)
                st.session_state['ventas_facel'] = df_temp_ventas
                st.success(f"¡{len(df_temp_ventas)} ventas inyectadas!")
            else:
                st.warning("Por favor, sube el archivo primero.")

    with col2:
        st.subheader("🏭 Planta (PRODUCCIÓN)")
        archivo_prod = st.file_uploader("Arrastra reporte de planta", type=["xlsx", "xls"], key="up_prod")
        if st.button("Procesar Producción"):
            if archivo_prod is not None:
                # Al procesar, lo guardamos en la memoria global
                df_temp_prod = pd.read_excel(archivo_prod)
                st.session_state['planta_produccion'] = df_temp_prod
                st.success(f"¡{len(df_temp_prod)} registros inyectados!")
            else:
                st.warning("Por favor, sube el archivo primero.")

# --- PESTAÑA: VENTAS ---
elif menu == "💰 1. Ventas & Analítica":
    st.title("💰 Análisis de Ventas y Fidelización")
    
    # El sistema ahora lee de la memoria, no de la caja de subida
    if st.session_state['ventas_facel'] is not None:
        df_ventas = st.session_state['ventas_facel']
        
        # --- AQUÍ ABAJO PEGA EL CÓDIGO DE TUS GRÁFICOS DE VENTAS ---
        st.write("Vista previa de base de datos activa:")
        st.dataframe(df_ventas.head())
        
    else:
        st.warning("Sube datos de ventas en la pestaña 'Carga de Datos' para ver la analítica.")

# --- PESTAÑA: PRODUCCIÓN ---
elif menu == "🏭 2. Producción & MRP":
    st.title("🏭 Analítica de Producción")
    
    # El sistema ahora lee de la memoria
    if st.session_state['planta_produccion'] is not None:
        df_prod = st.session_state['planta_produccion']
        
        # --- AQUÍ ABAJO PEGA TU CÓDIGO DE LOS GRÁFICOS AZULES (JABON LIMON, ETC) ---
        st.write("Vista previa de base de datos activa:")
        st.dataframe(df_prod.head())
        
    else:
        st.warning("Sube datos de Producción en la pestaña 'Carga de Datos'.")

# --- OTRAS PESTAÑAS (Estructura base) ---
elif menu == "📈 3. Finanzas & Costos":
    st.title("📈 Finanzas & Costos")
    st.info("Configura los gráficos usando los datos en memoria.")

elif menu == "📦 4. Inventario":
    st.title("📦 Inventario")
    st.info("Configura los gráficos usando los datos en memoria.")

elif menu == "🏪 5. Tienda Fiori (Unit Economics)":
    st.title("🏪 Tienda Fiori (Unit Economics)")
    st.info("Configura los gráficos usando los datos en memoria.")
