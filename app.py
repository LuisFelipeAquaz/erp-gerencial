import streamlit as st
import warnings

from core.auth import verificar_autenticacion
from core.database import conectar_supabase, cargar_memoria_nube
from modules import (
    dashboard,
    carga_datos,
    ventas_analitica,
    produccion_mrp,
    finanzas_costos,
    inventario,
    tienda_fiori,
    retencion_clientes,
    precios_quimaroma,
    explorador_ventas,
)

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(page_title="ERP Holding Gerencial", layout="wide", page_icon="🏢")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 2. SISTEMA DE SEGURIDAD
# ==========================================
verificar_autenticacion()

# ==========================================
# 3. CONEXIÓN A LA NUBE (SUPABASE)
# ==========================================
supabase = conectar_supabase()

# Arrancamos la descarga automática de datos al abrir la app
cargar_memoria_nube(supabase)

# ==========================================
# 4. BARRA LATERAL (SIDEBAR)
# ==========================================
st.sidebar.title("⚙️ Panel de Control")
empresa_activa = st.sidebar.selectbox("🏢 ENTORNO DE TRABAJO:", ["Aquaz (Planta/Mayorista)", "Quimaroma (Tienda Fiori)", "Consolidado Grupo"])
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegación Estratégica", [
    "📊 Inicio (Dashboard)",
    "📥 Carga de Datos",
    "🔎 Explorador de Ventas",
    "💰 1. Ventas & Analítica",
    "🏭 2. Producción & MRP",
    "⚖️ 3. Finanzas & Costos",
    "📦 4. Inventario",
    "🏬 5. Tienda Fiori (Unit Economics)",
    "👥 6. Retención de Clientes",
    "🏪 7. Precios Quimaroma"
])

# ==========================================
# 5. MÓDULOS DEL SISTEMA
# ==========================================
if menu == "📥 Carga de Datos":
    carga_datos.render(empresa_activa, supabase)

elif menu == "🔎 Explorador de Ventas":
    explorador_ventas.render(empresa_activa)

elif menu == "💰 1. Ventas & Analítica":
    ventas_analitica.render(empresa_activa)

elif menu == "🏭 2. Producción & MRP":
    produccion_mrp.render(empresa_activa)

elif menu == "⚖️ 3. Finanzas & Costos":
    finanzas_costos.render(empresa_activa)

elif menu == "📦 4. Inventario":
    inventario.render()

elif menu == "🏬 5. Tienda Fiori (Unit Economics)":
    tienda_fiori.render()

elif menu == "👥 6. Retención de Clientes":
    retencion_clientes.render(empresa_activa)

elif menu == "🏪 7. Precios Quimaroma":
    precios_quimaroma.render(supabase)

elif menu == "📊 Inicio (Dashboard)":
    dashboard.render(empresa_activa)
