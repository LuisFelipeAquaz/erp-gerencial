import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(page_title="ERP Holding Gerencial", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 2. SISTEMA DE SEGURIDAD
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔒 Acceso al ERP Gerencial")
    usuario = st.text_input("Usuario")
    clave = st.text_input("Contraseña", type="password")
    
    if st.button("Entrar"):
        if usuario in st.secrets["passwords"] and st.secrets["passwords"][usuario] == clave:
            st.session_state["autenticado"] = True
            st.rerun() 
        else:
            st.error("😕 Usuario o contraseña incorrectos")
    st.stop()

# ==========================================
# 3. MEMORIA DEL SISTEMA (CEREBRO)
# ==========================================
hojas_base = ['Ventas', 'Produccion', 'Gastos', 'Inventario', 'CxC', 'Logistica', 'Calidad', 'RRHH', 'Metas_Vendedores']

if 'dfs' not in st.session_state:
    st.session_state['dfs'] = {h: pd.DataFrame() for h in hojas_base}
if 'empresa_activa' not in st.session_state:
    st.session_state['empresa_activa'] = "Aquaz Perú SAC"
if 'dias_periodo' not in st.session_state:
    st.session_state['dias_periodo'] = 30

# ==========================================
# 4. BARRA LATERAL (MENÚ Y FILTROS)
# ==========================================
with st.sidebar:
    st.title("⚙️ Panel de Control")
    st.session_state['empresa_activa'] = st.selectbox("ENTORNO DE TRABAJO:", ["Aquaz Perú SAC", "Real Química", "Quimaroma", "Otra..."])
    
    st.markdown("---")
    st.markdown("**📅 FILTRO DE TIEMPO**")
    filtro_tiempo = st.selectbox("Rápido:", ["Todo el Historial", "Últimos 15 Días", "Mes Actual", "Rango Personalizado"])
    
    fecha_desde, fecha_hasta = None, None
    if filtro_tiempo == "Rango Personalizado":
        col1, col2 = st.columns(2)
        with col1:
            fecha_desde = st.date_input("Desde")
        with col2:
            fecha_hasta = st.date_input("Hasta")
            
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
# 5. LÓGICA DE FILTRADO GLOBAL
# ==========================================
dfs_filtrados = {h: pd.DataFrame() for h in hojas_base}
hoy = pd.Timestamp.today()

if filtro_tiempo == "Todo el Historial": 
    fecha_limite = pd.Timestamp.min
    st.session_state['dias_periodo'] = 30 # Por defecto
elif filtro_tiempo == "Últimos 15 Días": 
    fecha_limite = hoy - pd.Timedelta(days=15)
    st.session_state['dias_periodo'] = 15
elif filtro_tiempo == "Mes Actual": 
    fecha_limite = hoy.replace(day=1)
    st.session_state['dias_periodo'] = max(1, (hoy - fecha_limite).days + 1)

for nombre_hoja, df in st.session_state['dfs'].items():
    if not df.empty and 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        if filtro_tiempo == "Rango Personalizado" and fecha_desde and fecha_hasta:
            f_desde = pd.to_datetime(fecha_desde)
            f_hasta = pd.to_datetime(fecha_hasta) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            mask = (df['Fecha'] >= f_desde) & (df['Fecha'] <= f_hasta)
            st.session_state['dias_periodo'] = max(1, (f_hasta - f_desde).days)
        else:
            mask = df['Fecha'] >= fecha_limite
        dfs_filtrados[nombre_hoja] = df[mask].copy()
    else:
        dfs_filtrados[nombre_hoja] = df.copy()

# ==========================================
# 6. RUTAS DE LAS PESTAÑAS
# ==========================================

# --- INICIO ---
if menu == "🏠 Inicio (Dashboard)":
    st.title("📊 Panel de Control Principal")
    st.markdown(f"Resumen gerencial de **{st.session_state['empresa_activa']}**.")
    st.info("👈 Ve a la pestaña 'Carga de Datos' para inyectar tu Matriz y reportes.")

# --- CARGA DE DATOS ---
elif menu == "📥 Carga de Datos":
    st.title("📥 Inyección de Datos")
    
    if st.button("🗑️ Borrar toda la información en memoria y reiniciar", type="primary"):
        st.session_state['dfs'] = {h: pd.DataFrame() for h in hojas_base}
        st.success("¡Memoria borrada! El sistema está en cero.")
        st.rerun()

    st.markdown("---")
    st.subheader("1️⃣ Sube tu Matriz Principal")
    archivo_matriz = st.file_uploader("Arrastra tu archivo MATRIZ (Historial)", type=["xlsx", "xls"], key="up_matriz")
    
    if archivo_matriz is not None:
        try:
            xls = pd.ExcelFile(archivo_matriz)
            for h in st.session_state['dfs'].keys():
                if h in xls.sheet_names:
                    df_cargado = pd.read_excel(xls, h)
                    if df_cargado is not None:
                        df_cargado.columns = df_cargado.columns.str.strip()
                        st.session_state['dfs'][h] = df_cargado
            st.success("✅ Matriz cargada en memoria exitosamente.")
        except Exception as e:
            st.error(f"Error al leer Matriz: {e}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🚀 Ventas (FACEL)")
        archivo_ventas = st.file_uploader("Arrastra reporte de Facel", type=["xlsx", "xls"], key="up_ventas")
        if st.button("Procesar Ventas FACEL"):
            if archivo_ventas is not None:
                try:
                    xls_facel = pd.ExcelFile(archivo_ventas)
                    lista_ventas = [pd.read_excel(xls_facel, sheet_name=h, header=1) for h in ['FACTURAS', 'BOLETAS DE VENTAS', 'NOTAS DE VENTAS'] if h in xls_facel.sheet_names]
                    if lista_ventas:
                        df_bruto = pd.concat(lista_ventas, ignore_index=True)
                        df_bruto.columns = df_bruto.columns.str.strip()
                        df_v = pd.DataFrame()
                        df_v['Fecha'] = pd.to_datetime(df_bruto.get('FECHA EMISION', pd.Series(dtype=object)), format='%d/%m/%Y', errors='coerce')
                        df_v['Empresa'] = st.session_state['empresa_activa']
                        df_v['Cliente'] = df_bruto.get('CLIENTE NOMBRE', pd.Series(dtype=object)).fillna("CLIENTE VARIOS")
                        df_v['Vendedor'] = df_bruto.get('ATENDIDO POR', pd.Series(dtype=object)).fillna("TIENDA")
                        df_v['Producto'] = df_bruto.get('PRODUCTO/SERVICIO', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                        df_v['Cantidad'] = pd.to_numeric(df_bruto.get('CANTIDAD', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Precio_Venta'] = pd.to_numeric(df_bruto.get('PRECIO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Costo_Unitario'] = pd.to_numeric(df_bruto.get('COSTO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Descuento'] = pd.to_numeric(df_bruto.get('DESCUENTO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v = df_v[df_v['Cantidad'] > 0]
                        
                        st.session_state['dfs']['Ventas'] = pd.concat([st.session_state['dfs']['Ventas'], df_v], ignore_index=True)
                        st.success(f"¡Se inyectaron {len(df_v)} ventas nuevas!")
                except Exception as e:
                    st.error(f"Error procesando FACEL: {e}")

    with col2:
        st.subheader("🏭 Planta (PRODUCCIÓN)")
        archivo_prod = st.file_uploader("Arrastra reporte de planta", type=["xlsx", "xls"], key="up_prod")
        if st.button("Procesar Producción"):
            if archivo_prod is not None:
                try:
                    df_bruto_p = pd.read_excel(archivo_prod)
                    df_bruto_p.columns = df_bruto_p.columns.str.strip()
                    df_p = pd.DataFrame()
                    df_p['Fecha'] = pd.to_datetime(df_bruto_p.get('Fecha', pd.Series(dtype=object)), errors='coerce')
                    df_p['Empresa'] = st.session_state['empresa_activa']
                    df_p['Lote'] = df_bruto_p.get('Número de Lote (o de Orden)', pd.Series(dtype=object)).fillna("S/L")
                    df_p['Producto'] = df_bruto_p.get('Producto Fabricado', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                    df_p['Cantidad_Producida'] = pd.to_numeric(df_bruto_p.get('Cantidad Producida', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Costo_Materia_Prima'] = pd.to_numeric(df_bruto_p.get('Costo de Materia Prima', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Merma_Soles'] = pd.to_numeric(df_bruto_p.get('Merma', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Operario'] = df_bruto_p.get('Operario / Responsable', pd.Series(dtype=object)).fillna("No especificado")
                    
                    st.session_state['dfs']['Produccion'] = pd.concat([st.session_state['dfs']['Produccion'], df_p], ignore_index=True)
                    st.success(f"¡Se inyectaron {len(df_p)} registros de Planta!")
                except Exception as e:
                    st.error(f"Error procesando Producción: {e}")

# --- VENTAS & ANALÍTICA ---
elif menu == "💰 1. Ventas & Analítica":
    st.title("💰 Análisis de Ventas y Fidelización")
    df_v = dfs_filtrados['Ventas'].copy()
    
    if df_v.empty:
        st.warning("No hay datos de ventas para mostrar en este periodo.")
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Monto_Vendido'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento']
        df_v['Utilidad_Bruta'] = df_v['Monto_Vendido'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        
        # Gráficos
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), facecolor='#0f172a')
        
        # Gráfico 1: Termómetro Vendedores
        ventas_vendedor = df_v.groupby('Vendedor')['Monto_Vendido'].sum().reset_index()
        df_metas = st.session_state['dfs'].get('Metas_Vendedores', pd.DataFrame())
        
        if not df_metas.empty and 'Vendedor' in df_metas.columns and 'Meta' in df_metas.columns:
            ventas_vendedor['Vendedor'] = ventas_vendedor['Vendedor'].astype(str).str.strip()
            df_metas['Vendedor'] = df_metas['Vendedor'].astype(str).str.strip()
            ventas_vendedor = pd.merge(ventas_vendedor, df_metas, on='Vendedor', how='left')
            ventas_vendedor['Meta'] = pd.to_numeric(ventas_vendedor['Meta'], errors='coerce').fillna(1000) 
        else:
            ventas_vendedor['Meta'] = 1000

        ventas_vendedor['Porcentaje'] = (ventas_vendedor['Monto_Vendido'] / ventas_vendedor['Meta']) * 100
        ventas_vendedor = ventas_vendedor.sort_values(by='Porcentaje', ascending=True).tail(10)
        
        vendedores = ventas_vendedor['Vendedor']
        ventas = ventas_vendedor['Monto_Vendido']
        metas = ventas_vendedor['Meta']
        
        ax1.barh(vendedores, metas, color='#334155', label='Meta Asignada (S/)')
        ax1.barh(vendedores, ventas, color='#3b82f6', label='Venta Real (S/)', alpha=0.9)
        
        for i, (v, m, p) in enumerate(zip(ventas, metas, ventas_vendedor['Porcentaje'])):
            color_txt = '#10b981' if p >= 100 else ('#f59e0b' if p >= 70 else '#ef4444')
            pos_x = max(v, m) + (metas.max() * 0.05)
            ax1.text(pos_x, i, f"{p:.1f}%", color=color_txt, fontweight='bold', va='center')
            
        ax1.set_title('Termómetro de Metas por Vendedor', color='white', fontweight='bold')
        ax1.tick_params(colors='#94a3b8')
        ax1.legend(loc='lower right', facecolor='#1e293b', edgecolor='white', labelcolor='white')
        ax1.set_facecolor('#0f172a')

        # Gráfico 2: Top Clientes
        clientes = df_v.groupby('Cliente')['Utilidad_Bruta'].sum().sort_values().tail(5)
        ax2.barh(clientes.index, clientes.values, color='#10b981')
        ax2.set_title('Top 5 Clientes (Por Utilidad Neta)', color='white', fontweight='bold')
        ax2.tick_params(colors='#94a3b8')
        ax2.set_facecolor('#0f172a')
        
        plt.tight_layout()
        st.pyplot(fig)
        
        st.markdown("---")
        # Gráfico 3: Productos Ganadores
        st.subheader("🏆 PRODUCTOS GANADORES (Por Galones Vendidos)")
        top_productos = df_v.groupby('Producto')['Cantidad'].sum().sort_values().tail(10) 
        fig2, ax3 = plt.subplots(figsize=(10, 5), facecolor='#0f172a')
        ax3.barh(top_productos.index, top_productos.values, color='#f59e0b')
        ax3.tick_params(colors='#94a3b8')
        ax3.set_facecolor('#0f172a')
        plt.tight_layout()
        st.pyplot(fig2)

# --- PRODUCCIÓN & MRP ---
elif menu == "🏭 2. Producción & MRP":
    st.title("🏭 Analítica de Producción")
    df_p = dfs_filtrados['Produccion'].copy()
    
    if df_p.empty:
        st.warning("No hay datos de Planta para el periodo filtrado.")
    else:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5), facecolor='#0f172a')
        
        if 'Cantidad_Producida' in df_p.columns:
            prod_cant = df_p.groupby('Producto')['Cantidad_Producida'].sum().sort_values().tail(5)
            ax1.barh(prod_cant.index, prod_cant.values, color='#8b5cf6')
            ax1.set_title('Top Químicos Fabricados (Volumen)', color='white')
        
        ax1.tick_params(colors='#94a3b8')
        ax1.set_facecolor('#0f172a')

        if 'Merma_Soles' in df_p.columns:
            mermas = df_p.groupby('Producto')['Merma_Soles'].sum().sort_values(ascending=False).head(5)
            ax2.bar(mermas.index, mermas.values, color='#ef4444')
            ax2.set_title('Pérdidas por Merma (Soles)', color='white')
            ax2.tick_params(colors='white', labelrotation=15)
        
        ax2.set_facecolor('#0f172a')
        plt.tight_layout()
        st.pyplot(fig)

# --- FINANZAS & COSTOS ---
elif menu == "📈 3. Finanzas & Costos":
    st.title("📈 Finanzas & Costos")
    
    incluir_gastos = st.checkbox("✅ INCLUIR GASTOS Y MERMAS (Cálculo Neto)", value=True)
    st.markdown("---")
    
    df_v = dfs_filtrados['Ventas'].copy()
    df_g = dfs_filtrados['Gastos'].copy()
    df_p = dfs_filtrados['Produccion'].copy()

    if df_v.empty: utilidad_bruta = 0
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        utilidad_bruta = ((df_v['Cantidad'] * df_v['Precio_Venta']).sum() - df_v['Descuento'].sum()) - (df_v['Cantidad'] * df_v['Costo_Unitario']).sum()

    if incluir_gastos:
        mermas = df_p['Merma_Soles'].sum() if not df_p.empty and 'Merma_Soles' in df_p.columns else 0
        tiene_fecha_gastos = 'Fecha' in df_g.columns if not df_g.empty else False
        tiene_tipo = 'Tipo' in df_g.columns if not df_g.empty else False
        
        gastos_fijos_total = 0
        gastos_variables_total = 0

        if not df_g.empty:
            df_g['Monto'] = pd.to_numeric(df_g.get('Monto', 0), errors='coerce').fillna(0)
            if tiene_tipo:
                df_fijos = df_g[df_g['Tipo'].astype(str).str.upper() == 'FIJO']
                df_vars = df_g[df_g['Tipo'].astype(str).str.upper() == 'VARIABLE']
            else:
                df_fijos = df_g
                df_vars = pd.DataFrame({'Monto': []})

            if tiene_fecha_gastos:
                gastos_fijos_total = df_fijos['Monto'].sum()
                gastos_variables_total = df_vars['Monto'].sum()
                texto_gastos = (f"**3. GASTOS FIJOS (Del periodo):** - S/ {gastos_fijos_total:,.2f}  \n"
                                f"**4. GASTOS VARIABLES (Imprevistos):** - S/ {gastos_variables_total:,.2f}  \n")
            else:
                gastos_fijos_mes = df_fijos['Monto'].sum()
                gastos_fijos_total = (gastos_fijos_mes / 30.0) * st.session_state['dias_periodo']
                gastos_variables_total = df_vars['Monto'].sum()
                texto_gastos = (f"**3. GASTOS FIJOS (Proporcional {st.session_state['dias_periodo']} días):** - S/ {gastos_fijos_total:,.2f}  \n"
                                f"**4. GASTOS VARIABLES (Registrados):** - S/ {gastos_variables_total:,.2f}  \n")
        else:
            texto_gastos = ("**3. GASTOS FIJOS:** - S/ 0.00  \n"
                            "**4. GASTOS VARIABLES:** - S/ 0.00  \n")

        resultado_final = utilidad_bruta - gastos_fijos_total - gastos_variables_total - mermas
        texto_titulo = "BALANCE NETO (Con Gastos Operativos)"
    else:
        gastos_fijos_total, gastos_variables_total, mermas = 0, 0, 0
        resultado_final = utilidad_bruta
        texto_gastos = ("**3. GASTOS FIJOS (Ocultos):** - S/ 0.00  \n"
                        "**4. GASTOS VARIABLES (Ocultos):** - S/ 0.00  \n")
        texto_titulo = "BALANCE BRUTO (Solo Ventas Comerciales)"

    color_final = "#10b981" if resultado_final >= 0 else "#ef4444"
    
    st.markdown(f"### {texto_titulo}")
    st.info(f"**1. UTILIDAD DE VENTAS:** S/ {utilidad_bruta:,.2f}  \n"
            f"**2. PÉRDIDAS EN PLANTA (Mermas):** - S/ {mermas:,.2f}  \n"
            + texto_gastos)
    
    st.markdown(f"<h2 style='color: {color_final};'>💰 RESULTADO FINAL: S/ {resultado_final:,.2f}</h2>", unsafe_allow_html=True)

elif menu == "📦 4. Inventario":
    st.title("📦 Inventario")
    st.info("Módulo en construcción (Próximamente).")

elif menu == "🏪 5. Tienda Fiori (Unit Economics)":
    st.title("🏪 Tienda Fiori (Unit Economics)")
    st.info("Módulo en construcción (Próximamente).")
