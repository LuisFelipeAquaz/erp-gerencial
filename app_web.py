import streamlit as st
import pandas as pd
import numpy as np
import os
import warnings
import io
from supabase import create_client, Client

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(page_title="ERP Holding Gerencial", layout="wide", page_icon="🏢")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

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
# 3. CONEXIÓN A LA NUBE (SUPABASE)
# ==========================================
@st.cache_resource
def init_supabase():
    url = st.secrets["supabase"]["URL"]
    key = st.secrets["supabase"]["KEY"]
    return create_client(url, key)

try:
    supabase = init_supabase()
except Exception as e:
    st.error(f"Error conectando a la base de datos: {e}")
    supabase = None

def cargar_memoria_nube():
    if 'dfs' not in st.session_state:
        st.session_state.dfs = {
            'Ventas': pd.DataFrame(), 'Ventas_Quima': pd.DataFrame(), 
            'Produccion': pd.DataFrame(), 'Gastos': pd.DataFrame(), 
            'Inventario': pd.DataFrame(), 'Maestro_Costos': pd.DataFrame(), 
            'Gastos_Aquaz': pd.DataFrame(), 'Gastos_Quima': pd.DataFrame()    
        }
        if supabase:
            try:
                respuesta = supabase.table('base_datos_erp').select('*').execute()
                for fila in respuesta.data:
                    nombre = fila['nombre_tabla']
                    contenido = fila['contenido']
                    if contenido and contenido != "[]":
                        df = pd.read_json(io.StringIO(contenido), orient='records')
                        if 'Fecha' in df.columns:
                            df['Fecha'] = pd.to_datetime(df['Fecha'])
                        st.session_state.dfs[nombre] = df
            except Exception as e:
                pass # Pasa en blanco si la tabla es nueva

def guardar_en_nube(nombre_tabla, df):
    if supabase:
        try:
            json_str = df.to_json(orient='records', date_format='iso') if not df.empty else "[]"
            supabase.table('base_datos_erp').upsert({'nombre_tabla': nombre_tabla, 'contenido': json_str}).execute()
        except Exception as e:
            st.sidebar.error(f"Error al guardar {nombre_tabla} en la nube")

# Arrancamos la descarga automática de datos al abrir la app
cargar_memoria_nube()

def resaltar_stock_critico(fila):
    col_evaluar = 'Stock' if 'Stock' in fila else ('Cantidad' if 'Cantidad' in fila else None)
    if col_evaluar:
        valor = pd.to_numeric(fila[col_evaluar], errors='coerce')
        if pd.notna(valor) and valor < 0:
            return ['background-color: #fee2e2; color: #991b1b; font-weight: bold'] * len(fila)
    return [''] * len(fila)

# ==========================================
# 4. BARRA LATERAL (SIDEBAR)
# ==========================================
st.sidebar.title("⚙️ Panel de Control")
empresa_activa = st.sidebar.selectbox("🏢 ENTORNO DE TRABAJO:", ["Aquaz (Planta/Mayorista)", "Quimaroma (Tienda Fiori)", "Consolidado Grupo"])
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegación Estratégica", [
    "📊 Inicio (Dashboard)", 
    "📥 Carga de Datos", 
    "💰 1. Ventas & Analítica", 
    "🏭 2. Producción & MRP", 
    "⚖️ 3. Finanzas & Costos",
    "📦 4. Inventario",
    "🏬 5. Tienda Fiori (Unit Economics)",
    "👥 6. Retención de Clientes"
])

# ==========================================
# 5. MÓDULOS DEL SISTEMA
# ==========================================
if menu == "📥 Carga de Datos":
    st.title("📥 Centro de Inyección de Datos (Nube)")
    
    if st.button(f"🗑️ Borrar datos de {empresa_activa}", type="primary"):
        if "Aquaz" in empresa_activa:
            st.session_state.dfs['Ventas'] = pd.DataFrame()
            st.session_state.dfs['Produccion'] = pd.DataFrame()
            st.session_state.dfs['Gastos_Aquaz'] = pd.DataFrame()
            guardar_en_nube('Ventas', pd.DataFrame())
            guardar_en_nube('Produccion', pd.DataFrame())
            guardar_en_nube('Gastos_Aquaz', pd.DataFrame())
            st.success("¡Datos de Aquaz borrados de la nube! Quimaroma sigue intacto.")
        elif "Quimaroma" in empresa_activa:
            st.session_state.dfs['Ventas_Quima'] = pd.DataFrame()
            st.session_state.dfs['Gastos_Quima'] = pd.DataFrame()
            guardar_en_nube('Ventas_Quima', pd.DataFrame())
            guardar_en_nube('Gastos_Quima', pd.DataFrame())
            st.success("¡Datos de Quimaroma borrados de la nube! Aquaz sigue intacto.")
        else:
            for k in st.session_state.dfs.keys():
                st.session_state.dfs[k] = pd.DataFrame()
                guardar_en_nube(k, pd.DataFrame())
            st.success("¡Memoria de todo el Holding borrada de la nube!")
        st.rerun()

    st.markdown("---")
    tab_bases, tab_ventas, tab_planta = st.tabs(["🗄️ 1. Bases y Gastos", "🚀 2. Ventas", "🏭 3. Producción"])
    
    with tab_bases:
        colA, colB = st.columns(2)
        with colA:
            st.markdown("#### 📘 Matriz Principal y Costos")
            archivo_matriz = st.file_uploader("Arrastra tu archivo MATRIZ", type=["xlsx", "xls"], key="matriz")
            if archivo_matriz:
                if st.button("Procesar Matriz", use_container_width=True):
                    try:
                        xls = pd.ExcelFile(archivo_matriz)
                        for h in st.session_state.dfs.keys():
                            if h in xls.sheet_names:
                                df_cargado = pd.read_excel(xls, h)
                                df_cargado.columns = df_cargado.columns.str.strip()
                                st.session_state.dfs[h] = df_cargado
                                guardar_en_nube(h, df_cargado)
                        st.success("✅ Matriz cargada y sincronizada en la nube.")
                    except Exception as e:
                        st.error(f"Error: {e}")
                        
        with colB:
            st.markdown("#### 💸 Gastos Operativos 2026")
            archivo_gastos = st.file_uploader("Arrastra tu Excel de GASTOS", type=["xlsx", "xls"], key="gastos")
            if archivo_gastos:
                if st.button("Procesar Gastos", use_container_width=True):
                    try:
                        xls_g = pd.ExcelFile(archivo_gastos)
                        if 'AQUAZ' in xls_g.sheet_names:
                            st.session_state.dfs['Gastos_Aquaz'] = pd.read_excel(xls_g, 'AQUAZ')
                            guardar_en_nube('Gastos_Aquaz', st.session_state.dfs['Gastos_Aquaz'])
                        if 'QUIMA' in xls_g.sheet_names:
                            st.session_state.dfs['Gastos_Quima'] = pd.read_excel(xls_g, 'QUIMA', header=6) 
                            guardar_en_nube('Gastos_Quima', st.session_state.dfs['Gastos_Quima'])
                        st.success("✅ Gastos guardados en la nube.")
                    except Exception as e:
                        st.error(f"Error: {e}")

    with tab_ventas:
        colV1, colV2 = st.columns(2)
        with colV1:
            st.markdown("#### 🔵 Ventas AQUAZ (FACEL)")
            archivo_facel = st.file_uploader("Arrastra reporte de Facel", type=["xlsx", "xls"], key="facel")
            if archivo_facel:
                if st.button("Procesar FACEL", use_container_width=True, type="primary"):
                    try:
                        xls_f = pd.ExcelFile(archivo_facel)
                        hojas_buscadas = ['FACTURAS', 'BOLETAS DE VENTAS', 'NOTAS DE VENTAS', 'VENTAS GENERAL']
                        lista_ventas = [pd.read_excel(xls_f, sheet_name=h, header=1) for h in hojas_buscadas if h in xls_f.sheet_names]
                        
                        if lista_ventas:
                            df_bruto = pd.concat(lista_ventas, ignore_index=True)
                            df_bruto.columns = df_bruto.columns.str.strip()
                            df_v = pd.DataFrame()
                            df_v['Fecha'] = pd.to_datetime(df_bruto.get('FECHA EMISION', pd.Series(dtype=object)), format='%d/%m/%Y', errors='coerce')
                            df_v['Empresa'] = 'Aquaz' 
                            df_v['Cliente'] = df_bruto.get('CLIENTE NOMBRE', pd.Series(dtype=object)).fillna("CLIENTE VARIOS")
                            df_v['Vendedor'] = df_bruto.get('ATENDIDO POR', pd.Series(dtype=object)).fillna("TIENDA")
                            df_v['Producto'] = df_bruto.get('PRODUCTO/SERVICIO', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                            df_v['Cantidad'] = pd.to_numeric(df_bruto.get('CANTIDAD', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_v['Precio_Venta'] = pd.to_numeric(df_bruto.get('PRECIO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_v['Descuento'] = pd.to_numeric(df_bruto.get('DESCUENTO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_v['Zona'] = "No registrada" 
                            
                            df_v['Costo_Unitario_Facel'] = pd.to_numeric(df_bruto.get('COSTO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                            df_costos = st.session_state.dfs.get('Maestro_Costos', pd.DataFrame())
                            if not df_costos.empty and 'Producto' in df_costos.columns and 'Costo_Real' in df_costos.columns:
                                df_v = df_v.merge(df_costos[['Producto', 'Costo_Real']], on='Producto', how='left')
                                df_v['Costo_Unitario'] = df_v['Costo_Real'].fillna(df_v['Costo_Unitario_Facel'])
                                df_v = df_v.drop(columns=['Costo_Real'])
                            else:
                                df_v['Costo_Unitario'] = df_v['Costo_Unitario_Facel']
                                
                            df_v = df_v[df_v['Cantidad'] > 0]
                            st.session_state.dfs['Ventas'] = pd.concat([st.session_state.dfs['Ventas'], df_v], ignore_index=True)
                            guardar_en_nube('Ventas', st.session_state.dfs['Ventas'])
                            st.success(f"¡{len(df_v)} ventas guardadas permanentemente!")
                        else:
                            st.warning(f"⚠️ Hojas no encontradas. Pestañas: {xls_f.sheet_names}")
                    except Exception as e:
                        st.error(f"Error: {e}")

        with colV2:
            st.markdown("#### 🟢 Ventas QUIMAROMA")
            archivo_quima = st.file_uploader("Arrastra reporte Tienda", type=["xlsx", "xls"], key="quima_ventas")
            if archivo_quima:
                if st.button("Procesar QUIMAROMA", use_container_width=True, type="primary"):
                    try:
                        df_q_bruto = pd.read_excel(archivo_quima)
                        df_q = pd.DataFrame()
                        df_q['Fecha'] = pd.to_datetime(df_q_bruto.get('FechaEmision', pd.Series(dtype=object)), format='%d/%m/%Y', errors='coerce')
                        df_q['Empresa'] = 'Quimaroma'
                        df_q['Cliente'] = df_q_bruto.get('RazonSocial', pd.Series(dtype=object)).fillna("CLIENTES VARIOS")
                        df_q['Vendedor'] = "MOSTRADOR"
                        df_q['Producto'] = df_q_bruto.get('DescripcionItem', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                        df_q['Cantidad'] = pd.to_numeric(df_q_bruto.get('Cantidad de Item', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_q['Precio_Venta'] = pd.to_numeric(df_q_bruto.get('PrecioUnitario', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_q['Descuento'] = pd.to_numeric(df_q_bruto.get('DescuentoItem', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_q['Zona'] = "Mostrador Tienda"
                        
                        df_costos = st.session_state.dfs.get('Maestro_Costos', pd.DataFrame())
                        if not df_costos.empty and 'Producto' in df_costos.columns and 'Costo_Real' in df_costos.columns:
                            df_q = df_q.merge(df_costos[['Producto', 'Costo_Real']], on='Producto', how='left')
                            df_q['Costo_Unitario'] = df_q['Costo_Real'].fillna(0) 
                            df_q = df_q.drop(columns=['Costo_Real'])
                        else:
                            df_q['Costo_Unitario'] = 0.0

                        df_q = df_q[df_q['Cantidad'] > 0]
                        st.session_state.dfs['Ventas_Quima'] = pd.concat([st.session_state.dfs['Ventas_Quima'], df_q], ignore_index=True)
                        guardar_en_nube('Ventas_Quima', st.session_state.dfs['Ventas_Quima'])
                        st.success(f"¡{len(df_q)} ventas guardadas permanentemente!")
                    except Exception as e:
                        st.error(f"Error: {e}")

    with tab_planta:
        st.markdown("#### 🏭 Reporte de Planta")
        archivo_prod = st.file_uploader("Arrastra reporte de planta", type=["xlsx", "xls"], key="prod")
        if archivo_prod:
            if st.button("Procesar Producción", use_container_width=True, type="primary"):
                try:
                    df_bruto = pd.read_excel(archivo_prod)
                    df_bruto.columns = df_bruto.columns.str.strip()
                    df_p = pd.DataFrame()
                    df_p['Fecha'] = pd.to_datetime(df_bruto.get('Fecha', pd.Series(dtype=object)), errors='coerce')
                    df_p['Empresa'] = "Aquaz"
                    df_p['Lote'] = df_bruto.get('Número de Lote (o de Orden)', pd.Series(dtype=object)).fillna("S/L")
                    df_p['Producto'] = df_bruto.get('Producto Fabricado', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                    df_p['Cantidad_Producida'] = pd.to_numeric(df_bruto.get('Cantidad Producida', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Costo_Materia_Prima'] = pd.to_numeric(df_bruto.get('Costo de Materia Prima', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Merma_Soles'] = pd.to_numeric(df_bruto.get('Merma', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Operario'] = df_bruto.get('Operario / Responsable', pd.Series(dtype=object)).fillna("No especificado")
                        
                    st.session_state.dfs['Produccion'] = pd.concat([st.session_state.dfs['Produccion'], df_p], ignore_index=True)
                    guardar_en_nube('Produccion', st.session_state.dfs['Produccion'])
                    st.success(f"¡{len(df_p)} registros guardados permanentemente!")
                except Exception as e:
                    st.error(f"Error: {e}")

elif menu == "💰 1. Ventas & Analítica":
    st.title("💰 Análisis de Ventas y Fidelización")
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    
    df = df_aquaz if "Aquaz" in empresa_activa else (df_quima if "Quimaroma" in empresa_activa else pd.concat([df_aquaz, df_quima], ignore_index=True))
    
    if df.empty:
        st.info(f"Sube datos de ventas para ver la analítica.")
    else:
        df['Descuento'] = pd.to_numeric(df.get('Descuento', 0), errors='coerce').fillna(0)
        df['Utilidad_Bruta'] = (df['Cantidad'] * df['Precio_Venta']) - df['Descuento'] - (df['Cantidad'] * df['Costo_Unitario'])
        
        c1, c2 = st.columns(2)
        c1.subheader("Rendimiento por Vendedor (S/)")
        c1.bar_chart(df.groupby('Vendedor')['Utilidad_Bruta'].sum().sort_values(ascending=False))
        c2.subheader("Top Clientes más Rentables")
        c2.bar_chart(df.groupby('Cliente')['Utilidad_Bruta'].sum().sort_values(ascending=False).head(10))

        st.markdown("---")
        st.subheader("📱 Métricas Digitales (ROI/CAC)")
        c1, c2, c3 = st.columns(3)
        v_cerradas = c1.number_input("Ventas por Redes (S/)", value=5000.0)
        sueldo = c2.number_input("Sueldo Vendedora (S/)", value=1025.0)
        inv = c3.number_input("Inversión Ads (S/)", value=300.0)
        roi = (v_cerradas - (sueldo + inv)) / (sueldo + inv) * 100 if (sueldo + inv) > 0 else 0
        st.metric("ROI Digital", f"{roi:.1f}%")

elif menu == "🏭 2. Producción & MRP":
    st.title("🏭 Eficiencia de Planta y MRP Predictivo")
    if empresa_activa == "Quimaroma (Tienda Fiori)":
        st.info("⚠️ El módulo de Producción es exclusivo de la fábrica.")
    else:
        df = st.session_state.dfs.get('Produccion', pd.DataFrame())
        if not df.empty:
            c1, c2 = st.columns(2)
            c1.subheader("Volumen Fabricado")
            if 'Cantidad_Producida' in df.columns: c1.bar_chart(df.groupby('Producto')['Cantidad_Producida'].sum().sort_values(ascending=False).head(10))
            c2.subheader("Pérdidas por Mermas (S/)")
            if 'Merma_Soles' in df.columns: c2.bar_chart(df.groupby('Producto')['Merma_Soles'].sum().sort_values(ascending=False).head(10), color="#ff4b4b")

elif menu == "⚖️ 3. Finanzas & Costos":
    st.title("⚖️ Balance Financiero y Costos Estratégicos")
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    df_v = df_aquaz if "Aquaz" in empresa_activa else (df_quima if "Quimaroma" in empresa_activa else pd.concat([df_aquaz, df_quima], ignore_index=True))
        
    df_g = st.session_state.dfs.get('Gastos_Aquaz', pd.DataFrame()) if "Aquaz" in empresa_activa else st.session_state.dfs.get('Gastos_Quima', pd.DataFrame())
    if empresa_activa == "Consolidado Grupo": df_g = pd.concat([st.session_state.dfs.get('Gastos_Aquaz', pd.DataFrame()), st.session_state.dfs.get('Gastos_Quima', pd.DataFrame())])
    
    df_p = st.session_state.dfs.get('Produccion', pd.DataFrame())

    ut_bruta = 0
    if not df_v.empty:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        ut_bruta = ((df_v['Cantidad'] * df_v['Precio_Venta']).sum() - df_v['Descuento'].sum()) - (df_v['Cantidad'] * df_v['Costo_Unitario']).sum()

    g_fijos = df_g['Monto'].sum() if not df_g.empty and 'Monto' in df_g.columns else 0
    mermas = df_p['Merma_Soles'].sum() if not df_p.empty and 'Merma_Soles' in df_p.columns else 0
    resultado = ut_bruta - g_fijos - mermas

    st.markdown("### Resumen Consolidado")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("1. Utilidad Comercial", f"S/ {ut_bruta:,.2f}")
    c2.metric("2. Pérdidas Planta", f"- S/ {mermas:,.2f}")
    c3.metric("3. Gastos Fijos", f"- S/ {g_fijos:,.2f}")
    c4.metric("💰 RESULTADO NETO", f"S/ {resultado:,.2f}")

elif menu == "📦 4. Inventario":
    st.title("📦 Inventario en Tiempo Real")
    df = st.session_state.dfs.get('Inventario', pd.DataFrame())
    if df.empty:
        st.info("No hay datos de Inventario en la matriz.")
    else:
        st.dataframe(df.style.apply(resaltar_stock_critico, axis=1), use_container_width=True, height=500)

elif menu == "🏬 5. Tienda Fiori (Unit Economics)":
    st.title("🏬 Rentabilidad por Cotización")
    cf_mensual = st.number_input("Costos Fijos Mensuales (S/)", value=2500.0)
    cuota_diaria = cf_mensual / 30
    
    if 'carrito_fiori' not in st.session_state:
        st.session_state['carrito_fiori'] = pd.DataFrame({"Producto": ["Texapon", ""], "Cantidad": [1, 0], "Costo Unitario (S/)": [15.0, 0.0], "Precio Venta Unit. (S/)": [20.0, 0.0]})

    df_pedido = st.data_editor(st.session_state['carrito_fiori'], num_rows="dynamic", use_container_width=True, hide_index=True)
    st.session_state['carrito_fiori'] = df_pedido 
    
    gasto_var = st.number_input("Gastos Extra (S/)", value=0.0)
    c_tot = (pd.to_numeric(df_pedido['Cantidad'], errors='coerce').fillna(0) * pd.to_numeric(df_pedido['Costo Unitario (S/)'], errors='coerce').fillna(0)).sum()
    v_tot = (pd.to_numeric(df_pedido['Cantidad'], errors='coerce').fillna(0) * pd.to_numeric(df_pedido['Precio Venta Unit. (S/)'], errors='coerce').fillna(0)).sum()
    
    ut_ped = v_tot - c_tot - gasto_var
    cob = (ut_ped / cuota_diaria * 100) if cuota_diaria > 0 else 0

    cA, cB, cC, cD = st.columns(4)
    cA.metric("Venta Bruta Total", f"S/ {v_tot:,.2f}")
    cB.metric("Utilidad Neta", f"S/ {ut_ped:,.2f}")
    cD.metric("Cobertura del Día", f"{cob:.1f}%")

elif menu == "👥 6. Retención de Clientes":
    st.title("👥 Radar de Retención y Valor de Cliente")
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    df_v = df_aquaz if "Aquaz" in empresa_activa else (df_quima if "Quimaroma" in empresa_activa else pd.concat([df_aquaz, df_quima], ignore_index=True))
    
    if df_v.empty:
        st.warning(f"⚠️ No hay datos cargados para {empresa_activa}.")
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Utilidad_Bruta'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        if 'Vendedor' not in df_v.columns: df_v['Vendedor'] = 'Sin Vendedor'
        if 'Zona' not in df_v.columns: df_v['Zona'] = 'No registrada'

        hoy = pd.Timestamp.today()
        df_clientes = df_v.groupby('Cliente').agg(
            Ultima_Compra=('Fecha', 'max'), Frecuencia_Compras=('Fecha', 'nunique'),
            Utilidad_Total=('Utilidad_Bruta', 'sum'), Vendedor=('Vendedor', 'last'), Zona=('Zona', 'last')
        ).reset_index()
        df_clientes['Días Sin Comprar'] = (hoy - df_clientes['Ultima_Compra']).dt.days
        
        vendedores = ['Todos'] + sorted(list(df_clientes['Vendedor'].astype(str).unique()))
        vendedor_sel = st.selectbox("Filtrar por Vendedor:", vendedores)
        if vendedor_sel != 'Todos': df_clientes = df_clientes[df_clientes['Vendedor'] == vendedor_sel]
        
        c20 = df_clientes[df_clientes['Días Sin Comprar'] <= 20].sort_values(by='Utilidad_Total', ascending=False)
        c30 = df_clientes[(df_clientes['Días Sin Comprar'] > 20) & (df_clientes['Días Sin Comprar'] <= 30)].sort_values(by='Utilidad_Total', ascending=False)
        c45 = df_clientes[(df_clientes['Días Sin Comprar'] > 30) & (df_clientes['Días Sin Comprar'] <= 45)].sort_values(by='Utilidad_Total', ascending=False)
        c60 = df_clientes[(df_clientes['Días Sin Comprar'] > 45) & (df_clientes['Días Sin Comprar'] <= 60)].sort_values(by='Utilidad_Total', ascending=False)
        c_dorm = df_clientes[df_clientes['Días Sin Comprar'] > 60].sort_values(by='Utilidad_Total', ascending=False)
        
        t1, t2, t3, t4, t5 = st.tabs(["🟢 < 20 días", "🟡 21-30 días", "🟠 31-45 días", "🔴 46-60 días", "⚫ +60 días"])
        cols = ['Cliente', 'Vendedor', 'Zona', 'Días Sin Comprar', 'Ultima_Compra', 'Utilidad_Total']
        
        @st.cache_data
        def convert_df_to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Retencion')
            return output.getvalue()
            
        with t1:
            st.subheader(f"🟢 Activos - Total: {len(c20)}")
            if not c20.empty: st.download_button("📥 Descargar (.xlsx)", data=convert_df_to_excel(c20[cols]), file_name=f'activos_{vendedor_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(c20[cols], use_container_width=True, hide_index=True)
            
        with t2:
            st.subheader(f"🟡 Regulares - Total: {len(c30)}")
            if not c30.empty: st.download_button("📥 Descargar (.xlsx)", data=convert_df_to_excel(c30[cols]), file_name=f'regulares_{vendedor_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(c30[cols], use_container_width=True, hide_index=True)
            
        with t3:
            st.subheader(f"🟠 Alerta Temprana - Total: {len(c45)}")
            if not c45.empty: st.download_button("📥 Descargar (.xlsx)", data=convert_df_to_excel(c45[cols]), file_name=f'alerta_{vendedor_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(c45[cols], use_container_width=True, hide_index=True)
            
        with t4:
            st.subheader(f"🔴 En Riesgo - Total: {len(c60)}")
            if not c60.empty: st.download_button("📥 Descargar (.xlsx)", data=convert_df_to_excel(c60[cols]), file_name=f'riesgo_{vendedor_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(c60[cols], use_container_width=True, hide_index=True)
            
        with t5:
            st.subheader(f"⚫ Dormidos - Total: {len(c_dorm)}")
            if not c_dorm.empty: st.download_button("📥 Descargar (.xlsx)", data=convert_df_to_excel(c_dorm[cols]), file_name=f'dormidos_{vendedor_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(c_dorm[cols], use_container_width=True, hide_index=True)

elif menu == "📊 Inicio (Dashboard)":
    st.title("📊 Panel de Control Principal")
    st.write(f"Resumen gerencial de **{empresa_activa}**.")
    if st.session_state.dfs.get('Ventas', pd.DataFrame()).empty:
        st.warning("Base de datos conectada. 👈 Ve a 'Carga de Datos' y procesa tus archivos por primera vez para subirlos a la nube.")
    else:
        st.success("✅ Base de datos en la nube perfectamente sincronizada.")
