import streamlit as st
import pandas as pd
import numpy as np
import os
import warnings

# --- INICIO DE SISTEMA DE SEGURIDAD ---
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
# --- FIN DE SISTEMA DE SEGURIDAD ---

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 1. CONFIGURACIÓN Y MEMORIA DEL SISTEMA
# ==========================================
st.set_page_config(page_title="ERP Holding Gerencial", layout="wide", page_icon="🏢")

# Inicializamos la memoria ampliada para el Holding
if 'dfs' not in st.session_state:
    st.session_state.dfs = {
        'Ventas': pd.DataFrame(), 
        'Ventas_Quima': pd.DataFrame(), 
        'Produccion': pd.DataFrame(), 
        'Gastos': pd.DataFrame(), 
        'Inventario': pd.DataFrame(), 
        'CxC': pd.DataFrame(), 
        'Logistica': pd.DataFrame(), 
        'Calidad': pd.DataFrame(), 
        'RRHH': pd.DataFrame(),
        'Maestro_Costos': pd.DataFrame(), 
        'Gastos_Aquaz': pd.DataFrame(),   
        'Gastos_Quima': pd.DataFrame()    
    }

def resaltar_stock_critico(fila):
    col_evaluar = 'Stock' if 'Stock' in fila else ('Cantidad' if 'Cantidad' in fila else None)
    if col_evaluar:
        valor = pd.to_numeric(fila[col_evaluar], errors='coerce')
        if pd.notna(valor) and valor < 0:
            return ['background-color: #fee2e2; color: #991b1b; font-weight: bold'] * len(fila)
    return [''] * len(fila)

# ==========================================
# 2. BARRA LATERAL (SIDEBAR MULTISOCIEDAD)
# ==========================================
st.sidebar.title("⚙️ Panel de Control")
empresa_activa = st.sidebar.selectbox("🏢 ENTORNO DE TRABAJO:", ["Aquaz (Planta/Mayorista)", "Quinearoma (Tienda Fiori)", "Consolidado Grupo"])
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
# 3. PANTALLA: CARGA DE DATOS (NUEVA TUBERÍA)
# ==========================================
if menu == "📥 Carga de Datos":
    st.title("📥 Centro de Inyección de Datos (Holding)")
    st.write("Sube los archivos oficiales. El sistema cruzará costos y ventas automáticamente.")
    
    if st.button("🗑️ Borrar toda la información en memoria y reiniciar", type="primary"):
        st.session_state.dfs = {k: pd.DataFrame() for k in st.session_state.dfs.keys()}
        st.success("¡Memoria borrada! El sistema está en cero.")
        st.rerun()

    st.markdown("---")
    
    # Nombre corregido a "Aquaz / Quimaroma"
    tab_bases, tab_ventas, tab_planta = st.tabs(["🗄️ 1. Bases y Gastos (Matriz)", "🚀 2. Ventas (Aquaz / Quimaroma)", "🏭 3. Producción"])
    
    with tab_bases:
        colA, colB = st.columns(2)
        with colA:
            st.markdown("#### 📘 Matriz Principal y Costos")
            st.info("💡 Asegúrate de tener una hoja llamada 'Maestro_Costos' en tu Excel.")
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
                        st.success("✅ Matriz y Diccionario de Costos cargados.")
                    except Exception as e:
                        st.error(f"Error cargando Matriz: {e}")
                        
        with colB:
            st.markdown("#### 💸 Gastos Operativos 2026")
            archivo_gastos = st.file_uploader("Arrastra tu Excel de GASTOS", type=["xlsx", "xls"], key="gastos")
            if archivo_gastos:
                if st.button("Procesar Gastos", use_container_width=True):
                    try:
                        xls_g = pd.ExcelFile(archivo_gastos)
                        if 'AQUAZ' in xls_g.sheet_names:
                            st.session_state.dfs['Gastos_Aquaz'] = pd.read_excel(xls_g, 'AQUAZ')
                        if 'QUIMA' in xls_g.sheet_names:
                            st.session_state.dfs['Gastos_Quima'] = pd.read_excel(xls_g, 'QUIMA', header=6) 
                        st.success("✅ Gastos de Aquaz y Quimaroma inyectados.")
                    except Exception as e:
                        st.error(f"Error cargando Gastos: {e}")

    with tab_ventas:
        colV1, colV2 = st.columns(2)
        with colV1:
            st.markdown("#### 🔵 Ventas AQUAZ (FACEL)")
            archivo_facel = st.file_uploader("Arrastra reporte de Facel", type=["xlsx", "xls"], key="facel")
            if archivo_facel:
                if st.button("Procesar FACEL", use_container_width=True, type="primary"):
                    try:
                        xls_f = pd.ExcelFile(archivo_facel)
                        lista_ventas = [pd.read_excel(xls_f, sheet_name=h, header=1) for h in ['FACTURAS', 'BOLETAS DE VENTAS', 'NOTAS DE VENTAS'] if h in xls_f.sheet_names]
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
                            df_v['Zona'] = "No registrada" # Espacio listo para cuando lo agregues
                            
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
                            st.success(f"¡{len(df_v)} ventas de Aquaz inyectadas con éxito!")
                    except Exception as e:
                        st.error(f"Error procesando Facel: {e}")

        with colV2:
            st.markdown("#### 🟢 Ventas QUIMAROMA (Tienda)")
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
                        st.success(f"¡{len(df_q)} ventas de Quimaroma inyectadas!")
                    except Exception as e:
                        st.error(f"Error procesando Quimaroma: {e}")

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
                    st.success(f"¡{len(df_p)} registros de planta inyectados!")
                except Exception as e:
                    st.error(f"Error procesando Producción: {e}")

# ==========================================
# 4. PANTALLA: VENTAS Y ANALÍTICA (ACTUALIZADO A SELECTOR)
# ==========================================
elif menu == "💰 1. Ventas & Analítica":
    st.title("💰 Análisis de Ventas y Fidelización")
    
    # Logica inteligente que respeta tu menú lateral
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    
    if empresa_activa == "Aquaz (Planta/Mayorista)":
        df = df_aquaz
    elif empresa_activa == "Quinearoma (Tienda Fiori)":
        df = df_quima
    else:
        df = pd.concat([df_aquaz, df_quima], ignore_index=True)
    
    if df is None or df.empty:
        st.info(f"Sube datos de ventas para {empresa_activa} para ver la analítica.")
    else:
        df['Descuento'] = pd.to_numeric(df.get('Descuento', 0), errors='coerce').fillna(0)
        df['Utilidad_Bruta'] = (df['Cantidad'] * df['Precio_Venta']) - df['Descuento'] - (df['Cantidad'] * df['Costo_Unitario'])
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Rendimiento por Vendedor (Soles)")
            st.bar_chart(df.groupby('Vendedor')['Utilidad_Bruta'].sum().sort_values(ascending=False))
        with col2:
            st.subheader("Top Clientes más Rentables")
            st.bar_chart(df.groupby('Cliente')['Utilidad_Bruta'].sum().sort_values(ascending=False).head(10))

        st.markdown("---")
        st.subheader("📱 Métricas Digitales (ROI/CAC)")
        c1, c2, c3 = st.columns(3)
        ventas_cerradas = c1.number_input("Ventas por Redes (S/)", value=5000.0)
        sueldo = c2.number_input("Sueldo Vendedora (S/)", value=1025.0)
        inversion = c3.number_input("Inversión Ads (S/)", value=300.0)
        roi = (ventas_cerradas - (sueldo + inversion)) / (sueldo + inversion) * 100 if (sueldo + inversion) > 0 else 0
        st.metric("ROI Digital", f"{roi:.1f}%")

# ==========================================
# 5. PANTALLA: PRODUCCIÓN Y MRP (EXCLUSIVO AQUAZ)
# ==========================================
elif menu == "🏭 2. Producción & MRP":
    st.title("🏭 Eficiencia de Planta y MRP Predictivo")
    
    if empresa_activa == "Quinearoma (Tienda Fiori)":
        st.info("⚠️ El módulo de Producción y MRP es exclusivo de la fábrica (AQUAZ). Por favor cambia tu 'Entorno de Trabajo' en el panel izquierdo.")
    else:
        df = st.session_state.dfs.get('Produccion')
        if df is not None and not df.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Volumen Fabricado")
                if 'Cantidad_Producida' in df.columns:
                    st.bar_chart(df.groupby('Producto')['Cantidad_Producida'].sum().sort_values(ascending=False).head(10))
            with col2:
                st.subheader("Pérdidas por Mermas (Soles)")
                if 'Merma_Soles' in df.columns:
                    st.bar_chart(df.groupby('Producto')['Merma_Soles'].sum().sort_values(ascending=False).head(10), color="#ff4b4b")

        st.markdown("---")
        st.subheader("⚙️ Planificación de Producción (MRP)")
        c1, c2, c3, c4 = st.columns(4)
        ventas_promedio = c1.number_input("Ventas Semanales Promedio", value=500)
        stock_seguridad = c2.number_input("Stock Seguridad", value=200)
        stock_actual = c3.number_input("Stock Actual", value=150)
        lote_optimo = c4.number_input("Lote Óptimo de Máquina", value=200)
        
        necesidad_bruta = (ventas_promedio + stock_seguridad) - stock_actual
        orden = np.ceil(necesidad_bruta / lote_optimo) * lote_optimo if lote_optimo > 0 else 0
        st.success(f"**Orden de Producción Sugerida:** {orden:,.0f} Unidades")

# ==========================================
# 6. PANTALLA: FINANZAS Y COSTOS (INTACTO)
# ==========================================
elif menu == "⚖️ 3. Finanzas & Costos":
    st.title("⚖️ Balance Financiero y Costos Estratégicos")
    
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    
    if empresa_activa == "Aquaz (Planta/Mayorista)":
        df_v = df_aquaz
    elif empresa_activa == "Quinearoma (Tienda Fiori)":
        df_v = df_quima
    else:
        df_v = pd.concat([df_aquaz, df_quima], ignore_index=True)
        
    df_g = st.session_state.dfs.get('Gastos', pd.DataFrame())
    df_p = st.session_state.dfs.get('Produccion', pd.DataFrame())

    utilidad_bruta = 0
    if not df_v.empty:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        utilidad_bruta = ((df_v['Cantidad'] * df_v['Precio_Venta']).sum() - df_v['Descuento'].sum()) - (df_v['Cantidad'] * df_v['Costo_Unitario']).sum()

    gastos_fijos = df_g['Monto'].sum() if not df_g.empty else 0
    mermas = df_p['Merma_Soles'].sum() if not df_p.empty else 0
    resultado_final = utilidad_bruta - gastos_fijos - mermas

    st.markdown("### Resumen Consolidado")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("1. Utilidad Comercial", f"S/ {utilidad_bruta:,.2f}")
    col2.metric("2. Pérdidas Planta", f"- S/ {mermas:,.2f}")
    col3.metric("3. Gastos Fijos", f"- S/ {gastos_fijos:,.2f}")
    col4.metric("💰 RESULTADO NETO", f"S/ {resultado_final:,.2f}")

    st.markdown("---")
    colA, colB = st.columns(2)
    
    with colA:
        st.subheader("⚖️ Punto de Equilibrio")
        margen_pond = st.slider("Margen de Contribución Promedio (%)", 1, 100, 35) / 100
        pe = gastos_fijos / margen_pond if margen_pond > 0 else 0
        st.metric("Meta Mínima para no perder dinero", f"S/ {pe:,.2f}")
    
    with colB:
        st.subheader("🚦 Semáforo de Rentabilidad por Producto")
        if not df_v.empty:
            df_semaforo = df_v.groupby('Producto').agg({'Precio_Venta': 'mean', 'Costo_Unitario': 'mean'}).reset_index()
            df_semaforo['Margen (S/)'] = df_semaforo['Precio_Venta'] - df_semaforo['Costo_Unitario']
            
            def semaforo(fila):
                return ['background-color: #fee2e2; color: #991b1b; font-weight: bold' if fila['Margen (S/)'] < 0 else '' for _ in fila]
            
            st.dataframe(df_semaforo.style.apply(semaforo, axis=1), use_container_width=True)
        else:
            st.info("Sube datos para calcular rentabilidad.")

# ==========================================
# 7. PANTALLA: INVENTARIO (INTACTO)
# ==========================================
elif menu == "📦 4. Inventario":
    st.title("📦 Inventario en Tiempo Real")
    df = st.session_state.dfs.get('Inventario')
    
    if df is None or df.empty:
        st.info("No hay datos de Inventario en la matriz.")
    else:
        col_evaluar = 'Stock' if 'Stock' in df.columns else ('Cantidad' if 'Cantidad' in df.columns else None)
        if col_evaluar:
            valores_numericos = pd.to_numeric(df[col_evaluar], errors='coerce')
            criticos = df[valores_numericos < 0]
            if not criticos.empty:
                st.error(f"⚠️ Alertas de Stock Crítico: **{len(criticos)}** insumos en negativo. ↑ Revisar urgente.")
            else:
                st.success("✅ Stock saludable.")
        st.dataframe(df.style.apply(resaltar_stock_critico, axis=1), use_container_width=True, height=500)

# ==========================================
# 8. PANTALLA: TIENDA FIORI (INTACTO)
# ==========================================
elif menu == "🏬 5. Tienda Fiori (Unit Economics)":
    st.title("🏬 Rentabilidad Diaria y por Cotización - Quinearoma Fiori")
    st.write("Simulador de rentabilidad multiproducto para tickets de mostrador.")

    st.subheader("🏢 1. Costo Operativo del Local")
    cf_mensual = st.number_input("Costos Fijos del Local (Mensual) (S/)", value=2500.0, step=100.0)
    cuota_diaria = cf_mensual / 30
    st.info(f"💡 Tu local necesita generar **S/ {cuota_diaria:,.2f}** de utilidad bruta diaria (Punto de Equilibrio) para no perder dinero.")

    st.markdown("---")
    st.subheader("🛒 2. Simulador de Pedido (Calculadora por Ítem)")
    
    if 'carrito_fiori' not in st.session_state:
        st.session_state['carrito_fiori'] = pd.DataFrame({
            "Producto": ["Texapon", "", ""],
            "Cantidad": [1, 0, 0],
            "Costo Unitario (S/)": [15.0, 0.0, 0.0],
            "Precio Venta Unit. (S/)": [20.0, 0.0, 0.0]
        })

    st.write("Escribe los productos del pedido. **Para agregar más filas, simplemente haz clic en la última fila en blanco o arrastra hacia abajo.**")
    
    df_pedido = st.data_editor(
        st.session_state['carrito_fiori'],
        num_rows="dynamic", 
        use_container_width=True,
        hide_index=True
    )
    
    st.session_state['carrito_fiori'] = df_pedido 
    
    gasto_variable = st.number_input("Gastos Extra del Pedido (Bolsas, pasajes, comisión, etc.) (S/)", value=0.0, step=5.0)

    df_pedido['Cantidad'] = pd.to_numeric(df_pedido['Cantidad'], errors='coerce').fillna(0)
    df_pedido['Costo Unitario (S/)'] = pd.to_numeric(df_pedido['Costo Unitario (S/)'], errors='coerce').fillna(0)
    df_pedido['Precio Venta Unit. (S/)'] = pd.to_numeric(df_pedido['Precio Venta Unit. (S/)'], errors='coerce').fillna(0)

    costo_total = (df_pedido['Cantidad'] * df_pedido['Costo Unitario (S/)']).sum()
    venta_total = (df_pedido['Cantidad'] * df_pedido['Precio Venta Unit. (S/)']).sum()
    
    utilidad_pedido = venta_total - costo_total - gasto_variable
    margen_pedido = (utilidad_pedido / venta_total * 100) if venta_total > 0 else 0
    cobertura = (utilidad_pedido / cuota_diaria * 100) if cuota_diaria > 0 else 0

    st.markdown("---")
    
    st.markdown("#### 📊 Rentabilidad de la Cotización:")
    cA, cB, cC, cD = st.columns(4)
    cA.metric("Venta Bruta Total", f"S/ {venta_total:,.2f}")
    cB.metric("Utilidad Neta (Bolsillo)", f"S/ {utilidad_pedido:,.2f}")
    cC.metric("Margen de Ganancia", f"{margen_pedido:.1f}%")
    cD.metric("Cobertura del Día", f"{cobertura:.1f}%")

    if utilidad_pedido > 0:
        st.progress(min(cobertura / 100, 1.0))
        if cobertura >= 100:
            st.success("🎉 ¡Boom! Este ticket por sí solo acaba de pagar el costo operativo del día entero en la tienda.")
        else:
            st.warning(f"Rentable. Este ticket cubrió un tramo, te falta cubrir el **{max(0, 100 - cobertura):.1f}%** de los gastos operativos de hoy con otras ventas.")
    elif utilidad_pedido < 0:
        st.error("⚠️ ¡ALTO! Este pedido te está generando pérdidas. Revisa tus precios de venta o gastos extra.")

# ==========================================
# 9. PANTALLA: RETENCIÓN DE CLIENTES (ACTUALIZADA)
# ==========================================
elif menu == "👥 6. Retención de Clientes":
    st.title("👥 Radar de Retención y Valor de Cliente")
    st.write("Clasificación automática basada en la última fecha de compra.")
    
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    
    if empresa_activa == "Aquaz (Planta/Mayorista)":
        df_v = df_aquaz
    elif empresa_activa == "Quinearoma (Tienda Fiori)":
        df_v = df_quima
    else:
        df_v = pd.concat([df_aquaz, df_quima], ignore_index=True)
    
    if df_v is None or df_v.empty:
        st.warning(f"⚠️ No hay datos cargados para {empresa_activa}.")
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Utilidad_Bruta'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        
        # Validar columnas Vendedor y Zona por si acaso
        if 'Vendedor' not in df_v.columns: df_v['Vendedor'] = 'Sin Vendedor'
        if 'Zona' not in df_v.columns: df_v['Zona'] = 'No registrada'

        hoy = pd.Timestamp.today()
        
        # Agrupación con Vendedor y Zona
        df_clientes = df_v.groupby('Cliente').agg(
            Ultima_Compra=('Fecha', 'max'),
            Frecuencia_Compras=('Fecha', 'nunique'),
            Utilidad_Total=('Utilidad_Bruta', 'sum'),
            Vendedor=('Vendedor', 'last'),
            Zona=('Zona', 'last')
        ).reset_index()
        
        df_clientes['Días Sin Comprar'] = (hoy - df_clientes['Ultima_Compra']).dt.days
        
        # Los 5 bloques de tiempo exactos que pediste
        c20 = df_clientes[df_clientes['Días Sin Comprar'] <= 20].sort_values(by='Utilidad_Total', ascending=False)
        c30 = df_clientes[(df_clientes['Días Sin Comprar'] > 20) & (df_clientes['Días Sin Comprar'] <= 30)].sort_values(by='Utilidad_Total', ascending=False)
        c45 = df_clientes[(df_clientes['Días Sin Comprar'] > 30) & (df_clientes['Días Sin Comprar'] <= 45)].sort_values(by='Utilidad_Total', ascending=False)
        c60 = df_clientes[(df_clientes['Días Sin Comprar'] > 45) & (df_clientes['Días Sin Comprar'] <= 60)].sort_values(by='Utilidad_Total', ascending=False)
        c_dorm = df_clientes[df_clientes['Días Sin Comprar'] > 60].sort_values(by='Utilidad_Total', ascending=False)
        
        t1, t2, t3, t4, t5 = st.tabs(["🟢 < 20 días", "🟡 21-30 días", "🟠 31-45 días", "🔴 46-60 días", "⚫ +60 días (Dormidos)"])
        
        # Columnas a mostrar incluyendo Vendedor y Zona
        cols_vista = ['Cliente', 'Vendedor', 'Zona', 'Días Sin Comprar', 'Ultima_Compra', 'Utilidad_Total']
        
        with t1:
            st.subheader(f"🟢 Clientes Activos (Hace 20 días o menos) - Total: {len(c20)}")
            st.dataframe(c20[cols_vista], use_container_width=True, hide_index=True)
            
        with t2:
            st.subheader(f"🟡 Clientes Regulares (Entre 21 y 30 días) - Total: {len(c30)}")
            st.dataframe(c30[cols_vista], use_container_width=True, hide_index=True)
            
        with t3:
            st.subheader(f"🟠 Alerta Temprana (Entre 31 y 45 días) - Total: {len(c45)}")
            st.dataframe(c45[cols_vista], use_container_width=True, hide_index=True)
            
        with t4:
            st.subheader(f"🔴 En Riesgo (Entre 46 y 60 días) - Total: {len(c60)}")
            st.write("¡Recomendación: Asignar a su vendedor correspondiente para llamarlos hoy!")
            st.dataframe(c60[cols_vista], use_container_width=True, hide_index=True)
            
        with t5:
            st.subheader(f"⚫ Clientes Dormidos (Más de 60 días) - Total: {len(c_dorm)}")
            st.dataframe(c_dorm[cols_vista], use_container_width=True, hide_index=True)

# ==========================================
# 10. DASHBOARD INICIAL
# ==========================================
elif menu == "📊 Inicio (Dashboard)":
    st.title("📊 Panel de Control Principal")
    st.write(f"Resumen gerencial de **{empresa_activa}**.")
    df_v = st.session_state.dfs.get('Ventas', pd.DataFrame())
    if not df_v.empty:
        st.success("Sistema conectado. Usa el menú lateral para navegar por los módulos.")
    else:
        st.warning("👈 Ve a la pestaña 'Carga de Datos' para arrancar el sistema.")
