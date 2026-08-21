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

# Inicializamos la memoria a largo plazo blindada
if 'dfs' not in st.session_state:
    st.session_state.dfs = {h: pd.DataFrame() for h in ['Ventas', 'Produccion', 'Gastos', 'Inventario', 'CxC', 'Logistica', 'Calidad', 'RRHH']}

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
# 3. PANTALLA: CARGA DE DATOS
# ==========================================
if menu == "📥 Carga de Datos":
    st.title("📥 Inyección de Datos")
    st.write(f"Aquí actualizas la base de datos de **{empresa_activa}**.")
    
    if st.button("🗑️ Borrar toda la información en memoria y reiniciar", type="primary"):
        st.session_state.dfs = {h: pd.DataFrame() for h in ['Ventas', 'Produccion', 'Gastos', 'Inventario', 'CxC', 'Logistica', 'Calidad', 'RRHH']}
        st.success("¡Memoria borrada! El sistema está en cero.")
        st.rerun()

    st.markdown("---")

    st.markdown("### 1️⃣ Sube tu Matriz Principal")
    archivo_matriz = st.file_uploader("Arrastra tu archivo MATRIZ (Historial)", type=["xlsx", "xls"], key="matriz")

    # LA SOLUCIÓN ESTÁ AQUÍ: Ahora la Matriz requiere que presiones un botón para procesarse
    if archivo_matriz:
        if st.button("Procesar Matriz", use_container_width=True):
            try:
                xls = pd.ExcelFile(archivo_matriz)
                for h in st.session_state.dfs.keys():
                    if h in xls.sheet_names:
                        df_cargado = pd.read_excel(xls, h)
                        df_cargado.columns = df_cargado.columns.str.strip()
                        st.session_state.dfs[h] = df_cargado
                st.success("✅ Matriz cargada en memoria exitosamente.")
            except Exception as e:
                st.error(f"Error cargando Matriz: {e}")
            
    st.markdown("---")
        
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 🚀 Ventas (FACEL)")
        archivo_facel = st.file_uploader("Arrastra reporte de Facel", type=["xlsx", "xls"], key="facel")
        if archivo_facel:
            if st.button("Procesar Ventas", use_container_width=True, type="primary"):
                try:
                    xls_f = pd.ExcelFile(archivo_facel)
                    lista_ventas = [pd.read_excel(xls_f, sheet_name=h, header=1) for h in ['FACTURAS', 'BOLETAS DE VENTAS', 'NOTAS DE VENTAS'] if h in xls_f.sheet_names]
                    if lista_ventas:
                        df_bruto = pd.concat(lista_ventas, ignore_index=True)
                        df_bruto.columns = df_bruto.columns.str.strip()
                        df_v = pd.DataFrame()
                        df_v['Fecha'] = pd.to_datetime(df_bruto.get('FECHA EMISION', pd.Series(dtype=object)), format='%d/%m/%Y', errors='coerce')
                        df_v['Empresa'] = empresa_activa 
                        df_v['Cliente'] = df_bruto.get('CLIENTE NOMBRE', pd.Series(dtype=object)).fillna("CLIENTE VARIOS")
                        df_v['Vendedor'] = df_bruto.get('ATENDIDO POR', pd.Series(dtype=object)).fillna("TIENDA")
                        df_v['Producto'] = df_bruto.get('PRODUCTO/SERVICIO', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                        df_v['Cantidad'] = pd.to_numeric(df_bruto.get('CANTIDAD', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Precio_Venta'] = pd.to_numeric(df_bruto.get('PRECIO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Costo_Unitario'] = pd.to_numeric(df_bruto.get('COSTO UNITARIO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v['Descuento'] = pd.to_numeric(df_bruto.get('DESCUENTO', pd.Series(dtype=float)), errors='coerce').fillna(0)
                        df_v = df_v[df_v['Cantidad'] > 0]
                        
                        st.session_state.dfs['Ventas'] = pd.concat([st.session_state.dfs['Ventas'], df_v], ignore_index=True)
                        st.success(f"¡{len(df_v)} ventas inyectadas con éxito!")
                except Exception as e:
                    st.error(f"Error procesando Facel: {e}")

    with col2:
        st.markdown("#### 🏭 Planta (PRODUCCIÓN)")
        archivo_prod = st.file_uploader("Arrastra reporte de planta", type=["xlsx", "xls"], key="prod")
        if archivo_prod:
            if st.button("Procesar Producción", use_container_width=True, type="primary"):
                try:
                    df_bruto = pd.read_excel(archivo_prod)
                    df_bruto.columns = df_bruto.columns.str.strip()
                    df_p = pd.DataFrame()
                    df_p['Fecha'] = pd.to_datetime(df_bruto.get('Fecha', pd.Series(dtype=object)), errors='coerce')
                    df_p['Empresa'] = empresa_activa
                    df_p['Lote'] = df_bruto.get('Número de Lote (o de Orden)', pd.Series(dtype=object)).fillna("S/L")
                    df_p['Producto'] = df_bruto.get('Producto Fabricado', pd.Series(dtype=object)).fillna("SIN NOMBRE")
                    df_p['Cantidad_Producida'] = pd.to_numeric(df_bruto.get('Cantidad Producida', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Costo_Materia_Prima'] = pd.to_numeric(df_bruto.get('Costo de Materia Prima', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Merma_Soles'] = pd.to_numeric(df_bruto.get('Merma', pd.Series(dtype=float)), errors='coerce').fillna(0)
                    df_p['Operario'] = df_bruto.get('Operario / Responsable', pd.Series(dtype=object)).fillna("No especificado")
                        
                    st.session_state.dfs['Produccion'] = pd.concat([st.session_state.dfs['Produccion'], df_p], ignore_index=True)
                    st.success(f"¡{len(df_p)} registros inyectados con éxito!")
                except Exception as e:
                    st.error(f"Error procesando Producción: {e}")

# ==========================================
# 4. PANTALLA: VENTAS Y ANALÍTICA CHURN
# ==========================================
elif menu == "💰 1. Ventas & Analítica":
    st.title("💰 Análisis de Ventas y Fidelización")
    df = st.session_state.dfs.get('Ventas')
    
    if df is None or df.empty:
        st.info("Sube datos de ventas para ver la analítica.")
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
        st.subheader("🚨 Alarma de Deserción (Clientes inactivos > 60 días)")
        if 'Fecha' in df.columns:
            hoy = pd.Timestamp.today()
            df_clientes = df.groupby('Cliente')['Fecha'].max().reset_index()
            df_clientes['Días Inactivos'] = (hoy - df_clientes['Fecha']).dt.days
            
            def resaltar_churn(fila):
                return ['background-color: #fee2e2; color: #991b1b; font-weight: bold' if fila['Días Inactivos'] > 60 else '' for _ in fila]
            
            st.dataframe(df_clientes.sort_values(by='Días Inactivos', ascending=False).style.apply(resaltar_churn, axis=1), use_container_width=True)
        
        st.markdown("---")
        st.subheader("📱 Métricas Digitales (ROI/CAC)")
        c1, c2, c3 = st.columns(3)
        ventas_cerradas = c1.number_input("Ventas por Redes (S/)", value=5000.0)
        sueldo = c2.number_input("Sueldo Vendedora (S/)", value=1025.0)
        inversion = c3.number_input("Inversión Ads (S/)", value=300.0)
        roi = (ventas_cerradas - (sueldo + inversion)) / (sueldo + inversion) * 100 if (sueldo + inversion) > 0 else 0
        st.metric("ROI Digital", f"{roi:.1f}%")

# ==========================================
# 5. PANTALLA: PRODUCCIÓN Y MRP
# ==========================================
elif menu == "🏭 2. Producción & MRP":
    st.title("🏭 Eficiencia de Planta y MRP Predictivo")
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
# 6. PANTALLA: FINANZAS Y COSTOS 
# ==========================================
elif menu == "⚖️ 3. Finanzas & Costos":
    st.title("⚖️ Balance Financiero y Costos Estratégicos")
    
    df_v = st.session_state.dfs.get('Ventas', pd.DataFrame())
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
            st.info("Sube datos de Facel para calcular rentabilidad.")

# ==========================================
# 7. PANTALLA: INVENTARIO
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
# 8. PANTALLA: TIENDA FIORI (Unit Economics)
# ==========================================
elif menu == "🏬 5. Tienda Fiori (Unit Economics)":
    st.title("🏬 Rentabilidad Automatizada - Quinearoma Fiori")
    st.write("Análisis de rentabilidad por ticket, extrayendo datos directamente de FACEL.")

    # 1. Configuración de Gastos Fijos
    st.subheader("🏢 1. Costo Operativo del Local")
    cf_mensual = st.number_input("Costos Fijos del Local (Mensual) (S/)", value=2500.0, step=100.0)
    cuota_diaria = cf_mensual / 30
    st.info(f"💡 Tu local necesita generar **S/ {cuota_diaria:,.2f}** de utilidad bruta diaria para cubrir sus gastos.")

    st.markdown("---")
    
    # Llamamos a la base de datos de FACEL que ya subiste
    df_v = st.session_state.dfs.get('Ventas')
    
    if df_v is None or df_v.empty:
        st.warning("⚠️ No hay ventas cargadas. Ve a 'Carga de Datos' y sube tu archivo FACEL para automatizar este panel.")
    else:
        st.subheader("🛒 2. Extracción Automática de Pedidos")
        
        # Filtros inteligentes para encontrar el pedido rápido
        col1, col2 = st.columns(2)
        fechas_disponibles = df_v['Fecha'].dt.date.dropna().unique()
        
        with col1:
            fecha_seleccionada = st.selectbox("📅 Fecha de Venta:", sorted(fechas_disponibles, reverse=True))
        
        # Filtrar la base de datos solo por la fecha elegida
        df_dia = df_v[df_v['Fecha'].dt.date == fecha_seleccionada]
        
        with col2:
            clientes_disponibles = df_dia['Cliente'].unique()
            cliente_seleccionado = st.selectbox("👤 Selecciona el Cliente:", clientes_disponibles)
            
        # Extraer el ticket exacto de ese cliente
        df_ticket = df_dia[df_dia['Cliente'] == cliente_seleccionado].copy()
        
        st.write(f"**Detalle exacto del pedido de {cliente_seleccionado}:**")
        # Mostramos la tabla jalada automáticamente
        st.dataframe(df_ticket[['Producto', 'Cantidad', 'Precio_Venta', 'Costo_Unitario', 'Descuento']], use_container_width=True, hide_index=True)
        
        # Cálculos matemáticos automáticos
        venta_total = (df_ticket['Cantidad'] * df_ticket['Precio_Venta']).sum() - df_ticket['Descuento'].sum()
        costo_total = (df_ticket['Cantidad'] * df_ticket['Costo_Unitario']).sum()
        
        # Casilla opcional por si gastaste en un taxi o flete especial para este pedido
        gasto_variable = st.number_input("Gastos Extra (Opcional - Flete, empaque) (S/)", value=0.0, step=5.0)
        
        utilidad_pedido = venta_total - costo_total - gasto_variable
        margen_pedido = (utilidad_pedido / venta_total * 100) if venta_total > 0 else 0
        cobertura = (utilidad_pedido / cuota_diaria * 100) if cuota_diaria > 0 else 0

        st.markdown("---")
        
        # Panel de Resultados Rápidos
        st.markdown("#### 📊 Rentabilidad Real de este Ticket:")
        cA, cB, cC, cD = st.columns(4)
        cA.metric("Venta Bruta Total", f"S/ {venta_total:,.2f}")
        cB.metric("Utilidad Neta (Bolsillo)", f"S/ {utilidad_pedido:,.2f}")
        cC.metric("Margen de Ganancia", f"{margen_pedido:.1f}%")
        cD.metric("Cobertura del Día", f"{cobertura:.1f}%")

        # Termómetro visual de éxito
        if utilidad_pedido > 0:
            st.progress(min(cobertura / 100, 1.0))
            if cobertura >= 100:
                st.success("🎉 ¡Excelente! Este ticket por sí solo ya pagó la cuota operativa del local de hoy.")
            else:
                st.info(f"Ticket rentable. Ayudó a cubrir el **{cobertura:.1f}%** de los gastos operativos diarios.")
        elif utilidad_pedido < 0:
            st.error("⚠️ Alerta: Pérdida detectada en este ticket. Los costos superan a la venta.")
# ==========================================
# 9. PANTALLA: RETENCIÓN DE CLIENTES (CRM)
# ==========================================
elif menu == "👥 6. Retención de Clientes":
    st.title("👥 Radar de Retención y Valor de Cliente")
    st.write("Clasificación automática basada en la última fecha de compra en FACEL.")
    
    df_v = st.session_state.dfs.get('Ventas')
    
    if df_v is None or df_v.empty:
        st.warning("No hay datos de Facel. Sube tus ventas en la pestaña 'Carga de Datos'.")
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Utilidad_Bruta'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        
        hoy = pd.Timestamp.today()
        df_clientes = df_v.groupby('Cliente').agg(
            Ultima_Compra=('Fecha', 'max'),
            Frecuencia_Compras=('Fecha', 'nunique'),
            Utilidad_Total=('Utilidad_Bruta', 'sum')
        ).reset_index()
        
        df_clientes['Días Sin Comprar'] = (hoy - df_clientes['Ultima_Compra']).dt.days
        
        activos = df_clientes[df_clientes['Días Sin Comprar'] <= 30].sort_values(by='Utilidad_Total', ascending=False)
        en_riesgo = df_clientes[(df_clientes['Días Sin Comprar'] > 30) & (df_clientes['Días Sin Comprar'] <= 60)].sort_values(by='Días Sin Comprar')
        dormidos = df_clientes[df_clientes['Días Sin Comprar'] > 60].sort_values(by='Utilidad_Total', ascending=False)
        
        tab1, tab2, tab3 = st.tabs(["🟢 ACTIVOS (Últimos 30 días)", "🟡 EN RIESGO (31 - 60 días)", "🔴 DORMIDOS (+60 días)"])
        
        with tab1:
            st.subheader(f"🟢 Clientes Top Activos ({len(activos)})")
            st.write("Tus mejores clientes actuales. Cuídalos.")
            st.dataframe(activos[['Cliente', 'Ultima_Compra', 'Frecuencia_Compras', 'Utilidad_Total']], use_container_width=True)
            
        with tab2:
            st.subheader(f"🟡 ¡Alerta! Clientes En Riesgo ({len(en_riesgo)})")
            st.write("Han dejado de comprar este último mes. **Recomendación:** Asignar a un vendedor para llamarlos hoy mismo.")
            st.dataframe(en_riesgo[['Cliente', 'Días Sin Comprar', 'Ultima_Compra', 'Utilidad_Total']], use_container_width=True)
            
        with tab3:
            st.subheader(f"🔴 Clientes Dormidos ({len(dormidos)})")
            st.write("Hace más de 2 meses que no te compran. Ordenados por la plata que te dejaban antes de irse.")
            st.dataframe(dormidos[['Cliente', 'Días Sin Comprar', 'Utilidad_Total', 'Frecuencia_Compras']], use_container_width=True)

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
