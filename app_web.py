# ==========================================
# 10. PANTALLA: RETENCIÓN DE CLIENTES (CRM)
# ==========================================
elif menu == "👥 6. Retención de Clientes":
    st.title("👥 Radar de Retención y Valor de Cliente")
    st.write("Clasificación automática basada en la última fecha de compra en FACEL.")
    
    df_v = st.session_state.dfs.get('Ventas')
    
    if df_v is None or df_v.empty:
        st.warning("No hay datos de Facel. Sube tus ventas en la pestaña 'Carga de Datos'.")
    else:
        # Calcular Utilidad por si no estaba calculada
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Utilidad_Bruta'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        
        # Agrupar la historia de cada cliente
        hoy = pd.Timestamp.today()
        df_clientes = df_v.groupby('Cliente').agg(
            Ultima_Compra=('Fecha', 'max'),
            Frecuencia_Compras=('Fecha', 'nunique'),
            Utilidad_Total=('Utilidad_Bruta', 'sum')
        ).reset_index()
        
        # Calcular días inactivos
        df_clientes['Días Sin Comprar'] = (hoy - df_clientes['Ultima_Compra']).dt.days
        
        # Dividir a los clientes en los 3 grupos
        activos = df_clientes[df_clientes['Días Sin Comprar'] <= 30].sort_values(by='Utilidad_Total', ascending=False)
        en_riesgo = df_clientes[(df_clientes['Días Sin Comprar'] > 30) & (df_clientes['Días Sin Comprar'] <= 60)].sort_values(by='Días Sin Comprar')
        dormidos = df_clientes[df_clientes['Días Sin Comprar'] > 60].sort_values(by='Utilidad_Total', ascending=False)
        
        # Mostrar los paneles
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
