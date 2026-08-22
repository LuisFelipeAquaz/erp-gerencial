# ==========================================
# 9. PANTALLA: RETENCIÓN DE CLIENTES (ACTUALIZADA CON FILTROS Y DESCARGAS)
# ==========================================
elif menu == "👥 6. Retención de Clientes":
    st.title("👥 Radar de Retención y Valor de Cliente")
    st.write("Clasificación automática basada en la última fecha de compra.")
    
    df_aquaz = st.session_state.dfs.get('Ventas', pd.DataFrame())
    df_quima = st.session_state.dfs.get('Ventas_Quima', pd.DataFrame())
    
    if empresa_activa == "Aquaz (Planta/Mayorista)":
        df_v = df_aquaz
    elif empresa_activa == "Quimaroma (Tienda Fiori)":
        df_v = df_quima
    else:
        df_v = pd.concat([df_aquaz, df_quima], ignore_index=True)
    
    if df_v is None or df_v.empty:
        st.warning(f"⚠️ No hay datos cargados para {empresa_activa}.")
    else:
        df_v['Descuento'] = pd.to_numeric(df_v.get('Descuento', 0), errors='coerce').fillna(0)
        df_v['Utilidad_Bruta'] = (df_v['Cantidad'] * df_v['Precio_Venta']) - df_v['Descuento'] - (df_v['Cantidad'] * df_v['Costo_Unitario'])
        
        if 'Vendedor' not in df_v.columns: df_v['Vendedor'] = 'Sin Vendedor'
        if 'Zona' not in df_v.columns: df_v['Zona'] = 'No registrada'

        hoy = pd.Timestamp.today()
        
        df_clientes = df_v.groupby('Cliente').agg(
            Ultima_Compra=('Fecha', 'max'),
            Frecuencia_Compras=('Fecha', 'nunique'),
            Utilidad_Total=('Utilidad_Bruta', 'sum'),
            Vendedor=('Vendedor', 'last'),
            Zona=('Zona', 'last')
        ).reset_index()
        
        df_clientes['Días Sin Comprar'] = (hoy - df_clientes['Ultima_Compra']).dt.days
        
        # --- NUEVO: FILTRO POR VENDEDOR ---
        st.markdown("---")
        st.subheader("🎯 Filtro Gerencial para Asignación de Tareas")
        vendedores_disponibles = ['Todos'] + sorted(list(df_clientes['Vendedor'].unique()))
        vendedor_seleccionado = st.selectbox("Selecciona un Vendedor para filtrar toda la lista:", vendedores_disponibles)
        
        if vendedor_seleccionado != 'Todos':
            df_clientes = df_clientes[df_clientes['Vendedor'] == vendedor_seleccionado]
            st.info(f"Mostrando únicamente los clientes asignados a: **{vendedor_seleccionado}**")
        st.markdown("---")
        
        c20 = df_clientes[df_clientes['Días Sin Comprar'] <= 20].sort_values(by='Utilidad_Total', ascending=False)
        c30 = df_clientes[(df_clientes['Días Sin Comprar'] > 20) & (df_clientes['Días Sin Comprar'] <= 30)].sort_values(by='Utilidad_Total', ascending=False)
        c45 = df_clientes[(df_clientes['Días Sin Comprar'] > 30) & (df_clientes['Días Sin Comprar'] <= 45)].sort_values(by='Utilidad_Total', ascending=False)
        c60 = df_clientes[(df_clientes['Días Sin Comprar'] > 45) & (df_clientes['Días Sin Comprar'] <= 60)].sort_values(by='Utilidad_Total', ascending=False)
        c_dorm = df_clientes[df_clientes['Días Sin Comprar'] > 60].sort_values(by='Utilidad_Total', ascending=False)
        
        t1, t2, t3, t4, t5 = st.tabs(["🟢 < 20 días", "🟡 21-30 días", "🟠 31-45 días", "🔴 46-60 días", "⚫ +60 días (Dormidos)"])
        
        cols_vista = ['Cliente', 'Vendedor', 'Zona', 'Días Sin Comprar', 'Ultima_Compra', 'Utilidad_Total']
        
        # Función para convertir dataframe a CSV para descargar
        @st.cache_data
        def convert_df(df):
            return df.to_csv(index=False).encode('utf-8')
            
        with t1:
            st.subheader(f"🟢 Clientes Activos (Hace 20 días o menos) - Total: {len(c20)}")
            if not c20.empty:
                st.download_button(label="📥 Descargar lista (.csv)", data=convert_df(c20[cols_vista]), file_name=f'activos_20dias_{vendedor_seleccionado}.csv', mime='text/csv')
            st.dataframe(c20[cols_vista], use_container_width=True, hide_index=True)
            
        with t2:
            st.subheader(f"🟡 Clientes Regulares (Entre 21 y 30 días) - Total: {len(c30)}")
            if not c30.empty:
                st.download_button(label="📥 Descargar lista (.csv)", data=convert_df(c30[cols_vista]), file_name=f'regulares_30dias_{vendedor_seleccionado}.csv', mime='text/csv')
            st.dataframe(c30[cols_vista], use_container_width=True, hide_index=True)
            
        with t3:
            st.subheader(f"🟠 Alerta Temprana (Entre 31 y 45 días) - Total: {len(c45)}")
            if not c45.empty:
                st.download_button(label="📥 Descargar lista (.csv)", data=convert_df(c45[cols_vista]), file_name=f'alerta_45dias_{vendedor_seleccionado}.csv', mime='text/csv')
            st.dataframe(c45[cols_vista], use_container_width=True, hide_index=True)
            
        with t4:
            st.subheader(f"🔴 En Riesgo (Entre 46 y 60 días) - Total: {len(c60)}")
            st.write("¡Recomendación: Asignar a su vendedor correspondiente para llamarlos hoy!")
            if not c60.empty:
                st.download_button(label="📥 Descargar lista (.csv)", data=convert_df(c60[cols_vista]), file_name=f'riesgo_60dias_{vendedor_seleccionado}.csv', mime='text/csv')
            st.dataframe(c60[cols_vista], use_container_width=True, hide_index=True)
            
        with t5:
            st.subheader(f"⚫ Clientes Dormidos (Más de 60 días) - Total: {len(c_dorm)}")
            if not c_dorm.empty:
                st.download_button(label="📥 Descargar lista (.csv)", data=convert_df(c_dorm[cols_vista]), file_name=f'dormidos_mas60dias_{vendedor_seleccionado}.csv', mime='text/csv')
            st.dataframe(c_dorm[cols_vista], use_container_width=True, hide_index=True)
