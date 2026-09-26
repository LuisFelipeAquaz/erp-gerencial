import streamlit as st
import pandas as pd


def render():
    st.title("🏬 Rentabilidad por Cotización")
    cf_mensual = st.number_input("Costos Fijos Mensuales (S/)", value=2500.0)
    cuota_diaria = cf_mensual / 30

    if 'carrito_fiori' not in st.session_state:
        st.session_state['carrito_fiori'] = pd.DataFrame({
            "Producto": ["Texapon", ""],
            "Cantidad": [1, 0],
            "Costo Unitario (S/)": [15.0, 0.0],
            "Precio Venta Unit. (S/)": [20.0, 0.0]
        })

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
