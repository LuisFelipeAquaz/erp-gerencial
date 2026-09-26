import streamlit as st
import pandas as pd

from core.utils import convert_to_excel, filtrar_empresa, norm_texto

MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def _unir(serie):
    return ", ".join(sorted({str(x).strip() for x in serie if str(x).strip() not in ("", "nan", "None")}))


def construir_base_clientes(df_v: pd.DataFrame) -> pd.DataFrame:
    """Una fila por cliente, agrupando por RUC/DNI (o por nombre si no tiene documento válido)."""
    ventas = df_v[df_v["Venta_Neta"] > 0]  # compras reales (sin notas de crédito ni gratuitos)
    tiene_comp = df_v["Comprobante"].astype(str).str.len() > 0

    base = df_v.groupby("Cliente_ID").agg(
        Cliente=("Cliente", lambda s: s.value_counts().index[0]),
        Doc=("Cliente_Doc", "first"),
        Tipo_Doc=("Tipo_Doc", "first"),
        Empresa=("Empresa", _unir),
        Vendedor=("Vendedor", lambda s: s.value_counts().index[0]),
        Zona=("Zona", "last"),
        Venta_Neta=("Venta_Neta", "sum"),
        Utilidad_Total=("Utilidad_Bruta", "sum"),
    )
    compras = ventas.groupby("Cliente_ID").agg(
        Primera_Compra=("Fecha", "min"),
        Ultima_Compra=("Fecha", "max"),
    )
    # Frecuencia = comprobantes distintos (datos antiguos sin comprobante: días distintos)
    frec_comp = ventas[tiene_comp.loc[ventas.index]].groupby("Cliente_ID")["Comprobante"].nunique()
    frec_dias = ventas[~tiene_comp.loc[ventas.index]].groupby("Cliente_ID")["Fecha"].nunique()
    base = base.join(compras)
    base["Compras"] = frec_comp.reindex(base.index).fillna(0) + frec_dias.reindex(base.index).fillna(0)
    base["Compras"] = base["Compras"].astype(int)
    if "Placa" in df_v.columns:
        base["Placas"] = df_v.groupby("Cliente_ID")["Placa"].agg(_unir).reindex(base.index).fillna("")
    base["Ticket_Promedio"] = (base["Venta_Neta"] / base["Compras"].where(base["Compras"] > 0)).round(2)
    hoy = pd.Timestamp.today().normalize()
    base["Días Sin Comprar"] = (hoy - base["Ultima_Compra"]).dt.days.astype("Int64")
    base["Venta_Neta"] = base["Venta_Neta"].round(2)
    base["Utilidad_Total"] = base["Utilidad_Total"].round(2)
    base["Doc"] = base["Doc"].where(base["Tipo_Doc"] != "SIN DOC", "")
    return base.reset_index(drop=True).sort_values("Utilidad_Total", ascending=False)


def render(empresa_activa):
    st.title("👥 Radar de Retención y Valor de Cliente")
    df_v = filtrar_empresa(empresa_activa)

    if df_v.empty:
        st.warning(f"⚠️ No hay datos cargados para {empresa_activa}.")
        return

    df_clientes = construir_base_clientes(df_v)

    vendedores = ['Todos'] + sorted(df_clientes['Vendedor'].astype(str).unique())
    c_f1, c_f2 = st.columns([1, 2])
    vendedor_sel = c_f1.selectbox("Filtrar por Vendedor:", vendedores)
    buscar = c_f2.text_input("Buscar cliente (nombre, RUC/DNI o placa)")
    if vendedor_sel != 'Todos':
        df_clientes = df_clientes[df_clientes['Vendedor'] == vendedor_sel]
    if buscar:
        q = norm_texto(buscar)
        texto = (df_clientes['Cliente'] + " " + df_clientes['Doc'] + " " +
                 df_clientes.get('Placas', pd.Series("", index=df_clientes.index))).map(norm_texto)
        df_clientes = df_clientes[texto.str.contains(q, regex=False)]

    sin_doc = (df_clientes['Tipo_Doc'] == 'SIN DOC').sum()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Clientes", f"{len(df_clientes):,}")
    m2.metric("Con RUC / DNI", f"{(df_clientes['Tipo_Doc'] == 'RUC').sum()} / {(df_clientes['Tipo_Doc'] == 'DNI').sum()}")
    m3.metric("Sin documento válido", int(sin_doc))
    m4.metric("Dormidos (+60 días)", int((df_clientes['Días Sin Comprar'] > 60).sum()))
    if sin_doc:
        st.caption("Los clientes sin RUC/DNI válido (vacío, '-', códigos internos o teléfonos) se agrupan por nombre; "
                   "si el nombre se escribe distinto en otra venta, aparecerán separados.")

    dias = df_clientes['Días Sin Comprar']
    tramos = [
        ("🟢 < 20 días", "🟢 Activos", "activos", dias <= 20),
        ("🟡 21-30 días", "🟡 Regulares", "regulares", (dias > 20) & (dias <= 30)),
        ("🟠 31-45 días", "🟠 Alerta Temprana", "alerta", (dias > 30) & (dias <= 45)),
        ("🔴 46-60 días", "🔴 En Riesgo", "riesgo", (dias > 45) & (dias <= 60)),
        ("⚫ +60 días", "⚫ Dormidos", "dormidos", dias > 60),
    ]
    cols = ['Cliente', 'Doc', 'Vendedor', 'Zona', 'Días Sin Comprar', 'Ultima_Compra', 'Compras', 'Utilidad_Total']
    cols_base = ['Tipo_Doc', 'Doc', 'Cliente', 'Empresa', 'Placas', 'Vendedor', 'Zona', 'Compras', 'Primera_Compra',
                 'Ultima_Compra', 'Días Sin Comprar', 'Venta_Neta', 'Ticket_Promedio', 'Utilidad_Total']
    cols_base = [c for c in cols_base if c in df_clientes.columns]

    tabs = st.tabs([t[0] for t in tramos] + ["📋 Base completa"])
    formato = {"Ultima_Compra": st.column_config.DateColumn("Última compra", format="DD/MM/YYYY"),
               "Primera_Compra": st.column_config.DateColumn("Primera compra", format="DD/MM/YYYY")}

    for tab, (_, titulo, archivo, filtro) in zip(tabs, tramos):
        with tab:
            grupo = df_clientes[filtro]
            st.subheader(f"{titulo} - Total: {len(grupo)}")
            if not grupo.empty:
                st.download_button("📥 Descargar (.xlsx)", data=convert_to_excel(grupo[cols], sheet_name='Retencion'),
                                   file_name=f'{archivo}_{vendedor_sel}.xlsx', mime=MIME_XLSX, key=f"dl_{archivo}")
            st.dataframe(grupo[cols], use_container_width=True, hide_index=True, column_config=formato)

    with tabs[-1]:
        st.subheader(f"📋 Base de clientes - Total: {len(df_clientes)}")
        st.download_button("📥 Descargar base de clientes (.xlsx)",
                           data=convert_to_excel(df_clientes[cols_base], sheet_name='Clientes'),
                           file_name=f"base_clientes_{pd.Timestamp.today():%Y-%m-%d}.xlsx", mime=MIME_XLSX,
                           key="dl_base")
        st.dataframe(df_clientes[cols_base], use_container_width=True, hide_index=True, column_config=formato)
