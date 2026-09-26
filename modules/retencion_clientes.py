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
        Utilidad_Total=("Utilidad_Bruta", lambda s: s.sum(min_count=1)),  # en blanco si no hay costos
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

    # Qué compraba habitualmente (top 3 por unidades) y cuál fue su última compra
    if not ventas.empty:
        top = (ventas.groupby(["Cliente_ID", "Producto"])["Cantidad"].sum().reset_index()
               .sort_values(["Cliente_ID", "Cantidad"], ascending=[True, False]))
        top = top[top["Producto"] != "SIN NOMBRE"].groupby("Cliente_ID").head(3)
        top["txt"] = top["Producto"] + " (" + top["Cantidad"].map(lambda x: f"{x:,.0f}") + ")"
        base["Productos_Habituales"] = top.groupby("Cliente_ID")["txt"].agg(", ".join).reindex(base.index).fillna("")
        ult = ventas.sort_values("Fecha").groupby("Cliente_ID").tail(1).set_index("Cliente_ID")
        base["Ultimo_Comprobante"] = ult["Comprobante"].reindex(base.index).fillna("")
    else:
        base["Productos_Habituales"] = ""
        base["Ultimo_Comprobante"] = ""
    return base.reset_index(drop=True).sort_values("Venta_Neta", ascending=False)


COLS_REPORTE = ['Cliente', 'Tipo_Doc', 'Doc', 'Empresa', 'Vendedor', 'Placas', 'Días Sin Comprar', 'Ultima_Compra',
                'Ultimo_Comprobante', 'Primera_Compra', 'Compras', 'Venta_Neta', 'Ticket_Promedio', 'Utilidad_Total',
                'Productos_Habituales']
NOMBRES_EXCEL = {'Tipo_Doc': 'Tipo doc', 'Doc': 'RUC / DNI', 'Ultima_Compra': 'Última compra',
                 'Ultimo_Comprobante': 'Último comprobante', 'Primera_Compra': 'Cliente desde',
                 'Compras': 'N° compras', 'Venta_Neta': 'Compró en total (S/ sin IGV)',
                 'Ticket_Promedio': 'Ticket promedio', 'Utilidad_Total': 'Utilidad',
                 'Productos_Habituales': 'Lo que más compraba (unidades)'}


def _excel(df: pd.DataFrame, hoja: str) -> bytes:
    cols = [c for c in COLS_REPORTE if c in df.columns]
    return convert_to_excel(df[cols].rename(columns=NOMBRES_EXCEL), sheet_name=hoja)


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
    buscar = c_f2.text_input("Buscar cliente (nombre, RUC/DNI, placa o producto)")
    if vendedor_sel != 'Todos':
        df_clientes = df_clientes[df_clientes['Vendedor'] == vendedor_sel]
    if buscar:
        q = norm_texto(buscar)
        texto = (df_clientes['Cliente'] + " " + df_clientes['Doc'] + " " +
                 df_clientes.get('Placas', pd.Series("", index=df_clientes.index)) + " " +
                 df_clientes['Productos_Habituales']).map(norm_texto)
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

    formato = {"Ultima_Compra": st.column_config.DateColumn("Última compra", format="DD/MM/YYYY"),
               "Primera_Compra": st.column_config.DateColumn("Cliente desde", format="DD/MM/YYYY"),
               "Venta_Neta": st.column_config.NumberColumn("Compró en total", format="%.2f"),
               "Productos_Habituales": st.column_config.TextColumn("Lo que más compraba", width="large")}
    hoy = f"{pd.Timestamp.today():%Y-%m-%d}"

    dias = df_clientes['Días Sin Comprar']
    tramos = [
        ("🟢 < 20 días", "🟢 Activos", "activos", dias <= 20),
        ("🟡 21-30 días", "🟡 Regulares", "regulares", (dias > 20) & (dias <= 30)),
        ("🟠 31-45 días", "🟠 Alerta Temprana", "alerta", (dias > 30) & (dias <= 45)),
        ("🔴 46-60 días", "🔴 En Riesgo", "riesgo", (dias > 45) & (dias <= 60)),
        ("⚫ +60 días", "⚫ Dormidos", "dormidos", dias > 60),
    ]
    tabs = st.tabs(["⏱️ Dejaron de comprar"] + [t[0] for t in tramos] + ["📋 Base completa"])

    # ---------- Periodo libre: el usuario decide cuántos días
    with tabs[0]:
        st.subheader("Clientes que dejaron de comprar")
        a1, a2, a3 = st.columns(3)
        desde = a1.number_input("Sin comprar hace MÁS de (días)", min_value=0, value=30, step=5)
        limitar = a2.checkbox("Poner un tope", help="Ej.: entre 30 y 90 días, para no incluir clientes perdidos hace años.")
        hasta = a2.number_input("…y MENOS de (días)", min_value=int(desde) + 1, value=max(int(desde) + 60, 90),
                                step=5, disabled=not limitar)
        orden = a3.selectbox("Ordenar por", ["Lo que compró en total (mayor primero)", "Más días sin comprar primero",
                                             "Más compras realizadas primero", "Nombre A-Z"])
        a4, a5 = st.columns(2)
        min_compras = a4.number_input("Que hayan comprado al menos (veces)", min_value=1, value=1,
                                      help="Sube este número para ver solo clientes que eran recurrentes.")
        min_monto = a5.number_input("Que hayan comprado en total al menos (S/)", min_value=0.0, value=0.0, step=100.0)

        filtro = (dias > desde) & (df_clientes['Compras'] >= min_compras) & (df_clientes['Venta_Neta'] >= min_monto)
        if limitar:
            filtro &= dias < hasta
        grupo = df_clientes[filtro.fillna(False)]
        columna, asc = {"Lo que compró en total (mayor primero)": ('Venta_Neta', False),
                        "Más días sin comprar primero": ('Días Sin Comprar', False),
                        "Más compras realizadas primero": ('Compras', False),
                        "Nombre A-Z": ('Cliente', True)}[orden]
        grupo = grupo.sort_values(columna, ascending=asc)

        k1, k2, k3 = st.columns(3)
        k1.metric("Clientes", f"{len(grupo):,}")
        k2.metric("Lo que compraban en total", f"S/ {grupo['Venta_Neta'].sum():,.2f}")
        k3.metric("Promedio de días sin comprar", f"{grupo['Días Sin Comprar'].mean():,.0f}" if len(grupo) else "—")

        if not grupo.empty:
            rango = f"{int(desde)}-{int(hasta)}" if limitar else f"mas_de_{int(desde)}"
            st.download_button("📥 Descargar reporte en Excel", data=_excel(grupo, "Dejaron de comprar"),
                               file_name=f"clientes_sin_comprar_{rango}_dias_{hoy}.xlsx", mime=MIME_XLSX,
                               type="primary", key="dl_libre")
        cols_v = [c for c in COLS_REPORTE if c in grupo.columns]
        st.dataframe(grupo[cols_v], use_container_width=True, hide_index=True, column_config=formato)

    # ---------- Tramos fijos
    for tab, (_, titulo, archivo, filtro) in zip(tabs[1:], tramos):
        with tab:
            grupo = df_clientes[filtro.fillna(False)].sort_values('Venta_Neta', ascending=False)
            st.subheader(f"{titulo} - Total: {len(grupo)}")
            if not grupo.empty:
                st.download_button("📥 Descargar (.xlsx)", data=_excel(grupo, 'Retencion'),
                                   file_name=f'{archivo}_{vendedor_sel}_{hoy}.xlsx', mime=MIME_XLSX, key=f"dl_{archivo}")
            st.dataframe(grupo[[c for c in COLS_REPORTE if c in grupo.columns]], use_container_width=True,
                         hide_index=True, column_config=formato)

    # ---------- Base completa
    with tabs[-1]:
        st.subheader(f"📋 Base de clientes - Total: {len(df_clientes)}")
        st.download_button("📥 Descargar base de clientes (.xlsx)", data=_excel(df_clientes, 'Clientes'),
                           file_name=f"base_clientes_{hoy}.xlsx", mime=MIME_XLSX, key="dl_base")
        st.dataframe(df_clientes[[c for c in COLS_REPORTE if c in df_clientes.columns]], use_container_width=True,
                     hide_index=True, column_config=formato)
