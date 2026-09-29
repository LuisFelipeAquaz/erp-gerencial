import streamlit as st
import pandas as pd

from core.utils import convert_to_excel, filtrar_empresa, norm_texto
from modules.explorador_ventas import _coincide, _norm_vec

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
    return base.reset_index().sort_values("Venta_Neta", ascending=False)  # conserva Cliente_ID


def compras_de_productos(df_v: pd.DataFrame, categorias: list, productos: list, texto: str) -> pd.DataFrame:
    """
    Por cliente: qué compraba de los productos elegidos y cuándo fue la última vez.
    Devuelve vacío si no hay ningún filtro de productos.
    """
    if not (categorias or productos or str(texto or "").strip()):
        return pd.DataFrame()
    v = df_v[df_v["Venta_Neta"] > 0].copy()
    v["Categoria"] = v["Categoria"].fillna("").astype(str).replace({"": "Sin categoría"}) if "Categoria" in v else "Sin categoría"
    m = pd.Series(True, index=v.index)
    if categorias:
        m &= v["Categoria"].isin(categorias)
    if productos:
        m &= v["Producto"].isin(productos)
    if str(texto or "").strip():
        m &= _coincide(_norm_vec(v["Producto"] + " " + v["Categoria"]), texto)
    v = v[m]
    if v.empty:
        return pd.DataFrame(columns=["Cliente_ID", "Ultima_Prod", "Compras_Prod", "Unidades_Prod", "Venta_Prod",
                                     "Productos_Seguidos"]).astype({"Ultima_Prod": "datetime64[ns]"})
    pen = v[v["Moneda"] == "PEN"]
    r = v.groupby("Cliente_ID").agg(
        Ultima_Prod=("Fecha", "max"),
        Compras_Prod=("Comprobante", "nunique"),
        Unidades_Prod=("Cantidad", "sum"),
    )
    r["Venta_Prod"] = pen.groupby("Cliente_ID")["Total_Linea"].sum().reindex(r.index).fillna(0).round(2)
    top = (v.groupby(["Cliente_ID", "Producto"])["Cantidad"].sum().reset_index()
           .sort_values(["Cliente_ID", "Cantidad"], ascending=[True, False]))
    top["txt"] = top["Producto"] + " (" + top["Cantidad"].map(lambda x: f"{x:,.0f}") + ")"
    r["Productos_Seguidos"] = top.groupby("Cliente_ID")["txt"].agg(lambda x: ", ".join(list(x)[:5]))
    return r.reset_index()


COLS_REPORTE = ['Cliente', 'Tipo_Doc', 'Doc', 'Empresa', 'Vendedor', 'Placas', 'Días Sin Comprar', 'Ultima_Compra',
                'Ultimo_Comprobante', 'Primera_Compra', 'Compras', 'Venta_Neta', 'Ticket_Promedio', 'Utilidad_Total',
                'Productos_Habituales']
NOMBRES_EXCEL = {'Tipo_Doc': 'Tipo doc', 'Doc': 'RUC / DNI', 'Ultima_Compra': 'Última compra',
                 'Ultimo_Comprobante': 'Último comprobante', 'Primera_Compra': 'Cliente desde',
                 'Compras': 'N° compras', 'Venta_Neta': 'Compró en total (S/ sin IGV)',
                 'Ticket_Promedio': 'Ticket promedio', 'Utilidad_Total': 'Utilidad',
                 'Productos_Habituales': 'Lo que más compraba (unidades)',
                 'Productos_Seguidos': 'Productos elegidos que compraba (unidades)',
                 'Ultima_Prod': 'Última compra de esos productos', 'Dias_Prod': 'Días sin comprar esos productos',
                 'Unidades_Prod': 'Unidades de esos productos', 'Venta_Prod': 'Compró de esos productos (S/ con IGV)',
                 'Compras_Prod': 'Veces que compró esos productos'}
COLS_SEGUIMIENTO = ['Cliente', 'Tipo_Doc', 'Doc', 'Empresa', 'Vendedor', 'Placas', 'Dias_Prod', 'Ultima_Prod',
                    'Productos_Seguidos', 'Unidades_Prod', 'Venta_Prod', 'Compras_Prod', 'Días Sin Comprar',
                    'Ultima_Compra', 'Venta_Neta', 'Productos_Habituales']


def _excel(df: pd.DataFrame, hoja: str, cols_base: list = None) -> bytes:
    cols = [c for c in (cols_base or COLS_REPORTE) if c in df.columns]
    return convert_to_excel(df[cols].rename(columns=NOMBRES_EXCEL), sheet_name=hoja)


def render(empresa_activa):
    st.title("👥 Radar de Retención y Valor de Cliente")
    df_v = filtrar_empresa(empresa_activa, aplicar_periodo=False)  # los días sin comprar se miden con todo el historial
    st.caption("Esta sección usa siempre todo tu historial de ventas, sin importar el periodo elegido en la barra lateral.")

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

        # --- 1. ¿Qué productos quieres reconquistar?
        with st.container(border=True):
            st.markdown("##### 🎯 ¿A qué clientes quieres hacer seguimiento?")
            st.caption("Elige productos o categorías ganadoras para ver SOLO a los clientes que los compraban. "
                       "Deja todo en 'Todos' para ver a todos los clientes.")
            todos_prod = df_v[df_v["Venta_Neta"] > 0]
            cats = sorted(todos_prod["Categoria"].fillna("").astype(str).replace({"": "Sin categoría"}).unique()) \
                if "Categoria" in todos_prod else []
            p1, p2 = st.columns([1, 2])
            sel_cat = p1.multiselect("🏷️ Categoría", ["✅ Todas"] + cats, default=["✅ Todas"], key="ret_cat")
            sel_prod = p2.multiselect("📦 Productos (elige uno, dos, tres o los que quieras)",
                                      ["✅ Todos"] + sorted(todos_prod["Producto"].unique()), default=["✅ Todos"],
                                      key="ret_prod")
            p3, p4 = st.columns([2, 2])
            texto_prod = p3.text_input("🧴 O escribe: producto contiene", key="ret_txt",
                                       placeholder="ej: detergente 30  |  suavizante, lejia",
                                       help="Varias palabras = todas. Coma = cualquiera. '-bebe' excluye.")
            criterio = p4.radio("¿Qué cuenta como 'dejó de comprar'?",
                                ["Dejó de comprar ESOS productos (aunque me compre otros)",
                                 "Dejó de comprarme TODO"], key="ret_crit")
        categorias = [c for c in sel_cat if c != "✅ Todas"]
        productos = [p for p in sel_prod if p != "✅ Todos"]
        seguimiento = compras_de_productos(df_v, categorias, productos, texto_prod)
        hay_filtro_prod = not seguimiento.empty or bool(categorias or productos or texto_prod.strip())

        # --- 2. ¿Hace cuánto dejaron de comprar?
        a1, a2, a3 = st.columns(3)
        desde = a1.number_input("Sin comprar hace MÁS de (días)", min_value=0, value=30, step=5)
        limitar = a2.checkbox("Poner un tope", help="Ej.: entre 30 y 90 días, para no incluir clientes perdidos hace años.")
        hasta = a2.number_input("…y MENOS de (días)", min_value=int(desde) + 1, value=max(int(desde) + 60, 90),
                                step=5, disabled=not limitar)
        opciones_orden = ["Lo que compró en total (mayor primero)", "Más días sin comprar primero",
                          "Menos días sin comprar primero", "Más compras realizadas primero", "Nombre A-Z"]
        if hay_filtro_prod:
            opciones_orden = ["Lo que compró de esos productos (mayor primero)",
                              "Más unidades de esos productos primero"] + opciones_orden
        orden = a3.selectbox("Ordenar por", opciones_orden)
        a4, a5 = st.columns(2)
        min_compras = a4.number_input("Que hayan comprado al menos (veces)", min_value=1, value=1,
                                      help="Sube este número para ver solo clientes que eran recurrentes.")
        min_monto = a5.number_input("Que hayan comprado en total al menos (S/)", min_value=0.0, value=0.0, step=100.0)

        base = df_clientes
        if hay_filtro_prod:
            base = df_clientes.merge(seguimiento, on="Cliente_ID", how="inner")
            base["Dias_Prod"] = (pd.Timestamp.today().normalize() - base["Ultima_Prod"]).dt.days.astype("Int64")
            dias_eval = base["Dias_Prod"] if criterio.startswith("Dejó de comprar ESOS") else base["Días Sin Comprar"]
            compras_eval, monto_eval = base["Compras_Prod"], base["Venta_Prod"]
        else:
            dias_eval, compras_eval, monto_eval = base["Días Sin Comprar"], base["Compras"], base["Venta_Neta"]

        filtro = (dias_eval > desde) & (compras_eval >= min_compras) & (monto_eval >= min_monto)
        if limitar:
            filtro &= dias_eval < hasta
        grupo = base[filtro.fillna(False)].copy()
        grupo["_dias"] = dias_eval[filtro.fillna(False)]
        columna, asc = {"Lo que compró de esos productos (mayor primero)": ('Venta_Prod', False),
                        "Más unidades de esos productos primero": ('Unidades_Prod', False),
                        "Lo que compró en total (mayor primero)": ('Venta_Neta', False),
                        "Más días sin comprar primero": ('_dias', False),
                        "Menos días sin comprar primero": ('_dias', True),
                        "Más compras realizadas primero": ('Compras_Prod' if hay_filtro_prod else 'Compras', False),
                        "Nombre A-Z": ('Cliente', True)}[orden]
        grupo = grupo.sort_values(columna, ascending=asc, na_position="last")

        k1, k2, k3 = st.columns(3)
        k1.metric("Clientes a reconquistar", f"{len(grupo):,}")
        if hay_filtro_prod:
            k2.metric("Compraban de esos productos", f"S/ {grupo['Venta_Prod'].sum():,.2f}",
                      help="Con IGV, en todo su historial.")
        else:
            k2.metric("Lo que compraban en total", f"S/ {grupo['Venta_Neta'].sum():,.2f}")
        k3.metric("Promedio de días sin comprar", f"{grupo['_dias'].mean():,.0f}" if len(grupo) else "—")
        if hay_filtro_prod and criterio.startswith("Dejó de comprar ESOS"):
            siguen = (grupo["Días Sin Comprar"] <= desde).sum()
            if siguen:
                st.info(f"💡 {siguen} de estos clientes dejaron esos productos pero te siguen comprando otras cosas: "
                        "son los más fáciles de reconquistar (mira 'Última compra').")

        cols_v = COLS_SEGUIMIENTO if hay_filtro_prod else COLS_REPORTE
        cols_v = [c for c in cols_v if c in grupo.columns]
        if not grupo.empty:
            rango = f"{int(desde)}-{int(hasta)}" if limitar else f"mas_de_{int(desde)}"
            st.download_button("📥 Descargar reporte en Excel", data=_excel(grupo, "Reconquistar", cols_v),
                               file_name=f"clientes_reconquistar_{rango}_dias_{hoy}.xlsx", mime=MIME_XLSX,
                               type="primary", key="dl_libre")
        fmt = dict(formato)
        fmt.update({"Ultima_Prod": st.column_config.DateColumn("Última compra de esos productos", format="DD/MM/YYYY"),
                    "Dias_Prod": st.column_config.NumberColumn("Días sin comprar esos productos"),
                    "Productos_Seguidos": st.column_config.TextColumn("Productos elegidos que compraba", width="large"),
                    "Unidades_Prod": st.column_config.NumberColumn("Unidades de esos productos"),
                    "Venta_Prod": st.column_config.NumberColumn("Compró de esos productos (S/)", format="%.2f"),
                    "Compras_Prod": st.column_config.NumberColumn("Veces que los compró")})
        st.dataframe(grupo[cols_v], use_container_width=True, hide_index=True, column_config=fmt)

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
