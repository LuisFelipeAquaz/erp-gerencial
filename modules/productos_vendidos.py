"""
📦 Productos vendidos (se muestra dentro de "1. Ventas & Analítica")

Lista detallada de todos los productos vendidos, con filtros combinables:
  - ¿Cuántos detergentes se vendieron del 23 al 30?          -> Fechas + Producto "detergente"
  - ¿Qué vendió Juan en la zona norte?                        -> Vendedor + Zona
  - ¿Quién vendió más detergente en septiembre?               -> Producto "detergente" + pestaña "Por vendedor"
  - Tres productos de un vendedor                              -> Productos específicos + Vendedor
Los gráficos quedan ocultos hasta que se pidan.
"""
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from core.utils import convert_to_excel
from modules.explorador_ventas import aplicar_filtros, preparar

MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
K = "pv_"  # prefijo de claves para no chocar con otras secciones


def _dinero_por_moneda(res: pd.DataFrame, indice) -> pd.DataFrame:
    """Venta con y sin IGV por moneda (soles y dólares nunca se suman)."""
    salida = pd.DataFrame(index=res.groupby(indice).size().index)
    for mon, simbolo in [("PEN", "S/"), ("USD", "US$")]:
        g = res[res["Moneda"] == mon]
        if g["Total_Linea"].abs().sum() == 0 and mon == "USD":
            continue
        salida[f"Venta con IGV ({simbolo})"] = g.groupby(indice)["Total_Linea"].sum()
        salida[f"Venta sin IGV ({simbolo})"] = g.groupby(indice)["Venta_Neta"].sum()
    return salida.fillna(0).round(2)


def tabla_productos(res: pd.DataFrame) -> pd.DataFrame:
    ventas = res[res["Venta_Neta"] > 0]
    base = res.groupby("Producto").agg(
        Codigo=("Codigo", "first"),
        Categoria=("Categoria", "first"),
        Unidades=("Cantidad", "sum"),
    )
    base["Comprobantes"] = ventas.groupby("Producto")["Comprobante_Clave"].nunique()
    base["Clientes"] = ventas.groupby("Producto")["Cliente_ID"].nunique()
    base["Vendedores"] = ventas.groupby("Producto")["Vendedor"].nunique()
    base = base.join(_dinero_por_moneda(res, "Producto"))
    base[["Comprobantes", "Clientes", "Vendedores"]] = base[["Comprobantes", "Clientes", "Vendedores"]].fillna(0).astype(int)
    return base.reset_index().sort_values("Unidades", ascending=False)


ORDENES_CLIENTES = {
    "Dejó de comprar hace más tiempo primero (última compra más antigua)": ("Última compra", True),
    "Compró más recientemente primero": ("Última compra", False),
    "Más unidades primero": ("Unidades", False),
    "Más venta primero": ("Venta con IGV (S/)", False),
    "Más veces que compró primero": ("N° compras", False),
    "Nombre A-Z": ("Cliente", True),
}


def tabla_clientes(res: pd.DataFrame, df_todo: pd.DataFrame) -> pd.DataFrame:
    """Quiénes son los clientes que compraron lo filtrado, con su última compra."""
    ventas = res[res["Venta_Neta"] > 0]
    if ventas.empty:
        return pd.DataFrame()
    base = ventas.groupby("Cliente_ID").agg(
        Cliente=("Cliente", lambda x: x.value_counts().index[0]),
        Doc=("Cliente_Doc", "first"),
        Tipo_Doc=("Tipo_Doc", "first"),
        Vendedor=("Vendedor", lambda x: x.value_counts().index[0]),
        Zona=("Zona", "last"),
        Productos=("Producto", lambda x: ", ".join(sorted(x.unique())[:5])),
        Primera=("Fecha", "min"),
        Ultima=("Fecha", "max"),
        Compras=("Comprobante_Clave", "nunique"),
    )
    base["Unidades"] = res.groupby("Cliente_ID")["Cantidad"].sum()
    pen = res[res["Moneda"] == "PEN"]
    base["Venta con IGV (S/)"] = pen.groupby("Cliente_ID")["Total_Linea"].sum().reindex(base.index).fillna(0).round(2)
    hoy = pd.Timestamp.today().normalize()
    base["Días desde esa compra"] = (hoy - base["Ultima"]).dt.days.astype("Int64")
    # Última compra de CUALQUIER producto: para saber si dejó solo este producto o dejó de comprar del todo
    todo = df_todo[df_todo["Venta_Neta"] > 0].groupby("Cliente_ID")["Fecha"].max()
    base["Última compra de cualquier producto"] = todo.reindex(base.index)
    base["Doc"] = base["Doc"].where(base["Tipo_Doc"] != "SIN DOC", "")
    base = base.rename(columns={"Primera": "Primera compra", "Ultima": "Última compra", "Compras": "N° compras",
                                "Tipo_Doc": "Tipo doc", "Doc": "RUC / DNI", "Productos": "Productos que compró"})
    cols = ["Cliente", "Tipo doc", "RUC / DNI", "Vendedor", "Zona", "Unidades", "Venta con IGV (S/)", "N° compras",
            "Primera compra", "Última compra", "Días desde esa compra", "Última compra de cualquier producto",
            "Productos que compró"]
    return base.reset_index(drop=True)[cols]


def tabla_vendedores(res: pd.DataFrame) -> pd.DataFrame:
    ventas = res[res["Venta_Neta"] > 0]
    base = res.groupby("Vendedor").agg(Unidades=("Cantidad", "sum"))
    total = base["Unidades"].sum()
    base["% de unidades"] = (base["Unidades"] / total * 100).round(1) if total else 0
    base["Productos distintos"] = ventas.groupby("Vendedor")["Producto"].nunique()
    base["Comprobantes"] = ventas.groupby("Vendedor")["Comprobante_Clave"].nunique()
    base["Clientes"] = ventas.groupby("Vendedor")["Cliente_ID"].nunique()
    base = base.join(_dinero_por_moneda(res, "Vendedor"))
    base = base.fillna(0).reset_index().sort_values("Unidades", ascending=False)
    base.insert(0, "Puesto", range(1, len(base) + 1))
    return base


def tabla_cruce(res: pd.DataFrame, columna: str) -> pd.DataFrame:
    t = res.pivot_table(index="Producto", columns=columna, values="Cantidad", aggfunc="sum", fill_value=0)
    t["TOTAL"] = t.sum(axis=1)
    t = t.sort_values("TOTAL", ascending=False)
    t.loc["TOTAL"] = t.sum()
    return t.round(2)


TODOS = "✅ Todos"


def _multi(contenedor, etiqueta, opciones, clave):
    """Selector con opción 'Todos' (marcada por defecto). Si eliges otros, se usan solo esos."""
    sel = contenedor.multiselect(etiqueta, [TODOS] + sorted(opciones), default=[TODOS], key=K + clave,
                                 help="Deja '✅ Todos' o elige uno, dos o los que quieras.")
    elegidos = [x for x in sel if x != TODOS]
    return elegidos  # vacío = todos


def _filtros(df: pd.DataFrame) -> dict:
    f = {}
    with st.container(border=True):
        st.markdown("##### 🔎 Combina los filtros que quieras")
        fechas = df["Fecha"].dropna()
        d_min = fechas.min().date() if not fechas.empty else date.today()
        d_max = fechas.max().date() if not fechas.empty else date.today()
        c1, c2, c3 = st.columns([2, 2, 3])
        rango = c1.date_input("📅 Desde - hasta", value=(d_min, d_max), min_value=d_min, max_value=d_max,
                              format="DD/MM/YYYY", key=K + "fechas",
                              help="Elige el primer día y luego el último. Ej.: del 23 al 30.")
        if isinstance(rango, (tuple, list)) and len(rango) == 2:
            f["desde"], f["hasta"] = rango
        elif isinstance(rango, (tuple, list)) and len(rango) == 1:
            f["desde"] = f["hasta"] = rango[0]
        f["producto"] = c2.text_input("🧴 Producto contiene", key=K + "prod",
                                      placeholder="detergente  |  jabon, suavizante",
                                      help="Varias palabras = todas. Coma = cualquiera. '-bebe' excluye. "
                                           "No importan tildes ni mayúsculas.")
        f["productos"] = c3.multiselect("📦 O elige productos exactos (uno o varios)",
                                        sorted(df["Producto"].unique()), key=K + "prods")

        c4, c5, c6, c7 = st.columns(4)
        f["vendedores"] = _multi(c4, "🧑‍💼 Vendedor", df["Vendedor"].unique(), "vend")
        f["zonas"] = _multi(c5, "📍 Zona", df["Zona"].unique(), "zona")
        f["tipos"] = _multi(c6, "🧾 Tipo de comprobante", df["Tipo_Comprobante"].unique(), "tipo")
        f["categorias"] = _multi(c7, "🏷️ Categoría", df["Categoria"].unique(), "cat")
        c8, c9 = st.columns([3, 1])
        f["cliente"] = c8.text_input("👤 Cliente contiene (nombre, RUC/DNI o placa)", key=K + "cli")
        f["incluir_nc"] = c9.toggle("Restar devoluciones", value=True, key=K + "nc",
                                    help="Descuenta las notas de crédito de las unidades y montos.")
    return f


def render(df_base: pd.DataFrame):
    """df_base: ventas ya preparadas por filtrar_empresa (respeta el periodo de la barra lateral)."""
    st.subheader("📦 Productos vendidos")
    st.caption("Lista detallada de todo lo que se vendió. Combina fechas, productos, vendedores y zonas; "
               "los resultados, las tablas y los gráficos se ajustan a tu combinación.")
    if df_base.empty:
        return
    df = preparar(df_base)
    if (df["Zona"].nunique() == 1):
        st.caption(f"ℹ️ Tus reportes no traen zona (todas figuran como '{df['Zona'].iloc[0]}'). "
                   "El filtro de zona funcionará cuando se asigne una zona a cada cliente.")

    f = _filtros(df)
    res = aplicar_filtros(df, f)
    if res.empty:
        st.info("Ninguna venta cumple esta combinación. Quita algún filtro o amplía las fechas.")
        return

    # ---------- Totales de la combinación
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Unidades", f"{res['Cantidad'].sum():,.0f}")
    pen = res[res["Moneda"] == "PEN"]
    k2.metric("Venta con IGV (S/)", f"S/ {pen['Total_Linea'].sum():,.2f}")
    k3.metric("Productos distintos", f"{res['Producto'].nunique():,}")
    k4.metric("Vendedores", f"{res['Vendedor'].nunique():,}")
    k5.metric("Clientes", f"{res.loc[res['Venta_Neta'] > 0, 'Cliente_ID'].nunique():,}")
    usd = res[res["Moneda"] == "USD"]
    if usd["Total_Linea"].abs().sum() > 0:
        st.caption(f"Además hay ventas en dólares: US$ {usd['Total_Linea'].sum():,.2f} con IGV (no se suman a los soles).")
    if res["Producto"].nunique() > 1:
        st.caption("Las unidades suman presentaciones distintas (galón, bidón, litro…). Mira el detalle por producto.")

    # ---------- Tablas
    t_prod = tabla_productos(res)
    t_vend = tabla_vendedores(res)
    formato = {c: st.column_config.NumberColumn(c, format="%.2f") for c in t_prod.columns if "Venta" in c}
    tab1, tab_cli, tab2, tab3, tab4 = st.tabs(["📦 Por producto", "👥 Clientes que compraron",
                                               "🧑‍💼 Por vendedor (ranking)", "🔀 Producto × Vendedor",
                                               "📍 Producto × Zona"])
    with tab_cli:
        t_cli = tabla_clientes(res, df)
        if t_cli.empty:
            st.info("No hay clientes con compras en esta combinación.")
        else:
            o1, o2 = st.columns([3, 1])
            orden = o1.selectbox("Ordenar clientes por", list(ORDENES_CLIENTES.keys()), key=K + "orden")
            col_orden, asc = ORDENES_CLIENTES[orden]
            t_cli = t_cli.sort_values(col_orden, ascending=asc, na_position="last")
            o2.metric("Clientes", f"{len(t_cli):,}")
            st.caption("'Última compra' es la última vez que compró lo que filtraste. 'Última compra de cualquier "
                       "producto' te dice si dejó solo ese producto o dejó de comprarte del todo.")
            fmt = {"Primera compra": st.column_config.DateColumn(format="DD/MM/YYYY"),
                   "Última compra": st.column_config.DateColumn(format="DD/MM/YYYY"),
                   "Última compra de cualquier producto": st.column_config.DateColumn(format="DD/MM/YYYY"),
                   "Venta con IGV (S/)": st.column_config.NumberColumn(format="%.2f"),
                   "Productos que compró": st.column_config.TextColumn(width="large")}
            st.dataframe(t_cli, use_container_width=True, hide_index=True, column_config=fmt,
                         height=min(560, 40 + 35 * len(t_cli)))
            st.download_button("📥 Descargar lista de clientes (.xlsx)", data=convert_to_excel(t_cli, "Clientes", "Clientes que compraron los productos filtrados"),
                               file_name=f"clientes_filtrados_{date.today():%Y-%m-%d}.xlsx", mime=MIME_XLSX,
                               key=K + "dlcli")
    with tab1:
        st.caption(f"{len(t_prod)} productos · ordenados por unidades vendidas")
        st.dataframe(t_prod, use_container_width=True, hide_index=True, column_config=formato,
                     height=min(560, 40 + 35 * len(t_prod)))
        st.download_button("📥 Descargar lista de productos (.xlsx)", data=convert_to_excel(t_prod, "Productos"),
                           file_name=f"productos_vendidos_{date.today():%Y-%m-%d}.xlsx", mime=MIME_XLSX, key=K + "dl1")
    with tab2:
        if len(t_vend):
            lider = t_vend.iloc[0]
            st.success(f"🏆 Quien más vendió con esta combinación: **{lider['Vendedor']}** "
                       f"con {lider['Unidades']:,.0f} unidades ({lider['% de unidades']:.1f}%).")
        st.dataframe(t_vend, use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%.2f") for c in t_vend.columns if "Venta" in c})
        st.download_button("📥 Descargar ranking (.xlsx)", data=convert_to_excel(t_vend, "Vendedores"),
                           file_name=f"ranking_vendedores_{date.today():%Y-%m-%d}.xlsx", mime=MIME_XLSX, key=K + "dl2")
    with tab3:
        st.caption("Unidades de cada producto vendidas por cada vendedor.")
        cruce = tabla_cruce(res, "Vendedor")
        st.dataframe(cruce, use_container_width=True, height=min(560, 40 + 35 * len(cruce)))
        st.download_button("📥 Descargar cruce (.xlsx)", data=convert_to_excel(cruce.reset_index(), "Producto x Vendedor"),
                           file_name="producto_x_vendedor.xlsx", mime=MIME_XLSX, key=K + "dl3")
    with tab4:
        st.caption("Unidades de cada producto vendidas en cada zona.")
        cruce_z = tabla_cruce(res, "Zona")
        st.dataframe(cruce_z, use_container_width=True, height=min(560, 40 + 35 * len(cruce_z)))
        st.download_button("📥 Descargar cruce (.xlsx)", data=convert_to_excel(cruce_z.reset_index(), "Producto x Zona"),
                           file_name="producto_x_zona.xlsx", mime=MIME_XLSX, key=K + "dl4")

    # ---------- Gráficos ocultos
    with st.expander("📊 Ver gráficos de esta combinación", expanded=False):
        g1, g2 = st.columns(2)
        medida = g1.radio("Medir por:", ["Unidades", "Venta con IGV (S/)"], horizontal=True, key=K + "gm")
        top = g2.slider("Cuántos productos mostrar", 5, 40, 15, key=K + "gtop")
        datos = res if medida == "Unidades" else pen
        col = "Cantidad" if medida == "Unidades" else "Total_Linea"

        tp = datos.groupby("Producto")[col].sum().nlargest(top).sort_values().reset_index()
        fig = px.bar(tp, x=col, y="Producto", orientation="h", text_auto=",.0f", title=f"Productos ({medida.lower()})")
        fig.update_layout(height=max(360, 26 * len(tp)), yaxis_title="", xaxis_title=medida)
        st.plotly_chart(fig, use_container_width=True, key=K + "f1")

        tv = datos.groupby("Vendedor")[col].sum().sort_values().reset_index()
        fig = px.bar(tv, x=col, y="Vendedor", orientation="h", text_auto=",.0f", title=f"Vendedores ({medida.lower()})")
        fig.update_layout(height=max(300, 34 * len(tv)), yaxis_title="", xaxis_title=medida)
        st.plotly_chart(fig, use_container_width=True, key=K + "f2")

        ev = datos.groupby([datos["Fecha"].dt.date, "Vendedor"])[col].sum().reset_index()
        ev.columns = ["Día", "Vendedor", medida]
        fig = px.bar(ev, x="Día", y=medida, color="Vendedor", title=f"Día por día ({medida.lower()})")
        fig.update_layout(height=380, legend_title_text="", xaxis_title="")
        st.plotly_chart(fig, use_container_width=True, key=K + "f3")
