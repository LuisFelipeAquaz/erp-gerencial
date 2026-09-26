"""
🔎 Explorador de Ventas

Responde cualquier pregunta sobre las ventas cargadas combinando filtros:
  - ¿Cuántos jabones manzana se vendieron?            -> Producto: "jabon manzana"
  - ¿Cuántos suavizantes en factura y cuántos en boleta? -> Producto: "suavizante" + tabla por Tipo de comprobante
  - ¿Cuánto detergente compró Lavandería Ema?          -> Cliente: "ema" + Producto: "detergente"
"""
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from core.utils import convert_to_excel, filtrar_empresa, cobertura_costos, texto_utilidad
from core.resumen import mostrar_resumen

MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

# Dimensiones por las que se puede agrupar / cruzar
DIMENSIONES = {
    "Producto": "Producto",
    "Categoría": "Categoria",
    "Cliente": "Cliente_Etiqueta",
    "Tipo de comprobante": "Tipo_Comprobante",
    "Vendedor": "Vendedor",
    "Empresa": "Empresa",
    "Mes": "Mes",
    "Semana": "Semana",
    "Día": "Dia",
    "Día de la semana": "Dia_Semana",
    "Tipo de documento": "Tipo_Doc",
    "Placa": "Placa",
    "Moneda": "Moneda",
}

# Medidas que se pueden calcular
MEDIDAS = {
    "Unidades vendidas": ("Cantidad", "sum"),
    "Venta neta (sin IGV)": ("Venta_Neta", "sum"),
    "Total con IGV": ("Total_Linea", "sum"),
    "Utilidad bruta": ("Utilidad_Bruta", "sum"),
    "N° de comprobantes": ("Comprobante_Clave", "nunique"),
    "N° de clientes": ("Cliente_ID", "nunique"),
}

DIAS = ["1-Lun", "2-Mar", "3-Mié", "4-Jue", "5-Vie", "6-Sáb", "7-Dom"]


# ==========================================
# PREPARACIÓN (una sola vez por conjunto de datos)
# ==========================================
def _norm_vec(serie: pd.Series) -> pd.Series:
    """Normaliza texto en bloque: sin tildes, mayúsculas, espacios simples."""
    return (serie.fillna("").astype(str).str.normalize("NFD")
            .str.replace(r"[\u0300-\u036f]", "", regex=True)
            .str.upper().str.replace(r"\s+", " ", regex=True).str.strip())


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col, defecto in [("Producto", "SIN NOMBRE"), ("Categoria", "Sin categoría"), ("Tipo_Comprobante", "Sin dato"),
                         ("Placa", ""), ("Moneda", "PEN"), ("Empresa", ""), ("Codigo", "")]:
        if col not in df.columns:
            df[col] = defecto
        df[col] = df[col].fillna("").astype(str).str.strip().replace({"": defecto, "nan": defecto})
    if "Total_Linea" not in df.columns:  # datos de Quimaroma o antiguos: sin desglose de IGV
        df["Total_Linea"] = df["Venta_Neta"]
    df["Total_Linea"] = pd.to_numeric(df["Total_Linea"], errors="coerce").fillna(0)
    df["Gratuito"] = df["Gratuito"].fillna(False).astype(bool) if "Gratuito" in df.columns else False
    df["Es_NC"] = _norm_vec(df["Tipo_Comprobante"]).str.contains("CREDITO")

    # Comprobante único incluso en datos sin número (se usa fecha + cliente)
    sin_comp = df["Comprobante"].astype(str).str.len() == 0
    df["Comprobante_Clave"] = df["Comprobante"].astype(str)
    df.loc[sin_comp, "Comprobante_Clave"] = (df.loc[sin_comp, "Fecha"].astype(str) + "|" +
                                             df.loc[sin_comp, "Cliente_ID"].astype(str))

    doc = df["Cliente_Doc"].where(df["Tipo_Doc"] != "SIN DOC", "")
    df["Cliente_Etiqueta"] = df["Cliente"] + doc.map(lambda d: f" ({d})" if d else "")

    f = df["Fecha"]
    df["Mes"] = f.dt.strftime("%Y-%m").fillna("Sin fecha")
    df["Semana"] = (f - pd.to_timedelta(f.dt.weekday, unit="D")).dt.strftime("%Y-%m-%d").fillna("Sin fecha")
    df["Dia"] = f.dt.strftime("%Y-%m-%d").fillna("Sin fecha")
    df["Dia_Semana"] = f.dt.weekday.map(lambda d: DIAS[int(d)] if pd.notna(d) else "Sin fecha")

    # Textos normalizados para búsquedas rápidas
    df["_n_prod"] = _norm_vec(df["Producto"] + " " + df["Codigo"] + " " + df["Categoria"])
    df["_n_cli"] = _norm_vec(df["Cliente"] + " " + df["Cliente_Doc"].astype(str) + " " + df["Placa"])
    df["_n_todo"] = _norm_vec(df["_n_prod"] + " " + df["_n_cli"] + " " + df["Tipo_Comprobante"] + " " +
                              df["Vendedor"] + " " + df["Comprobante"].astype(str) + " " + df["Empresa"])
    return df


def _medidas(df: pd.DataFrame) -> list:
    """La utilidad solo se ofrece si hay costos cargados."""
    return [m for m in MEDIDAS if m != "Utilidad bruta" or df["Utilidad_Bruta"].notna().any()]


def _coincide(texto_norm: pd.Series, consulta: str) -> pd.Series:
    """
    Búsqueda flexible:
      "jabon manzana"            -> contiene JABON y MANZANA
      "suavizante, detergente"   -> contiene SUAVIZANTE o DETERGENTE
      "jabon -bebe"              -> contiene JABON pero no BEBE
    """
    consulta = str(consulta or "").strip()
    if not consulta:
        return pd.Series(True, index=texto_norm.index)
    resultado = pd.Series(False, index=texto_norm.index)
    for grupo in consulta.split(","):
        palabras = _norm_vec(pd.Series([grupo])).iloc[0].split()
        if not palabras:
            continue
        cumple = pd.Series(True, index=texto_norm.index)
        for p in palabras:
            if p.startswith("-") and len(p) > 1:
                cumple &= ~texto_norm.str.contains(p[1:], regex=False)
            else:
                cumple &= texto_norm.str.contains(p, regex=False)
        resultado |= cumple
    return resultado


def aplicar_filtros(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    m = pd.Series(True, index=df.index)
    if f.get("desde") is not None:
        m &= df["Fecha"] >= pd.Timestamp(f["desde"])
    if f.get("hasta") is not None:
        m &= df["Fecha"] <= pd.Timestamp(f["hasta"])
    m &= _coincide(df["_n_prod"], f.get("producto"))
    m &= _coincide(df["_n_cli"], f.get("cliente"))
    m &= _coincide(df["_n_todo"], f.get("general"))
    for clave, col in [("clientes", "Cliente_Etiqueta"), ("tipos", "Tipo_Comprobante"), ("categorias", "Categoria"),
                       ("vendedores", "Vendedor"), ("empresas", "Empresa"), ("productos", "Producto"),
                       ("tipos_doc", "Tipo_Doc"), ("monedas", "Moneda")]:
        if f.get(clave):
            m &= df[col].isin(f[clave])
    if not f.get("incluir_nc", True):
        m &= ~df["Es_NC"]
    if not f.get("incluir_gratis", True):
        m &= ~df["Gratuito"]
    return df[m]


def resumir(df: pd.DataFrame, filas: str, columnas: str = None, medida: str = "Unidades vendidas") -> pd.DataFrame:
    col_valor, funcion = MEDIDAS[medida]
    claves = [DIMENSIONES[filas]] + ([DIMENSIONES[columnas]] if columnas else [])
    t = df.groupby(claves)[col_valor].agg(funcion)
    if columnas:
        t = t.unstack(fill_value=0)
        t["TOTAL"] = t.sum(axis=1) if funcion == "sum" else df.groupby(claves[0])[col_valor].agg(funcion)
        t = t.sort_values("TOTAL", ascending=False)
    else:
        t = t.sort_values(ascending=False).to_frame(medida)
    t.index.name = filas
    return t.round(2)


# ==========================================
# PANTALLA
# ==========================================
def render(empresa_activa):
    st.title("🔎 Explorador de Ventas")
    st.caption("Combina filtros para responder cualquier pregunta: qué se vendió, a quién, en qué comprobante y cuándo.")

    base = filtrar_empresa(empresa_activa)
    if base.empty:
        st.warning(f"⚠️ No hay ventas cargadas para {empresa_activa}.")
        return
    df = preparar(base)

    f = _panel_filtros(df)
    res = aplicar_filtros(df, f)

    if res.empty:
        st.info("Ninguna venta cumple estos filtros. Prueba con menos palabras o quita algún filtro.")
        return

    mostrar_resumen(res, clave="explorador", con_grafico=False)
    _indicadores(res)

    with st.expander("📊 Ver gráficos", expanded=False):
        _graficos(res)

    tab_det, tab_din = st.tabs(["📋 Detalle de ventas", "🧮 Tabla dinámica"])
    with tab_det:
        _detalle(res)
    with tab_din:
        _tabla_dinamica(res)


def _panel_filtros(df: pd.DataFrame) -> dict:
    f = {}
    with st.container(border=True):
        st.markdown("##### Filtros")
        c1, c2, c3 = st.columns([2, 2, 2])
        f["producto"] = c1.text_input("🧴 Producto contiene", placeholder="ej: jabon manzana  |  suavizante, detergente",
                                      help="Varias palabras = deben estar todas. Coma = cualquiera de los grupos. "
                                           "'-bebe' excluye esa palabra. No importan tildes ni mayúsculas.")
        f["cliente"] = c2.text_input("👤 Cliente contiene (nombre, RUC/DNI o placa)", placeholder="ej: servilimpio  |  ema")
        f["general"] = c3.text_input("🔍 Buscar en todo", placeholder="ej: N001-00014170, PIÑANGO...")

        fechas = df["Fecha"].dropna()
        d_min = fechas.min().date() if not fechas.empty else date.today()
        d_max = fechas.max().date() if not fechas.empty else date.today()
        c4, c5, c6 = st.columns([2, 2, 2])
        rango = c4.date_input("📅 Periodo", value=(d_min, d_max), min_value=d_min, max_value=d_max, format="DD/MM/YYYY")
        if isinstance(rango, (tuple, list)) and len(rango) == 2:
            f["desde"], f["hasta"] = rango
        f["tipos"] = c5.multiselect("🧾 Tipo de comprobante", sorted(df["Tipo_Comprobante"].unique()))
        f["categorias"] = c6.multiselect("🏷️ Categoría", sorted(df["Categoria"].unique()))

        with st.expander("Más filtros"):
            e1, e2 = st.columns(2)
            f["clientes"] = e1.multiselect("Clientes específicos", sorted(df["Cliente_Etiqueta"].unique()))
            f["productos"] = e2.multiselect("Productos específicos", sorted(df["Producto"].unique()))
            e3, e4, e5, e6 = st.columns(4)
            f["vendedores"] = e3.multiselect("Vendedor", sorted(df["Vendedor"].unique()))
            f["empresas"] = e4.multiselect("Empresa", sorted(df["Empresa"].unique()))
            f["tipos_doc"] = e5.multiselect("Tipo de documento", sorted(df["Tipo_Doc"].unique()))
            f["monedas"] = e6.multiselect("Moneda", sorted(df["Moneda"].unique()))
            t1, t2 = st.columns(2)
            f["incluir_nc"] = t1.toggle("Incluir notas de crédito (devoluciones)", value=True)
            f["incluir_gratis"] = t2.toggle("Incluir bonificaciones gratuitas", value=True)
    return f


def _indicadores(res: pd.DataFrame):
    ventas = res[~res["Es_NC"]]
    k1, k4, k5, k6 = st.columns(4)
    k1.metric("Unidades", f"{res['Cantidad'].sum():,.0f}",
              help="Unidades vendidas menos devoluciones (notas de crédito). Incluye bonificaciones.")
    cob = cobertura_costos(res)
    k4.metric("Utilidad bruta", texto_utilidad(res),
              help="En blanco si no hay costos cargados." if cob == 0 else f"Calculada sobre el {cob:.0%} de la venta que tiene costo.")
    k5.metric("Comprobantes", f"{ventas['Comprobante_Clave'].nunique():,}")
    k6.metric("Clientes", f"{ventas['Cliente_ID'].nunique():,}")
    if res["Moneda"].nunique() > 1:
        st.caption("ℹ️ En gráficos y tabla dinámica, los montos en soles y dólares se suman juntos: "
                   "para separarlos usa 'Más filtros → Moneda'. El resumen de arriba ya los separa.")
    gratis = res.loc[res["Gratuito"], "Cantidad"].sum()
    devol = -res.loc[res["Es_NC"], "Cantidad"].sum()
    notas = []
    if gratis:
        notas.append(f"{gratis:,.0f} unidades fueron bonificación gratuita")
    if devol:
        notas.append(f"{devol:,.0f} unidades se devolvieron con nota de crédito")
    if res["Producto"].nunique() > 1:
        notas.append(f"se suman {res['Producto'].nunique()} productos distintos (con presentaciones distintas: "
                     "revisa el detalle por producto)")
    if notas:
        st.caption("ℹ️ " + "; ".join(notas) + ".")


def _graficos(res: pd.DataFrame):
    medida = st.radio("Medir por:", _medidas(res), horizontal=True, key="g_medida")
    col_valor, funcion = MEDIDAS[medida]
    top_n = st.slider("Cuántos mostrar en los rankings", 5, 30, 10, key="g_top")

    g1, g2 = st.columns(2)
    por_tipo = res.groupby("Tipo_Comprobante")[col_valor].agg(funcion).reset_index()
    fig = px.pie(por_tipo, names="Tipo_Comprobante", values=col_valor, hole=0.45, title=f"{medida} por tipo de comprobante")
    fig.update_traces(textinfo="percent+value")
    g1.plotly_chart(fig, use_container_width=True)

    periodo = "Dia" if res["Dia"].nunique() <= 62 else ("Semana" if res["Semana"].nunique() <= 60 else "Mes")
    evo = res.groupby([periodo, "Tipo_Comprobante"])[col_valor].agg(funcion).reset_index()
    fig = px.bar(evo, x=periodo, y=col_valor, color="Tipo_Comprobante", title=f"{medida} en el tiempo")
    fig.update_layout(legend_title_text="", xaxis_title="")
    g2.plotly_chart(fig, use_container_width=True)

    g3, g4 = st.columns(2)
    for contenedor, dim, titulo in [(g3, "Producto", "Top productos"), (g4, "Cliente_Etiqueta", "Top clientes")]:
        top = res.groupby(dim)[col_valor].agg(funcion).nlargest(top_n).sort_values().reset_index()
        fig = px.bar(top, x=col_valor, y=dim, orientation="h", title=f"{titulo} por {medida.lower()}", text_auto=".3s")
        fig.update_layout(yaxis_title="", height=max(350, 28 * len(top)))
        contenedor.plotly_chart(fig, use_container_width=True)

    if res["Categoria"].nunique() > 1:
        cat = res.groupby(["Categoria", "Producto"])[col_valor].agg(funcion).reset_index()
        cat = cat[cat[col_valor] > 0]
        if not cat.empty:
            fig = px.treemap(cat, path=["Categoria", "Producto"], values=col_valor,
                             title=f"Mapa de {medida.lower()} por categoría y producto (haz clic para entrar)")
            st.plotly_chart(fig, use_container_width=True)


def _tabla_dinamica(res: pd.DataFrame):
    st.caption("Arma tu propio cruce. Ejemplo: Filas = Producto, Columnas = Tipo de comprobante, "
               "Valor = Unidades vendidas → cuántas unidades de cada producto salieron en factura, boleta o nota de venta.")
    d1, d2, d3, d4 = st.columns(4)
    dims = list(DIMENSIONES.keys())
    filas = d1.selectbox("Filas", dims, index=0, key="p_filas")
    opciones_col = ["(ninguna)"] + [d for d in dims if d != filas]
    defecto = opciones_col.index("Tipo de comprobante") if "Tipo de comprobante" in opciones_col else 0
    columnas = d2.selectbox("Columnas", opciones_col, index=defecto, key="p_cols")
    medida = d3.selectbox("Valor", _medidas(res), key="p_medida")
    grafico = d4.selectbox("Gráfico", ["Barras", "Barras apiladas", "Mapa de calor", "Circular", "Líneas"], key="p_graf")
    columnas = None if columnas == "(ninguna)" else columnas

    tabla = resumir(res, filas, columnas, medida)
    st.dataframe(tabla, use_container_width=True, height=min(600, 38 + 35 * len(tabla)))
    st.download_button("📥 Descargar tabla (.xlsx)", data=convert_to_excel(tabla.reset_index(), sheet_name="Dinamica"),
                       file_name=f"dinamica_{filas}_{medida}.xlsx".replace(" ", "_"), mime=MIME_XLSX, key="dl_din")

    with st.expander("📊 Ver gráfico de esta tabla", expanded=False):
        top = st.slider("Filas a graficar", 5, 50, 15, key="p_top")
        datos = tabla.head(top).drop(columns=["TOTAL"], errors="ignore")
        if columnas:
            largo = datos.reset_index().melt(id_vars=filas, var_name=columnas, value_name=medida)
        else:
            largo = datos.reset_index()
        color = columnas or None
        if grafico == "Mapa de calor" and columnas:
            fig = px.imshow(datos, text_auto=".3s", aspect="auto", color_continuous_scale="Teal")
        elif grafico == "Circular":
            base = tabla["TOTAL"] if columnas else tabla[medida]
            fig = px.pie(base.head(top).reset_index(), names=filas, values=base.name, hole=0.4)
        elif grafico == "Líneas":
            fig = px.line(largo, x=filas, y=medida, color=color, markers=True)
        else:
            fig = px.bar(largo, x=filas, y=medida, color=color,
                         barmode="stack" if grafico == "Barras apiladas" else "group", text_auto=".3s")
        fig.update_layout(height=520, legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)


def _detalle(res: pd.DataFrame):
    cols = ["Fecha", "Empresa", "Tipo_Comprobante", "Comprobante", "Cliente", "Cliente_Doc", "Placa", "Vendedor",
            "Codigo", "Producto", "Categoria", "Moneda", "Cantidad", "Precio_Venta", "Venta_Neta", "Total_Linea",
            "Costo_Unitario", "Utilidad_Bruta", "Gratuito"]
    cols = [c for c in cols if c in res.columns]
    tabla = res[cols].sort_values("Fecha", ascending=False)
    st.caption(f"{len(tabla):,} líneas. Haz clic en un encabezado para ordenar.")
    st.dataframe(tabla, use_container_width=True, hide_index=True, height=520,
                 column_config={"Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                "Precio_Venta": st.column_config.NumberColumn("Precio sin IGV", format="%.2f"),
                                "Venta_Neta": st.column_config.NumberColumn("Venta neta", format="%.2f"),
                                "Total_Linea": st.column_config.NumberColumn("Total con IGV", format="%.2f"),
                                "Utilidad_Bruta": st.column_config.NumberColumn("Utilidad", format="%.2f")})
    st.download_button("📥 Descargar detalle filtrado (.xlsx)", data=convert_to_excel(tabla, sheet_name="Ventas"),
                       file_name=f"ventas_filtradas_{date.today():%Y-%m-%d}.xlsx", mime=MIME_XLSX, key="dl_det")
