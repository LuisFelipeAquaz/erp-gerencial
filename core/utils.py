import io
import re
import unicodedata

import pandas as pd
import streamlit as st


# ==========================================
# TEXTO Y FECHAS
# ==========================================
def norm_texto(valor) -> str:
    """Mayúsculas, sin tildes y con espacios simples. Sirve para comparar nombres."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = unicodedata.normalize("NFD", str(valor))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).upper().strip()


def parse_fecha(serie: pd.Series) -> pd.Series:
    """
    Convierte fechas en formato peruano (dd/mm/aaaa), con o sin hora,
    o fechas reales de Excel. Nunca confunde día con mes.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    texto = serie.astype(str).str.strip()
    fecha = pd.to_datetime(texto.str.slice(0, 10), format="%d/%m/%Y", errors="coerce")
    faltan = fecha.isna() & serie.notna()
    if faltan.any():
        # Fechas guardadas como fecha real de Excel u otros formatos ISO (aaaa-mm-dd)
        fecha[faltan] = pd.to_datetime(serie[faltan], errors="coerce")
    return fecha.dt.normalize()


def a_numero(serie: pd.Series) -> pd.Series:
    """Convierte a número; '-', vacíos y textos quedan en 0."""
    return pd.to_numeric(serie, errors="coerce").fillna(0.0)


# ==========================================
# CLIENTES
# ==========================================
def limpiar_doc(valor) -> str:
    """Documento solo con dígitos. Recupera el cero inicial de DNIs que Excel pierde."""
    s = str(valor if valor is not None else "").strip()
    if re.fullmatch(r"\d+\.0", s):  # número leído como decimal (75770151.0)
        s = s[:-2]
    d = re.sub(r"\D", "", s)
    if len(d) == 7:
        d = d.zfill(8)
    return d


def tipo_doc(doc: str) -> str:
    # Relleno para "clientes varios" (88888888, 00000000...) o códigos internos (00000293) no identifican a nadie
    if not doc or len(set(doc)) == 1 or doc.startswith("0000"):
        return "SIN DOC"
    if len(doc) == 11:
        return "RUC"
    if len(doc) == 8:
        return "DNI"
    return "SIN DOC"


def id_cliente(doc: str, nombre: str) -> str:
    """Clave única del cliente: su RUC/DNI si es válido, si no, su nombre normalizado."""
    return doc if tipo_doc(doc) != "SIN DOC" else "NOM:" + norm_texto(nombre)


# ==========================================
# VENTAS
# ==========================================
def preparar_ventas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve una COPIA de ventas con las columnas que usan los módulos
    (Venta_Neta, Costo_Total, Utilidad_Bruta, Cliente_ID...).
    Funciona con datos nuevos y con los cargados con la versión anterior.
    """
    df = df.copy()
    if df.empty:
        return df

    for col, defecto in [("Cantidad", 0.0), ("Precio_Venta", 0.0), ("Descuento", 0.0), ("Costo_Unitario", 0.0)]:
        df[col] = a_numero(df[col]) if col in df.columns else defecto

    if "Venta_Neta" in df.columns:
        df["Venta_Neta"] = a_numero(df["Venta_Neta"])
    else:  # datos antiguos
        df["Venta_Neta"] = df["Cantidad"] * df["Precio_Venta"] - df["Descuento"]

    # Sin costo conocido => costo y utilidad EN BLANCO (no se inventa una utilidad = venta)
    if "Sin_Costo" in df.columns:
        sin_costo = df["Sin_Costo"].astype("boolean").fillna(False).astype(bool)
    else:  # datos antiguos: costo 0 significa que no se conocía
        sin_costo = df["Costo_Unitario"] == 0
    df["Sin_Costo"] = sin_costo
    df.loc[sin_costo, "Costo_Unitario"] = float("nan")
    df["Costo_Total"] = df["Cantidad"] * df["Costo_Unitario"]
    df["Utilidad_Bruta"] = df["Venta_Neta"] - df["Costo_Total"]

    for col, defecto in [("Vendedor", "Sin Vendedor"), ("Zona", "No registrada"), ("Cliente", "CLIENTES VARIOS")]:
        if col not in df.columns:
            df[col] = defecto
        df[col] = df[col].fillna(defecto).astype(str).str.strip()

    if "Cliente_Doc" not in df.columns:
        df["Cliente_Doc"] = ""
    df["Cliente_Doc"] = df["Cliente_Doc"].fillna("").astype(str)
    df["Tipo_Doc"] = df["Cliente_Doc"].map(tipo_doc)
    df["Cliente_ID"] = [id_cliente(d, n) for d, n in zip(df["Cliente_Doc"], df["Cliente"])]

    if "Comprobante" not in df.columns:
        df["Comprobante"] = ""
    df["Comprobante"] = df["Comprobante"].fillna("").astype(str)
    if "Cuenta" not in df.columns:
        df["Cuenta"] = ""
    df["Cuenta"] = df["Cuenta"].fillna("").astype(str)
    if df.loc[df["Comprobante"] != "", "Cuenta"].nunique() > 1:
        # Dos cuentas de FACEL pueden tener el mismo número (ej. N001-00001677): se identifican por cuenta
        tiene = (df["Comprobante"] != "") & (df["Cuenta"] != "")
        df.loc[tiene, "Comprobante"] = df.loc[tiene, "Cuenta"] + " · " + df.loc[tiene, "Comprobante"]
    if "Moneda" not in df.columns:
        df["Moneda"] = "PEN"
    df["Moneda"] = df["Moneda"].fillna("PEN").astype(str).replace({"": "PEN"})
    df["Total_Linea"] = a_numero(df["Total_Linea"]) if "Total_Linea" in df.columns else df["Venta_Neta"]
    if "Tipo_Comprobante" not in df.columns:
        df["Tipo_Comprobante"] = ""
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    return df


def cobertura_costos(df: pd.DataFrame) -> float:
    """Porcentaje (0 a 1) de la venta neta que tiene costo conocido."""
    if df.empty or "Utilidad_Bruta" not in df.columns:
        return 0.0
    ventas = df.loc[df["Venta_Neta"] > 0]
    total = ventas["Venta_Neta"].sum()
    return float(ventas.loc[ventas["Utilidad_Bruta"].notna(), "Venta_Neta"].sum() / total) if total else 0.0


def texto_utilidad(df: pd.DataFrame) -> str:
    """Utilidad formateada, o '—' si no hay costos."""
    if df["Utilidad_Bruta"].notna().any():
        return f"S/ {df['Utilidad_Bruta'].sum():,.2f}"
    return "—"


def filtrar_empresa(empresa_activa: str, aplicar_periodo: bool = True) -> pd.DataFrame:
    """
    Ventas de la empresa elegida en la barra lateral (siempre una copia).
    Si aplicar_periodo=True, se limita al mes elegido en "📅 PERIODO A REVISAR".
    """
    df_aquaz = st.session_state.dfs.get("Ventas", pd.DataFrame())
    df_quima = st.session_state.dfs.get("Ventas_Quima", pd.DataFrame())
    if "Aquaz" in empresa_activa:
        df = df_aquaz
    elif "Quimaroma" in empresa_activa:
        df = df_quima
    else:
        df = pd.concat([df_aquaz, df_quima], ignore_index=True)
    df = preparar_ventas(df)
    periodo = st.session_state.get("periodo", "TODO")
    if aplicar_periodo and periodo != "TODO" and not df.empty and "Fecha" in df.columns:
        df = df[df["Fecha"].dt.strftime("%Y-%m") == periodo]
    return df


def cruzar_costos(df: pd.DataFrame, df_costos: pd.DataFrame) -> pd.Series:
    """
    Busca el costo de cada producto en el Maestro de Costos comparando nombres
    normalizados (sin importar mayúsculas, tildes o espacios de más).
    Devuelve una serie con el costo o NaN si no lo encuentra.
    """
    if df_costos is None or df_costos.empty or "Producto" not in df_costos.columns or "Costo_Real" not in df_costos.columns:
        return pd.Series(float("nan"), index=df.index)
    maestro = df_costos[["Producto", "Costo_Real"]].copy()
    maestro["_clave"] = maestro["Producto"].map(norm_texto)
    maestro["Costo_Real"] = pd.to_numeric(maestro["Costo_Real"], errors="coerce")
    maestro = maestro.dropna(subset=["Costo_Real"]).drop_duplicates("_clave", keep="last")
    mapa = dict(zip(maestro["_clave"], maestro["Costo_Real"]))
    return df["Producto"].map(norm_texto).map(mapa)


# ==========================================
# FORMATO Y EXPORTACIÓN
# ==========================================
def resaltar_stock_critico(fila):
    col_evaluar = 'Stock' if 'Stock' in fila else ('Cantidad' if 'Cantidad' in fila else None)
    if col_evaluar:
        valor = pd.to_numeric(fila[col_evaluar], errors='coerce')
        if pd.notna(valor) and valor < 0:
            return ['background-color: #fee2e2; color: #991b1b; font-weight: bold'] * len(fila)
    return [''] * len(fila)


TITULOS_EXCEL = {
    "Dinamica": "Tabla dinámica de ventas",
    "Ventas": "Detalle de ventas",
    "Productos": "Productos vendidos",
    "Vendedores": "Ranking de vendedores",
    "Producto x Vendedor": "Unidades por producto y vendedor",
    "Producto x Zona": "Unidades por producto y zona",
    "Resumen mensual": "Resumen de ventas por mes",
    "Reconquistar": "Clientes a reconquistar",
    "Retencion": "Retención de clientes",
    "Top Clientes": "Top clientes",
}


def convert_to_excel(df_export, sheet_name='Datos', titulo=None, periodo=None):
    """
    Excel con encabezado: título del reporte, empresa, periodo, fecha de generación y
    encabezados de columna con formato. Se usa en todos los botones de descarga del ERP.
    """
    empresa = str(st.session_state.get("empresa_activa", "") or "")
    periodo = periodo or str(st.session_state.get("periodo", "TODO") or "TODO")
    titulo = titulo or TITULOS_EXCEL.get(sheet_name, sheet_name)
    return _excel_bytes(df_export, sheet_name, titulo, empresa, periodo,
                        pd.Timestamp.now(tz="America/Lima").strftime("%d/%m/%Y %H:%M"))


def _texto_periodo(periodo: str) -> str:
    if periodo in ("", "TODO"):
        return "Todo el periodo"
    if not re.fullmatch(r"\d{4}-\d{2}", periodo):
        return periodo
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
             "Octubre", "Noviembre", "Diciembre"]
    try:
        anio, mes = periodo.split("-")
        return f"{meses[int(mes) - 1]} {anio}"
    except (ValueError, IndexError):
        return periodo


@st.cache_data(show_spinner=False)
def _excel_bytes(df_export, sheet_name, titulo, empresa, periodo, generado):
    output = io.BytesIO()
    df = df_export.copy()
    # Índices con nombre (ej. 'Producto' en tablas cruzadas) se vuelven columna; los numéricos se descartan
    df = df.reset_index() if any(n is not None for n in df.index.names) else df.reset_index(drop=True)
    df.columns = [str(c) for c in df.columns]
    hoja_nombre = re.sub(r"[\[\]\:\*\?\/\\\\]", " ", str(sheet_name))[:31] or "Datos"
    fila_tabla = 5  # filas 0-3: encabezado del reporte; fila 5: nombres de columna

    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format="dd/mm/yyyy",
                        date_format="dd/mm/yyyy") as writer:
        df.to_excel(writer, index=False, sheet_name=hoja_nombre, startrow=fila_tabla)
        libro = writer.book
        hoja = writer.sheets[hoja_nombre]
        n_cols = max(len(df.columns), 1)

        f_titulo = libro.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E5A"})
        f_info = libro.add_format({"font_size": 10, "font_color": "#44546A"})
        f_encabezado = libro.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E5A",
                                         "border": 1, "border_color": "#BFBFBF", "text_wrap": True,
                                         "valign": "vcenter", "align": "center"})
        f_moneda = libro.add_format({"num_format": "#,##0.00"})
        f_entero = libro.add_format({"num_format": "#,##0"})

        hoja.write(0, 0, titulo.upper(), f_titulo)
        hoja.write(1, 0, f"Empresa: {empresa or '—'}      Periodo: {_texto_periodo(periodo)}", f_info)
        hoja.write(2, 0, f"Generado: {generado}      Registros: {len(df):,}", f_info)
        hoja.write(3, 0, "ERP Gerencial · Holding Aquaz / Quimaroma", f_info)

        for i, col in enumerate(df.columns):
            hoja.write(fila_tabla, i, col, f_encabezado)
            serie = df[col]
            # Ancho según el contenido; columnas vacías no rompen el cálculo
            largos = [len(str(x)) for x in serie if not (x is None or (isinstance(x, float) and pd.isna(x)))]
            ancho = min(max(len(col) * 0.9, max(largos) if largos else 0) + 2, 55)
            if pd.api.types.is_float_dtype(serie):
                hoja.set_column(i, i, max(ancho, 12), f_moneda)
            elif pd.api.types.is_integer_dtype(serie):
                hoja.set_column(i, i, max(ancho, 8), f_entero)
            elif pd.api.types.is_datetime64_any_dtype(serie):
                hoja.set_column(i, i, 12)
            else:
                hoja.set_column(i, i, ancho)
        hoja.set_row(fila_tabla, 32)
        hoja.freeze_panes(fila_tabla + 1, 0)
        if len(df):
            hoja.autofilter(fila_tabla, 0, fila_tabla + len(df), n_cols - 1)
        hoja.set_landscape()
        hoja.fit_to_pages(1, 0)
        hoja.repeat_rows(fila_tabla)
    return output.getvalue()
