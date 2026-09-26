"""
Lector universal de reportes de ventas.

Recibe cualquier Excel que se suba en "Carga de Datos", reconoce qué reporte es
leyendo sus encabezados (no el nombre de la hoja ni la fila donde empiezan) y lo
convierte al formato estándar del ERP.

Reportes reconocidos:
  - FACEL detallado (Aquaz): SERIE, NÚMERO, CLIENTE DOC, ... , ATENDIDO POR
  - Tienda Quimaroma: FechaEmision, RazonSocial, DescripcionItem, ...
  - FACEL resumido: se reconoce para avisar que no sirve (no trae RUC/DNI).
"""
import pandas as pd

from core.utils import norm_texto, parse_fecha, a_numero, limpiar_doc, cruzar_costos

# Encabezados que identifican cada reporte (ya normalizados: sin tildes y en mayúsculas)
FIRMAS = {
    "facel_detallado": {"SERIE", "NUMERO", "CLIENTE DOC", "CLIENTE NOMBRE", "PRODUCTO/SERVICIO", "TOTAL LINEA"},
    "quimaroma": {"FECHAEMISION", "RAZONSOCIAL", "DESCRIPCIONITEM"},
    "facel_resumido": {"FECHA", "COMPROBANTE", "REFERENCIA", "PRODUCTO (CANTIDAD)", "ESTADO PAGO"},
}

NOMBRES = {
    "facel_detallado": "Reporte detallado de FACEL (Aquaz)",
    "quimaroma": "Reporte de ventas de la tienda Quimaroma",
    "facel_resumido": "Reporte resumido de FACEL",
}


def _buscar_encabezado(df_raw: pd.DataFrame, max_filas: int = 15):
    """Devuelve (tipo, fila) del reporte que mejor coincide en esta hoja, o (None, -1)."""
    mejor = (None, -1, 0)
    for i in range(min(len(df_raw), max_filas)):
        celdas = {norm_texto(c) for c in df_raw.iloc[i].tolist()}
        for tipo, firma in FIRMAS.items():
            aciertos = len(firma & celdas)
            if aciertos == len(firma) and aciertos > mejor[2]:
                mejor = (tipo, i, aciertos)
    return mejor[0], mejor[1]


def _hoja_con_encabezado(df_raw: pd.DataFrame, fila: int) -> pd.DataFrame:
    df = df_raw.iloc[fila + 1:].copy()
    df.columns = [norm_texto(c) for c in df_raw.iloc[fila].tolist()]
    df = df.loc[:, [c != "" for c in df.columns]]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    return df.dropna(how="all").reset_index(drop=True)


def leer_reporte(archivo, df_costos: pd.DataFrame = None) -> dict:
    """
    Lee el archivo y devuelve:
      tipo:    'facel_detallado' | 'quimaroma' | 'facel_resumido' | None
      nombre:  descripción legible del reporte
      datos:   DataFrame en formato estándar (vacío si no aplica)
      hojas:   hojas usadas
      avisos:  lista de (nivel, mensaje) con nivel 'info' | 'warning' | 'error'
    """
    hojas = pd.read_excel(archivo, sheet_name=None, header=None, dtype=object)

    encontrados = {}  # tipo -> [(nombre_hoja, df)]
    for nombre_hoja, df_raw in hojas.items():
        tipo, fila = _buscar_encabezado(df_raw)
        if tipo:
            encontrados.setdefault(tipo, []).append((nombre_hoja, _hoja_con_encabezado(df_raw, fila)))

    if "facel_detallado" in encontrados:
        return _procesar_facel(encontrados["facel_detallado"], df_costos)
    if "quimaroma" in encontrados:
        return _procesar_quimaroma(encontrados["quimaroma"], df_costos)
    if "facel_resumido" in encontrados:
        return {
            "tipo": "facel_resumido", "nombre": NOMBRES["facel_resumido"], "datos": pd.DataFrame(),
            "hojas": [h for h, _ in encontrados["facel_resumido"]],
            "avisos": [("error", "Este es el reporte RESUMIDO de FACEL: no trae RUC/DNI del cliente ni "
                                 "el detalle por producto. Exporta el reporte DETALLADO (el que tiene las "
                                 "columnas SERIE, NÚMERO, CLIENTE DOC...) y súbelo aquí.")],
        }
    return {
        "tipo": None, "nombre": "Formato no reconocido", "datos": pd.DataFrame(), "hojas": [],
        "avisos": [("error", "No reconocí este archivo. Hojas encontradas: " + ", ".join(map(str, hojas.keys()))
                             + ". Sube el reporte detallado de FACEL o el reporte de la tienda Quimaroma.")],
    }


# ==========================================
# FACEL DETALLADO (AQUAZ)
# ==========================================
def _procesar_facel(hojas: list, df_costos: pd.DataFrame) -> dict:
    avisos = []
    partes = []
    vistos = set()
    for nombre_hoja, df in hojas:
        df = df[df["SERIE"].notna() & (df["SERIE"].astype(str).str.strip() != "")].copy()
        if df.empty:
            continue
        # Llave SERIE-NÚMERO normalizada: '00014155', 14155 y '14155' son el mismo comprobante
        numero = df["NUMERO"].astype(str).str.replace(r"\.0$", "", regex=True).str.replace(r"\D", "", regex=True)
        df["_comp"] = df["SERIE"].astype(str).str.strip().str.upper() + "-" + numero.str.zfill(8)
        # Si una hoja repite comprobantes de otra (ej. una hoja "VENTAS GENERAL"), no se cuentan dos veces
        repetidos = df["_comp"].isin(vistos)
        if repetidos.any():
            avisos.append(("info", f"La hoja '{nombre_hoja}' repetía {df.loc[repetidos, '_comp'].nunique()} "
                                   "comprobantes de otras hojas; se contaron una sola vez."))
            df = df[~repetidos]
        vistos.update(df["_comp"].unique())
        partes.append(df)

    if not partes:
        return {"tipo": "facel_detallado", "nombre": NOMBRES["facel_detallado"], "datos": pd.DataFrame(),
                "hojas": [h for h, _ in hojas], "avisos": [("warning", "El reporte no tiene ventas.")]}

    b = pd.concat(partes, ignore_index=True)
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)
    texto = lambda c, defecto="": col(c).fillna(defecto).astype(str).str.strip().replace({"-": defecto, "nan": defecto})

    tipo_comp = texto("TIPO COMPROBANTE")
    es_nc = tipo_comp.map(norm_texto).str.contains("CREDITO")
    signo = es_nc.map({True: -1.0, False: 1.0})
    gratuito = texto("TIPO IMPUESTO").map(norm_texto).str.contains("GRATUITO")

    v = pd.DataFrame(index=b.index)
    v["Fecha"] = parse_fecha(col("FECHA EMISION"))
    v["Empresa"] = "Aquaz"
    v["Tipo_Comprobante"] = tipo_comp
    v["Comprobante"] = b["_comp"]
    v["Cliente_Doc"] = col("CLIENTE DOC").map(limpiar_doc)
    v["Cliente"] = texto("CLIENTE NOMBRE", "CLIENTE VARIOS").replace({"": "CLIENTE VARIOS"})
    v["Placa"] = texto("PLACA VEHICULO").str.upper()
    v["Vendedor"] = texto("ATENDIDO POR", "TIENDA").replace({"": "TIENDA"})
    v["Codigo"] = texto("CODIGO")
    v["Producto"] = texto("PRODUCTO/SERVICIO", "SIN NOMBRE").replace({"": "SIN NOMBRE"})
    v["Categoria"] = texto("CATEGORIA")
    v["Moneda"] = texto("MONEDA", "PEN").replace({"": "PEN"})
    v["Gratuito"] = gratuito
    v["Cantidad"] = a_numero(col("CANTIDAD")) * signo
    v["Precio_Venta"] = a_numero(col("PRECIO NETO"))            # sin IGV
    v["Precio_Con_IGV"] = a_numero(col("PRECIO UNITARIO"))
    v["Descuento"] = a_numero(col("DESCUENTO"))
    # VALOR VENTA = venta sin IGV y ya con descuento. Lo gratuito no es ingreso.
    v["Venta_Neta"] = (a_numero(col("VALOR VENTA")) * signo).where(~gratuito, 0.0)
    v["IGV"] = (a_numero(col("IGV")) * signo).where(~gratuito, 0.0)
    v["Total_Linea"] = (a_numero(col("TOTAL LINEA")) * signo).where(~gratuito, 0.0)
    v["Zona"] = "No registrada"

    # Costo: primero el Maestro de Costos, luego el costo de FACEL si es mayor a 0
    costo_maestro = cruzar_costos(v, df_costos)
    costo_facel = a_numero(col("COSTO UNITARIO")).where(lambda s: s > 0)
    v["Costo_Unitario"] = costo_maestro.fillna(costo_facel)
    v["Sin_Costo"] = v["Costo_Unitario"].isna()
    v["Costo_Unitario"] = v["Costo_Unitario"].fillna(0.0)

    v = v[(v["Cantidad"] != 0) | (v["Total_Linea"] != 0)].reset_index(drop=True)

    # ---- Avisos para el usuario
    n_comp = v["Comprobante"].nunique()
    avisos.insert(0, ("info", f"{len(v)} líneas de {n_comp} comprobantes, del "
                              f"{v['Fecha'].min():%d/%m/%Y} al {v['Fecha'].max():%d/%m/%Y}."
                      if v["Fecha"].notna().any() else f"{len(v)} líneas de {n_comp} comprobantes."))
    nc = v[v["Tipo_Comprobante"].map(norm_texto).str.contains("CREDITO")]
    if not nc.empty:
        avisos.append(("info", f"{nc['Comprobante'].nunique()} notas de crédito restan "
                               f"S/ {-nc['Venta_Neta'].sum():,.2f} (sin IGV) de las ventas."))
    if v["Gratuito"].any():
        avisos.append(("info", f"{v['Gratuito'].sum()} líneas gratuitas (bonificaciones): se registra su costo "
                               "pero no cuentan como ingreso."))
    if v["Fecha"].isna().any():
        avisos.append(("warning", f"{v['Fecha'].isna().sum()} líneas sin fecha válida."))
    usd = v[v["Moneda"] != "PEN"]
    if not usd.empty:
        avisos.append(("warning", f"{len(usd)} líneas en otra moneda ({', '.join(usd['Moneda'].unique())}), "
                                  f"comprobantes {', '.join(usd['Comprobante'].unique()[:5])}. Se suman sin convertir."))
    sin_costo = v[v["Sin_Costo"] & (v["Cantidad"] > 0)]
    if not sin_costo.empty:
        avisos.append(("info", f"{sin_costo['Producto'].nunique()} productos sin costo registrado "
                                  f"({len(sin_costo)} líneas): su costo y utilidad quedarán en blanco. Todo lo demás "
                                  "(unidades, ventas, clientes, vendedores) funciona normal."))
    # Comprobantes cuyo total no coincide con la suma de sus líneas (líneas faltantes en el reporte)
    tc = b.assign(_tc=a_numero(col("TOTAL COMPROBANTE")), _tl=a_numero(col("TOTAL LINEA")))
    tc = tc[~gratuito].groupby("_comp").agg(total=("_tc", "first"), lineas=("_tl", "sum"))
    descuadre = tc[(tc["total"] > 0) & ((tc["total"] - tc["lineas"]).abs() > 0.05)]
    if not descuadre.empty:
        avisos.append(("warning", "Comprobantes cuyo total no coincide con sus líneas (revísalos en FACEL): "
                                  + ", ".join(f"{c} (total {r.total:,.2f} vs líneas {r.lineas:,.2f})"
                                              for c, r in descuadre.head(5).iterrows())))

    modelo = ("Modelo VENTAS GENERAL (todo en una hoja)" if len(hojas) == 1
              else "Modelo por tipo de comprobante (facturas, boletas, notas...)")
    return {"tipo": "facel_detallado", "nombre": f"{NOMBRES['facel_detallado']} · {modelo}", "datos": v,
            "hojas": [h for h, _ in hojas], "avisos": avisos}


def combinar_facel(actual: pd.DataFrame, nuevo: pd.DataFrame):
    """
    Une un reporte nuevo de FACEL con lo que ya está guardado, sin borrar lo demás.

    La llave es SERIE-NÚMERO: da igual si el comprobante vino en el modelo VENTAS GENERAL
    o en el modelo por hojas (FACTURAS, BOLETAS...). Si ya existía, se reemplaza por la versión
    nueva; si no, se agrega. Todo lo que no viene en el archivo nuevo se queda como estaba.

    Devuelve (resultado, n_nuevos, n_actualizados).
    """
    comps = set(nuevo["Comprobante"].unique())
    if actual is None or actual.empty:
        return nuevo.reset_index(drop=True), len(comps), 0
    ya_estaban = actual["Comprobante"].isin(comps)
    actualizados = actual.loc[ya_estaban, "Comprobante"].nunique()
    resultado = pd.concat([actual[~ya_estaban], nuevo], ignore_index=True)
    return resultado, len(comps) - actualizados, actualizados


# ==========================================
# TIENDA QUIMAROMA
# ==========================================
def _procesar_quimaroma(hojas: list, df_costos: pd.DataFrame) -> dict:
    b = pd.concat([df for _, df in hojas], ignore_index=True)
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)

    q = pd.DataFrame(index=b.index)
    q["Fecha"] = parse_fecha(col("FECHAEMISION"))
    q["Empresa"] = "Quimaroma"
    q["Cliente"] = col("RAZONSOCIAL").fillna("CLIENTES VARIOS").astype(str).str.strip()
    q["Cliente_Doc"] = ""
    q["Vendedor"] = "MOSTRADOR"
    q["Producto"] = col("DESCRIPCIONITEM").fillna("SIN NOMBRE").astype(str).str.strip()
    q["Cantidad"] = a_numero(col("CANTIDAD DE ITEM"))
    q["Precio_Venta"] = a_numero(col("PRECIOUNITARIO"))
    q["Descuento"] = a_numero(col("DESCUENTOITEM"))
    q["Zona"] = "Mostrador Tienda"
    costo = cruzar_costos(q, df_costos)
    q["Sin_Costo"] = costo.isna()
    q["Costo_Unitario"] = costo.fillna(0.0)
    q = q[q["Cantidad"] > 0].reset_index(drop=True)

    avisos = []
    if q["Fecha"].notna().any():
        avisos.append(("info", f"{len(q)} líneas del {q['Fecha'].min():%d/%m/%Y} al {q['Fecha'].max():%d/%m/%Y}."))
    if q["Fecha"].isna().any():
        avisos.append(("warning", f"{q['Fecha'].isna().sum()} líneas sin fecha válida."))
    if q["Sin_Costo"].any():
        avisos.append(("warning", f"{q.loc[q['Sin_Costo'], 'Producto'].nunique()} productos no están en el "
                                  "Maestro de Costos: su costo y utilidad quedarán en blanco."))
    return {"tipo": "quimaroma", "nombre": NOMBRES["quimaroma"], "datos": q,
            "hojas": [h for h, _ in hojas], "avisos": avisos}
