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
import re

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
    "quimaroma": "Reporte de Ventas Detallado de Quimaroma",
    "facel_resumido": "Informe de Ventas de FACEL (Aquaz)",
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
        return _procesar_informe(encontrados["facel_resumido"])
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
    # TOTAL con IGV de la línea = VALOR VENTA + IGV. No se usa "TOTAL LINEA" directamente porque en el
    # modelo VENTAS GENERAL de FACEL esa columna puede repetir montos y duplicar/triplicar las ventas.
    total_calc = a_numero(col("VALOR VENTA")) + a_numero(col("IGV"))
    total_facel = a_numero(col("TOTAL LINEA"))
    v["Total_Linea"] = (total_calc.where(total_calc != 0, total_facel) * signo).where(~gratuito, 0.0)
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
    no_cuadra = ((total_facel - total_calc).abs() > 0.05) & (total_calc != 0)
    if no_cuadra.sum() > 0:
        avisos.append(("info", f"En {int(no_cuadra.sum())} líneas la columna TOTAL LINEA de FACEL no coincide con "
                               "VALOR VENTA + IGV; se usó VALOR VENTA + IGV para no inflar las ventas."))
    # Control: la suma de líneas calculada por el ERP vs. el TOTAL COMPROBANTE que informa FACEL
    linea_ok = total_calc.where(total_calc != 0, total_facel)
    tc = b.assign(_tc=a_numero(col("TOTAL COMPROBANTE")).abs(), _tl=linea_ok.abs())
    tc = tc[~gratuito].groupby("_comp").agg(total=("_tc", "first"), lineas=("_tl", "sum"))
    tc = tc[tc["total"] > 0]
    descuadre = tc[(tc["total"] - tc["lineas"]).abs() > 0.05]
    cuadran = len(tc) - len(descuadre)
    if len(tc):
        avisos.append(("success" if descuadre.empty else "info",
                       f"Control de totales: {cuadran:,} de {len(tc):,} comprobantes suman exactamente su TOTAL "
                       f"COMPROBANTE de FACEL (S/ {tc['lineas'].sum():,.2f} vs S/ {tc['total'].sum():,.2f})."))
    if not descuadre.empty:
        avisos.append(("warning", f"{len(descuadre)} comprobantes no cuadran con su total en FACEL (puede faltar alguna "
                                  "línea en el reporte; revísalos allá): "
                                  + ", ".join(f"{c} (total {r.total:,.2f} vs líneas {r.lineas:,.2f})"
                                              for c, r in descuadre.head(5).iterrows())))

    modelo = ("Modelo VENTAS GENERAL (todo en una hoja)" if len(hojas) == 1
              else "Modelo por tipo de comprobante (facturas, boletas, notas...)")
    return {"tipo": "facel_detallado", "nombre": f"{NOMBRES['facel_detallado']} · {modelo}", "datos": v,
            "hojas": [h for h, _ in hojas], "avisos": avisos}


def combinar_por_comprobante(actual: pd.DataFrame, nuevo: pd.DataFrame):
    """
    Une un reporte nuevo de FACEL con lo que ya está guardado, sin borrar lo demás.

    La llave es CUENTA + SERIE-NÚMERO (dos cuentas de FACEL pueden repetir numeración): da igual si el comprobante vino en el modelo VENTAS GENERAL
    o en el modelo por hojas (FACTURAS, BOLETAS...). Si ya existía, se reemplaza por la versión
    nueva; si no, se agrega. Todo lo que no viene en el archivo nuevo se queda como estaba.

    Devuelve (resultado, n_nuevos, n_actualizados).
    """
    def llave(df):
        cuenta = df["Cuenta"].fillna("").astype(str) if "Cuenta" in df.columns else pd.Series("", index=df.index)
        return cuenta + "|" + df["Comprobante"].astype(str)

    comps = set(llave(nuevo).unique())
    if actual is None or actual.empty:
        return nuevo.reset_index(drop=True), len(comps), 0
    llaves = llave(actual)
    ya_estaban = llaves.isin(comps)
    actualizados = llaves[ya_estaban].nunique()
    resultado = pd.concat([actual[~ya_estaban], nuevo], ignore_index=True)
    return resultado, len(comps) - actualizados, actualizados


# ==========================================
# FACEL "INFORME DE VENTAS" (una fila por comprobante)
# ==========================================
_RE_ITEM = re.compile(r"^(.*\S)\s*\((-?[\d.,]+)\)\s*$")


def _items(texto: str, cantidad_total: float):
    """'DETERGENTE ... (3)\\nSUAVIZANTE ... (1)' -> [(producto, cantidad), ...]"""
    partes = [p.strip() for p in str(texto or "").replace("\\n", "\n").split("\n") if p.strip()]
    salida = []
    for p in partes:
        m = _RE_ITEM.match(p)
        if m:
            try:
                salida.append((m.group(1).strip(), float(m.group(2).replace(",", ""))))
                continue
            except ValueError:
                pass
        salida.append((p, None))
    if not salida:
        return [("SIN NOMBRE", cantidad_total or 0.0)]
    if all(q is None for _, q in salida) and len(salida) == 1:
        return [(salida[0][0], cantidad_total or 0.0)]
    return [(n, q if q is not None else 0.0) for n, q in salida]


def _procesar_informe(hojas: list) -> dict:
    b = pd.concat([df for _, df in hojas], ignore_index=True)
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)
    ref = col("REFERENCIA").fillna("").astype(str).str.strip()
    tipo = col("COMPROBANTE").fillna("").astype(str).str.strip()
    validos = (ref != "") & (ref != "-") & (tipo != "") & (tipo.str.lower() != "nan")
    avisos = []
    if (~validos).sum():
        avisos.append(("info", f"{(~validos).sum()} registros sin comprobante emitido (sin serie-número) no se cargaron."))
    b = b[validos].copy()
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)

    total = a_numero(col("TOTAL"))
    es_nc = total < 0  # FACEL muestra las notas de crédito como el comprobante original con monto negativo
    partes = col("REFERENCIA").astype(str).str.strip().str.upper().str.extract(r"^([A-Z0-9]+)-0*(\d+)$")
    clave = partes[0].fillna("") + "-" + partes[1].fillna("").str.zfill(8)
    clave = clave.where(partes[0].notna(), col("REFERENCIA").astype(str).str.strip().str.upper())
    # Solo si la nota de crédito repite el número del comprobante original se le agrega "-NC"
    choca = clave.duplicated(keep=False)
    clave = clave.where(~(es_nc & choca), clave + "-NC")
    fecha = parse_fecha(col("FECHA"))
    tipo_txt = col("COMPROBANTE").astype(str).str.strip().str.upper()
    tipo_txt = tipo_txt.where(~es_nc, "NOTA DE CRÉDITO (" + tipo_txt + ")")
    estado = col("ESTADO PAGO").fillna("").astype(str).str.strip()
    saldo = a_numero(col("BALANCE"))
    cant_col = a_numero(col("CANTIDAD"))
    cliente = col("CLIENTE").fillna("CLIENTES VARIOS").astype(str).str.strip().replace({"": "CLIENTES VARIOS"})

    filas = []
    for i in b.index:
        items = _items(b.at[i, "PRODUCTO (CANTIDAD)"] if "PRODUCTO (CANTIDAD)" in b.columns else "", cant_col[i])
        suma = sum(abs(q) for _, q in items)
        for n_item, (prod, q) in enumerate(items):
            peso = abs(q) / suma if suma else 1 / len(items)
            filas.append((i, n_item, prod, q, total[i] * peso))
    L = pd.DataFrame(filas, columns=["_i", "_n", "Producto", "Cantidad", "Total_Linea"])
    v = pd.DataFrame({
        "Fecha": fecha.reindex(L["_i"]).values,
        "Empresa": "Aquaz",
        "Tipo_Comprobante": tipo_txt.reindex(L["_i"]).values,
        "Comprobante": clave.reindex(L["_i"]).values,
        "Cliente_Doc": "",
        "Cliente": cliente.reindex(L["_i"]).values,
        "Placa": "",
        "Vendedor": "No registrado",
        "Codigo": "",
        "Producto": L["Producto"].values,
        "Categoria": "",
        "Moneda": "PEN",
        "Gratuito": False,
        "Cantidad": L["Cantidad"].values,
        "Total_Linea": L["Total_Linea"].round(4).values,
        "Estado_Pago": estado.reindex(L["_i"]).values,
        # El saldo es del comprobante: se guarda solo en su primera línea para no sumarlo varias veces
        "Saldo_Pendiente": [saldo[i] if n == 0 else 0.0 for i, n in zip(L["_i"], L["_n"])],
        "Zona": "No registrada",
        "Monto_Estimado": True,
        "Cuenta": col("FACTURADOR").fillna("").astype(str).str.strip().str.upper().reindex(L["_i"]).values,
    })
    v["Venta_Neta"] = (v["Total_Linea"] / 1.18).round(4)
    v["IGV"] = v["Total_Linea"] - v["Venta_Neta"]
    v["Precio_Con_IGV"] = (v["Total_Linea"] / v["Cantidad"].where(v["Cantidad"] != 0)).abs().round(4).fillna(0)
    v["Precio_Venta"] = (v["Precio_Con_IGV"] / 1.18).round(4)
    v["Descuento"] = 0.0
    v["Costo_Unitario"] = 0.0
    v["Sin_Costo"] = True

    n_comp = v["Comprobante"].nunique()
    avisos.insert(0, ("info", f"{len(v):,} líneas de {n_comp:,} comprobantes, del "
                              f"{v['Fecha'].min():%d/%m/%Y} al {v['Fecha'].max():%d/%m/%Y}."))
    if es_nc.any():
        avisos.append(("info", f"{int(es_nc.sum())} notas de crédito (montos negativos) restan "
                               f"S/ {-total[es_nc].sum():,.2f} con IGV."))
    avisos.append(("warning", "Este informe trae el TOTAL de cada comprobante, pero no el precio de cada producto: "
                              "el monto por producto se reparte según las cantidades (es aproximado). Los totales por "
                              "comprobante, cliente, día y mes sí son exactos. Tampoco trae RUC/DNI, vendedor ni moneda: "
                              "todo se toma en soles y el IGV se calcula como 18%."))
    dia = fecha.dt.strftime("%Y-%m-%d")
    grupos = pd.DataFrame({"dia": dia, "cli": cliente, "tot": total.round(2), "tipo": tipo_txt})
    grupos = grupos[~es_nc.values]
    g = grupos.groupby(["dia", "cli", "tot"])["tipo"].agg(lambda s: set(s))
    pares = g[g.map(lambda s: any("NOTA DE VENTA" in t for t in s) and any(("BOLETA" in t) or ("FACTURA" in t) for t in s))]
    if len(pares):
        avisos.append(("warning", f"{len(pares)} ventas aparecen dos veces: una nota de venta y una boleta/factura del "
                                  f"mismo cliente, el mismo día y por el mismo monto (S/ {float(pd.Series(pares.index.get_level_values('tot')).sum()):,.2f}). "
                                  "Probablemente son notas de venta convertidas a comprobante. Se cargaron tal cual."))
    deuda = saldo[~es_nc].sum()
    if deuda:
        avisos.append(("info", f"Saldo pendiente de cobro en este informe: S/ {deuda:,.2f}."))
    return {"tipo": "facel_resumido", "nombre": NOMBRES["facel_resumido"], "datos": v,
            "hojas": [h for h, _ in hojas], "avisos": avisos}


# ==========================================
# TIENDA QUIMAROMA ("Reporte de Ventas Detallado")
# ==========================================
MONEDAS = {"SOLES": "PEN", "PEN": "PEN", "DOLARES AMERICANOS": "USD", "DOLARES": "USD", "USD": "USD"}


def _procesar_quimaroma(hojas: list, df_costos: pd.DataFrame) -> dict:
    b = pd.concat([df for _, df in hojas], ignore_index=True)
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)
    texto = lambda c, defecto="": col(c).fillna(defecto).astype(str).str.strip().replace({"nan": defecto, "-": defecto})

    b = b[texto("SERIE") != ""].copy()
    col = lambda c: b[c] if c in b.columns else pd.Series([None] * len(b), index=b.index)
    texto = lambda c, defecto="": col(c).fillna(defecto).astype(str).str.strip().replace({"nan": defecto, "-": defecto})

    tipo_comp = texto("COMPROBANTE")
    es_nc = tipo_comp.map(norm_texto).str.contains("CREDITO")
    cant = a_numero(col("CANTIDAD DE ITEM")).abs()
    valor_u = a_numero(col("VALORUNITARIO"))      # sin IGV (en notas de crédito ya viene negativo)
    precio_u = a_numero(col("PRECIOUNITARIO"))    # con IGV
    correl = texto("CORRELATIVO").str.replace(r"\.0$", "", regex=True).str.replace(r"\D", "", regex=True)

    q = pd.DataFrame(index=b.index)
    q["Fecha"] = parse_fecha(col("FECHAEMISION"))
    q["Empresa"] = "Quimaroma"
    q["Tipo_Comprobante"] = tipo_comp
    q["Comprobante"] = texto("SERIE").str.upper() + "-" + correl.str.zfill(8)
    q["Cliente_Doc"] = col("DOCUMENTORECEPTOR").map(limpiar_doc)
    q["Cliente"] = texto("RAZONSOCIAL", "CLIENTES VARIOS").replace({"": "CLIENTES VARIOS"})
    q["Placa"] = texto("PLACA").str.upper()
    q["Vendedor"] = "MOSTRADOR"
    q["Codigo"] = texto("CODIGOITEM")
    q["Producto"] = texto("DESCRIPCIONITEM", "SIN NOMBRE").replace({"": "SIN NOMBRE"})
    q["Categoria"] = texto("FAMILIA").replace({"SIN FAMILIA": ""})
    q["Moneda"] = texto("MONEDA", "SOLES").map(lambda m: MONEDAS.get(norm_texto(m), norm_texto(m) or "PEN"))
    q["Gratuito"] = False
    q["Cantidad"] = cant.where(~es_nc, -cant)     # devoluciones restan unidades
    q["Precio_Venta"] = valor_u.abs()
    q["Precio_Con_IGV"] = precio_u.abs()
    q["Descuento"] = a_numero(col("DESCUENTOITEM"))
    signo = es_nc.map({True: -1.0, False: 1.0})
    q["Venta_Neta"] = (cant * valor_u.abs() - q["Descuento"]) * signo
    q["Total_Linea"] = cant * precio_u.abs() * signo
    q["IGV"] = q["Total_Linea"] - q["Venta_Neta"]
    q["Zona"] = "Mostrador Tienda"
    costo = cruzar_costos(q, df_costos)
    q["Sin_Costo"] = costo.isna()
    q["Costo_Unitario"] = costo.fillna(0.0)
    q = q[(q["Cantidad"] != 0) | (q["Total_Linea"] != 0)].reset_index(drop=True)

    avisos = []
    n_comp = q["Comprobante"].nunique()
    if q["Fecha"].notna().any():
        avisos.append(("info", f"{len(q)} líneas de {n_comp} comprobantes, del "
                               f"{q['Fecha'].min():%d/%m/%Y} al {q['Fecha'].max():%d/%m/%Y}."))
    if q["Fecha"].isna().any():
        avisos.append(("warning", f"{q['Fecha'].isna().sum()} líneas sin fecha válida."))
    nc = q[q["Tipo_Comprobante"].map(norm_texto).str.contains("CREDITO")]
    if not nc.empty:
        avisos.append(("info", f"{nc['Comprobante'].nunique()} notas de crédito restan de las ventas y las unidades."))
    for moneda, grupo in q.groupby("Moneda"):
        if moneda != "PEN":
            avisos.append(("warning", f"{len(grupo)} líneas en {moneda} ({grupo['Comprobante'].nunique()} comprobantes, "
                                      f"{grupo['Venta_Neta'].sum():,.2f} {moneda} sin IGV). Se guardan en su moneda: "
                                      "en el Explorador filtra por Moneda para no mezclarlas con soles."))
    if q["Sin_Costo"].any():
        avisos.append(("info", f"{q.loc[q['Sin_Costo'], 'Producto'].nunique()} productos sin costo registrado: "
                               "su costo y utilidad quedarán en blanco. Todo lo demás funciona normal."))
    return {"tipo": "quimaroma", "nombre": NOMBRES["quimaroma"], "datos": q,
            "hojas": [h for h, _ in hojas], "avisos": avisos}


# Nombre anterior, por compatibilidad
combinar_facel = combinar_por_comprobante
