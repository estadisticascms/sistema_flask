"""Clientes por producto: quién compró cada producto en el período, cuánto, en cuántas facturas y en qué meses.

- Filtros: período, bodegas, proveedor (última compra del catálogo), marcas, producto y cliente.
- Bodega = dónde está HOY el producto (existencia > 0 en esa bodega); las ventas cuentan de cualquier bodega.
- Ventas: ventas_por_producto sin anuladas; venta = Sub-Total 2 (neto de descuento, sin IVA). Las devoluciones no se restan.
- Facturas: facturas distintas (tipo + número) en que aparece el producto para ese cliente.
- Meses: por mes del período, o por mes del año (Ene–Dic, sumando los años) para ver en qué meses compra más.
- Ingresos: entradas por orden de compra (transacciones_inventario ENTRADA / ORDEN COMPRA, como Seguimiento de ingresos).
  Total del período seleccionado en unidades y costo C$; «último ingreso» es la fecha más reciente, aunque sea antes del período.
- Existencia por bodega: cada bodega donde hay existencia, con su cantidad (las filtradas resaltadas, las demás como «otras»).
- Agrupación por «Familia | Subfamilia» (clasificación de productos, tabla clase_prod); lo que no tiene clasificación va al final
  como «Sin clasificar». Dentro de cada grupo se respeta el orden escogido. El segmento queda solo como columna en el Excel."""
import io
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.existencias_bodegas import lista_bodegas

ventas_producto_cliente_bp = Blueprint("ventas_producto_cliente", __name__)
VISTA = "ventas_producto_cliente"
VERDE, NARANJA = "027266", "FE6412"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MAX_PRODUCTOS_PANTALLA = 150
SIN_CLASIFICAR = "Sin clasificar"
POCO_INFORMATIVOS = {"", "Sin definir", "No aplica", "Por clasificar"}
MAX_CLIENTES_PANTALLA = 60
LOTE_CODIGOS = 1000               # códigos por consulta IN (...)
ORDENES = {"venta": "Mayor venta", "unidades": "Más unidades", "clientes": "Más clientes", "existencia": "Mayor existencia", "codigo": "Código"}


# -----------------------------
# Filtros
# -----------------------------
def fecha(valor):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return None


def ultima_venta():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Con el índice por fecha, buscar primero en los últimos 60 días es inmediato; si no hay nada, en toda la tabla
        cursor.execute("SELECT MAX(Fecha) FROM ventas_por_producto WHERE Fecha >= %s", ((date.today() - timedelta(days=60)).isoformat(),))
        v = cursor.fetchone()[0]
        if not v:
            cursor.execute("SELECT MAX(Fecha) FROM ventas_por_producto")
            v = cursor.fetchone()[0]
    finally:
        cursor.close()
        conn.close()
    return date.fromisoformat(str(v)[:10]) if v else date.today()


def leer_filtros(args):
    hasta = fecha(args.get("hasta")) or ultima_venta()
    mes = hasta.month - 11
    desde = fecha(args.get("desde")) or date(hasta.year + (mes - 1) // 12, (mes - 1) % 12 + 1, 1)   # 12 meses completos
    if desde > hasta:
        desde, hasta = hasta, desde
    n_meses = (hasta.year - desde.year) * 12 + hasta.month - desde.month + 1
    meses = args.get("meses") if args.get("meses") in ("periodo", "anio") else ("periodo" if n_meses <= 13 else "anio")
    return {
        "desde": desde, "hasta": hasta, "n_meses": n_meses,
        "bodegas": sorted({b for b in args.getlist("bodega") if b.strip()}),
        "proveedores": sorted({p.strip() for p in args.getlist("proveedor") if p.strip()}),
        "marcas": sorted({m.strip() for m in args.getlist("marca") if m.strip()}),
        "q": (args.get("q") or "").strip(), "cliente": (args.get("cliente") or "").strip(),
        "meses": meses, "medida": "monto" if args.get("medida") == "monto" else "unidades",
        "orden": args.get("orden") if args.get("orden") in ORDENES else "venta",
        "sin_venta": args.get("sin_venta") == "1",
    }


def params_url(f):
    p = [("desde", f["desde"].isoformat()), ("hasta", f["hasta"].isoformat()), ("q", f["q"]), ("cliente", f["cliente"]),
         ("meses", f["meses"]), ("medida", f["medida"]), ("orden", f["orden"])]
    p += [("bodega", b) for b in f["bodegas"]] + [("proveedor", x) for x in f["proveedores"]] + [("marca", m) for m in f["marcas"]]
    if f["sin_venta"]:
        p.append(("sin_venta", "1"))
    return [(k, v) for k, v in p if v]


def descripcion_filtros(f):
    partes = [f"Del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y}"]
    if f["bodegas"]:
        partes.append("Productos que hoy están en: " + ", ".join(f["bodegas"]))
    if f["proveedores"]:
        partes.append("Proveedor: " + ", ".join(f["proveedores"]))
    if f["marcas"]:
        partes.append("Marca: " + ", ".join(f["marcas"]))
    if f["q"]:
        partes.append(f"Producto: «{f['q']}»")
    if f["cliente"]:
        partes.append(f"Cliente: «{f['cliente']}»")
    return " · ".join(partes)


def listas_catalogo():
    """(marcas, proveedores) del catálogo con cuántos productos tienen."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TRIM(Marca), COUNT(*) FROM catalogo_productos WHERE Marca IS NOT NULL AND TRIM(Marca) <> '' GROUP BY 1 ORDER BY 1")
        marcas = cursor.fetchall()
        cursor.execute("SELECT TRIM(`Proveedor Ultima Compra`), COUNT(*) FROM catalogo_productos "
                       "WHERE `Proveedor Ultima Compra` IS NOT NULL AND TRIM(`Proveedor Ultima Compra`) <> '' GROUP BY 1 ORDER BY 1")
        return marcas, cursor.fetchall()
    finally:
        cursor.close()
        conn.close()


# -----------------------------
# Cálculo
# -----------------------------
def claves_meses(f):
    """[(clave, etiqueta)] de las columnas de meses."""
    if f["meses"] == "anio":
        return [(m, MESES[m - 1]) for m in range(1, 13)]
    claves, a, m = [], f["desde"].year, f["desde"].month
    while (a, m) <= (f["hasta"].year, f["hasta"].month):
        claves.append((f"{a}-{m:02d}", f"{MESES[m - 1]} {str(a)[2:]}"))
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return claves


def calcular(f):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        # 1) Productos del catálogo que cumplen los filtros de producto
        where, params = ["`Codigo Producto` IS NOT NULL"], []
        if f["marcas"]:
            where.append(f"TRIM(Marca) IN ({', '.join(['%s'] * len(f['marcas']))})"); params += f["marcas"]
        if f["proveedores"]:
            where.append(f"TRIM(`Proveedor Ultima Compra`) IN ({', '.join(['%s'] * len(f['proveedores']))})"); params += f["proveedores"]
        for palabra in f["q"].split():
            where.append("(`Codigo Producto` LIKE %s OR Nombre LIKE %s OR `Numero Parte` LIKE %s)"); params += [f"%{palabra}%"] * 3
        cursor.execute(f"SELECT `Codigo Producto`, Nombre, `Numero Parte`, Marca, TRIM(`Proveedor Ultima Compra`) FROM catalogo_productos "
                       f"WHERE {' AND '.join(where)}", params)
        catalogo = {r[0]: {"codigo": r[0], "nombre": (r[1] or "").strip(), "parte": (r[2] or "").strip(), "marca": (r[3] or "").strip(),
                           "proveedor": r[4] or ""} for r in cursor.fetchall()}
        cursor.execute("SELECT codigo, segmento, familia, subfamilia FROM clase_prod")
        for cod, seg, fam, sub in cursor.fetchall():
            if cod in catalogo:
                catalogo[cod].update(segmento=(seg or "").strip(), familia=(fam or "").strip(), subfamilia=(sub or "").strip())

        # 2) Existencia actual (todas las bodegas y las filtradas); con bodegas, solo productos que hoy están ahí
        cursor.execute("SELECT Codigo_producto, Bodega, SUM(Cant_Final) FROM existencias_bodega GROUP BY Codigo_producto, Bodega")
        existencia, en_bodegas, detalle_bod = {}, {}, {}
        for cod, bod, cant in cursor.fetchall():
            cant = float(cant or 0)
            existencia[cod] = existencia.get(cod, 0) + cant
            if cant:
                detalle_bod.setdefault(cod, []).append((bod, cant))
            if bod in f["bodegas"]:
                en_bodegas[cod] = en_bodegas.get(cod, 0) + cant
        if f["bodegas"]:
            catalogo = {c: x for c, x in catalogo.items() if en_bodegas.get(c, 0) > 0}

        # 3) Ventas del período. Con filtros de producto se consultan solo esos códigos (índice código + fecha);
        #    sin filtros, una pasada por fecha y se filtra en Python
        sql = ("SELECT Codigo, `Cod. CLiente`, Cliente, `Tipo Factura`, `No. Factura`, Fecha, `Cantidad Vendida`, `Sub-Total 2`, Vendedor "
               "FROM ventas_por_producto WHERE Fecha BETWEEN %s AND %s AND Estatus <> 'Anulado'")
        params = [f["desde"].isoformat(), f["hasta"].isoformat()]
        for palabra in f["cliente"].split():
            sql += " AND Cliente LIKE %s"; params.append(f"%{palabra}%")
        filtrado = bool(f["marcas"] or f["proveedores"] or f["q"] or f["bodegas"])
        if filtrado and not catalogo:
            ventas = []
        elif filtrado and len(catalogo) <= LOTE_CODIGOS * 5:
            ventas, codigos = [], sorted(catalogo)
            for i in range(0, len(codigos), LOTE_CODIGOS):
                lote = codigos[i:i + LOTE_CODIGOS]
                cursor.execute(sql + f" AND Codigo IN ({', '.join(['%s'] * len(lote))})", params + lote)
                ventas += cursor.fetchall()
        else:
            cursor.execute(sql, params)
            ventas = cursor.fetchall()

        # 4) Ingresos por orden de compra: total del período y fecha del último ingreso (de cualquier fecha)
        sql_ing = ("SELECT codigo_producto, SUM(CASE WHEN fecha BETWEEN %s AND %s THEN cantidad ELSE 0 END), "
                   "SUM(CASE WHEN fecha BETWEEN %s AND %s THEN total_mn ELSE 0 END), "
                   "COUNT(DISTINCT CASE WHEN fecha BETWEEN %s AND %s THEN CONCAT(fecha, id_documento_origen) END), MAX(fecha) "
                   "FROM transacciones_inventario WHERE tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' {IN} "
                   "GROUP BY codigo_producto")
        p_ing = [f["desde"], f["hasta"]] * 3
        filas_ing = []
        if filtrado and len(catalogo) <= LOTE_CODIGOS * 5:
            codigos = sorted(catalogo)
            for i in range(0, len(codigos), LOTE_CODIGOS):
                lote = codigos[i:i + LOTE_CODIGOS]
                cursor.execute(sql_ing.replace("{IN}", f"AND codigo_producto IN ({', '.join(['%s'] * len(lote))})"), p_ing + lote)
                filas_ing += cursor.fetchall()
        elif catalogo:
            cursor.execute(sql_ing.replace("{IN}", ""), p_ing)
            filas_ing = cursor.fetchall()
        ingresos = {c: (float(u or 0), float(m or 0), int(d or 0), ult) for c, u, m, d, ult in filas_ing}
    finally:
        cursor.close()
        conn.close()

    por_anio = f["meses"] == "anio"
    productos = {}
    for cod, cod_cli, cliente, tipo, factura, fch, cant, monto, vendedor in ventas:
        if cod not in catalogo:
            continue
        cant, monto = float(cant or 0), float(monto or 0)
        fch = str(fch)[:10]
        mes = int(fch[5:7]) if por_anio else fch[:7]
        fac = f"{tipo}-{int(factura) if factura is not None else ''}"
        p = productos.setdefault(cod, {"unidades": 0.0, "venta": 0.0, "facturas": set(), "meses": {}, "clientes": {}})
        p["unidades"] += cant; p["venta"] += monto; p["facturas"].add(fac)
        pm = p["meses"].setdefault(mes, [0.0, 0.0]); pm[0] += cant; pm[1] += monto
        cli_id = int(cod_cli) if cod_cli is not None else (cliente or "")
        c = p["clientes"].setdefault(cli_id, {"codigo": cli_id if isinstance(cli_id, int) else None, "cliente": (cliente or "").strip(),
                                             "unidades": 0.0, "venta": 0.0, "facturas": set(), "ultima": fch, "primera": fch, "meses": {},
                                             "vendedor": ""})
        c["unidades"] += cant; c["venta"] += monto; c["facturas"].add(fac)
        if fch >= c["ultima"] or not c["vendedor"]:
            c["vendedor"] = vendedor_corto(vendedor) or c["vendedor"]     # el de la compra más reciente
        c["ultima"], c["primera"] = max(c["ultima"], fch), min(c["primera"], fch)
        cm = c["meses"].setdefault(mes, [0.0, 0.0]); cm[0] += cant; cm[1] += monto

    etiquetas = dict(claves_meses(f))
    i = 1 if f["medida"] == "monto" else 0

    def mes_mayor(meses):
        positivos = {k: v[i] for k, v in meses.items() if v[i] > 0}
        if not positivos:
            return None
        k = max(positivos, key=positivos.get)
        return etiquetas.get(k, str(k))

    filas = []
    codigos = set(productos) | (set(catalogo) if f["sin_venta"] else set())
    for cod in codigos:
        x = catalogo[cod]
        p = productos.get(cod, {"unidades": 0.0, "venta": 0.0, "facturas": set(), "meses": {}, "clientes": {}})
        clientes = sorted(p["clientes"].values(), key=lambda c: (-c["venta"], c["cliente"]))
        for c in clientes:
            c["n_facturas"] = len(c["facturas"])
            c["mes_mayor"] = mes_mayor(c["meses"])
            c["ultima"] = date.fromisoformat(c["ultima"])
            c["primera"] = date.fromisoformat(c["primera"])
            c["pct"] = c["venta"] / p["venta"] * 100 if p["venta"] else 0
        ing_u, ing_m, ing_d, ult_ing = ingresos.get(cod, (0.0, 0.0, 0, None))
        filas.append({**x, "ingresos": ing_u, "ingresos_costo": ing_m, "ingresos_docs": ing_d, "ultimo_ingreso": ult_ing,
                      "existencia": existencia.get(cod, 0.0), "en_bodegas": en_bodegas.get(cod, 0.0) if f["bodegas"] else None,
                      "bodegas_txt": " · ".join(f"{b}: {c:,.0f}" for b, c in sorted(detalle_bod.get(cod, []), key=lambda d: -d[1])),
                      "por_bodega": sorted(detalle_bod.get(cod, []), key=lambda d: (-d[1], d[0])),
                      "segmento": x.get("segmento") or SIN_CLASIFICAR, "familia": x.get("familia") or SIN_CLASIFICAR,
                      "subfamilia": x.get("subfamilia") or "",
                      "unidades": p["unidades"], "venta": p["venta"], "n_facturas": len(p["facturas"]), "facturas": p["facturas"], "meses": p["meses"],
                      "mes_mayor": mes_mayor(p["meses"]), "clientes": clientes, "n_clientes": len(clientes)})
    orden = {"venta": lambda r: (-r["venta"], r["codigo"]), "unidades": lambda r: (-r["unidades"], r["codigo"]),
             "clientes": lambda r: (-r["n_clientes"], -r["venta"]), "existencia": lambda r: (-r["existencia"], r["codigo"]),
             "codigo": lambda r: r["codigo"]}[f["orden"]]
    filas.sort(key=orden)
    totales = {"productos": len(filas), "con_venta": sum(1 for r in filas if r["n_clientes"]), "venta": sum(r["venta"] for r in filas),
               "unidades": sum(r["unidades"] for r in filas), "existencia": sum(r["existencia"] for r in filas),
               "ingresos": sum(r["ingresos"] for r in filas), "ingresos_costo": sum(r["ingresos_costo"] for r in filas),
               "facturas": len(set().union(*(productos[c]["facturas"] for c in productos if c in codigos))) if productos else 0,
               "clientes": len({cid for r in filas for cid in productos.get(r["codigo"], {"clientes": {}})["clientes"]})}
    return filas, totales


def etiqueta_familia(familia, subfamilia):
    """«Rodamientos y chumaceras | Rodamientos de bolas»; sin la subfamilia si se repite o no dice nada."""
    if not subfamilia or subfamilia in POCO_INFORMATIVOS or subfamilia == familia:
        return familia
    return f"{familia} | {subfamilia}"


def clave_texto(texto):
    """Orden alfabético con lo que no está clasificado al final."""
    return (texto in POCO_INFORMATIVOS or texto == SIN_CLASIFICAR, texto.upper())


def agrupar(filas):
    """[{etiqueta «Familia | Subfamilia», filas, totales}] respetando el orden de las filas dentro de cada grupo."""
    def totales(lista):
        return {"productos": len(lista), "unidades": sum(r["unidades"] for r in lista), "venta": sum(r["venta"] for r in lista),
                "existencia": sum(r["existencia"] for r in lista), "ingresos": sum(r["ingresos"] for r in lista),
                "clientes": len({c["codigo"] if c["codigo"] is not None else c["cliente"] for r in lista for c in r["clientes"]})}
    familias = {}
    for r in filas:
        familias.setdefault(etiqueta_familia(r["familia"], r["subfamilia"]), []).append(r)
    return [{"etiqueta": e, "filas": familias[e], "totales": totales(familias[e])}
            for e in sorted(familias, key=lambda e: clave_texto(e.split(" | ")[0]) + (e.upper(),))]


def vendedor_corto(nombre):
    """Nombre y primer apellido del vendedor, solo informativo: «Glen G. Mayorga C.» → «Glen Mayorga»,
    «JOSE ADALID PORRAS GARCIA» → «Jose Porras», «Eddy . Chevez .» → «Eddy Chevez», «MOSTRADOR . . .» → «Mostrador»."""
    partes = (nombre or "").split()
    inicial = lambda t: t == "." or (len(t) <= 2 and (t.endswith(".") or len(t) == 1))
    while partes and inicial(partes[0]) and any(not inicial(t) for t in partes[1:]):
        partes.pop(0)                                    # «A. Javier Torres M.»: la inicial del primer nombre no cuenta
    if not partes:
        return ""
    primero, resto = partes[0], partes[1:]
    completos = [t for t in resto if not inicial(t)]
    if not completos:
        apellido = ""
    elif resto and inicial(resto[0]):
        apellido = completos[0]                          # segundo nombre abreviado o vacío: lo que sigue es el apellido
    else:
        apellido = completos[1] if len(completos) >= 2 else completos[0]   # segundo nombre + apellido(s)
    return " ".join(w.capitalize() if w.isupper() or w.islower() else w for w in (primero, apellido) if w)


def intensidad(valor, maximo):
    """Opacidad del color de fondo de la celda del mes (0 a 0.75)."""
    return round(0.08 + 0.67 * valor / maximo, 2) if maximo and valor > 0 else 0


# -----------------------------
# Vistas
# -----------------------------
@ventas_producto_cliente_bp.route("/ventas_producto_cliente")
@requiere_acceso(VISTA)
def reporte():
    f = leer_filtros(request.args)
    marcas, proveedores = listas_catalogo()
    generado = request.args.get("generar") == "1"
    filas, t = calcular(f) if generado else ([], None)
    return render_template("ventas_producto_cliente.html", f=f, filas=filas[:MAX_PRODUCTOS_PANTALLA], t=t, generado=generado,
                           grupos=agrupar(filas[:MAX_PRODUCTOS_PANTALLA]),
                           bodegas=lista_bodegas(), marcas=marcas, proveedores=proveedores, ordenes=ORDENES,
                           meses=claves_meses(f), params=params_url(f), filtros_txt=descripcion_filtros(f),
                           intensidad=intensidad, max_clientes=MAX_CLIENTES_PANTALLA, max_productos=MAX_PRODUCTOS_PANTALLA)


@ventas_producto_cliente_bp.route("/ventas_producto_cliente/excel")
@requiere_acceso(VISTA)
def excel():
    f = leer_filtros(request.args)
    filas, t = calcular(f)
    meses = claves_meses(f)
    i = 1 if f["medida"] == "monto" else 0
    medida = "C$" if i else "unid."
    con_bod = bool(f["bodegas"])
    wb = Workbook()

    def encabezar(ws, fila, columnas):
        for n, (nombre, ancho) in enumerate(columnas, start=1):
            c = ws.cell(row=fila, column=n, value=nombre)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NARANJA if nombre.startswith(("Venta", "Facturas")) else VERDE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(n)].width = ancho
        ws.row_dimensions[fila].height = 32

    def titulo(ws, texto):
        ws["A1"] = texto
        ws["A1"].font = Font(bold=True, size=13, color=VERDE)
        ws["A2"] = (f"{descripcion_filtros(f)} · Venta = sub-total neto sin IVA, sin anuladas · Meses en {medida}"
                    f"{' por mes del año (Ene–Dic, suma de los años)' if f['meses'] == 'anio' else ''} · Generado {datetime.now():%d/%m/%Y %H:%M}")
        ws["A2"].font = Font(size=9, color="52514E")

    base = [("Código", 12), ("Nº parte", 16), ("Descripción", 36), ("Marca", 14), ("Proveedor", 24), ("Existencia actual", 11)]
    base += [("Existencia en bodegas filtradas", 12)] if con_bod else []
    base += [("Segmento", 14), ("Familia", 24), ("Subfamilia", 26)]

    def datos_producto(r):
        return ([r["codigo"], r["parte"], r["nombre"], r["marca"], r["proveedor"], r["existencia"]] + ([r["en_bodegas"]] if con_bod else [])
                + [r["segmento"], r["familia"], r["subfamilia"] or None])

    # Productos en el orden de los grupos (familia | subfamilia) y, dentro, en el orden escogido
    grupos_cls = agrupar(filas)
    filas = [r for fam in grupos_cls for r in fam["filas"]]

    def escribir(hoja, fila, valores, formatos):
        for col, v in enumerate(valores, start=1):
            c = hoja.cell(row=fila, column=col, value=v)
            if isinstance(v, str) and v.startswith("="):
                c.data_type = "s"          # texto del catálogo, no fórmula
            elif col in formatos:
                c.number_format = formatos[col]

    num_mes = "#,##0.00" if i else "#,##0.##"
    k0 = len(base)
    fmt_base = {6: "#,##0.##", 7: "#,##0.##"} if con_bod else {6: "#,##0.##"}
    formatos_cli = {**fmt_base, k0 + 3: "#,##0.##", k0 + 4: "#,##0.00", k0 + 6: "0.0%", k0 + 7: "dd/mm/yyyy", k0 + 8: "dd/mm/yyyy",
                    **{k0 + 10 + j: num_mes for j in range(len(meses))}}
    formatos_prod = {**fmt_base, k0 + 2: "#,##0.##", k0 + 3: "#,##0.00", k0 + 7: "#,##0.##", k0 + 8: "#,##0.00", k0 + 9: "dd/mm/yyyy",
                     **{k0 + 10 + j: num_mes for j in range(len(meses))}}

    # Hoja 1: un renglón por producto y cliente
    ws = wb.active
    ws.title = "Clientes por producto"
    titulo(ws, "Grupo Comersa · Clientes por producto")
    cols = base + [("Vendedor", 16), ("Cliente", 34), ("Unidades", 10), ("Venta C$", 14), ("Facturas", 9), ("% de la venta del producto", 10),
                   ("Primera compra", 11), ("Última compra", 11), ("Mes que más compra", 11)] + [(e, 9) for _, e in meses]
    encabezar(ws, 4, cols)
    n = 5
    for r in filas:
        for c in r["clientes"] or [None]:
            valores = datos_producto(r)
            if c:
                valores += [c["vendedor"] or None, c["cliente"], c["unidades"], round(c["venta"], 2), c["n_facturas"], c["pct"] / 100, c["primera"], c["ultima"],
                            c["mes_mayor"]] + [round(c["meses"][m][i], 2) if m in c["meses"] else None for m, _ in meses]
            else:
                valores += [None, "Sin ventas en el período"]
            escribir(ws, n, valores, formatos_cli)
            n += 1
    ws.freeze_panes = ws.cell(row=5, column=k0 + 3)
    ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{max(n - 1, 4)}"

    # Hoja 2: resumen por producto
    hp = wb.create_sheet("Por producto")
    titulo(hp, "Grupo Comersa · Ventas por producto")
    cols = base + [("Bodegas con existencia", 30), ("Unidades vendidas", 10), ("Venta C$", 14), ("Facturas", 9), ("Clientes", 9),
                   ("Mes que más se vende", 11), ("Ingresos del período (unid.)", 11), ("Costo de los ingresos C$", 14),
                   ("Último ingreso", 11)] + [(e, 9) for _, e in meses]
    encabezar(hp, 4, cols)
    for n, r in enumerate(filas, start=5):
        valores = datos_producto(r) + [r["bodegas_txt"], r["unidades"], round(r["venta"], 2), r["n_facturas"], r["n_clientes"], r["mes_mayor"],
                                       r["ingresos"] or None, round(r["ingresos_costo"], 2) or None, r["ultimo_ingreso"]]
        valores += [round(r["meses"][m][i], 2) if m in r["meses"] else None for m, _ in meses]
        escribir(hp, n, valores, formatos_prod)
    tot = len(filas) + 5
    hp.cell(row=tot, column=1, value=f"Total · {t['productos']:,} productos").font = Font(bold=True)
    for col in [6] + ([7] if con_bod else []) + [k0 + 2, k0 + 3, k0 + 7, k0 + 8]:
        letra = get_column_letter(col)
        c = hp.cell(row=tot, column=col, value=f"=SUBTOTAL(9,{letra}5:{letra}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00" if col in (k0 + 3, k0 + 8) else "#,##0.##"
    hp.freeze_panes = "D5"
    hp.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{max(tot - 1, 4)}"

    # Hoja 3: resumen por número de parte y descripción; debajo de cada uno, sus clientes (filas agrupadas, se pliegan con +/−)
    hr = wb.create_sheet("Resumen por Nº parte")
    titulo(hr, "Grupo Comersa · Resumen por número de parte y sus clientes")
    etiquetas = dict(meses)

    def mes_mayor(datos):
        positivos = {k: v[i] for k, v in datos.items() if v[i] > 0}
        return etiquetas.get(max(positivos, key=positivos.get)) if positivos else None

    def grupos_parte(filas_grupo):
        """Números de parte (con sus clientes) de un grupo de productos."""
        grupos = {}
        for r in filas_grupo:
            agregar_parte(grupos, r)
        return sorted(grupos.values(), key=lambda g: (g["parte"] == "(sin Nº parte)", g["parte"].upper(), g["nombre"].upper()))

    def agregar_parte(grupos, r):
        g = grupos.setdefault((r["parte"].upper() or "(SIN Nº PARTE)", r["nombre"].upper()),
                              {"parte": r["parte"] or "(sin Nº parte)", "nombre": r["nombre"], "codigos": [], "marcas": set(), "ultimo": None,
                               "ingresos": 0.0, "ingresos_costo": 0.0, "existencia": 0.0, "en_bodegas": 0.0, "unidades": 0.0, "venta": 0.0,
                               "facturas": set(), "meses": {}, "clientes": {}, "bodegas": {}})
        for b, cant in r["por_bodega"]:
            g["bodegas"][b] = g["bodegas"].get(b, 0) + cant
        g["codigos"].append(r["codigo"])
        g["marcas"].add(r["marca"])
        if r["ultimo_ingreso"] and (not g["ultimo"] or r["ultimo_ingreso"] > g["ultimo"]):
            g["ultimo"] = r["ultimo_ingreso"]
        for k in ("ingresos", "ingresos_costo", "existencia", "unidades", "venta"):
            g[k] += r[k]
        g["en_bodegas"] += r["en_bodegas"] or 0
        g["facturas"] |= r["facturas"]
        for m, (u, v_) in r["meses"].items():
            gm = g["meses"].setdefault(m, [0.0, 0.0]); gm[0] += u; gm[1] += v_
        for c in r["clientes"]:
            clave_cli = c["codigo"] if c["codigo"] is not None else c["cliente"]
            gc = g["clientes"].setdefault(clave_cli, {"codigo": c["codigo"], "cliente": c["cliente"], "unidades": 0.0, "venta": 0.0,
                                                      "facturas": set(), "meses": {}, "ultima": c["ultima"], "vendedor": c["vendedor"]})
            gc["unidades"] += c["unidades"]; gc["venta"] += c["venta"]; gc["facturas"] |= c["facturas"]
            if c["ultima"] >= gc["ultima"]:
                gc["vendedor"] = c["vendedor"] or gc["vendedor"]
            gc["ultima"] = max(gc["ultima"], c["ultima"])
            for m, (u, v_) in c["meses"].items():
                cm = gc["meses"].setdefault(m, [0.0, 0.0]); cm[0] += u; cm[1] += v_

    # Un solo grupo de columnas para el número de parte (sus totales) y para cada cliente (filas de abajo). Blanco y negro para imprimir.
    cols = [("Nº parte", 16), ("Descripción", 34), ("Códigos", 14), ("Marca", 11), ("Fecha último ingreso", 11),
            ("Total ingresos del período (unid.)", 11), ("Costo de los ingresos C$", 13), ("Existencia actual", 10)]
    cols += [("Existencia bodega seleccionada", 11)] if con_bod else []
    cols += [("Existencia en otras bodegas" if con_bod else "Existencia por bodega", 26)]
    cols += [("Cliente", 34), ("Unidades vendidas", 10), ("Facturas", 9), ("Venta C$", 13), ("Mes que más compra", 10), ("Última compra", 11)]
    k = 10 if con_bod else 9                # última columna de existencia (la de bodegas, en texto)
    c_bod = k
    c_cli, c_uni, c_fac, c_ven, c_mes, c_ult = range(k + 1, k + 7)
    fmts = {5: "dd/mm/yyyy", 6: "#,##0.##", 7: "#,##0.00", 8: "#,##0.##", 9: "#,##0.##", c_uni: "#,##0.##", c_fac: "#,##0",
            c_ven: "#,##0.00", c_ult: "dd/mm/yyyy"}
    if not con_bod:
        fmts[9] = None
    fino, medio = Side(style="thin", color="000000"), Side(style="medium", color="000000")
    punteado = Side(style="hair", color="000000")
    negro = "000000"
    for col, (nombre, ancho) in enumerate(cols, start=1):
        c = hr.cell(row=4, column=col, value=nombre)
        c.font = Font(bold=True, color=negro, size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=medio, bottom=medio, left=fino, right=fino)
        hr.column_dimensions[get_column_letter(col)].width = ancho
    hr.row_dimensions[4].height = 36
    hr["A1"].font = Font(bold=True, size=13, color=negro)
    hr["A2"].font = Font(size=8, color=negro)

    def escribir_bn(fila, valores, negrita, borde_arriba):
        for col, v in enumerate(valores, start=1):
            c = hr.cell(row=fila, column=col, value=v)
            if isinstance(v, str) and v.startswith("="):
                c.data_type = "s"
            elif fmts.get(col):
                c.number_format = fmts[col]
            c.font = Font(bold=negrita, size=9 if negrita else 8, color=negro)
            c.border = Border(top=fino if borde_arriba else punteado, bottom=punteado, left=fino, right=fino)
            c.alignment = Alignment(vertical="top", wrap_text=col in (2, 3, c_bod, c_cli))

    def texto_bodegas(bodegas):
        lista = sorted(((b, c) for b, c in bodegas.items() if c and b not in f["bodegas"]), key=lambda x: (-x[1], x[0]))
        return " · ".join(f"{b}: {c:,.0f}" for b, c in lista) or None

    def fila_grupo(fila, texto, tot, nivel):
        """Fila de familia | subfamilia (nivel 0) con sus totales."""
        valores = [texto] + [None] * (len(cols) - 1)
        valores[5], valores[7] = tot["ingresos"] or None, tot["existencia"]
        valores[c_cli - 1] = f"{tot['productos']:,} producto(s) · {tot['clientes']:,} cliente(s)"
        valores[c_uni - 1], valores[c_ven - 1] = tot["unidades"] or None, round(tot["venta"], 2) or None
        escribir_bn(fila, valores, True, True)
        for col in range(1, len(cols) + 1):
            celda = hr.cell(row=fila, column=col)
            celda.font = Font(bold=True, size=11 if nivel == 0 else 10, color=negro)
            if nivel == 0:
                celda.fill = PatternFill("solid", fgColor="D9D9D9")
                celda.border = Border(top=medio, bottom=fino, left=fino, right=fino)
        hr.cell(row=fila, column=1).alignment = Alignment(vertical="top", wrap_text=False)
        hr.row_dimensions[fila].outlineLevel = nivel

    orden_g = []
    n = 5
    for fam in grupos_cls:
        fila_grupo(n, fam["etiqueta"].upper(), fam["totales"], 0)
        n += 1
        for g in grupos_parte(fam["filas"]):
            orden_g.append(g)
            clientes = sorted(g["clientes"].values(), key=lambda c: (-c["venta"], c["cliente"]))
            escribir_bn(n, [g["parte"], g["nombre"], ", ".join(g["codigos"]), ", ".join(sorted(m for m in g["marcas"] if m)), g["ultimo"],
                            g["ingresos"] or None, round(g["ingresos_costo"], 2) or None, g["existencia"]]
                        + ([g["en_bodegas"]] if con_bod else []) + [texto_bodegas(g["bodegas"])]
                        + [f"{len(clientes)} cliente{'s' if len(clientes) != 1 else ''}" if clientes else "Sin ventas en el período",
                           g["unidades"] or None, len(g["facturas"]) or None, round(g["venta"], 2) or None, mes_mayor(g["meses"]),
                           max((c["ultima"] for c in clientes), default=None)], True, True)
            hr.row_dimensions[n].outlineLevel = 1
            n += 1
            for c in clientes:
                escribir_bn(n, [None] * k + ["   " + c["cliente"] + (f" · {c['vendedor']}" if c["vendedor"] else ""), c["unidades"],
                                             len(c["facturas"]), round(c["venta"], 2), mes_mayor(c["meses"]), c["ultima"]], False, False)
                hr.row_dimensions[n].outlineLevel = 2
                n += 1
    # Totales calculados aquí (solo filas de número de parte): una fórmula sumaría también las filas de clientes
    hr.cell(row=n, column=1, value=f"TOTAL · {len(orden_g):,} números de parte")
    totales_fila = {6: sum(g["ingresos"] for g in orden_g), 7: round(sum(g["ingresos_costo"] for g in orden_g), 2),
                    8: sum(g["existencia"] for g in orden_g), c_uni: sum(g["unidades"] for g in orden_g),
                    c_ven: round(sum(g["venta"] for g in orden_g), 2), c_cli: f"{t['clientes']:,} clientes"}
    if con_bod:
        totales_fila[9] = sum(g["en_bodegas"] for g in orden_g)
    for col, v in totales_fila.items():
        hr.cell(row=n, column=col, value=v)
    for col in range(1, len(cols) + 1):
        c = hr.cell(row=n, column=col)
        c.font = Font(bold=True, size=9, color=negro)
        c.border = Border(top=medio, bottom=medio, left=fino, right=fino)
        if fmts.get(col):
            c.number_format = fmts[col]
    hr.sheet_properties.outlinePr.summaryBelow = False          # el +/− queda en la fila del grupo o del número de parte
    hr.freeze_panes = "C5"
    hr.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{max(n - 1, 4)}"
    # Impresión: horizontal, a lo ancho de la hoja, encabezado repetido, sin colores
    hr.print_title_rows = "4:4"
    hr.page_setup.orientation = "landscape"
    hr.page_setup.paperSize = hr.PAPERSIZE_LETTER
    hr.page_setup.fitToWidth, hr.page_setup.fitToHeight = 1, 0
    hr.sheet_properties.pageSetUpPr.fitToPage = True
    hr.print_options.horizontalCentered = True
    hr.page_setup.blackAndWhite = True
    hr.page_margins.left = hr.page_margins.right = 0.3
    hr.page_margins.top, hr.page_margins.bottom = 0.5, 0.6
    hr.oddFooter.center.text = "Página &P de &N"
    hr.oddFooter.center.size = 8

    # Hoja 4: existencia por bodega, una columna por cada bodega donde hay existencia de estos productos
    he = wb.create_sheet("Existencia por bodega")
    titulo(he, "Grupo Comersa · Existencia actual por bodega")
    todas = sorted({b for r in filas for b, _ in r["por_bodega"]}, key=lambda b: (b not in f["bodegas"], b))
    cols_e = [("Código", 12), ("Nº parte", 16), ("Descripción", 36), ("Marca", 14), ("Segmento", 14), ("Familia", 24), ("Subfamilia", 26),
              ("Existencia total", 11)] + [(b, 11) for b in todas]
    encabezar(he, 4, cols_e)
    for col, b in enumerate(todas, start=9):
        if b in f["bodegas"]:
            he.cell(row=4, column=col).fill = PatternFill("solid", fgColor=NARANJA)     # bodegas filtradas, primero y en naranja
    for n, r in enumerate(filas, start=5):
        cantidades = dict(r["por_bodega"])
        escribir(he, n, [r["codigo"], r["parte"], r["nombre"], r["marca"], r["segmento"], r["familia"], r["subfamilia"] or None, r["existencia"]]
                 + [cantidades.get(b) or None for b in todas], {col: "#,##0.##" for col in range(8, len(cols_e) + 1)})
    tot = len(filas) + 5
    he.cell(row=tot, column=1, value=f"Total · {len(filas):,} productos").font = Font(bold=True)
    for col in range(8, len(cols_e) + 1):
        letra = get_column_letter(col)
        c = he.cell(row=tot, column=col, value=f"=SUBTOTAL(9,{letra}5:{letra}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.##"
    he.freeze_panes = "D5"
    he.auto_filter.ref = f"A4:{get_column_letter(len(cols_e))}{max(tot - 1, 4)}"

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"clientes_por_producto_{f['desde']:%Y%m%d}_{f['hasta']:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
