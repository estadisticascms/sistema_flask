"""Clientes por producto: quién compró cada producto en el período, cuánto, en cuántas facturas y en qué meses.

- Filtros: período, bodegas, proveedor (última compra del catálogo), marcas, producto y cliente.
- Bodega = dónde está HOY el producto (existencia > 0 en esa bodega); las ventas cuentan de cualquier bodega.
- Ventas: ventas_por_producto sin anuladas; venta = Sub-Total 2 (neto de descuento, sin IVA). Las devoluciones no se restan.
- Facturas: facturas distintas (tipo + número) en que aparece el producto para ese cliente.
- Meses: por mes del período, o por mes del año (Ene–Dic, sumando los años) para ver en qué meses compra más."""
import io
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.existencias_bodegas import lista_bodegas

ventas_producto_cliente_bp = Blueprint("ventas_producto_cliente", __name__)
VISTA = "ventas_producto_cliente"
VERDE, NARANJA = "027266", "FE6412"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MAX_PRODUCTOS_PANTALLA = 150
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
        sql = ("SELECT Codigo, `Cod. CLiente`, Cliente, `Tipo Factura`, `No. Factura`, Fecha, `Cantidad Vendida`, `Sub-Total 2` "
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
    finally:
        cursor.close()
        conn.close()

    por_anio = f["meses"] == "anio"
    productos = {}
    for cod, cod_cli, cliente, tipo, factura, fch, cant, monto in ventas:
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
                                             "unidades": 0.0, "venta": 0.0, "facturas": set(), "ultima": fch, "primera": fch, "meses": {}})
        c["unidades"] += cant; c["venta"] += monto; c["facturas"].add(fac)
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
            c["n_facturas"] = len(c.pop("facturas"))
            c["mes_mayor"] = mes_mayor(c["meses"])
            c["ultima"] = date.fromisoformat(c["ultima"])
            c["primera"] = date.fromisoformat(c["primera"])
            c["pct"] = c["venta"] / p["venta"] * 100 if p["venta"] else 0
        filas.append({**x, "existencia": existencia.get(cod, 0.0), "en_bodegas": en_bodegas.get(cod, 0.0) if f["bodegas"] else None,
                      "bodegas_txt": " · ".join(f"{b}: {c:,.0f}" for b, c in sorted(detalle_bod.get(cod, []), key=lambda d: -d[1])),
                      "unidades": p["unidades"], "venta": p["venta"], "n_facturas": len(p["facturas"]), "meses": p["meses"],
                      "mes_mayor": mes_mayor(p["meses"]), "clientes": clientes, "n_clientes": len(clientes)})
    orden = {"venta": lambda r: (-r["venta"], r["codigo"]), "unidades": lambda r: (-r["unidades"], r["codigo"]),
             "clientes": lambda r: (-r["n_clientes"], -r["venta"]), "existencia": lambda r: (-r["existencia"], r["codigo"]),
             "codigo": lambda r: r["codigo"]}[f["orden"]]
    filas.sort(key=orden)
    totales = {"productos": len(filas), "con_venta": sum(1 for r in filas if r["n_clientes"]), "venta": sum(r["venta"] for r in filas),
               "unidades": sum(r["unidades"] for r in filas), "existencia": sum(r["existencia"] for r in filas),
               "facturas": len(set().union(*(productos[c]["facturas"] for c in productos if c in codigos))) if productos else 0,
               "clientes": len({cid for r in filas for cid in productos.get(r["codigo"], {"clientes": {}})["clientes"]})}
    return filas, totales


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

    def datos_producto(r):
        return [r["codigo"], r["parte"], r["nombre"], r["marca"], r["proveedor"], r["existencia"]] + ([r["en_bodegas"]] if con_bod else [])

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
    formatos_prod = {**fmt_base, k0 + 2: "#,##0.##", k0 + 3: "#,##0.00", **{k0 + 7 + j: num_mes for j in range(len(meses))}}

    # Hoja 1: un renglón por producto y cliente
    ws = wb.active
    ws.title = "Clientes por producto"
    titulo(ws, "Grupo Comersa · Clientes por producto")
    cols = base + [("Cód. cliente", 10), ("Cliente", 34), ("Unidades", 10), ("Venta C$", 14), ("Facturas", 9), ("% de la venta del producto", 10),
                   ("Primera compra", 11), ("Última compra", 11), ("Mes que más compra", 11)] + [(e, 9) for _, e in meses]
    encabezar(ws, 4, cols)
    n = 5
    for r in filas:
        for c in r["clientes"] or [None]:
            valores = datos_producto(r)
            if c:
                valores += [c["codigo"], c["cliente"], c["unidades"], round(c["venta"], 2), c["n_facturas"], c["pct"] / 100, c["primera"], c["ultima"],
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
                   ("Mes que más se vende", 11)] + [(e, 9) for _, e in meses]
    encabezar(hp, 4, cols)
    for n, r in enumerate(filas, start=5):
        valores = datos_producto(r) + [r["bodegas_txt"], r["unidades"], round(r["venta"], 2), r["n_facturas"], r["n_clientes"], r["mes_mayor"]]
        valores += [round(r["meses"][m][i], 2) if m in r["meses"] else None for m, _ in meses]
        escribir(hp, n, valores, formatos_prod)
    tot = len(filas) + 5
    hp.cell(row=tot, column=1, value=f"Total · {t['productos']:,} productos").font = Font(bold=True)
    for col in [6] + ([7] if con_bod else []) + [k0 + 2, k0 + 3]:
        letra = get_column_letter(col)
        c = hp.cell(row=tot, column=col, value=f"=SUBTOTAL(9,{letra}5:{letra}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00" if col == k0 + 3 else "#,##0.##"
    hp.freeze_panes = "D5"
    hp.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{max(tot - 1, 4)}"

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"clientes_por_producto_{f['desde']:%Y%m%d}_{f['hasta']:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
