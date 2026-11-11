"""Análisis de proveedor: de las marcas de un proveedor, qué ingresó desde una fecha, qué se movió (vendió) desde esa fecha,
cuánto hay en inventario y en NHE, la última venta, la temporada de más venta y un comentario (sin venta, lento, agotado…).

- Proveedor: «Proveedor Ultima Compra» del catálogo (desde 2023 el sistema no trae el proveedor en las transacciones).
- Ingresado: entradas por ORDEN COMPRA desde la fecha, valorado a FOB.
- Ventas: unidades y venta neta en C$ (Sub-Total 2 de ventas_por_producto: sin IVA, con descuento, sin anuladas) desde la fecha.
- FOB (ingresado, inventario, NHE): cantidades × «Ultimo Precio FOB» del catálogo (US$), igual que el sugerido de compras; sin FOB no suma.
- Rotación: el estado de cada producto (se vende / lento / agotado / sin venta) y el resumen general por estado.
- Rentabilidad: utilidad bruta y margen de lo vendido (costo de ventas por producto), rotación del inventario (costo de lo vendido
  al año ÷ inventario a último costo) y GMROI (utilidad bruta al año por cada C$ 1 de inventario a costo), con un veredicto.
- Inventario: existencia por bodega; las bodegas NHE (mercadería que no está) se muestran aparte y no cuentan como disponible.
- Temporada: con la venta de los últimos 24 meses, los 3 meses seguidos del año que más venden."""
import io
from collections import defaultdict
from datetime import date, datetime

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.ventas_bodegas import en_cache

analisis_proveedor_bp = Blueprint("analisis_proveedor", __name__)

HINT = "/*+ INDEX(transacciones_inventario idx_transacciones_producto_fecha) */"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MESES_TEMPORADA = 24
MIN_UNIDADES_TEMPORADA = 6          # con menos venta en 24 meses no se habla de temporada
MAX_FILAS_PANTALLA = 1500
VERDE, NARANJA = "027266", "FE6412"
ESTADOS = {"": "Todos", "sin_venta": "Sin venta", "lento": "Lento", "rota": "Se vende bien", "agotado": "Agotado", "nhe": "Con NHE",
           "ingreso": "Ingresaron desde la fecha"}
ORDENES = {"inventario": "Mayor inventario FOB", "ingreso": "Mayor ingreso FOB", "vendido": "Mayor venta neta",
           "sin_venta": "Más tiempo sin venta", "marca": "Marca y descripción"}


def restar_meses(d, n):
    m = d.month - 1 - n
    return date(d.year + m // 12, m % 12 + 1, 1)


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def lista_proveedores():
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT TRIM(`Proveedor Ultima Compra`), COUNT(*) FROM catalogo_productos "
                           "WHERE `Proveedor Ultima Compra` <> '' GROUP BY 1 ORDER BY 1")
            return [(p, n) for p, n in cursor.fetchall() if p]
        finally:
            cursor.close()
            conn.close()
    return en_cache("analisis_proveedor_lista", leer)


def marcas_del_proveedor(proveedor):
    """Marcas del proveedor con cuántos productos tiene y la existencia total."""
    if not proveedor:
        return []
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COALESCE(NULLIF(TRIM(Marca), ''), 'SIN MARCA'), COUNT(*), SUM(COALESCE(`Existencia Actual`, 0)) "
                       "FROM catalogo_productos WHERE TRIM(`Proveedor Ultima Compra`) = %s GROUP BY 1 ORDER BY 2 DESC", (proveedor,))
        return [{"marca": m, "productos": n, "existencia": float(e or 0)} for m, n, e in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()


def leer_filtros(args):
    hoy = date.today()
    proveedor = (args.get("proveedor") or "").strip()
    nombres = {p for p, _ in lista_proveedores()}
    if proveedor and proveedor not in nombres:   # escrito a mano: se busca sin importar mayúsculas
        proveedor = next((p for p in nombres if p.upper() == proveedor.upper()), proveedor)
    return {"proveedor": proveedor, "marcas": sorted({m.strip() for m in args.getlist("marca") if m.strip()}),
            "desde": min(fecha_param(args.get("desde"), restar_meses(hoy, 6)), hoy), "hoy": hoy,
            "q": (args.get("q") or "").strip(), "estado": args.get("estado") if args.get("estado") in ESTADOS else "",
            "orden": args.get("orden") if args.get("orden") in ORDENES else "inventario",
            "existencia": "con" if args.get("existencia") == "con" else ""}


ROTACION = [("rota", "Se vende (rotación buena)", "#027266"), ("lento", "Lenta (más de 12 meses de inventario)", "#D9A441"),
            ("agotado", "Agotado (se vende, sin existencia)", "#4F6D8A"), ("sin_venta", "Sin venta (nula)", "#FE6412")]


def ventas_netas(inicio, desde):
    """Venta por producto y mes desde `inicio`: {código: [(mes, desde_la_fecha, unidades, venta_neta)]}.
    ventas_por_producto no tiene índices: se recorre una vez y se guarda 10 minutos (sirve para cualquier proveedor)."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT Codigo, LEFT(Fecha, 7), Fecha >= %s, SUM(`Cantidad Vendida`), SUM(`Sub-Total 2`),
                       SUM(`Cantidad Vendida` * `Costo Unit.`)
                FROM ventas_por_producto WHERE Fecha >= %s AND Estatus <> 'Anulado'
                GROUP BY Codigo, LEFT(Fecha, 7), Fecha >= %s
            """, (desde.isoformat(), inicio.isoformat(), desde.isoformat()))
            por_codigo = defaultdict(list)
            for cod, mes, desde_fecha, cant, monto, costo in cursor.fetchall():
                por_codigo[cod].append((mes, bool(desde_fecha), float(cant or 0), float(monto or 0), float(costo or 0)))
            return dict(por_codigo)
        finally:
            cursor.close()
            conn.close()
    return en_cache(f"analisis_proveedor_ventas2_{inicio}_{desde}", leer)


# Veredicto de rentabilidad por GMROI (utilidad bruta al año por cada C$ 1 de inventario a costo)
GMROI_BUENO, GMROI_ACEPTABLE = 1.5, 0.8


def rentabilidad(x, dias):
    """Métricas de rentabilidad de un total (proveedor o marca) con las ventas de `dias` días."""
    venta, costo, inv = x["venta_neta"], x["costo_vendido"], x["inventario_costo"]
    utilidad = venta - costo
    anual = 365 / dias if dias else 0
    r = {"utilidad": utilidad, "margen": utilidad / venta * 100 if venta else None,
         "rotacion": costo * anual / inv if inv > 0 else None,               # veces al año que se vende el inventario
         "gmroi": utilidad * anual / inv if inv > 0 else None}
    r["meses_inventario"] = 12 / r["rotacion"] if r["rotacion"] else None
    if venta <= 0:
        r["veredicto"], r["nivel"] = ("Sin ventas en el período: el inventario no está generando utilidad", "malo") if inv > 0 else ("Sin ventas ni inventario", "")
    elif r["gmroi"] is None:
        r["veredicto"], r["nivel"] = "Se vende sin inventario disponible: rentable si se repone", "bueno"
    elif r["gmroi"] >= GMROI_BUENO:
        r["veredicto"], r["nivel"] = f"Rentable: cada C$ 1 en inventario deja C$ {r['gmroi']:.2f} de utilidad al año", "bueno"
    elif r["gmroi"] >= GMROI_ACEPTABLE:
        r["veredicto"], r["nivel"] = f"Aceptable: deja C$ {r['gmroi']:.2f} al año por cada C$ 1 en inventario; mejorar la rotación", "medio"
    else:
        r["veredicto"], r["nivel"] = f"Poco rentable: solo C$ {r['gmroi']:.2f} al año por cada C$ 1 en inventario (inventario lento)", "malo"
    return r


def en_lista(cursor, sql, codigos, extra=()):
    filas, codigos = [], list(codigos)
    for i in range(0, len(codigos), 800):
        bloque = codigos[i:i + 800]
        cursor.execute(sql.replace("{IN}", ", ".join(["%s"] * len(bloque))), tuple(bloque) + tuple(extra))
        filas += cursor.fetchall()
    return filas


def temporada(por_mes_anio):
    """{1..12: unidades} -> (texto «Mar–May», % de la venta del año) con los 3 meses seguidos que más venden."""
    total = sum(por_mes_anio.values())
    if total < MIN_UNIDADES_TEMPORADA:
        return None, None
    mejor, inicio = -1, 1
    for m in range(1, 13):
        suma = sum(por_mes_anio.get((m + k - 1) % 12 + 1, 0) for k in range(3))
        if suma > mejor:
            mejor, inicio = suma, m
    fin = (inicio + 1) % 12 + 1
    return f"{MESES[inicio - 1]}–{MESES[fin - 1]}", mejor / total * 100


def meses_txt(dias):
    meses = dias / 30.44
    return f"{meses:.0f} meses" if meses < 24 else f"{meses / 12:.1f} años"


def calcular(f):
    vacio = {"filas": [], "marcas": [], "total": {}, "meses": [], "temporada": (None, None), "sin_actividad": 0}
    if not f["proveedor"]:
        return vacio
    hoy = f["hoy"]
    inicio_hist = min(f["desde"], restar_meses(hoy, MESES_TEMPORADA - 1))
    conn = get_connection()
    cursor = conn.cursor()
    try:
        where, params = ["TRIM(p.`Proveedor Ultima Compra`) = %s"], [f["proveedor"]]
        if f["marcas"]:
            where.append(f"COALESCE(NULLIF(TRIM(p.Marca), ''), 'SIN MARCA') IN ({', '.join(['%s'] * len(f['marcas']))})")
            params += f["marcas"]
        cursor.execute(f"""
            SELECT /*+ NO_INDEX(c) */ p.`Codigo Producto`, p.Nombre, p.`Numero Parte`, COALESCE(NULLIF(TRIM(p.Marca), ''), 'SIN MARCA'),
                   p.Categoria, p.`Ultimo Precio FOB`, p.`Fecha Ultima Compra`, c.segmento, c.familia, c.subfamilia, p.`Ultimo Costo`
            FROM catalogo_productos p LEFT JOIN clase_prod c ON c.codigo = p.`Codigo Producto`
            WHERE {' AND '.join(where)}
        """, params)
        productos = {r[0]: {"codigo": r[0], "nombre": r[1] or "", "parte": r[2] or "", "marca": r[3], "categoria": r[4] or "",
                            "fob": float(r[5]) if r[5] else None, "ultima_compra": r[6],
                            "segmento": r[7] or "Sin clasificar", "familia": r[8] or "Sin clasificar", "subfamilia": r[9] or "Sin clasificar",
                            "costo": float(r[10] or 0)}
                     for r in cursor.fetchall()}
        if not productos:
            return vacio
        codigos = list(productos)
        # Movimientos por producto y mes (desde la fecha y los últimos 24 meses para la temporada)
        desde_sql = f"'{f['desde'].isoformat()}'"     # fecha (objeto date): segura para ir en el texto de la consulta
        movs = en_lista(cursor, f"""
            SELECT {HINT} codigo_producto, DATE_FORMAT(fecha, '%Y-%m'), fecha >= {desde_sql},
                   SUM(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' THEN cantidad ELSE 0 END),
                   SUM(CASE WHEN tipo_transaccion = 'SALIDA' AND tipo_documento_origen LIKE 'FACTURA%%' THEN cantidad ELSE 0 END),
                   SUM(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'DEVOLUCION' THEN cantidad ELSE 0 END),
                   MAX(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' THEN fecha END)
            FROM transacciones_inventario
            WHERE codigo_producto IN ({{IN}}) AND fecha >= %s
            GROUP BY codigo_producto, DATE_FORMAT(fecha, '%Y-%m'), fecha >= {desde_sql}
        """, codigos, (inicio_hist,))
        ultima_venta = dict(en_lista(cursor, f"SELECT {HINT} codigo_producto, MAX(fecha) FROM transacciones_inventario "
                                             "WHERE codigo_producto IN ({IN}) AND tipo_transaccion = 'SALIDA' "
                                             "AND tipo_documento_origen LIKE 'FACTURA%%' GROUP BY codigo_producto", codigos))
        existencias = {r[0]: (float(r[1] or 0), float(r[2] or 0)) for r in en_lista(cursor,
            "SELECT Codigo_producto, SUM(Cant_Final), SUM(CASE WHEN UPPER(Bodega) LIKE '%NHE%' THEN Cant_Final ELSE 0 END) "
            "FROM existencias_bodega WHERE Codigo_producto IN ({IN}) GROUP BY Codigo_producto", codigos)}
    finally:
        cursor.close()
        conn.close()

    # Ingresos (de las transacciones de inventario)
    ingreso, ultimo_ingreso = defaultdict(float), {}
    for cod, mes, desde_fecha, ing, sal, dev, ult_ing in movs:
        if desde_fecha:
            ingreso[cod] += float(ing or 0)
        if ult_ing and (cod not in ultimo_ingreso or ult_ing > ultimo_ingreso[cod]):
            ultimo_ingreso[cod] = ult_ing
    # Ventas a precio de venta neto (de ventas por producto)
    vendido, venta_neta, costo_vendido = defaultdict(float), defaultdict(float), defaultdict(float)
    venta_mes = defaultdict(lambda: defaultdict(float))     # codigo -> 'AAAA-MM' -> unidades
    monto_mes = defaultdict(lambda: defaultdict(float))     # codigo -> 'AAAA-MM' -> venta neta C$
    todas = ventas_netas(inicio_hist, f["desde"])
    for cod in productos:
        for mes, desde_fecha, cant, monto, costo in todas.get(cod, ()):
            venta_mes[cod][mes] += cant
            monto_mes[cod][mes] += monto
            if desde_fecha:
                vendido[cod] += cant
                venta_neta[cod] += monto
                costo_vendido[cod] += costo

    meses = []
    d = restar_meses(hoy, MESES_TEMPORADA - 1)
    while d <= hoy:
        meses.append(f"{d.year}-{d.month:02d}")
        d = restar_meses(d, -1)
    desde_12 = meses[-12]

    filas, sin_actividad = [], 0
    for cod, p in productos.items():
        total, nhe = existencias.get(cod, (0.0, 0.0))
        disponible = total - nhe
        ing, ven = ingreso.get(cod, 0.0), vendido.get(cod, 0.0)
        ult = ultima_venta.get(cod)
        ult = ult.date() if isinstance(ult, datetime) else ult
        ventas_24 = sum(v for m, v in venta_mes[cod].items() if m >= meses[0])
        if not (ing or ven or abs(total) > 1e-9 or ventas_24):
            sin_actividad += 1
            continue
        if f["existencia"] == "con" and disponible <= 0:   # «Solo con existencia»: disponible (sin NHE) mayor que cero
            continue
        por_mes_anio = defaultdict(float)
        for m, v in venta_mes[cod].items():
            if m >= meses[0]:
                por_mes_anio[int(m[5:])] += v
        temp, temp_pct = temporada(por_mes_anio)
        mensual = sum(v for m, v in venta_mes[cod].items() if m >= desde_12) / 12
        cobertura = disponible / mensual if mensual > 0 else None
        dias_sin = (hoy - ult).days if ult else None
        # Comentario
        if not ult:
            estado, comentario = "sin_venta", "Sin venta: nunca se ha vendido"
        elif dias_sin > 365:
            estado, comentario = "sin_venta", f"Sin venta hace {meses_txt(dias_sin)} (última {ult:%d/%m/%Y})"
        elif ing > 0 and ven <= 0:
            estado, comentario = "sin_venta", "Ingresó y aún no se ha vendido"
        elif ven <= 0:
            estado, comentario = "sin_venta", f"Sin venta desde el {f['desde']:%d/%m/%Y}"
        elif disponible <= 0:
            estado, comentario = "agotado", "Agotado: se vende, revisar reposición"
        elif cobertura is not None and cobertura > 12:
            plazo = "más de 10 años" if cobertura > 120 else (f"{cobertura / 12:.1f} años" if cobertura > 24 else f"{cobertura:.0f} meses")
            estado, comentario = "lento", f"Lento: inventario para {plazo}"
        else:
            estado, comentario = "rota", f"Se vende: inventario para {cobertura:.1f} meses" if cobertura is not None else "Se vende"
        if nhe > 0:
            comentario += f" · {nhe:,.0f} en NHE"
        fob = p["fob"]
        filas.append({**p, "ingreso": ing, "ingreso_fob": ing * fob if fob else None,
                      "vendido": ven, "venta_neta": venta_neta.get(cod, 0.0), "costo_vendido": costo_vendido.get(cod, 0.0),
                      "margen": (venta_neta[cod] - costo_vendido[cod]) / venta_neta[cod] * 100 if venta_neta.get(cod, 0) > 0 else None,
                      "inventario_costo": max(disponible, 0) * p["costo"],
                      "existencia": total, "nhe": nhe, "disponible": disponible,
                      "inventario_fob": disponible * fob if fob else None, "nhe_fob": nhe * fob if fob and nhe else None,
                      "ultima_venta": ult, "dias_sin_venta": dias_sin, "ultimo_ingreso": ultimo_ingreso.get(cod),
                      "temporada": temp, "temporada_pct": temp_pct, "mensual": mensual, "cobertura": cobertura,
                      "estado": estado, "comentario": comentario,
                      "venta_meses": [venta_mes[cod].get(m, 0.0) for m in meses],
                      "monto_meses": [monto_mes[cod].get(m, 0.0) for m in meses]})

    # Resumen por marca y total (antes de los filtros de estado / búsqueda, para que cuadre con el proveedor)
    def acumular(destino, x):
        for campo in ("ingreso", "vendido", "disponible", "nhe"):
            destino[campo] += x[campo]
        for campo in ("ingreso_fob", "venta_neta", "costo_vendido", "inventario_costo", "inventario_fob", "nhe_fob"):
            destino[campo] += x[campo] or 0
        destino["productos"] += 1
        destino["sin_venta"] += x["estado"] == "sin_venta"
        destino["sin_venta_fob"] += (x["inventario_fob"] or 0) if x["estado"] == "sin_venta" else 0
        destino["sin_fob"] += x["fob"] is None
    nuevo = lambda: defaultdict(float)
    total, por_marca = nuevo(), defaultdict(nuevo)
    venta_total_mes = [0.0] * len(meses)
    venta_total_mes_monto = [0.0] * len(meses)
    por_estado = {k: {"estado": k, "nombre": nombre, "color": color, "productos": 0, "disponible": 0.0, "inventario_fob": 0.0,
                      "vendido": 0.0, "venta_neta": 0.0} for k, nombre, color in ROTACION}
    for x in filas:
        acumular(total, x)
        acumular(por_marca[x["marca"]], x)
        for i, (v, mto) in enumerate(zip(x["venta_meses"], x["monto_meses"])):
            venta_total_mes[i] += v
            venta_total_mes_monto[i] += mto
        e = por_estado[x["estado"]]
        e["productos"] += 1
        e["disponible"] += max(x["disponible"], 0)
        e["inventario_fob"] += max(x["inventario_fob"] or 0, 0)
        e["vendido"] += x["vendido"]
        e["venta_neta"] += x["venta_neta"]
    rotacion = list(por_estado.values())
    total_productos = sum(e["productos"] for e in rotacion)
    total_inv = sum(e["inventario_fob"] for e in rotacion)
    for e in rotacion:
        e["pct_productos"] = e["productos"] / total_productos * 100 if total_productos else 0
        e["pct_inventario"] = e["inventario_fob"] / total_inv * 100 if total_inv else 0
    dias = (hoy - f["desde"]).days + 1
    marcas = sorted(({"marca": m, **v} for m, v in por_marca.items()), key=lambda x: -x["inventario_fob"])
    for m in marcas:
        m["rent"] = rentabilidad(m, dias)
    rent_total = rentabilidad(total, dias) if filas else None
    por_mes_anio = defaultdict(float)
    for m, v in zip(meses, venta_total_mes_monto):
        por_mes_anio[int(m[5:])] += v
    temp_total = temporada(por_mes_anio) if sum(venta_total_mes) >= MIN_UNIDADES_TEMPORADA else (None, None)

    # Filtros de la tabla
    q = f["q"].upper()
    if q:
        filas = [x for x in filas if q in x["codigo"].upper() or q in x["parte"].upper() or q in x["nombre"].upper()]
    if f["estado"] == "nhe":
        filas = [x for x in filas if x["nhe"] > 0]
    elif f["estado"] == "ingreso":
        filas = [x for x in filas if x["ingreso"] > 0]
    elif f["estado"]:
        filas = [x for x in filas if x["estado"] == f["estado"]]
    clave = {"inventario": lambda x: -(x["inventario_fob"] or 0), "ingreso": lambda x: -(x["ingreso_fob"] or 0),
             "vendido": lambda x: -(x["venta_neta"] or 0),
             "sin_venta": lambda x: (x["dias_sin_venta"] is not None, -(x["dias_sin_venta"] or 0), -(x["inventario_fob"] or 0)),
             "marca": lambda x: (x["marca"], x["nombre"])}[f["orden"]]
    filas.sort(key=clave)
    return {"filas": filas, "marcas": marcas, "total": total, "meses": meses, "etiquetas": [f"{MESES[int(m[5:]) - 1]} {m[2:4]}" for m in meses],
            "venta_mes": venta_total_mes, "venta_mes_monto": venta_total_mes_monto, "temporada": temp_total, "sin_actividad": sin_actividad,
            "rotacion": rotacion, "rent": rent_total, "dias": dias}


@analisis_proveedor_bp.route("/analisis_proveedor")
@requiere_acceso("analisis_proveedor")
def reporte():
    f = leer_filtros(request.args)
    marcas_prov = marcas_del_proveedor(f["proveedor"])
    d = calcular(f)
    return render_template("analisis_proveedor.html", f=f, d=d, proveedores=lista_proveedores(), marcas_prov=marcas_prov,
                           estados=ESTADOS, ordenes=ORDENES, max_filas=MAX_FILAS_PANTALLA)


@analisis_proveedor_bp.route("/analisis_proveedor/excel")
@requiere_acceso("analisis_proveedor")
def excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    wb = Workbook()
    nota = (f"Proveedor: {f['proveedor']} · Marcas: {', '.join(f['marcas']) or 'todas'}{' · Solo con existencia' if f['existencia'] else ''} · Ingresos y ventas desde el {f['desde']:%d/%m/%Y} · "
            f"Ventas = venta neta C$ (sin IVA, con descuento, sin anuladas) · FOB = cantidad × último precio FOB del catálogo (US$) · "
            f"Inventario disponible sin bodegas NHE · Generado {datetime.now():%d/%m/%Y %H:%M}")

    def hoja(ws, titulo, columnas):
        ws["A1"], ws["A2"] = titulo, nota
        ws["A1"].font = Font(bold=True, size=14, color=VERDE)
        ws["A2"].font = Font(size=9, color="52514E")
        for i, (nombre, ancho) in enumerate(columnas, start=1):
            c = ws.cell(row=4, column=i, value=nombre)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NARANJA if "NHE" in nombre or "Comentario" in nombre else VERDE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.row_dimensions[4].height = 32
        ws.freeze_panes = "B5"

    ws = wb.active
    ws.title = "Por marca"
    cols = [("Marca", 22), ("Productos", 10), ("Ingresado (unid.)", 12), ("Ingresado FOB US$", 14), ("Vendido (unid.)", 12),
            ("Venta neta C$", 14), ("Inventario disponible (unid.)", 13), ("Inventario FOB US$", 14), ("NHE (unid.)", 10),
            ("NHE FOB US$", 12), ("Productos sin venta", 11), ("Inventario sin venta FOB US$", 14), ("Costo de lo vendido C$", 14),
            ("Utilidad bruta C$", 14), ("Margen %", 9), ("Inventario a costo C$", 14), ("Rotación (veces al año)", 10),
            ("GMROI (C$ por C$ 1 de inventario al año)", 13), ("Veredicto", 60)]
    hoja(ws, f"Análisis de proveedor: {f['proveedor']}", cols)
    for n, m in enumerate(d["marcas"], start=5):
        valores = [m["marca"], m["productos"], m["ingreso"], round(m["ingreso_fob"], 2), m["vendido"], round(m["venta_neta"], 2), m["disponible"],
                   round(m["inventario_fob"], 2), m["nhe"], round(m["nhe_fob"], 2), m["sin_venta"], round(m["sin_venta_fob"], 2),
                   round(m["costo_vendido"], 2), round(m["rent"]["utilidad"], 2),
                   m["rent"]["margen"] / 100 if m["rent"]["margen"] is not None else None, round(m["inventario_costo"], 2),
                   round(m["rent"]["rotacion"], 2) if m["rent"]["rotacion"] is not None else None,
                   round(m["rent"]["gmroi"], 2) if m["rent"]["gmroi"] is not None else None, m["rent"]["veredicto"]]
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=n, column=i, value=v)
            if cols[i - 1][0] == "Margen %":
                c.number_format = "0.0%"
                continue
            c.number_format = "#,##0.00" if ("FOB" in cols[i - 1][0] or "C$" in cols[i - 1][0]) else "#,##0.##"
    tot = 5 + len(d["marcas"])
    ws.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    for i in range(2, 15):   # sumas hasta la utilidad; margen, rotación y GMROI del total van en la fila siguiente
        l = get_column_letter(i)
        c = ws.cell(row=tot, column=i, value=f"=SUM({l}5:{l}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00" if ("FOB" in cols[i - 1][0] or "C$" in cols[i - 1][0]) else "#,##0.##"

    r = d.get("rent")
    if r:
        ws.cell(row=tot, column=15, value=r["margen"] / 100 if r["margen"] is not None else None).number_format = "0.0%"
        ws.cell(row=tot, column=16, value=f"=SUM(P5:P{tot - 1})").number_format = "#,##0.00"
        ws.cell(row=tot, column=17, value=round(r["rotacion"], 2) if r["rotacion"] is not None else None)
        ws.cell(row=tot, column=18, value=round(r["gmroi"], 2) if r["gmroi"] is not None else None)
        ws.cell(row=tot, column=19, value=r["veredicto"]).font = Font(bold=True)

    ws2 = wb.create_sheet("Detalle")
    cols2 = [("Código", 11), ("Nº parte", 18), ("Descripción", 40), ("Marca", 14), ("Segmento", 16), ("Familia", 22), ("Subfamilia", 24),
             ("Categoría", 16), ("Precio FOB US$", 11), ("Ingresado (unid.)", 11), ("Ingresado FOB US$", 13), ("Último ingreso", 12),
             ("Vendido (unid.)", 11), ("Venta neta C$", 13), ("Margen %", 9), ("Existencia total", 11), ("NHE (unid.)", 10), ("Inventario disponible", 11),
             ("Inventario FOB US$", 13), ("NHE FOB US$", 12), ("Venta mensual (12 meses)", 11), ("Meses de inventario", 10),
             ("Última venta", 12), ("Temporada de más venta", 13), ("% del año en esa temporada", 11), ("Comentario", 48)]
    hoja(ws2, f"Detalle por producto: {f['proveedor']}", cols2)
    formatos = {"Precio FOB US$": "#,##0.00##", "Ingresado FOB US$": "#,##0.00", "Venta neta C$": "#,##0.00", "Inventario FOB US$": "#,##0.00",
                "NHE FOB US$": "#,##0.00", "Último ingreso": "dd/mm/yyyy", "Última venta": "dd/mm/yyyy", "Venta mensual (12 meses)": "#,##0.0",
                "Meses de inventario": "#,##0.0", "% del año en esa temporada": "0%", "Margen %": "0.0%"}
    for n, x in enumerate(d["filas"], start=5):
        valores = [x["codigo"], x["parte"], x["nombre"], x["marca"], x["segmento"], x["familia"], x["subfamilia"], x["categoria"], x["fob"],
                   x["ingreso"] or None, x["ingreso_fob"] if x["ingreso"] else None, x["ultimo_ingreso"], x["vendido"] or None,
                   round(x["venta_neta"], 2) or None, x["margen"] / 100 if x["margen"] is not None else None, x["existencia"], x["nhe"] or None, x["disponible"], x["inventario_fob"],
                   x["nhe_fob"], round(x["mensual"], 2) or None, round(x["cobertura"], 1) if x["cobertura"] is not None else None,
                   x["ultima_venta"], x["temporada"], x["temporada_pct"] / 100 if x["temporada_pct"] else None, x["comentario"]]
        for i, v in enumerate(valores, start=1):
            c = ws2.cell(row=n, column=i, value=v)
            c.number_format = formatos.get(cols2[i - 1][0], "#,##0.##" if isinstance(v, (int, float)) else c.number_format)
    ws2.auto_filter.ref = f"A4:{get_column_letter(len(cols2))}{4 + max(len(d['filas']), 1)}"

    ws3 = wb.create_sheet("Venta por mes")
    hoja(ws3, "Ventas por mes (últimos 24 meses)", [("Mes", 10), ("Unidades vendidas", 12), ("Venta neta C$", 14)])
    for n, (e, u, v) in enumerate(zip(d.get("etiquetas", []), d.get("venta_mes", []), d.get("venta_mes_monto", [])), start=5):
        ws3.cell(row=n, column=1, value=e)
        ws3.cell(row=n, column=2, value=u).number_format = "#,##0"
        ws3.cell(row=n, column=3, value=round(v, 2)).number_format = "#,##0.00"

    ws4 = wb.create_sheet("Rotación")
    hoja(ws4, "Rotación general", [("Rotación", 36), ("Productos", 10), ("% de productos", 11), ("Inventario disponible (unid.)", 13),
                                   ("Inventario FOB US$", 14), ("% del inventario FOB", 11), ("Vendido (unid.)", 12), ("Venta neta C$", 14)])
    for n, e in enumerate(d.get("rotacion", []), start=5):
        valores = [e["nombre"], e["productos"], e["pct_productos"] / 100, e["disponible"], round(e["inventario_fob"], 2), e["pct_inventario"] / 100,
                   e["vendido"], round(e["venta_neta"], 2)]
        for i, v in enumerate(valores, start=1):
            ws4.cell(row=n, column=i, value=v).number_format = {3: "0.0%", 6: "0.0%", 5: "#,##0.00", 8: "#,##0.00"}.get(i, "#,##0.##")

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = "".join(ch if ch.isalnum() else "_" for ch in f["proveedor"])[:40] or "proveedor"
    return send_file(salida, as_attachment=True, download_name=f"analisis_proveedor_{nombre}_{f['desde']}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
