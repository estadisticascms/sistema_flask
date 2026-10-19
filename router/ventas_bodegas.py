"""Ventas por bodega: cantidad vendida, productos distintos y facturas emitidas (contado / crédito),
con detalle de unidades por producto en cada bodega."""
import time
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.dashboard import ultima_fecha_con_ventas
from router.ventas_vendedor import rangos_rapidos

ventas_bodega_bp = Blueprint("ventas_bodega", __name__)

# Bodegas que van en el gráfico y en la tabla principal (las demás se agrupan en "Otras bodegas")
PRINCIPALES = ["BODEGA #1", "BODEGA #2", "BODEGA #3", "BODEGA #4", "BODEGA #5 MONTOYA", "SUC. IVAN MONTENEGRO"]
TIPOS = ["CONTADO", "CREDITO"]
MAX_PRODUCTOS_BODEGA = 300   # filas de detalle por bodega en pantalla
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


_CACHE = {}


def en_cache(clave, calcular, segundos=600):
    """Listas para los filtros (marcas, vendedores): cambian poco, se guardan 10 minutos."""
    ahora = time.time()
    if clave not in _CACHE or ahora - _CACHE[clave][0] > segundos:
        _CACHE[clave] = (ahora, calcular())
    return _CACHE[clave][1]


def consulta(cursor, sql, params=()):
    cursor.execute(sql, params)
    return cursor.fetchall()


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def leer_filtros(args):
    ultima = ultima_fecha_con_ventas() or date.today()
    if isinstance(ultima, datetime):
        ultima = ultima.date()
    fecha_fin = fecha_param(args.get("fecha_fin"), ultima)
    fecha_inicio = min(fecha_param(args.get("fecha_inicio"), fecha_fin), fecha_fin)  # por defecto, un día
    tipo = (args.get("tipo") or "").upper()
    dias = (fecha_fin - fecha_inicio).days + 1
    detalle = args.get("detalle")
    if detalle not in ("dia", "semana", "mes") or (detalle == "dia" and dias > 31):
        detalle = "dia" if dias <= 31 else ("semana" if dias <= 120 else "mes")
    return {"fecha_inicio": fecha_inicio, "fecha_fin": fecha_fin, "ultima": ultima, "dias": dias,
            "marca": (args.get("marca") or "").strip(), "vendedor": (args.get("vendedor") or "").strip(),
            "tipo": tipo if tipo in TIPOS else "", "detalle": detalle}


def periodo_de(fecha, detalle):
    """Clave ordenable y etiqueta de la columna de detalle a la que pertenece una fecha."""
    if detalle == "dia":
        return fecha.isoformat(), f"{fecha.day:02d}/{fecha.month:02d}"
    if detalle == "semana":  # semanas de lunes a domingo
        lunes = fecha - timedelta(days=fecha.weekday())
        domingo = lunes + timedelta(days=6)
        return lunes.isoformat(), f"{lunes.day:02d}/{lunes.month:02d}–{domingo.day:02d}/{domingo.month:02d}"
    return f"{fecha.year}-{fecha.month:02d}", f"{MESES[fecha.month - 1]} {fecha.year}"


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()


def obtener_ventas_bodega(f):
    where = ["Fecha BETWEEN %s AND %s", "Estatus <> 'Anulado'"]
    params = [f["fecha_inicio"].isoformat(), f["fecha_fin"].isoformat()]  # Fecha es texto 'YYYY-MM-DD'
    if f["marca"]:
        where.append("Marca = %s"); params.append(f["marca"])
    if f["vendedor"]:
        where.append("Vendedor = %s"); params.append(f["vendedor"])
    if f["tipo"]:
        where.append("`Contado/Credito` = %s"); params.append(f["tipo"])

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Una sola lectura de ventas_por_producto (no tiene índices); todo se agrega en Python
        cursor.execute(f"""
            SELECT COALESCE(NULLIF(TRIM(Bodega), ''), 'Sin bodega'), UPPER(`Contado/Credito`), Codigo, Numero_parte,
                   `Descrip Prod.`, Marca, Fecha, `No. Factura`, SUM(`Cantidad Vendida`)
            FROM ventas_por_producto WHERE {' AND '.join(where)}
            GROUP BY 1, 2, Codigo, Numero_parte, `Descrip Prod.`, Marca, Fecha, `No. Factura`
        """, params)
        filas = cursor.fetchall()
        vendedores = en_cache("vendedores", lambda: [r[0] for r in consulta(cursor,
            "SELECT DISTINCT vendedor FROM ventas_historicas WHERE Fecha >= %s ORDER BY vendedor",
            (date.today() - timedelta(days=365),)) if r[0]])
        marcas = en_cache("marcas", lambda: [r[0] for r in consulta(cursor,
            "SELECT DISTINCT Marca FROM catalogo_productos WHERE Marca IS NOT NULL AND Marca <> '' ORDER BY Marca")])
    finally:
        cursor.close()
        conn.close()

    resumen = {}    # bodega -> tipo -> {cantidad, productos:set, facturas:set}
    detalle = {}    # bodega -> codigo -> {datos, periodos:{clave: cant}, total}
    periodos = {}   # clave -> etiqueta
    productos_global, facturas_global = set(), set()
    for bodega, tipo, codigo, parte, desc, marca, fecha, factura, cant in filas:
        cant = float(cant or 0)
        tipo = tipo if tipo in TIPOS else "CONTADO"
        r = resumen.setdefault(bodega, {}).setdefault(tipo, {"cantidad": 0.0, "productos": set(), "facturas": set()})
        r["cantidad"] += cant
        r["productos"].add(codigo)
        r["facturas"].add(factura)
        productos_global.add(codigo)
        facturas_global.add((bodega, factura))

        clave, etiqueta = periodo_de(a_fecha(fecha), f["detalle"])
        periodos[clave] = etiqueta
        p = detalle.setdefault(bodega, {}).setdefault(codigo, {"codigo": codigo, "numero_parte": parte, "descripcion": desc,
                                                                "marca": marca, "periodos": {}, "total": 0.0})
        p["periodos"][clave] = p["periodos"].get(clave, 0.0) + cant
        p["total"] += cant

    vacio = {"cantidad": 0.0, "productos": set(), "facturas": set()}

    def fila_resumen(nombre, por_tipo):
        """En "Otras bodegas" los productos llegan como (bodega, código): cada bodega cuenta aparte en su fila."""
        fila = {"bodega": nombre, "_conjuntos": {}}
        union = set()
        for t in TIPOS:
            d = por_tipo.get(t, vacio)
            fila[t] = {"cantidad": d["cantidad"], "productos": len(d["productos"]), "facturas": len(d["facturas"])}
            fila["_conjuntos"][t] = set(d["productos"])
            union |= d["productos"]
        # Productos del total = distintos (un producto vendido al contado y a crédito cuenta una vez)
        fila["total"] = {"cantidad": sum(fila[t]["cantidad"] for t in TIPOS), "productos": len(union),
                         "facturas": sum(fila[t]["facturas"] for t in TIPOS)}
        fila["_conjuntos"]["total"] = union
        return fila

    principales = sorted((fila_resumen(b, resumen.get(b, {})) for b in PRINCIPALES),
                         key=lambda x: -x["total"]["cantidad"])
    otras_nombres = sorted(b for b in resumen if b not in PRINCIPALES)
    otras = None
    if otras_nombres:
        agregadas = {}
        for b in otras_nombres:
            for t, d in resumen[b].items():
                a = agregadas.setdefault(t, {"cantidad": 0.0, "productos": set(), "facturas": set()})
                a["cantidad"] += d["cantidad"]
                a["productos"] |= {(b, x) for x in d["productos"]}
                a["facturas"] |= {(b, x) for x in d["facturas"]}
        otras = fila_resumen(f"Otras bodegas ({len(otras_nombres)})", agregadas)
        otras["nombres"] = otras_nombres

    def suma(filas_):
        """Cantidad y facturas se suman; productos = distintos entre las bodegas sumadas."""
        total = {"bodega": "Total"}
        for t in TIPOS + ["total"]:
            productos = set()
            for x in filas_:  # (bodega, código) de "Otras bodegas" -> código, para no contar dos veces un producto
                productos |= {p[1] if isinstance(p, tuple) else p for p in x["_conjuntos"][t]}
            total[t] = {"cantidad": sum(x[t]["cantidad"] for x in filas_), "productos": len(productos),
                        "facturas": sum(x[t]["facturas"] for x in filas_)}
        return total

    con_venta = [x for x in principales if x["total"]["cantidad"] > 0]
    claves = sorted(periodos)
    bloques = []
    for bodega in PRINCIPALES + otras_nombres:
        if bodega not in detalle:
            continue
        prods = sorted(detalle[bodega].values(), key=lambda p: (-p["total"], p["descripcion"] or ""))
        bloques.append({"bodega": bodega, "principal": bodega in PRINCIPALES, "productos": prods[:MAX_PRODUCTOS_BODEGA],
                        "total_productos": len(prods), "unidades": sum(p["total"] for p in prods),
                        "por_periodo": {c: sum(p["periodos"].get(c, 0) for p in prods) for c in claves}})

    return {
        "principales": principales, "otras": otras,
        "total_principales": suma(principales),
        "total_general": suma(principales + ([otras] if otras else [])),
        "mas_carga": con_venta[0] if con_venta else None,
        "menos_carga": con_venta[-1] if len(con_venta) > 1 else None,
        "unidades": sum(x["total"]["cantidad"] for x in principales) + (otras["total"]["cantidad"] if otras else 0),
        "productos_distintos": len(productos_global), "facturas": len(facturas_global),
        "bloques": bloques, "periodos": [(c, periodos[c]) for c in claves],
        "vendedores": vendedores, "marcas": marcas,
    }


@ventas_bodega_bp.route("/ventas_bodega", methods=["GET", "POST"])
@requiere_acceso("ventas_bodega")
def reporte_ventas_bodega():
    f = leer_filtros(request.values)  # acepta enlaces viejos con POST
    datos = obtener_ventas_bodega(f)
    # El gráfico va en el orden fijo de las bodegas (como el reporte original); la tabla, por cantidad
    por_nombre = {b["bodega"]: b for b in datos["principales"]}
    paneles = [{"bodega": n, "CONTADO": por_nombre[n]["CONTADO"], "CREDITO": por_nombre[n]["CREDITO"]} for n in PRINCIPALES]
    return render_template("reporte_ventas_bodega.html", f=f, rangos=rangos_rapidos(f["ultima"]),
                           max_productos=MAX_PRODUCTOS_BODEGA, paneles=paneles, **datos)
