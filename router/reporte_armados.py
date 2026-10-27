"""Reporte de armado de mangueras (reemplaza a "Bienes y mixtas").

Sale de ventas_armados (router/cargar_armados.py): cuántas mangueras se arman por día, la venta, la mano de obra (MO) y lo
cobrado de soldadura (Otros). El reporte de armados no trae el costo, así que la ganancia y el margen se toman de la factura en
ventas_historicas (misma factura y fecha): utilidad = venta neta (con descuento) − costo de los materiales."""
import re
from datetime import date, datetime, timedelta
from decimal import Decimal

from flask import Blueprint, render_template, request

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_armados import tipo_linea
from router.ventas_vendedor import rangos_rapidos

reporte_armados_bp = Blueprint("reporte_armados", __name__)

# Medida de la manguera en la descripción: 1/4, 3/8, 1 1/4 o 1" (pulgada entera)
MEDIDA = re.compile(r"(?<![\w/.])(\d \d/\d{1,2}|\d/\d{1,2}|\d(?=\s*\"))(?![\w/])")
# o en código de dieciseisavos de pulgada: 100R1-04, R02-12, #10
DIECISEISAVOS = re.compile(r"(?:R\d{1,2}-|#)(\d{2})\b")
DASH = {"03": '3/16"', "04": '1/4"', "05": '5/16"', "06": '3/8"', "08": '1/2"', "10": '5/8"', "12": '3/4"', "16": '1"',
        "20": '1-1/4"', "24": '1-1/2"', "32": '2"'}


def medida(producto):
    m = MEDIDA.search(producto or "")
    if m:
        return m.group(1).replace(" ", "-") + '"'
    m = DIECISEISAVOS.search(producto or "")
    return DASH.get(m.group(1), "Sin medida") if m else "Sin medida"


def _pulgadas(texto):
    """'1-1/4"' -> 1.25 para ordenar las medidas; 'Sin medida' al final."""
    try:
        return sum(float(a) / float(b or 1) for a, _, b in (p.partition("/") for p in texto.rstrip('"').split("-")))
    except ValueError:
        return 99.0


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def _consulta(sql, params=()):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(sql, params)
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()


def leer_filtros(args):
    ultima = _consulta("SELECT MAX(fecha) AS f FROM ventas_armados")[0]["f"] or date.today()
    hasta = fecha_param(args.get("hasta"), ultima)
    desde = min(fecha_param(args.get("desde"), ultima.replace(day=1)), hasta)
    return {"desde": desde, "hasta": hasta, "ultima": ultima}


def _acumulador(**extra):
    return {"facturas": 0, "mangueras": Decimal(0), "servicios": Decimal(0), "subtotal": Decimal(0), "mano_obra": Decimal(0),
            "otros": Decimal(0), "neto": Decimal(0), "costo": Decimal(0), "utilidad": Decimal(0), "neto_con_costo": Decimal(0),
            "credito": Decimal(0), "contado": Decimal(0), "mang_credito": Decimal(0), "mang_contado": Decimal(0), **extra}


def _sumar(acc, f):
    acc["facturas"] += 1
    for c in ("mangueras", "servicios", "subtotal", "mano_obra", "otros", "neto"):
        acc[c] += f[c]
    acc["credito" if f["tipo_venta"] == "CREDITO" else "contado"] += f["neto"]
    acc["mang_credito" if f["tipo_venta"] == "CREDITO" else "mang_contado"] += f["mangueras"]
    if f["costo"] is not None:   # el margen solo usa facturas con costo (las que están en Ventas del día)
        acc["costo"] += f["costo"]
        acc["utilidad"] += f["utilidad"]
        acc["neto_con_costo"] += f["neto"]


def _cerrar(acc):
    acc["margen"] = acc["utilidad"] / acc["neto_con_costo"] * 100 if acc["neto_con_costo"] else None
    return acc


def calcular(f):
    lineas = _consulta("SELECT * FROM ventas_armados WHERE fecha BETWEEN %s AND %s ORDER BY fecha, nfactura, id",
                       (f["desde"], f["hasta"]))
    ventas = {}
    for v in _consulta("SELECT factura, Fecha, tipo_venta, `tipo factura` AS tipo, SubTotal, Descuento, neto, `Costo Total` AS costo, "
                       "utilidad FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s AND estado_factura <> 'Anulado' "
                       "AND (`tipo factura` = 'Mixtos' OR factura IN (SELECT nfactura FROM ventas_armados WHERE fecha BETWEEN %s AND %s))",
                       (f["desde"], f["hasta"], f["desde"], f["hasta"])):
        ventas.setdefault((v["factura"], v["Fecha"]), []).append(v)

    facturas, medidas = {}, {}
    for l in lineas:
        clave = (l["nfactura"], l["fecha"])
        fac = facturas.get(clave)
        if fac is None:
            fac = facturas[clave] = {"nfactura": l["nfactura"], "fecha": l["fecha"], "tipo_venta": l["tipo_venta"], "cliente": l["cliente"],
                                     "facturador": l["facturador"] or "Sin dato", "vendedor": l["vendedor"] or "Sin dato",
                                     "estatus": l["estatus"], "mangueras": Decimal(0), "servicios": Decimal(0), "subtotal": Decimal(0),
                                     "mano_obra": Decimal(0), "otros": Decimal(0), "productos": []}
        tipo = tipo_linea(l["producto"], l["mano_obra"])
        if tipo == "manguera":
            fac["mangueras"] += l["cantidad"]
            m = medidas.setdefault(medida(l["producto"]), {"medida": medida(l["producto"]), "mangueras": Decimal(0), "subtotal": Decimal(0)})
            m["mangueras"] += l["cantidad"]
            m["subtotal"] += l["subtotal"]
        elif tipo == "servicio":
            fac["servicios"] += l["cantidad"]
        for c in ("subtotal", "mano_obra", "otros"):
            fac[c] += l[c]
        fac["productos"].append({**l, "tipo": tipo})

    sin_costo = []
    for clave, fac in facturas.items():
        candidatas = ventas.pop(clave, [])
        v = next((c for c in candidatas if c["tipo_venta"] == fac["tipo_venta"]), candidatas[0] if candidatas else None)
        if v:
            fac.update(neto=v["neto"], descuento=v["Descuento"], costo=v["costo"], utilidad=v["utilidad"])
        else:   # aún no se cargó la venta: se usa el subtotal del armado y queda fuera del margen
            fac.update(neto=fac["subtotal"], descuento=None, costo=None, utilidad=None)
            sin_costo.append(fac)
        fac["margen"] = fac["utilidad"] / fac["neto"] * 100 if v and fac["neto"] else None

    total, dias, por_facturador, por_vendedor = _acumulador(), {}, {}, {}
    for fac in facturas.values():
        _sumar(total, fac)
        _sumar(dias.setdefault(fac["fecha"], _acumulador(fecha=fac["fecha"])), fac)
        _sumar(por_facturador.setdefault(fac["facturador"], _acumulador(nombre=fac["facturador"])), fac)
        _sumar(por_vendedor.setdefault(fac["vendedor"], _acumulador(nombre=fac["vendedor"])), fac)
    _cerrar(total)
    dias = [_cerrar(d) for _, d in sorted(dias.items())]
    # Facturas mixtas que no aparecen en el reporte de armados (falta cargarlo o no son de armado)
    mixtas_sin_armado = sorted((v for vs in ventas.values() for v in vs if v["tipo"] == "Mixtos"), key=lambda v: (v["Fecha"], v["factura"]))
    return {
        "total": total, "dias": dias,
        "promedio_dia": total["mangueras"] / len(dias) if dias else 0,
        "mejor_dia": max(dias, key=lambda d: d["mangueras"]) if dias else None,
        "por_facturador": sorted((_cerrar(a) for a in por_facturador.values()), key=lambda a: -a["mangueras"]),
        "por_vendedor": sorted((_cerrar(a) for a in por_vendedor.values()), key=lambda a: -a["neto"]),
        "medidas": sorted(medidas.values(), key=lambda m: _pulgadas(m["medida"])),
        "facturas": sorted(facturas.values(), key=lambda x: (x["fecha"], x["nfactura"])),
        "sin_costo": sin_costo, "mixtas_sin_armado": mixtas_sin_armado,
    }


def armados_del_dia(dia):
    """Mangueras armadas del día y acumulado del mes (para la tarjeta de ventas del día por vendedor).
    None si no hay armados cargados en el mes; 'pendiente' = el día todavía no se ha cargado."""
    inicio = dia.replace(day=1)
    lineas = _consulta("SELECT fecha, nfactura, tipo_venta, producto, cantidad, mano_obra FROM ventas_armados WHERE fecha BETWEEN %s AND %s",
                       (inicio, dia))
    if not lineas:
        return None
    r = {"dia": Decimal(0), "credito": Decimal(0), "contado": Decimal(0), "mes": Decimal(0), "facturas": set(), "dias_mes": set()}
    for l in lineas:
        if tipo_linea(l["producto"], l["mano_obra"]) != "manguera":
            continue
        r["mes"] += l["cantidad"]
        r["dias_mes"].add(l["fecha"])
        if l["fecha"] == dia:
            r["dia"] += l["cantidad"]
            r["credito" if l["tipo_venta"] == "CREDITO" else "contado"] += l["cantidad"]
            r["facturas"].add(l["nfactura"])
    r["facturas"], r["dias_mes"] = len(r["facturas"]), len(r["dias_mes"])
    r["promedio"] = r["mes"] / r["dias_mes"] if r["dias_mes"] else 0
    r["pendiente"] = max(l["fecha"] for l in lineas) < dia
    return r


@reporte_armados_bp.route("/reporte_armados")
@requiere_acceso("reporte_armados")
def reporte():
    f = leer_filtros(request.args)
    return render_template("reporte_armados.html", f=f, d=calcular(f), rangos=rangos_rapidos(f["ultima"]),
                           params={"desde": f["desde"].isoformat(), "hasta": f["hasta"].isoformat()})


# ---------------------------------------------------------------------------
# Resumen para gerencia (una hoja A4 que se descarga en PDF)
# ---------------------------------------------------------------------------
MARGEN_BAJO = 20   # % de margen debajo del cual una factura se señala


def periodo_anterior(desde, hasta):
    """Mes en curso (del 1 al día X) -> mismos días del mes anterior; otro rango -> el mismo número de días justo antes."""
    if desde.day == 1 and (desde.year, desde.month) == (hasta.year, hasta.month):
        fin_mes = desde - timedelta(days=1)
        return fin_mes.replace(day=1), fin_mes.replace(day=min(hasta.day, fin_mes.day))
    dias = (hasta - desde).days + 1
    return desde - timedelta(days=dias), desde - timedelta(days=1)


def nombre_limpio(nombre):
    """'GEYCEL . ORTIZ .' -> 'Geycel Ortiz' (los nombres del sistema traen puntos sueltos); conserva iniciales como 'E.'."""
    return re.sub(r"\s+\.(?=\s|$)", "", nombre or "").strip().title() or "Sin dato"


def _variacion(actual, anterior):
    return (float(actual) / float(anterior) - 1) * 100 if anterior else None


def puntos_de_atencion(d):
    """Frases cortas para gerencia, de lo más importante a lo menos."""
    t, puntos = d["total"], []
    if d["por_facturador"] and t["mangueras"]:
        p = d["por_facturador"][0]
        puntos.append(f"{p['nombre']} armó {float(p['mangueras'] / t['mangueras'] * 100):.0f}% de las mangueras ({p['mangueras']:,.0f}).")
    con_medida = [m for m in d["medidas"] if m["medida"] != "Sin medida"]
    if con_medida:
        top = sorted(con_medida, key=lambda m: -m["mangueras"])[:3]
        puntos.append("Medidas más pedidas: " + ", ".join(f"{m['medida']} ({m['mangueras']:,.0f})" for m in top) + ".")
    bajas = [x for x in d["facturas"] if x["margen"] is not None and x["margen"] < MARGEN_BAJO]
    if bajas:
        puntos.append(f"{len(bajas)} factura(s) con margen menor a {MARGEN_BAJO}%: "
                      + ", ".join(f"#{x['nfactura']} {(x['cliente'] or '').title()[:28]} ({x['margen']:.0f}%)"
                                  for x in sorted(bajas, key=lambda x: x["margen"])[:3]) + ("…" if len(bajas) > 3 else "."))
    if t["mangueras"]:
        puntos.append(f"Se cobra en promedio C$ {float(t['mano_obra'] / t['mangueras']):,.0f} de mano de obra por manguera "
                      f"y C$ {float(t['otros']):,.0f} de soldadura en el período.")
    if d["sin_costo"]:
        puntos.append(f"{len(d['sin_costo'])} factura(s) sin costo en Ventas del día: no entran en el margen.")
    if d["mixtas_sin_armado"]:
        puntos.append(f"{len(d['mixtas_sin_armado'])} factura(s) mixtas no aparecen en el reporte de armados (revisar carga).")
    return puntos


@reporte_armados_bp.route("/reporte_armados/gerencia")
@requiere_acceso("reporte_armados")
def gerencia():
    f = leer_filtros(request.args)
    d = calcular(f)
    for p in d["por_facturador"] + d["por_vendedor"]:
        p["nombre"] = nombre_limpio(p["nombre"])
    desde_ant, hasta_ant = periodo_anterior(f["desde"], f["hasta"])
    a = calcular({"desde": desde_ant, "hasta": hasta_ant})["total"]
    t = d["total"]
    comparacion = None
    if a["facturas"]:
        comparacion = {"desde": desde_ant, "hasta": hasta_ant, "mangueras": _variacion(t["mangueras"], a["mangueras"]),
                       "neto": _variacion(t["neto"], a["neto"]), "utilidad": _variacion(t["utilidad"], a["utilidad"]),
                       "mano_obra": _variacion(t["mano_obra"], a["mano_obra"]),
                       "margen": float(t["margen"] - a["margen"]) if t["margen"] is not None and a["margen"] is not None else None,
                       "anterior": a}
    return render_template("reporte_armados_gerencia.html", f=f, d=d, c=comparacion, puntos=puntos_de_atencion(d),
                           params={"desde": f["desde"].isoformat(), "hasta": f["hasta"].isoformat()})
