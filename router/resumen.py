"""Resumen general: una sola pantalla con el acumulado del año (enero a la fecha) contra el mismo corte del año anterior."""
from flask import Blueprint, render_template, request
from datetime import date, datetime, timedelta
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.dashboard import con_cache, millones, venta_neta_por_mes

resumen_bp = Blueprint("resumen", __name__)

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
TOP_MARCAS = 5


def un_anio_antes(fecha):
    # El 29 de febrero no existe el año anterior
    return fecha.replace(year=fecha.year - 1, day=28 if (fecha.month, fecha.day) == (2, 29) else fecha.day)


def fecha_larga(fecha):
    return f"{fecha.day} {MESES[fecha.month - 1].lower()} {fecha.year}"


def leer_fecha(texto, defecto):
    try:
        return datetime.strptime(texto or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


# -----------------------------
# Último día con ventas cargadas (los datos no siempre llegan a hoy)
# -----------------------------
def ultima_venta(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s", (desde, hasta))
    ultima = cursor.fetchone()[0]
    cursor.close()
    conn.close()
    return ultima


# -----------------------------
# Indicadores de un período (sin anuladas; devoluciones cerradas sin IVA)
# -----------------------------
def indicadores(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT COUNT(*) AS facturas, COUNT(DISTINCT cliente) AS clientes,
               COALESCE(SUM(neto), 0) AS facturado, COALESCE(SUM(utilidad), 0) AS utilidad,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Credito' THEN neto END), 0) AS credito,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Contado' THEN neto END), 0) AS contado
        FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s AND estado_factura <> 'Anulado'
    """, (desde, hasta))
    m = {k: float(v or 0) for k, v in cursor.fetchone().items()}
    cursor.execute("""
        SELECT COALESCE(SUM(totaldv), 0) / 1.15 AS total,
               COALESCE(SUM(CASE WHEN tipofact = 'Credito' THEN totaldv END), 0) / 1.15 AS credito
        FROM devoluciones WHERE fechadv BETWEEN %s AND %s AND estado = 'Cerrado'
    """, (desde, hasta))
    dev = cursor.fetchone()
    cursor.close()
    conn.close()

    devoluciones, dev_credito = float(dev["total"] or 0), float(dev["credito"] or 0)
    m["neto"] = m["facturado"] - devoluciones
    m["credito"] -= dev_credito
    m["contado"] -= devoluciones - dev_credito
    m["margen"] = m["utilidad"] / m["facturado"] * 100 if m["facturado"] else None
    return m


# -----------------------------
# Recuperación mensual desde recibos
# -----------------------------
def recuperacion_por_mes(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DATE_FORMAT(fecharc, '%Y-%m'), SUM(total) FROM reciboscaja
        WHERE fecharc BETWEEN %s AND %s GROUP BY 1
    """, (desde, hasta))
    filas = {mes: float(total or 0) for mes, total in cursor.fetchall()}
    cursor.close()
    conn.close()
    return filas


# -----------------------------
# Venta neta por tipo de vendedor (el Ejército va aparte, sin importar quién lo atiende)
# -----------------------------
def ventas_por_tipo(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT grupo, SUM(venta) AS venta FROM (
            SELECT CASE WHEN v.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJÉRCITO DE NICARAGUA'
                        ELSE COALESCE(e.tipo, 'SIN TIPO') END AS grupo,
                   SUM(v.neto) AS venta
            FROM ventas_historicas v
            LEFT JOIN (SELECT tipo, nombre FROM cat_vendedor GROUP BY tipo, nombre) e ON e.nombre = v.Vendedor
            WHERE v.fecha BETWEEN %s AND %s AND v.estado_factura <> 'Anulado'
            GROUP BY grupo
            UNION ALL
            SELECT CASE WHEN r.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJÉRCITO DE NICARAGUA'
                        ELSE COALESCE(e.tipo, 'SIN TIPO') END AS grupo,
                   -SUM(r.totaldv) / 1.15 AS venta
            FROM devoluciones r
            LEFT JOIN (SELECT tipo, nombre FROM cat_vendedor GROUP BY tipo, nombre) e ON e.nombre = r.Vendedor
            WHERE r.fechadv BETWEEN %s AND %s AND r.estado = 'Cerrado'
            GROUP BY grupo
        ) datos
        GROUP BY grupo ORDER BY venta DESC
    """, (desde, hasta, desde, hasta))
    filas = [{"tipo": tipo, "venta": float(venta or 0)} for tipo, venta in cursor.fetchall()]
    cursor.close()
    conn.close()
    total = sum(f["venta"] for f in filas)
    for f in filas:
        f["pct"] = f["venta"] / total * 100 if total else 0
    return filas, total


# -----------------------------
# Marcas con más venta en el período
# -----------------------------
def marcas_top(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Marca, SUM(`Sub-Total 2`) AS venta, SUM(`Cantidad Vendida` * `Costo Unit.`) AS costo
        FROM ventas_por_producto
        WHERE Estatus <> 'Anulado' AND Fecha BETWEEN %s AND %s
        GROUP BY Marca ORDER BY venta DESC
    """, (str(desde), str(hasta)))  # Fecha es texto 'YYYY-MM-DD' en esta tabla
    filas = [{"marca": marca or "Sin marca", "venta": float(venta or 0), "costo": float(costo or 0)} for marca, venta, costo in cursor.fetchall()]
    cursor.close()
    conn.close()
    total = sum(f["venta"] for f in filas)
    for f in filas:
        f["pct"] = f["venta"] / total * 100 if total else 0
        f["margen"] = (f["venta"] - f["costo"]) / f["venta"] * 100 if f["venta"] else None
    top = filas[:TOP_MARCAS]
    venta_top, costo_top = sum(f["venta"] for f in top), sum(f["costo"] for f in top)
    return {"filas": top, "venta": venta_top, "pct": venta_top / total * 100 if total else 0,
            "margen": (venta_top - costo_top) / venta_top * 100 if venta_top else None}


def meses_entre(desde, hasta):
    meses, anio, mes = [], desde.year, desde.month
    while (anio, mes) <= (hasta.year, hasta.month):
        meses.append(f"{anio}-{mes:02d}")
        anio, mes = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
    return meses


def variacion(actual, anterior):
    return (actual / anterior - 1) * 100 if anterior else None


def calcular(desde, hasta):
    corte = min(hasta, ultima_venta(desde, hasta) or hasta)
    desde_ant, corte_ant = un_anio_antes(desde), un_anio_antes(corte)
    actual, anterior = indicadores(desde, corte), indicadores(desde_ant, corte_ant)

    meses = meses_entre(desde, corte)
    por_mes = venta_neta_por_mes(desde, corte)
    por_mes_ant = venta_neta_por_mes(desde_ant, corte_ant)
    recuperado = recuperacion_por_mes(desde, corte)
    # El mes del corte queda incompleto si el corte no es fin de mes
    incompleto = (corte + timedelta(days=1)).month == corte.month
    varios_anios = desde.year != corte.year
    etiquetas = [MESES[int(m[5:]) - 1] + (f" {m[2:4]}" if varios_anios else "") for m in meses]
    ventas = [float(por_mes.get(m) or 0) for m in meses]
    ventas_ant = [float(por_mes_ant.get(f"{int(m[:4]) - 1}{m[4:]}") or 0) for m in meses]
    recuperacion = [recuperado.get(m, 0.0) for m in meses]

    completos = [(m, v) for m, v in zip(etiquetas, ventas) if v][:-1 if incompleto else None]
    con_recuperacion = [(m, v) for m, v in zip(etiquetas, recuperacion) if v]
    tipos, total_tipos = ventas_por_tipo(desde, corte)
    return {
        "corte": corte, "actual": actual, "anterior": anterior, "anio_ant": corte_ant.year,
        "var_venta": variacion(actual["neto"], anterior["neto"]),
        "var_utilidad": variacion(actual["utilidad"], anterior["utilidad"]),
        "var_clientes": variacion(actual["clientes"], anterior["clientes"]),
        "dif_margen": actual["margen"] - anterior["margen"] if actual["margen"] is not None and anterior["margen"] is not None else None,
        "pct_credito": actual["credito"] / actual["neto"] * 100 if actual["neto"] else None,
        "total_recuperado": sum(recuperacion),
        "mejor_mes": max(completos, key=lambda x: x[1], default=None),
        "peor_mes": min(completos, key=lambda x: x[1], default=None),
        "mejor_recuperacion": max(con_recuperacion, key=lambda x: x[1], default=None),
        "etiquetas": etiquetas, "ventas": ventas, "ventas_ant": ventas_ant, "recuperacion": recuperacion,
        "mes_incompleto": incompleto, "tipos": tipos, "total_tipos": total_tipos, "marcas": marcas_top(desde, corte),
    }


# -----------------------------
# Ruta principal del resumen
# -----------------------------
@resumen_bp.route("/resumen", methods=["GET"])
@requiere_acceso("resumen")
def resumen_dashboard():
    hoy = date.today()
    desde = leer_fecha(request.args.get("fecha_inicio"), hoy.replace(month=1, day=1))
    hasta = leer_fecha(request.args.get("fecha_fin"), hoy)
    r = con_cache(f"resumen_{desde}_{hasta}", lambda: calcular(desde, hasta))
    return render_template("resumen_dashboard.html", fecha_inicio=desde, fecha_fin=hasta, r=r, millones=millones,
                           desde_txt=fecha_larga(desde), corte_txt=fecha_larga(r["corte"]))
