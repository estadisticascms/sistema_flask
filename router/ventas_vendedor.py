"""Datos del reporte de ventas para vendedores (/rep_vend_totales): solo ventas, sin márgenes."""
from datetime import date, timedelta
from utils.db import get_connection

EJERCITO = "EJERCITO DE NICARAGUA"  # sus ventas se reportan como un "vendedor" aparte


def filtro_vendedor(vendedor, col_vendedor, col_cliente):
    """Condición SQL equivalente a la regla del reporte: el cliente Ejército cuenta como su propio vendedor."""
    if not vendedor:
        return "", []
    if vendedor == EJERCITO:
        return f" AND {col_cliente} = %s", [EJERCITO]
    return f" AND {col_vendedor} = %s AND {col_cliente} <> %s", [vendedor, EJERCITO]


def productos_periodo(fecha_inicio, fecha_fin, vendedor=None):
    """Venta y unidades por producto en el período (sin anuladas), con su segmento."""
    extra, params = filtro_vendedor(vendedor, "v.Vendedor", "COALESCE(v.Cliente, '')")
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute(f"""
        SELECT v.Codigo AS codigo, v.Numero_parte AS numero_parte, v.`Descrip Prod.` AS descripcion,
               v.Marca AS marca, c.segmento, c.familia,
               SUM(v.`Cantidad Vendida`) AS unidades, SUM(v.`Sub-Total 2`) AS venta
        FROM ventas_por_producto v
        LEFT JOIN clase_prod c ON c.codigo = v.Codigo
        WHERE v.Fecha BETWEEN %s AND %s AND v.Estatus <> 'Anulado' {extra}
        GROUP BY v.Codigo, v.Numero_parte, v.`Descrip Prod.`, v.Marca, c.segmento, c.familia
        ORDER BY venta DESC
    """, [str(fecha_inicio), str(fecha_fin)] + params)  # Fecha es texto 'YYYY-MM-DD' en esta tabla
    filas = cursor.fetchall()
    cursor.close()
    conn.close()
    for f in filas:
        f["venta"] = float(f["venta"] or 0)
        f["unidades"] = float(f["unidades"] or 0)
        f["costo"] = 0.0  # este reporte no muestra márgenes
    return filas


def facturas_periodo(fecha_inicio, fecha_fin, vendedor=None):
    extra, params = filtro_vendedor(vendedor, "vendedor", "cliente")
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f"SELECT COUNT(*), COUNT(DISTINCT Fecha) FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s {extra}",
                   [fecha_inicio, fecha_fin] + params)
    facturas, dias = cursor.fetchone()
    cursor.close()
    conn.close()
    return int(facturas or 0), int(dias or 0)


def rangos_rapidos(ultima):
    """Atajos de período a partir del último día con ventas cargadas."""
    inicio_mes = ultima.replace(day=1)
    fin_mes_ant = inicio_mes - timedelta(days=1)
    return [
        ("Último día", ultima, ultima),
        ("Mes en curso", inicio_mes, ultima),
        ("Mes anterior", fin_mes_ant.replace(day=1), fin_mes_ant),
        ("Año", date(ultima.year, 1, 1), ultima),
    ]


def ranking_vendedores(datos):
    """Convierte el resultado de obtener_reporte_ventas en una lista ordenada con crédito/contado netos ajustados."""
    devoluciones = {v: float(d["devoluciones"]) for g in datos["grupos"].values() for v, d in g["vendedores"].items()}
    filas = []
    for nombre, grupo, credito, contado in zip(datos["bar_labels"], datos["bar_grupos"], datos["bar_credito"], datos["bar_contado"]):
        filas.append({"vendedor": nombre, "grupo": grupo or "Sin grupo", "credito": credito, "contado": contado,
                      "total": credito + contado, "devoluciones": devoluciones.get(nombre, 0.0)})
    filas.sort(key=lambda f: -f["total"])
    total = sum(f["total"] for f in filas)
    for i, f in enumerate(filas, start=1):
        f["posicion"] = i
        f["pct"] = f["total"] / total * 100 if total else 0
    return filas


def resumen_grupos(vendedores):
    grupos = {}
    for v in vendedores:
        g = grupos.setdefault(v["grupo"], {"grupo": v["grupo"], "credito": 0.0, "contado": 0.0, "total": 0.0, "vendedores": 0})
        g["credito"] += v["credito"]
        g["contado"] += v["contado"]
        g["total"] += v["total"]
        g["vendedores"] += 1
    total = sum(g["total"] for g in grupos.values())
    filas = sorted(grupos.values(), key=lambda g: -g["total"])
    for g in filas:
        g["pct"] = g["total"] / total * 100 if total else 0
    return filas


def resumen_dia(dia, obtener_reporte):
    """Vista compacta para WhatsApp: venta del día por vendedor junto al acumulado del mes (del 1 al día elegido).
    obtener_reporte es obtener_reporte_ventas de app.py (misma regla de devoluciones y del cliente Ejército)."""
    inicio = dia.replace(day=1)
    del_dia = {v["vendedor"]: v for v in ranking_vendedores(obtener_reporte(dia.isoformat(), dia.isoformat()))}
    del_mes = {v["vendedor"]: v for v in ranking_vendedores(obtener_reporte(inicio.isoformat(), dia.isoformat()))}
    filas = []
    for nombre in set(del_dia) | set(del_mes):
        d, m = del_dia.get(nombre), del_mes.get(nombre)
        filas.append({"vendedor": nombre, "grupo": (d or m)["grupo"],
                      "dia": d["total"] if d else 0.0, "dia_credito": d["credito"] if d else 0.0, "dia_contado": d["contado"] if d else 0.0,
                      "mes": m["total"] if m else 0.0, "pos_mes": m["posicion"] if m else None, "pct_mes": m["pct"] if m else 0.0})
    # Orden: lo vendido hoy; los que no vendieron hoy van al final según su acumulado
    filas.sort(key=lambda f: (-f["dia"], -f["mes"]))
    pos = 0
    for f in filas:
        if f["dia"] > 0:
            pos += 1
            f["pos_dia"] = pos
        else:
            f["pos_dia"] = None
    facturas_dia, _ = facturas_periodo(dia, dia)
    facturas_mes, dias_mes = facturas_periodo(inicio, dia)
    total_dia = sum(f["dia"] for f in filas)
    total_mes = sum(f["mes"] for f in filas)
    promedio = total_mes / dias_mes if dias_mes else None
    con_venta = [f for f in filas if f["dia"] > 0]
    lider_mes = max(filas, key=lambda f: f["mes"], default=None)
    productos = productos_periodo(dia, dia)
    return {
        "productos": productos[:5], "total_productos": len(productos), "venta_productos": sum(p["venta"] for p in productos),
        "filas": filas, "inicio": inicio,
        "total_dia": total_dia, "credito_dia": sum(f["dia_credito"] for f in filas), "contado_dia": sum(f["dia_contado"] for f in filas),
        "total_mes": total_mes, "dias_mes": dias_mes, "facturas_dia": facturas_dia, "facturas_mes": facturas_mes,
        "promedio": promedio, "vs_promedio": (total_dia / promedio - 1) * 100 if promedio and total_dia else None,
        "mejor_dia": con_venta[0] if con_venta else None, "mejor_mes": lider_mes if lider_mes and lider_mes["mes"] > 0 else None,
        "vendieron": len(con_venta),
    }


def dias_vecinos(dia):
    """Día con ventas anterior y siguiente (para las flechas de la vista del día)."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT (SELECT MAX(Fecha) FROM ventas_historicas WHERE Fecha < %s), (SELECT MIN(Fecha) FROM ventas_historicas WHERE Fecha > %s)",
                   (dia, dia))
    anterior, siguiente = cursor.fetchone()
    cursor.close()
    conn.close()
    return anterior, siguiente
