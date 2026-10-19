
from flask import Blueprint, render_template, request
from datetime import datetime, timedelta
from utils.db import get_connection
from router.accesos import requiere_acceso

dashboard_bp = Blueprint("dashboard", __name__)

# -----------------------------
# Venta del día (Neto Ajustado)
# -----------------------------
def obtener_venta_dia(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT SUM(datos.Venta_Neta) - SUM(datos.DevTotal) AS Neto_Ajustado
        FROM (
            SELECT SUM(v.neto) AS Venta_Neta, 0 AS DevTotal
            FROM ventas_historicas v WHERE DATE(v.fecha) = %s
            UNION ALL
            SELECT 0 AS Venta_Neta, SUM(r.totaldv)/1.15 AS DevTotal
            FROM devoluciones r WHERE DATE(r.fechadv) = %s AND r.estado='Cerrado'
        ) datos;
    """
    cursor.execute(query, (fecha_dia, fecha_dia))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row["Neto_Ajustado"] or 0

# -----------------------------
# Recuperación del día
# -----------------------------
def obtener_recuperacion_dia(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = "SELECT SUM(total) AS TotalCordobas FROM reciboscaja WHERE DATE(fecharc) = %s;"
    cursor.execute(query, (fecha_dia,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row["TotalCordobas"] or 0

# -----------------------------
# Devoluciones del día
# -----------------------------
def obtener_devoluciones_dia(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT SUM(r.totaldv)/1.15 AS Devoluciones
        FROM devoluciones r
        WHERE DATE(r.fechadv) = %s AND r.estado='Cerrado';
    """
    cursor.execute(query, (fecha_dia,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row["Devoluciones"] or 0
# -----------------------------
# Ranking de vendedores del día
# -----------------------------
def obtener_ranking_vendedores(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT datos.Vendedor,
               SUM(datos.Credito) - SUM(datos.DevCredito) AS Credito,
               SUM(datos.Contado) - SUM(datos.DevContado) AS Contado,
               (SUM(datos.Credito) - SUM(datos.DevCredito)) + (SUM(datos.Contado) - SUM(datos.DevContado)) AS Neto_Ajustado
        FROM (
            SELECT v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE DATE(v.fecha) = %s GROUP BY v.Vendedor
            UNION ALL
            SELECT r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE DATE(r.fechadv) = %s AND r.estado='Cerrado' GROUP BY r.Vendedor
        ) datos
        GROUP BY datos.Vendedor
        ORDER BY Neto_Ajustado DESC;
    """
    cursor.execute(query, (fecha_dia, fecha_dia))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

# -----------------------------
# Productos del día (una sola lectura de ventas_por_producto para top 10 y marcas)
# -----------------------------
def obtener_productos_dia(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT v.Codigo AS codigo, v.Numero_parte AS numero_parte, v.`Descrip Prod.` AS descripcion,
               v.Marca AS marca, c.segmento, c.familia,
               SUM(v.`Cantidad Vendida`) AS unidades,
               SUM(v.`Sub-Total 2`) AS venta,
               SUM(v.`Cantidad Vendida` * v.`Costo Unit.`) AS costo,
               COUNT(DISTINCT v.`No. Factura`) AS facturas
        FROM ventas_por_producto v
        LEFT JOIN clase_prod c ON c.codigo = v.Codigo
        WHERE v.Fecha = %s AND v.Estatus <> 'Anulado'
        GROUP BY v.Codigo, v.Numero_parte, v.`Descrip Prod.`, v.Marca, c.segmento, c.familia
        ORDER BY venta DESC
    """
    cursor.execute(query, (str(fecha_dia),))  # Fecha es texto 'YYYY-MM-DD' en esta tabla
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        r["venta"] = float(r["venta"] or 0)
        r["costo"] = float(r["costo"] or 0)
        r["unidades"] = float(r["unidades"] or 0)
        r["margen"] = (r["venta"] - r["costo"]) / r["venta"] * 100 if r["venta"] else None
    return rows


def ranking_marcas(productos, limite=10):
    """Agrupa los productos del día por marca; el resto se pliega en 'Otras marcas'."""
    marcas = {}
    for p in productos:
        m = marcas.setdefault(p["marca"] or "Sin marca", {"Marca": p["marca"] or "Sin marca", "VentasTotales": 0.0,
                                                          "Unidades": 0.0, "Costo": 0.0, "Productos": 0})
        m["VentasTotales"] += p["venta"]
        m["Unidades"] += p["unidades"]
        m["Costo"] += p["costo"]
        m["Productos"] += 1
    filas = sorted(marcas.values(), key=lambda m: -m["VentasTotales"])
    for m in filas:
        m["Margen"] = (m["VentasTotales"] - m["Costo"]) / m["VentasTotales"] * 100 if m["VentasTotales"] else None
    top, resto = filas[:limite], filas[limite:]
    otras = None
    if resto:
        venta = sum(m["VentasTotales"] for m in resto)
        costo = sum(m["Costo"] for m in resto)
        otras = {"Marca": f"Otras {len(resto)} marcas", "VentasTotales": venta, "Costo": costo,
                 "Unidades": sum(m["Unidades"] for m in resto), "Productos": sum(m["Productos"] for m in resto),
                 "Margen": (venta - costo) / venta * 100 if venta else None}
    return filas, top, otras


# -----------------------------
# Indicadores del día y acumulado del mes
# -----------------------------
def obtener_indicadores(fecha_dia):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT COUNT(*) AS facturas, COALESCE(SUM(neto), 0) AS neto, COALESCE(SUM(utilidad), 0) AS utilidad,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Credito' THEN neto END), 0) AS credito,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Contado' THEN neto END), 0) AS contado
        FROM ventas_historicas WHERE Fecha = %s
    """, (fecha_dia,))
    dia = cursor.fetchone()
    inicio_mes = fecha_dia.replace(day=1)
    cursor.execute("""
        SELECT COALESCE(SUM(neto), 0) AS neto, COUNT(DISTINCT Fecha) AS dias
        FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s
    """, (inicio_mes, fecha_dia))
    mes = cursor.fetchone()
    # Devoluciones cerradas (sin IVA), igual que en la venta neta del día y en el reporte mensual
    cursor.execute("""
        SELECT COALESCE(SUM(CASE WHEN fechadv = %s THEN totaldv END), 0) / 1.15 AS dia,
               COALESCE(SUM(totaldv), 0) / 1.15 AS mes
        FROM devoluciones WHERE fechadv BETWEEN %s AND %s AND estado = 'Cerrado'
    """, (fecha_dia, inicio_mes, fecha_dia))
    dev = cursor.fetchone()
    cursor.close()
    conn.close()

    dia = {k: float(v or 0) for k, v in dia.items()}
    dia["margen"] = dia["utilidad"] / dia["neto"] * 100 if dia["neto"] else None
    dia["ticket"] = dia["neto"] / dia["facturas"] if dia["facturas"] else None
    dev_dia, dev_mes = float(dev["dia"] or 0), float(dev["mes"] or 0)
    facturado_mes, dias = float(mes["neto"] or 0), int(mes["dias"] or 0)
    mes_neto = facturado_mes - dev_mes          # acumulado del mes neto de devoluciones
    dia_neto = dia["neto"] - dev_dia            # venta del día neta de devoluciones
    promedio = mes_neto / dias if dias else None
    # Comparar contra el promedio de los OTROS días del mes con venta (todo neto de devoluciones)
    promedio_otros = (mes_neto - dia_neto) / (dias - 1) if dias > 1 and dia["neto"] else None
    variacion = (dia_neto / promedio_otros - 1) * 100 if promedio_otros else None
    return dia, {"neto": mes_neto, "facturado": facturado_mes, "devoluciones": dev_mes, "dias": dias,
                 "promedio": promedio, "inicio": inicio_mes}, variacion


def obtener_resumen_mensual(fecha_inicio, fecha_fin, credito, contado):
    """Una fila por mes: crédito/contado netos de devoluciones (del reporte mensual) + días, facturas y margen."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT DATE_FORMAT(Fecha, '%Y-%m') AS Mes, COUNT(DISTINCT Fecha) AS dias, COUNT(*) AS facturas,
               COALESCE(SUM(neto), 0) AS neto, COALESCE(SUM(utilidad), 0) AS utilidad, MAX(Fecha) AS ultima
        FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s
        GROUP BY Mes ORDER BY Mes
    """, (fecha_inicio, fecha_fin))
    datos = {r["Mes"]: r for r in cursor.fetchall()}
    cursor.close()
    conn.close()

    filas, anterior = [], None
    for mes in sorted(set(credito) | set(contado) | set(datos)):
        d = datos.get(mes, {})
        cr, co = float(credito.get(mes, 0) or 0), float(contado.get(mes, 0) or 0)
        total, dias = cr + co, int(d.get("dias") or 0)
        promedio = total / dias if dias else None
        ultima = d.get("ultima")
        siguiente_dia = ultima + timedelta(days=1) if ultima else None
        filas.append({
            "mes": mes, "credito": cr, "contado": co, "total": total, "dias": dias,
            "facturas": int(d.get("facturas") or 0),
            "pct_credito": cr / total * 100 if total else None,
            "promedio": promedio,
            # Se compara el promedio diario: así el mes en curso (incompleto) es comparable
            "variacion": (promedio / anterior - 1) * 100 if promedio and anterior else None,
            "margen": float(d["utilidad"]) / float(d["neto"]) * 100 if d.get("neto") else None,
            "parcial_al": ultima.day if ultima and siguiente_dia.month == ultima.month
                          and ultima == datetime.strptime(fecha_fin, "%Y-%m-%d").date() else None,
        })
        anterior = promedio
    return filas


def ultima_fecha_con_ventas():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas")
    fecha = cursor.fetchone()[0]
    cursor.close()
    conn.close()
    return fecha
# -----------------------------
# Ventas mensuales solo Totales
# -----------------------------
def obtener_totales_vendedor_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT Mes, Vendedor,
               (SUM(Credito) - SUM(DevCredito)) + (SUM(Contado) - SUM(DevContado)) AS Total
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes, v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE v.fecha BETWEEN %s AND %s GROUP BY Mes, v.Vendedor
            UNION ALL
            SELECT DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes, r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado' GROUP BY Mes, r.Vendedor
        ) datos
        GROUP BY Mes, Vendedor
        ORDER BY Vendedor, Mes;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    ventas_por_vendedor = {}
    meses = sorted(set(r["Mes"] for r in rows))
    for r in rows:
        vendedor = r["Vendedor"]
        mes = r["Mes"]
        if vendedor not in ventas_por_vendedor:
            ventas_por_vendedor[vendedor] = {}
        ventas_por_vendedor[vendedor][mes] = r["Total"]

    vendedores_ordenados = sorted(
        ventas_por_vendedor.keys(),
        key=lambda v: sum(ventas_por_vendedor[v].values()),
        reverse=True
    )
    return ventas_por_vendedor, meses, vendedores_ordenados

# -----------------------------
# Ventas mensuales con Crédito/Contado
# -----------------------------
def obtener_ventas_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT Mes, Vendedor,
               SUM(Credito) - SUM(DevCredito) AS Credito,
               SUM(Contado) - SUM(DevContado) AS Contado
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes, v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE v.fecha BETWEEN %s AND %s GROUP BY Mes, v.Vendedor
            UNION ALL
            SELECT DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes, r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado' GROUP BY Mes, r.Vendedor
        ) datos
        GROUP BY Mes, Vendedor
        ORDER BY Mes, Vendedor;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    ventas_por_vendedor = {}
    meses = sorted(set(r["Mes"] for r in rows))
    for r in rows:
        vendedor = r["Vendedor"]
        mes = r["Mes"]
        if vendedor not in ventas_por_vendedor:
            ventas_por_vendedor[vendedor] = {}
        ventas_por_vendedor[vendedor][mes] = {
            "Credito": r["Credito"],
            "Contado": r["Contado"]
        }
    return ventas_por_vendedor, meses


# -----------------------------
# Recuperación mensual
# -----------------------------
def obtener_recuperacion_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT DATE_FORMAT(fecharc, '%Y-%m') AS Mes,
               COUNT(numerorc) AS Cantidad,
               SUM(total) AS TotalCordobas
        FROM reciboscaja
        WHERE fecharc BETWEEN %s AND %s
        GROUP BY Mes
        ORDER BY Mes;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

# -----------------------------
# Ruta principal del dashboard
# -----------------------------
@dashboard_bp.route("/dashboard", methods=["GET"])
@requiere_acceso("dashboard")
def dashboard_view():
    # Por defecto: el último día con ventas cargadas (no "hoy", que suele estar vacío hasta la carga)
    ultima = ultima_fecha_con_ventas()
    defecto = ultima or datetime.today().date()
    try:
        fecha_dia = datetime.strptime(request.args.get("fecha_fin") or str(defecto), "%Y-%m-%d").date()
    except ValueError:
        fecha_dia = defecto
    fecha_fin = fecha_dia.strftime("%Y-%m-%d")
    fecha_inicio = request.args.get("fecha_inicio") or f"{fecha_dia.year}-01-01"

    # Indicadores del día
    venta_dia = obtener_venta_dia(fecha_fin)
    recuperacion_dia = obtener_recuperacion_dia(fecha_fin)
    devoluciones_dia = obtener_devoluciones_dia(fecha_fin)
    indicadores, mes, variacion = obtener_indicadores(fecha_dia)

    # Rankings del día
    ranking_vendedores_dia = obtener_ranking_vendedores(fecha_fin)
    productos_dia = obtener_productos_dia(fecha_dia)
    top_productos = productos_dia[:10]
    ranking_marcas_dia, top_marcas, otras_marcas = ranking_marcas(productos_dia)
    venta_productos_dia = sum(p["venta"] for p in productos_dia)

    # Totales del día
    total_credito = sum(r["Credito"] or 0 for r in ranking_vendedores_dia)
    total_contado = sum(r["Contado"] or 0 for r in ranking_vendedores_dia)
    total_neto = sum(r["Neto_Ajustado"] or 0 for r in ranking_vendedores_dia)

    # Reporte mensual con Crédito/Contado
    ventas_detalladas, meses_detallados = obtener_ventas_mes(fecha_inicio, fecha_fin)

    # Reporte mensual solo Totales
    ventas_totales, meses_totales, vendedores_ordenados = obtener_totales_vendedor_mes(fecha_inicio, fecha_fin)

    # Totales generales por mes (para gráficos)
    totales_por_mes_credito = {
        mes: float(sum(ventas_detalladas[v].get(mes, {}).get("Credito", 0) or 0 for v in ventas_detalladas))
        for mes in meses_detallados
    }
    totales_por_mes_contado = {
        mes: float(sum(ventas_detalladas[v].get(mes, {}).get("Contado", 0) or 0 for v in ventas_detalladas))
        for mes in meses_detallados
    }


    # Recuperación mensual
    recuperacion_mes = obtener_recuperacion_mes(fecha_inicio, fecha_fin)
    total_general = sum(totales_por_mes_credito.values()) + sum(totales_por_mes_contado.values())
    resumen_mensual = obtener_resumen_mensual(fecha_inicio, fecha_fin, totales_por_mes_credito, totales_por_mes_contado)

    return render_template("dashboard.html",
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           fecha_dia=fecha_dia,
                           dia_anterior=(fecha_dia - timedelta(days=1)).strftime("%Y-%m-%d"),
                           dia_siguiente=(fecha_dia + timedelta(days=1)).strftime("%Y-%m-%d"),
                           ultima_fecha=ultima,
                           venta_dia=venta_dia,
                           recuperacion_dia=recuperacion_dia,
                           devoluciones_dia=devoluciones_dia,
                           indicadores=indicadores,
                           mes=mes,
                           variacion=variacion,
                           ranking_vendedores_dia=ranking_vendedores_dia,
                           ranking_marcas_dia=ranking_marcas_dia,
                           top_marcas=top_marcas,
                           otras_marcas=otras_marcas,
                           top_productos=top_productos,
                           total_productos_dia=len(productos_dia),
                           venta_productos_dia=venta_productos_dia,
                           total_general=total_general,
                           resumen_mensual=resumen_mensual,
                           total_credito=total_credito,
                           total_contado=total_contado,
                           total_neto=total_neto,
                           ventas_detalladas=ventas_detalladas,
                           meses_detallados=meses_detallados,
                           ventas_totales=ventas_totales,
                           meses_totales=meses_totales,
                           vendedores_ordenados=vendedores_ordenados,
                           totales_por_mes_credito=totales_por_mes_credito,
                           totales_por_mes_contado=totales_por_mes_contado,
                           recuperacion_mes=recuperacion_mes)
