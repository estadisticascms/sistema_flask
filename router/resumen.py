from flask import Blueprint, render_template, request
from datetime import datetime
from decimal import Decimal
from calendar import monthrange
from utils.db import get_connection
from router.reporte_top_marcas import obtener_top_marcas
from router.accesos import requiere_acceso

resumen_bp = Blueprint("resumen", __name__)

# -----------------------------
# Ventas mensuales por vendedores
# -----------------------------
def obtener_resumen_ventas_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT Mes,
               e.tipo AS Tipo_Vendedor,
               datos.Vendedor,
               SUM(datos.Venta_Neta) - SUM(datos.DevTotal) AS Neto_Ajustado
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes,
                   CASE 
                       WHEN v.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJERCITO DE NICARAGUA'
                       ELSE v.Vendedor
                   END AS Vendedor,
                   SUM(v.neto) AS Venta_Neta,
                   0 AS DevTotal
            FROM ventas_historicas v
            WHERE v.fecha BETWEEN %s AND %s
            GROUP BY Mes, v.vendedor, v.cliente

            UNION ALL

            SELECT DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes,
                   CASE 
                       WHEN r.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJERCITO DE NICARAGUA'
                       ELSE r.Vendedor
                   END AS Vendedor,
                   0 AS Venta_Neta,
                   SUM(r.totaldv)/1.15 AS DevTotal
            FROM devoluciones r
            WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado'
            GROUP BY Mes, r.vendedor, r.cliente
        ) datos
        LEFT JOIN (
            SELECT tipo, nombre FROM cat_vendedor GROUP BY tipo, nombre
        ) e ON datos.Vendedor = e.nombre
        GROUP BY e.tipo, datos.Mes, datos.Vendedor
        ORDER BY datos.Mes, e.tipo, datos.Vendedor;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    grupos = {}
    meses = set()
    for r in rows:
        tipo = r["Tipo_Vendedor"] or "SIN TIPO"
        vendedor = r["Vendedor"]
        mes = r["Mes"]
        neto = Decimal(r["Neto_Ajustado"] or 0)

        if tipo not in grupos:
            grupos[tipo] = {}
        if vendedor not in grupos[tipo]:
            grupos[tipo][vendedor] = {}

        grupos[tipo][vendedor][mes] = grupos[tipo][vendedor].get(mes, Decimal(0)) + neto
        meses.add(mes)

    # Convertimos a float
    grupos_float = {
        tipo: {v: {m: float(n) for m, n in ventas.items()} for v, ventas in vendedores.items()}
        for tipo, vendedores in grupos.items()
    }

    chart_labels = sorted(meses)
    totales_por_mes = {mes: sum(grupos_float[t][v].get(mes, 0) for t in grupos_float for v in grupos_float[t]) for mes in chart_labels}
    chart_values = [totales_por_mes[m] for m in chart_labels]

    return {
        "grupos": grupos_float,
        "totales_por_mes": totales_por_mes,
        "chart_labels": chart_labels,
        "chart_values": chart_values,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin
    }

# -----------------------------
# Recuperación mensual desde recibos
# -----------------------------
def obtener_recuperacion(fecha_inicio, fecha_fin, tipo_cambio=36.6243):
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

    for r in rows:
        subtotal_cordobas = float(r["TotalCordobas"]) / 1.15
        iva_cordobas = subtotal_cordobas * 0.15
        total_cordobas = subtotal_cordobas + iva_cordobas
        subtotal_usd = subtotal_cordobas / tipo_cambio

        r["SubtotalCordobas"] = subtotal_cordobas
        r["IVACordobas"] = iva_cordobas
        r["TotalCordobasCalc"] = total_cordobas
        r["SubtotalUSD"] = subtotal_usd

    labels = [r["Mes"] for r in rows]
    valores = [r["TotalCordobasCalc"] for r in rows]

    return {
        "rows": rows,
        "chart_labels": labels,
        "chart_values": valores
    }

# -----------------------------
# Datasets absolutos por tipo de vendedor
# -----------------------------
def preparar_datasets_vendedores(datos_ventas_mes):
    datasets = []
    colores = ["#FE6412","#027266","#666666","#4B9CD3","#9C27B0","#FFC107","#E91E63","#795548","#00BCD4","#8BC34A"]
    tipos = list(datos_ventas_mes["grupos"].keys())

    for i, tipo in enumerate(tipos):
        vendedores = datos_ventas_mes["grupos"][tipo]
        data = []
        for mes in datos_ventas_mes["chart_labels"]:
            subtotal = sum(v.get(mes, 0) for v in vendedores.values())
            data.append(subtotal)
        datasets.append({
            "label": tipo,
            "data": data,
            "backgroundColor": colores[i % len(colores)]
        })
    return datasets

# -----------------------------
# Participación mensual por tipo de vendedor (%)
# -----------------------------
def preparar_participacion_vendedores(datos_ventas_mes):
    chart_labels = datos_ventas_mes["chart_labels"]
    grupos = datos_ventas_mes["grupos"]

    participacion = {}
    for mes in chart_labels:
        total_mes = sum(
            sum(v.get(mes, 0) for v in vendedores.values())
            for vendedores in grupos.values()
        )
        for tipo, vendedores in grupos.items():
            subtotal = sum(v.get(mes, 0) for v in vendedores.values())
            porcentaje = (subtotal / total_mes * 100) if total_mes > 0 else 0
            participacion.setdefault(tipo, []).append(round(porcentaje, 2))

    return participacion

# -----------------------------
# Ruta principal del resumen
# -----------------------------
@resumen_bp.route("/resumen", methods=["GET"])
@requiere_acceso("resumen")
def resumen_dashboard():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-01"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))

    datos_ventas_mes = obtener_resumen_ventas_mes(fecha_inicio, fecha_fin)
    datos_recuperacion = obtener_recuperacion(fecha_inicio, fecha_fin)
    datasets_vendedores = preparar_datasets_vendedores(datos_ventas_mes)    
    participacion_vendedores = preparar_participacion_vendedores(datos_ventas_mes)

    # Último mes del rango para Top Marcas
    ultimo_mes = max(datos_ventas_mes["chart_labels"]) if datos_ventas_mes["chart_labels"] else None

    if ultimo_mes:
        year, month = map(int, ultimo_mes.split("-"))
        fecha_inicio_mes = f"{year}-{month:02d}-01"
        last_day = monthrange(year, month)[1]
        fecha_fin_mes = f"{year}-{month:02d}-{last_day:02d}"
        datos_marcas = obtener_top_marcas(fecha_inicio_mes, fecha_fin_mes)
    else:
        datos_marcas = {"top_detalle": [], "top_simple": []}

    return render_template("resumen_dashboard.html",
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           datos_ventas_mes=datos_ventas_mes,
                           datos_marcas=datos_marcas,
                           datos_recuperacion=datos_recuperacion,
                           datasets_vendedores=datasets_vendedores,
                           participacion_vendedores=participacion_vendedores)









