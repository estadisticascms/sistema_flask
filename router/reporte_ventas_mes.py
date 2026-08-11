from decimal import Decimal
from datetime import datetime
from flask import Blueprint, render_template, request
from utils.db import get_connection
from router.accesos import requiere_acceso

# Definir el Blueprint
reporte_ventas_mes_bp = Blueprint("reporte_ventas_mes", __name__)

def obtener_reporte_ventas_mes(fecha_inicio, fecha_fin):
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
			ELSE r.Vendedor END AS Vendedor,
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

    chart_labels = sorted(meses)

    return {
        "grupos": grupos,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin,
        "chart_labels": chart_labels
    }

# Ruta dentro del Blueprint
@reporte_ventas_mes_bp.route("/reporte_ventas_mes", methods=["GET"])
@requiere_acceso("reporte_ventas_mes")
def reporte_ventas_mes():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-%d"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    datos = obtener_reporte_ventas_mes(fecha_inicio, fecha_fin)
    return render_template("reporte_ventas_mes.html", **datos)







