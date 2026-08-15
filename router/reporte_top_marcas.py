from flask import Blueprint, render_template, request
from utils.db import get_connection
from datetime import datetime
from router.accesos import requiere_acceso

top_marcas_bp = Blueprint("top_marcas", __name__)

def obtener_top_marcas(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    # Todas las marcas con detalle
    query_detalle = """
        SELECT Marca,
               SUM(`Sub-Total 2`) AS VentasTotales,
               SUM(`Cantidad Vendida`) AS UnidadesVendidas,
               SUM(`Sub-Total 2`) AS Ingreso,
               SUM(`Cantidad Vendida` * `Costo Unit.`) AS CostoTotal,
               (SUM(`Sub-Total 2`) - SUM(`Cantidad Vendida` * `Costo Unit.`)) AS Ganancia,
               ROUND(((SUM(`Sub-Total 2`) - SUM(`Cantidad Vendida` * `Costo Unit.`)) / 
                      SUM(`Sub-Total 2`)) * 100, 2) AS MargenPorcentaje
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s  and marca !='MITROX'
        GROUP BY Marca
        ORDER BY VentasTotales DESC;
    """
    cursor.execute(query_detalle, (fecha_inicio, fecha_fin))
    detalle = cursor.fetchall()

    # Totales generales
    query_totales = """
        SELECT 
               SUM(`Sub-Total 2`) AS VentasTotales,
               SUM(`Cantidad Vendida`) AS UnidadesVendidas,
               SUM(`Sub-Total 2`) AS Ingreso,
               SUM(`Cantidad Vendida` * `Costo Unit.`) AS CostoTotal,
               (SUM(`Sub-Total 2`) - SUM(`Cantidad Vendida` * `Costo Unit.`)) AS Ganancia,
               ROUND(((SUM(`Sub-Total 2`) - SUM(`Cantidad Vendida` * `Costo Unit.`)) / 
                      SUM(`Sub-Total 2`)) * 100, 2) AS MargenPorcentaje
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s;
    """
    cursor.execute(query_totales, (fecha_inicio, fecha_fin))
    totales = cursor.fetchone()

    cursor.close()
    conn.close()

    return {"detalle": detalle, "totales": totales}


@top_marcas_bp.route("/top_marcas", methods=["GET", "POST"])
@requiere_acceso("top_marcas")
def reporte_top_marcas():
    if request.method == "POST":
        fecha_inicio = request.form.get("fecha_inicio", datetime.today().strftime("%Y-%m-01"))
        fecha_fin = request.form.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    else:
        fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-01"))
        fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))

    datos = obtener_top_marcas(fecha_inicio, fecha_fin)

    return render_template("reporte_top_marcas.html",
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           **datos)

