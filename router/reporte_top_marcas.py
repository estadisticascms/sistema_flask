from flask import Blueprint, render_template, request
from utils.db import get_connection
from datetime import datetime
from router.accesos import requiere_acceso

top_marcas_bp = Blueprint("top_marcas", __name__)

def obtener_top_marcas(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    # Top 15 marcas con detalle
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
          AND Fecha BETWEEN %s AND %s 
        GROUP BY Marca
        ORDER BY VentasTotales DESC
        LIMIT 15;
    """
    cursor.execute(query_detalle, (fecha_inicio, fecha_fin))
    top_detalle = cursor.fetchall()

    # Top 15 marcas solo ventas
    query_simple = """
        SELECT Marca,
               SUM(`Sub-Total 2`) AS VentasTotales,
               SUM(`Cantidad Vendida`) AS UnidadesVendidas
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s 
        GROUP BY Marca
        ORDER BY VentasTotales DESC
        LIMIT 15;
    """
    cursor.execute(query_simple, (fecha_inicio, fecha_fin))
    top_simple = cursor.fetchall()

    cursor.close()
    conn.close()

    return {"top_detalle": top_detalle, "top_simple": top_simple}

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

