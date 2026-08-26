from flask import Blueprint, render_template, request
from utils.db import get_connection
from router.accesos import requiere_acceso
import datetime

productos_agotarse_bp = Blueprint("productos_agotarse", __name__)

def obtener_productos_agotarse(filtro_cobertura=4, fecha_inicio="2026-01-01", fecha_fin=None):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if not fecha_fin:
        fecha_fin = datetime.date.today().strftime("%Y-%m-%d")

    query = """
    WITH ingresos AS (
        SELECT 
            codigo_producto,
            fecha,
            ROW_NUMBER() OVER (PARTITION BY codigo_producto ORDER BY fecha DESC) AS rn
        FROM transacciones_inventario
        WHERE tipo_transaccion IN ('ENTRADA','AJUSTE ENTRADA')
    ),
    ventas_mensuales AS (
        SELECT 
            t.codigo_producto,
            AVG(mensual.cant_mes) AS PromedioMensual,
            MAX(t.fecha) AS UltimaVenta
        FROM (
            SELECT 
                t.codigo_producto,
                DATE_FORMAT(t.fecha, '%Y-%m') AS Mes,
                SUM(CASE WHEN t.tipo_transaccion IN ('SALIDA','AJUSTE SALIDA') 
                         THEN t.cantidad ELSE 0 END) AS cant_mes
            FROM transacciones_inventario t
            WHERE t.tipo_transaccion IN ('SALIDA','AJUSTE SALIDA')
            GROUP BY t.codigo_producto, DATE_FORMAT(t.fecha, '%Y-%m')
        ) mensual
        JOIN transacciones_inventario t 
          ON mensual.codigo_producto = t.codigo_producto
        WHERE t.tipo_transaccion IN ('SALIDA','AJUSTE SALIDA')
        GROUP BY t.codigo_producto
    )
    SELECT 
        p.`Codigo Producto` AS Codigo,
        p.`numero Parte` AS NumeroParte,
        REPLACE(REPLACE(REPLACE(p.nombre, '"',''),',',''),"'",'') AS Descripcion,
        p.marca AS Marca,
        p.`existencia actual` AS ExistenciaActual,
        v.PromedioMensual,
        ROUND(p.`existencia actual` / v.PromedioMensual, 1) AS MesesCobertura,
        i1.fecha AS UltimaFechaIngreso,
        v.UltimaVenta AS UltimaFechaVenta,
        TIMESTAMPDIFF(MONTH, i1.fecha, v.UltimaVenta) AS MesesTranscurridos
    FROM catalogo_productos p
    JOIN ventas_mensuales v ON p.`Codigo Producto` = v.codigo_producto
    LEFT JOIN (SELECT codigo_producto, fecha FROM ingresos WHERE rn = 1) i1 
           ON p.`Codigo Producto` = i1.codigo_producto
    WHERE v.PromedioMensual > 0 and p.marca !='CL'
      AND (p.`existencia actual` / v.PromedioMensual) <= %s
      AND p.`existencia actual` >= 0
      AND i1.fecha BETWEEN %s AND %s
    ORDER BY MesesCobertura,p.marca, p.`Codigo Producto` ASC;
    """

    cursor.execute(query, (filtro_cobertura, fecha_inicio, fecha_fin))
    productos = cursor.fetchall()

    cursor.close()
    conn.close()
    return productos


@productos_agotarse_bp.route("/productos_agotarse")
@requiere_acceso("productos_agotarse")
def reporte_productos_agotarse():
    # valores por defecto
    filtro_cobertura = request.args.get("cobertura", 4)
    fecha_inicio = request.args.get("fecha_inicio", "2026-01-01")
    fecha_fin = request.args.get("fecha_fin", datetime.date.today().strftime("%Y-%m-%d"))

    productos = obtener_productos_agotarse(filtro_cobertura, fecha_inicio, fecha_fin)
    return render_template("reporte_productos_agotarse.html", 
                           productos=productos,
                           cobertura=filtro_cobertura,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin)

