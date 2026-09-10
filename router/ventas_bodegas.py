from flask import Blueprint, render_template, request
from utils.db import get_connection
from datetime import datetime
from router.accesos import requiere_acceso

ventas_bodega_bp = Blueprint("ventas_bodega", __name__)

def obtener_ventas_bodega(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    # --- Detalle diario ---
    query_dias = """
        SELECT 
            Bodega,
            Marca,
            DATE(Fecha) AS Dia,
            SUM(`Cantidad Vendida`) AS UnidadesVendidas
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s
        GROUP BY Bodega, Marca, Dia
        ORDER BY Bodega, Marca, Dia;
    """
    cursor.execute(query_dias, (fecha_inicio, fecha_fin))
    registros_dias = cursor.fetchall()

    # Organizar datos diarios
    data_dias, fechas = {}, set()
    for r in registros_dias:
        bodega, marca, dia, cantidad = r["Bodega"], r["Marca"], r["Dia"].strftime("%d/%m/%Y"), r["UnidadesVendidas"]
        fechas.add(dia)
        data_dias.setdefault(bodega, {}).setdefault(marca, {})[dia] = cantidad
    fechas = sorted(fechas)

    # Totales diarios por bodega
    totales_dias = {}
    for bodega, marcas in data_dias.items():
        totales_dias[bodega] = {fecha: sum(marcas.get(m, {}).get(fecha, 0) for m in marcas) for fecha in fechas}
        totales_dias[bodega]["Total"] = sum(totales_dias[bodega].values())

    # --- Resumen semanal ---
    query_semanal = """
        SELECT 
            TIMESTAMPDIFF(WEEK, DATE_SUB(Fecha, INTERVAL DAY(Fecha)-1 DAY), Fecha) + 1 AS SemanaMes,
            Bodega,
            Marca,
            SUM(`Cantidad Vendida`) AS UnidadesVendidas
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s
        GROUP BY SemanaMes, Bodega, Marca
        ORDER BY Bodega, Marca, SemanaMes;
    """
    cursor.execute(query_semanal, (fecha_inicio, fecha_fin))
    registros_sem = cursor.fetchall()

    cursor.close()
    conn.close()

    # Organizar datos semanales
    data_sem, semanas = {}, set()
    for r in registros_sem:
        bodega, marca, semana, cantidad = r["Bodega"], r["Marca"], f"Sem {r['SemanaMes']}", r["UnidadesVendidas"]
        semanas.add(semana)
        data_sem.setdefault(bodega, {}).setdefault(marca, {})[semana] = cantidad
    semanas = sorted(semanas)

    # Totales semanales por bodega
    totales_sem = {}
    for bodega, marcas in data_sem.items():
        totales_sem[bodega] = {semana: sum(marcas.get(m, {}).get(semana, 0) for m in marcas) for semana in semanas}
        totales_sem[bodega]["Total"] = sum(totales_sem[bodega].values())

    return {"data_dias": data_dias, "fechas": fechas, "totales_dias": totales_dias,
            "data_sem": data_sem, "semanas": semanas, "totales_sem": totales_sem}

@ventas_bodega_bp.route("/ventas_bodega", methods=["GET", "POST"])
@requiere_acceso("ventas_bodega")
def reporte_ventas_bodega():
    if request.method == "POST":
        fecha_inicio = request.form.get("fecha_inicio", datetime.today().strftime("%Y-%m-01"))
        fecha_fin = request.form.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    else:
        fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-01"))
        fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))

    datos = obtener_ventas_bodega(fecha_inicio, fecha_fin)

    return render_template("reporte_ventas_bodega.html",
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           **datos)


