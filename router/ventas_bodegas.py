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
            Numero_parte,
            `Descrip Prod.` AS DescripProd,
            Marca,
            DATE(Fecha) AS Dia,
            SUM(`Cantidad Vendida`) AS UnidadesVendidas
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s
        GROUP BY Bodega, Numero_parte, `Descrip Prod.`, Marca, Dia
        ORDER BY Bodega, Numero_parte, `Descrip Prod.`, Marca, Dia;
    """
    cursor.execute(query_dias, (fecha_inicio, fecha_fin))
    registros_dias = cursor.fetchall()

    data_dias, fechas = {}, set()
    for r in registros_dias:
        bodega = r["Bodega"]
        dia = r["Dia"].strftime("%d/%m/%Y")
        fechas.add(dia)

        producto = {
            "Numero_Parte": r["Numero_parte"],
            "DescripProd": r["DescripProd"],
            "Marca": r["Marca"],
            dia: r["UnidadesVendidas"]
        }

        productos = data_dias.setdefault(bodega, [])
        encontrado = next((p for p in productos if p["Numero_Parte"] == producto["Numero_Parte"]), None)
        if encontrado:
            encontrado[dia] = r["UnidadesVendidas"]
        else:
            productos.append(producto)

    fechas = sorted(fechas)

    # Calcular total por producto en diario
    for bodega, productos in data_dias.items():
        for p in productos:
            p["Total"] = sum(p.get(f, 0) for f in fechas)

    # Totales diarios por bodega
    totales_dias = {}
    for bodega, productos in data_dias.items():
        totales_dias[bodega] = {fecha: sum(p.get(fecha, 0) for p in productos) for fecha in fechas}
        totales_dias[bodega]["Total"] = sum(totales_dias[bodega].values())
    # --- Resumen semanal ---
    query_semanal = """
        SELECT 
            TIMESTAMPDIFF(WEEK, DATE_SUB(Fecha, INTERVAL DAY(Fecha)-1 DAY), Fecha) + 1 AS SemanaMes,
            Bodega,
            Numero_parte,
            `Descrip Prod.` AS DescripProd,
            Marca,
            SUM(`Cantidad Vendida`) AS UnidadesVendidas
        FROM ventas_por_producto
        WHERE Estatus != 'Anulado'
          AND Fecha BETWEEN %s AND %s
        GROUP BY SemanaMes, Bodega, Numero_parte, `Descrip Prod.`, Marca
        ORDER BY Bodega, Numero_parte, `Descrip Prod.`, Marca, SemanaMes;
    """
    cursor.execute(query_semanal, (fecha_inicio, fecha_fin))
    registros_sem = cursor.fetchall()

    cursor.close()
    conn.close()

    data_sem, semanas = {}, set()
    for r in registros_sem:
        bodega = r["Bodega"]
        semana = f"Sem {r['SemanaMes']}"
        semanas.add(semana)

        producto = {
            "Numero_Parte": r["Numero_parte"],
            "DescripProd": r["DescripProd"],
            "Marca": r["Marca"],
            semana: r["UnidadesVendidas"]
        }

        productos = data_sem.setdefault(bodega, [])
        encontrado = next((p for p in productos if p["Numero_Parte"] == producto["Numero_Parte"]), None)
        if encontrado:
            encontrado[semana] = r["UnidadesVendidas"]
        else:
            productos.append(producto)

    semanas = sorted(semanas)

    # Calcular total por producto en semanal
    for bodega, productos in data_sem.items():
        for p in productos:
            p["Total"] = sum(p.get(s, 0) for s in semanas)

    # Totales semanales por bodega
    totales_sem = {}
    for bodega, productos in data_sem.items():
        totales_sem[bodega] = {semana: sum(p.get(semana, 0) for p in productos) for semana in semanas}
        totales_sem[bodega]["Total"] = sum(totales_sem[bodega].values())

    return {
        "data_dias": data_dias, "fechas": fechas, "totales_dias": totales_dias,
        "data_sem": data_sem, "semanas": semanas, "totales_sem": totales_sem
    }

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
