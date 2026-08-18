
from flask import Flask, render_template, request, redirect, url_for, Response, flash, session
import mysql.connector
from datetime import datetime
from datetime import date
from collections import defaultdict
import csv
import pandas as pd
import os
from werkzeug.utils import secure_filename
from decimal import Decimal
from router.recibos import recibos_bp
from router.devoluciones import devoluciones_bp
from router.reporte_facturas import reporte_facturas_bp
from router.reporte_ventas_mes import reporte_ventas_mes_bp
from router.ventas_zonas_mes import ventas_zonas_mes_bp
from router.reporte_top_marcas import top_marcas_bp
from router.resumen import resumen_bp
from router.accesos import accesos_bp
from router.usuarios import usuarios_bp
from router.login import login_bp
from router.accesos import requiere_acceso
from router.ventas_perdidas import ventas_perdidas_bp


app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
app.secret_key = "clave_secreta_segura"

app.register_blueprint(recibos_bp)
app.register_blueprint(devoluciones_bp)
app.register_blueprint(reporte_facturas_bp)
app.register_blueprint(reporte_ventas_mes_bp) 
app.register_blueprint(ventas_zonas_mes_bp)
app.register_blueprint(top_marcas_bp)
app.register_blueprint(resumen_bp)
app.register_blueprint(accesos_bp)
app.register_blueprint(usuarios_bp)
app.register_blueprint(login_bp)
app.register_blueprint(ventas_perdidas_bp)


# -----------------------------
# Configuración de la base de datos
# -----------------------------
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': 'Lpsmqlg.',
    'database': 'comersa'
}

# Conexión global


conn = mysql.connector.connect(**db_config)

def get_cursor():
    """Devuelve un cursor con resultados como diccionarios"""
    return conn.cursor(dictionary=True)


# -----------------------------
# Función para agrupar datos
# -----------------------------
def agrupar_por_zona(data):
    resultado = []
    zonas = {}

    for row in data:
        zona = row["nombre_zona"]
        cliente = row["Cliente"]

        if zona not in zonas:
            zonas[zona] = {"nombre": zona, "clientes": [], "total": 0}
            resultado.append(zonas[zona])

        zona_obj = zonas[zona]
        zona_obj["total"] += float(row["Total"])

        cliente_obj = next((c for c in zona_obj["clientes"] if c["nombre"] == cliente), None)
        if not cliente_obj:
            cliente_obj = {"nombre": cliente, "filas": [], "total": 0}
            zona_obj["clientes"].append(cliente_obj)

        cliente_obj["filas"].append(row)
        cliente_obj["total"] += float(row["Total"])

    return resultado


# -----------------------------
# Filtros y utilidades
# -----------------------------
@app.template_filter("cordobas")
def cordobas_format(value):
    try:
        return f"C$ {float(value):,.2f}"
    except (ValueError, TypeError):
        return value

def calcular_totales(data):
    """Agrega zona_total y cliente_total a cada fila"""
    zona_totales = defaultdict(float)
    cliente_totales = defaultdict(float)

    for row in data:
        zona = row["nombre_zona"]
        cliente = row["Cliente"]
        total = float(row["Total"]) if row["Total"] is not None else 0.0
        zona_totales[zona] += total
        cliente_totales[(zona, cliente)] += total

    for row in data:
        zona = row["nombre_zona"]
        cliente = row["Cliente"]
        row["zona_total"] = zona_totales[zona]
        row["cliente_total"] = cliente_totales[(zona, cliente)]

    return data

# -----------------------------
# formateo de valores
# -----------------------------

@app.template_filter('cordobas')
def cordobas(value):
    try:
        return "C$ " + "{:,.2f}".format(float(value))
    except:
        return "C$ 0.00"

@app.template_filter('dolares')
def dolares(value):
    try:
        return "$ " + "{:,.2f}".format(float(value))
    except:
        return "$ 0.00"



# -----------------------------
# Funciones de construcción de queries
# -----------------------------
def build_query_detalle(years,order_by="zona_total"):
    select_parts = [
        "d.vendedor",
        "(case when c.direccion is null then 'CASUAL' ELSE z.nombre_zona END) AS nombre_zona",
        "d.Cliente",
        "UPPER(c.municipio) AS Dto_Munic",
        "c.direccion",
        "REPLACE(d.Numero_parte, '\"', '') AS NParte",
        "REPLACE(d.`Descrip Prod.`, '\"', '') AS Descripcion",
        "d.Marca"
    ]

    for year in years:
        select_parts.append(
            f"COALESCE(SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (1,2,3) "
            f"THEN d.`Sub-Total 2` ELSE 0 END),0) AS `1T {year}`"
        )
        select_parts.append(
            f"COALESCE(SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (4,5,6) "
            f"THEN d.`Sub-Total 2` ELSE 0 END),0) AS `2T {year}`"
        )
        select_parts.append(
            f"COALESCE(SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (7,8,9) "
            f"THEN d.`Sub-Total 2` ELSE 0 END),0) AS `3T {year}`"
        )
        select_parts.append(
            f"COALESCE(SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (10,11,12) "
            f"THEN d.`Sub-Total 2` ELSE 0 END),0) AS `4T {year}`"
        )

    select_parts.append("SUM(d.`Sub-Total 2`) AS Total")

       # lógica condicional para el ORDER BY
    if order_by == "zona_total":
        order_clause = "ORDER BY zona_total DESC"
    elif order_by == "Total":
        order_clause = "ORDER BY Total DESC"
    elif order_by == "SUBSTRING(z.grupozona,1,1)":
        order_clause = "ORDER BY SUBSTRING(z.grupozona,1,1) ASC"
    else:
        # valor por defecto
        order_clause = "ORDER BY zona_total DESC"

    return f"""
    SELECT {", ".join(select_parts)},
    SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY z.nombre_zona) AS zona_total,
    SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY z.nombre_zona, d.Cliente) AS cliente_total
    FROM ventas_por_producto d
    LEFT JOIN clientes c ON d.Cliente=c.nombre
    LEFT JOIN zona_cliente z ON c.municipio=z.departamento
    WHERE d.vendedor=%s AND d.fecha BETWEEN %s AND %s AND Estatus!='Anulado'
    GROUP BY d.vendedor, d.Cliente, c.municipio, d.Numero_parte, z.nombre_zona, c.direccion, d.`Descrip Prod.`, d.Marca,z.grupozona
    {order_clause};
    """


def build_query_resumen(years,order_by="zona_total"):
    select_parts = [
        "(case when c.direccion is null then 'CASUAL' ELSE z.nombre_zona END) AS nombre_zona",
        "d.Cliente",
        "UPPER(c.municipio) AS Dto_Munic",
    ]

    for year in years:
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (1,2,3) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `1T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (4,5,6) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `2T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (7,8,9) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `3T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (10,11,12) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `4T {year}`"
        )

    select_parts.append("SUM(d.`Sub-Total 2`) AS Total")

       # columna calculada para zona_total
    select_parts.append(
        "SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY SUBSTRING(z.grupozona,1,1),d.cliente) AS zona_total"
    )
    select_parts.append(
        "SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY d.cliente) AS cliente_total"
    )
   # lógica condicional para el ORDER BY
    if order_by == "zona_total":
        order_clause = "ORDER BY zona_total DESC"
    elif order_by == "Total":
        order_clause = "ORDER BY Total DESC"
    elif order_by == "SUBSTRING(z.grupozona,1,1)":
        order_clause = "ORDER BY SUBSTRING(z.grupozona,1,1) ASC"
    else:
        # valor por defecto
        order_clause = "ORDER BY zona_total DESC"


    return f"""
    SELECT {", ".join(select_parts)}
    FROM ventas_por_producto d
    LEFT JOIN clientes c ON d.Cliente=c.nombre
    LEFT JOIN zona_cliente z ON c.municipio=z.departamento
    WHERE d.vendedor=%s AND d.fecha BETWEEN %s AND %s AND Estatus!='Anulado'
    GROUP BY  SUBSTRING(z.grupozona,1,1),z.nombre_zona,d.Cliente,c.direccion,c.municipio
    {order_clause};
    """


def build_query_totales(years,order_by="zona_total"):
    select_parts = [
        "(case when c.direccion is null then 'CASUAL' ELSE z.nombre_zona END) AS nombre_zona",
        "d.Cliente",
        "UPPER(c.municipio) AS Dto_Munic",
        "d.Marca"
    ]

    for year in years:
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (1,2,3) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `1T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (4,5,6) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `2T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (7,8,9) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `3T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (10,11,12) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `4T {year}`"
        )

    select_parts.append("SUM(d.`Sub-Total 2`) AS Total")
           # columna calculada para zona_total
    select_parts.append(
        "SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY SUBSTRING(z.grupozona,1,1),d.cliente) AS zona_total"
    )
    select_parts.append(
        "SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY d.cliente) AS cliente_total"
    )

           # lógica condicional para el ORDER BY
    if order_by == "zona_total":
        order_clause = "ORDER BY zona_total DESC"
    elif order_by == "Total":
        order_clause = "ORDER BY Total DESC"
    elif order_by == "SUBSTRING(z.grupozona,1,1)":
        order_clause = "ORDER BY SUBSTRING(z.grupozona,1,1) ASC"
    else:
        # valor por defecto
        order_clause = "ORDER BY zona_total DESC"

    return f"""
    SELECT {", ".join(select_parts)}
    FROM ventas_por_producto d
    LEFT JOIN clientes c ON d.Cliente=c.nombre
    LEFT JOIN zona_cliente z ON c.municipio=z.departamento
    WHERE d.vendedor=%s AND d.fecha BETWEEN %s AND %s AND Estatus!='Anulado'
    GROUP BY SUBSTRING(z.grupozona,1,1),z.nombre_zona, d.Cliente,c.direccion,c.municipio, d.Marca
    {order_clause};
    """

def build_query_zonas_marcas(years, order_by="zona_total"):
    select_parts = ["(case when c.direccion is null then 'CASUAL' ELSE z.nombre_zona END) AS nombre_zona", "d.Marca"]
    for year in years:
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (1,2,3) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `1T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (4,5,6) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `2T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (7,8,9) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `3T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (10,11,12) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `4T {year}`"
        )
        select_parts.append("SUM(d.`Sub-Total 2`) AS Total")
            # columna calculada para zona_total

        select_parts.append("SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY z.nombre_zona) AS zona_total")


               # lógica condicional para el ORDER BY
    if order_by == "zona_total":
        order_clause = "ORDER BY zona_total DESC, Total desc"
    elif order_by == "Total":
        order_clause = "ORDER BY Total DESC"
    elif order_by == "SUBSTRING(z.grupozona,1,1)":
        order_clause = "ORDER BY SUBSTRING(z.grupozona,1,1) ASC, Total desc"
    else:
        # valor por defecto
        order_clause = "ORDER BY zona_total DESC"

    return f"""
    SELECT {", ".join(select_parts)}
    FROM ventas_por_producto d
    LEFT JOIN clientes c ON d.Cliente=c.nombre
    LEFT JOIN zona_cliente z ON c.municipio=z.departamento
    WHERE d.vendedor=%s AND d.fecha BETWEEN %s AND %s AND Estatus!='Anulado'
    GROUP BY SUBSTRING(z.grupozona,1,1), z.nombre_zona, c.direccion,d.Marca
     {order_clause};
    """



def build_query_zonas_totales(years, order_by="zona_total"):
    select_parts = ["(case when c.direccion is null then 'CASUAL' ELSE z.nombre_zona END) AS nombre_zona"]
    for year in years:
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (1,2,3) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `1T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (4,5,6) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `2T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (7,8,9) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `3T {year}`"
        )
        select_parts.append(
            f"SUM(CASE WHEN YEAR(d.fecha)={year} AND MONTH(d.fecha) IN (10,11,12) "
            f"THEN d.`Sub-Total 2` ELSE 0 END) AS `4T {year}`"
        )
    select_parts.append("SUM(d.`Sub-Total 2`) AS Total")

    return f"""
    SELECT {", ".join(select_parts)},
           SUM(SUM(d.`Sub-Total 2`)) OVER (PARTITION BY SUBSTRING(z.grupozona,1,1)) AS zona_total
    FROM ventas_por_producto d
    LEFT JOIN clientes c ON d.Cliente=c.nombre
    LEFT JOIN zona_cliente z ON c.municipio=z.departamento
    WHERE d.vendedor=%s AND d.fecha BETWEEN %s AND %s AND Estatus!='Anulado'
    GROUP BY d.Cliente,c.direccion,SUBSTRING(z.grupozona,1,1), z.nombre_zona
    ORDER BY SUBSTRING(z.grupozona,1,1) ASC, {order_by} DESC, Total DESC;
    """

# -----------------------------
# Menu de reportes
# -----------------------------
@app.route("/reportes")
@requiere_acceso("reportes")
def reportes():
    return render_template("reportes.html")
# -----------------------------
# Reporte de ventas
# -----------------------------
from collections import defaultdict
from decimal import Decimal
from datetime import datetime
from flask import render_template, request


def obtener_reporte_ventas(fecha_inicio, fecha_fin):
    
    cursor = get_cursor()

    query = """
        SELECT Mes,
              e.tipo AS Tipo_Vendedor,
               datos.Vendedor,
               datos.Venta,
               SUM(datos.Subtotal) AS Subtotal,
               SUM(datos.Descuento) AS Descuento,
               SUM(datos.Venta_Neta) AS Venta_Neta,
               SUM(datos.Iva) AS Iva,
               SUM(datos.Total) AS Total,
               SUM(datos.DevTotal) AS DevTotal
        FROM (
            -- Ventas
            SELECT 
                DATE_FORMAT(v.fecha, '%Y-%m') AS Mes,
                CASE 
					WHEN v.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJERCITO DE NICARAGUA'
					ELSE v.Vendedor
				END AS Vendedor,
                v.tipo_venta AS Venta,
                SUM(v.SubTotal) AS Subtotal,
                SUM(v.Descuento) AS Descuento,
                SUM(v.neto) AS Venta_Neta,
                SUM(v.iva) AS Iva,
                SUM(v.total) AS Total,
                0 AS DevTotal
            FROM ventas_historicas v
            WHERE v.fecha BETWEEN %s AND %s
            GROUP BY Mes, v.vendedor, v.cliente, v.tipo_venta

            UNION ALL

            -- Devoluciones
            SELECT
                DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes,
                CASE 
                    WHEN r.cliente = 'EJERCITO DE NICARAGUA' THEN 'EJERCITO DE NICARAGUA'
                ELSE r.Vendedor END AS Vendedor,
                r.tipofact AS Venta,
                0 AS Subtotal,
                0 AS Descuento,
                0 AS Venta_Neta,
                0 AS Iva,
                0 AS Total,
                SUM(r.totaldv)/1.15 AS DevTotal
            FROM devoluciones r
            WHERE r.fechadv BETWEEN %s AND %s and r.estado='Cerrado'
            GROUP BY Mes, r.vendedor, r.cliente, r.tipofact
        ) datos
        LEFT JOIN (
            SELECT tipo, nombre FROM cat_vendedor GROUP BY tipo, nombre
        ) e ON datos.Vendedor = e.nombre
        GROUP BY e.tipo, datos.Mes, datos.Vendedor, datos.Venta
        ORDER BY datos.Mes, e.tipo, datos.Vendedor;
    """


    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()

    # Totales generales
    total_subtotal = sum(Decimal(r["Subtotal"]) for r in rows)
    total_descuento = sum(Decimal(r["Descuento"]) for r in rows)
    total_dev = sum(Decimal(r["DevTotal"]) for r in rows)
    total_neto = sum(Decimal(r["Venta_Neta"]) - Decimal(r["DevTotal"]) for r in rows)
    total_iva = sum(Decimal(r["Iva"]) for r in rows)
    total_general = sum(Decimal(r["Total"]) for r in rows)

    # Calcular neto ajustado y porcentaje por fila
    for r in rows:
        dev = Decimal(r["DevTotal"])
        venta_neta = Decimal(r["Venta_Neta"])
        neto_ajustado = venta_neta - dev
        r["Neto_Ajustado"] = neto_ajustado
        r["Porcentaje"] = (neto_ajustado / total_neto * Decimal(100)) if total_neto else Decimal(0)

    # Agrupar por tipo y vendedor
    grupos = {}
    for r in rows:
        tipo = r["Tipo_Vendedor"]
        vendedor = r["Vendedor"]

        if tipo not in grupos:
            grupos[tipo] = {
                "vendedores": {},
                "subtotal": Decimal(0),
                "descuento": Decimal(0),
                "devoluciones": Decimal(0),
                "neto": Decimal(0),
                "iva": Decimal(0),
                "total": Decimal(0)
            }

        grupos[tipo]["subtotal"] += Decimal(r["Subtotal"])
        grupos[tipo]["descuento"] += Decimal(r["Descuento"])
        grupos[tipo]["devoluciones"] += Decimal(r["DevTotal"])
        grupos[tipo]["neto"] += r["Neto_Ajustado"]
        grupos[tipo]["iva"] += Decimal(r["Iva"])
        grupos[tipo]["total"] += Decimal(r["Total"])

        if vendedor not in grupos[tipo]["vendedores"]:
            grupos[tipo]["vendedores"][vendedor] = {
                "filas": [],
                "subtotal": Decimal(0),
                "descuento": Decimal(0),
                "devoluciones": Decimal(0),
                "neto": Decimal(0),
                "iva": Decimal(0),
                "total": Decimal(0)
            }

        grupos[tipo]["vendedores"][vendedor]["filas"].append(r)
        grupos[tipo]["vendedores"][vendedor]["subtotal"] += Decimal(r["Subtotal"])
        grupos[tipo]["vendedores"][vendedor]["descuento"] += Decimal(r["Descuento"])
        grupos[tipo]["vendedores"][vendedor]["devoluciones"] += Decimal(r["DevTotal"])
        grupos[tipo]["vendedores"][vendedor]["neto"] += r["Neto_Ajustado"]
        grupos[tipo]["vendedores"][vendedor]["iva"] += Decimal(r["Iva"])
        grupos[tipo]["vendedores"][vendedor]["total"] += Decimal(r["Total"])

    # Porcentaje por grupo y vendedor
    for tipo, g in grupos.items():
        g["porcentaje"] = (g["neto"] / total_neto * Decimal(100)) if total_neto else Decimal(0)
        for vendedor, data in g["vendedores"].items():
            data["porcentaje"] = (data["neto"] / total_neto * Decimal(100)) if total_neto else Decimal(0)

    # Ordenar grupos y vendedores
    grupos = dict(sorted(grupos.items(), key=lambda x: x[1]["neto"], reverse=True))
    for tipo, g in grupos.items():
        g["vendedores"] = dict(
            sorted(g["vendedores"].items(), key=lambda x: x[1]["neto"], reverse=True)
        )

    # Datos para gráficos
    chart_labels = list(grupos.keys())
    chart_values = [float(grupos[t]["neto"]) for t in chart_labels]

    vendedores_totales = []
    for tipo, g in grupos.items():
        for v, data in g["vendedores"].items():
            vendedores_totales.append((v, data["neto"], tipo))

    vendedores_ordenados = sorted(vendedores_totales, key=lambda x: x[1], reverse=True)
    bar_labels = [v[0] for v in vendedores_ordenados]
    bar_contado, bar_credito, bar_grupos = [], [], [v[2] for v in vendedores_ordenados]

    for v, _, tipo in vendedores_ordenados:
        ventas_vendedor = grupos[tipo]["vendedores"][v]["filas"]
        contado = sum(Decimal(r["Neto_Ajustado"]) for r in ventas_vendedor if r["Venta"].upper() == "CONTADO")
        credito = sum(Decimal(r["Neto_Ajustado"]) for r in ventas_vendedor if r["Venta"].upper() == "CREDITO")
        bar_contado.append(float(contado))
        bar_credito.append(float(credito))

    for tipo, g in grupos.items():
        for vendedor, data in g["vendedores"].items():
            # Crear un diccionario con neto por mes
            netos_por_mes = {}
            for fila in data["filas"]:
                mes = fila["Mes"]
                netos_por_mes[mes] = netos_por_mes.get(mes, Decimal(0)) + fila["Neto_Ajustado"]
            data["netos_por_mes"] = netos_por_mes


    # Resultado final
    resultado = {
        "grupos": grupos,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin,
        "total_subtotal": total_subtotal,
        "total_descuento": total_descuento,
        "total_dev": total_dev,
        "total_neto": total_neto,
        "total_iva": total_iva,
        "total_general": total_general,
        "chart_labels": chart_labels,
        "chart_values": chart_values,
        "bar_labels": bar_labels,
        "bar_contado": bar_contado,
        "bar_credito": bar_credito,
        "bar_grupos": bar_grupos
    }
    return resultado


@app.route("/reporte_ventas", methods=["GET"])
@requiere_acceso("reporte_ventas")
def reporte_ventas():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-%d"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    datos = obtener_reporte_ventas(fecha_inicio, fecha_fin)
    return render_template("reporte_ventas.html", **datos)


@app.route("/rep_vend_totales", methods=["GET"])
@requiere_acceso("rep_vend_totales")
def rep_vend_totales():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-%d"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    datos = obtener_reporte_ventas(fecha_inicio, fecha_fin)
    return render_template("RepVendTotales.html", **datos)

@app.route("/rep_vend_netocreco", methods=["GET"])
@requiere_acceso("rep_vend_netocreco")
def rep_vend_netocreco():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-%d"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    datos = obtener_reporte_ventas(fecha_inicio, fecha_fin)
    return render_template("RepVentasVendedor.html", **datos)

@app.route("/reporte_ventas_dia", methods=["GET"])
@requiere_acceso("reporte_ventas_dia")
def reporte_ventas_dia():
    fecha_inicio = request.args.get("fecha_inicio", datetime.today().strftime("%Y-%m-%d"))
    fecha_fin = request.args.get("fecha_fin", datetime.today().strftime("%Y-%m-%d"))
    tipo_cambio = request.args.get("tipo_cambio", "36.6243")  # valor por defecto
    tipo_cambio = Decimal(tipo_cambio)

    cursor = get_cursor() 

    query = """
    SELECT datos.Fecha,
        datos.TipoVenta,
        SUM(datos.Subtotal) AS Subtotal,
        SUM(datos.Descuento) AS Descuento,
        SUM(datos.Neto) AS Neto,
        SUM(datos.Iva) AS Iva,
        SUM(datos.Total) AS Total,
        SUM(datos.DevTotal) AS DevTotal
    FROM (
        -- Ventas
        SELECT v.fecha AS Fecha,
            v.tipo_venta AS TipoVenta,
            SUM(v.SubTotal) AS Subtotal,
            SUM(v.Descuento) AS Descuento,
            SUM(v.neto) AS Neto,
            SUM(v.iva) AS Iva,
            SUM(v.total) AS Total,
            0 AS DevTotal
        FROM ventas_historicas v
        WHERE v.fecha BETWEEN %s AND %s
        GROUP BY v.fecha, v.tipo_venta

        UNION ALL

        -- Devoluciones
        SELECT r.fechadv AS Fecha,
            r.tipofact AS TipoVenta,
            0 AS Subtotal,
            0 AS Descuento,
            0 AS Neto,
            0 AS Iva,
            0 AS Total,
            SUM(r.totaldv)/1.15 AS DevTotal
        FROM devoluciones r
        WHERE r.fechadv BETWEEN %s AND %s and r.estado='Cerrado'
        GROUP BY r.fechadv, r.tipofact
    ) datos
    GROUP BY datos.Fecha, datos.TipoVenta
    ORDER BY datos.Fecha, datos.TipoVenta;

    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
 

 # Agrupar por fecha
    reporte = {}
    for r in rows:
        fecha = r["Fecha"].strftime("%d-%b")
        if fecha not in reporte:
            reporte[fecha] = {"filas": [], "subtotal": Decimal(0), "descuento": Decimal(0),
                            "devoluciones": Decimal(0), "neto": Decimal(0),
                            "iva": Decimal(0), "total": Decimal(0)}
        
        dev = Decimal(r["DevTotal"] or 0)
        neto_val = Decimal(r["Neto"] or 0)
        neto_ajustado = neto_val - dev

        r["Neto_Ajustado"] = neto_ajustado
        r["NetoUSD"] = neto_ajustado / tipo_cambio

        reporte[fecha]["filas"].append(r)
        reporte[fecha]["subtotal"] += Decimal(r["Subtotal"] or 0)
        reporte[fecha]["descuento"] += Decimal(r["Descuento"] or 0)
        reporte[fecha]["neto"] += Decimal(r["Neto"] or 0)
        reporte[fecha]["devoluciones"] += dev
        reporte[fecha]["neto"] += neto_ajustado
        reporte[fecha]["iva"] += Decimal(r["Iva"] or 0)
        reporte[fecha]["total"] += Decimal(r["Total"] or 0)

    # Totales generales (seguro aunque no haya filas)
    total_subtotal = sum(Decimal(r["Subtotal"] or 0) for r in rows)
    total_descuento = sum(Decimal(r["Descuento"] or 0) for r in rows)
    total_subtotal2 = sum(Decimal(r["Neto"] or 0) for r in rows)
    total_dev = sum(Decimal(r["DevTotal"] or 0) for r in rows)
    total_neto = sum((Decimal(r["Neto"] or 0) - Decimal(r["DevTotal"] or 0)) for r in rows)
    total_iva = sum(Decimal(r["Iva"] or 0) for r in rows)
    total_general = sum(Decimal(r["Total"] or 0) for r in rows)
    total_usd = sum(grupo["neto"] / tipo_cambio for grupo in reporte.values())


    # Inicializar listas seguras
    dias = list(reporte.keys()) if reporte else []
    ventas_contado = []
    ventas_credito = []

    for fecha, grupo in reporte.items():
        neto_contado = sum(r["Neto"] or 0 for r in grupo["filas"] if r["TipoVenta"] == "CONTADO")
        neto_credito = sum(r["Neto"] or 0 for r in grupo["filas"] if r["TipoVenta"] == "CREDITO")
        ventas_contado.append(float(neto_contado))
        ventas_credito.append(float(neto_credito))

    return render_template("reporte_ventas_dia.html",
                           reporte=reporte,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           tipo_cambio=tipo_cambio,
                           total_subtotal=total_subtotal,
                           total_descuento=total_descuento,
                           total_subtotal2=total_subtotal2,
                           total_dev=total_dev,
                           total_neto=total_neto,
                           total_iva=total_iva,
                           total_general=total_general,
                           chart_labels=dias,
                           chart_contado=ventas_contado,
                           chart_credito=ventas_credito)


# -----------------------------
# Reporte de zonas
# -----------------------------

@app.route("/reporteZona", methods=["GET"])
@requiere_acceso("reporteZona")
def reporte():
    tipo = request.args.get("tipo", "detalle")
    vendedor = request.args.get("vendedor")
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    filtro = request.args.get("filtro", "").strip()
    orden = request.args.get("orden", "zona_total")
    print("Filtro recibido:", vendedor)

    if not fecha_inicio:
        fecha_inicio = "2025-01-01"
    if not fecha_fin:
        fecha_fin = "2026-12-31"

    start_year = datetime.strptime(fecha_inicio, "%Y-%m-%d").year
    end_year = datetime.strptime(fecha_fin, "%Y-%m-%d").year
    years = list(range(start_year, end_year + 1))

    cursor = get_cursor()

    if tipo == "totales":
        query = build_query_totales(years, order_by=orden)
    elif tipo == "resumen":
        query = build_query_resumen(years, order_by=orden)
    elif tipo == "zonas_marcas":
        query = build_query_zonas_marcas(years, order_by=orden)
    elif tipo == "zonas_totales":
        query = build_query_zonas_totales(years, order_by=orden)
    else:
        query = build_query_detalle(years, order_by=orden)

    # ✅ Siempre pasar los 3 parámetros
    cursor.execute(query, (vendedor, fecha_inicio, fecha_fin))

    rows = cursor.fetchall()
    data = [dict(row) for row in rows]

    if filtro:
        filtro_lower = filtro.lower()
        data = [
            row for row in data
            if filtro_lower in str(row.get("Cliente", "")).lower()
            or filtro_lower in str(row.get("Marca", "")).lower()
            or filtro_lower in str(row.get("nombre_zona", "")).lower()
            or filtro_lower in str(row.get("Dto_Munic", "")).lower()
        ]

    if tipo in ["zonas_marcas", "zonas_totales"]:
        return render_template("reporte_zonas.html",
                               data=data,
                               years=years,
                               tipo=tipo,
                               vendedor=vendedor,
                               fecha_inicio=fecha_inicio,
                               fecha_fin=fecha_fin,
                               filtro=filtro,
                               orden=orden)

    else:
        estructura = agrupar_por_zona(data)
        return render_template("ventas_vendedor_zona.html",
                               zonas=estructura,
                               years=years,
                               vendedor=vendedor,
                               fecha_inicio=fecha_inicio,
                               fecha_fin=fecha_fin,
                               tipo=tipo,
                               filtro=filtro,
                               orden=orden)
    


@app.route("/exportar", methods=["GET"])
def exportar():
    vendedor = request.args.get("vendedor", "ROGER . COLLADO .")
    fecha_inicio = request.args.get("fecha_inicio", "2020-01-01")
    fecha_fin = request.args.get("fecha_fin", "2026-12-31")
    tipo = request.args.get("tipo", "detalle")

    start_year = datetime.strptime(fecha_inicio, "%Y-%m-%d").year
    end_year = datetime.strptime(fecha_fin, "%Y-%m-%d").year
    years = list(range(start_year, end_year + 1))

    if tipo == "detalle":
        query = build_query_detalle(years)
    elif tipo == "resumen":
        query = build_query_resumen(years)
    else:
        query = build_query_totales(years)

    cursor = get_cursor()
    cursor.execute(query, (vendedor, fecha_inicio, fecha_fin))
    data = cursor.fetchall()

    def generate():
        if tipo == "detalle":
            headers = ["Vendedor","Zona","Departamento","Cliente","Dirección","N° Parte","Descripción","Marca"]
        elif tipo == "resumen":
            headers = ["Zona","Cliente"]
        else:
            headers = ["Zona","Cliente","Marca"]

        for year in years:
            headers += [f"1T {year}", f"2T {year}", f"3T {year}", f"4T {year}"]
        headers.append("Total")

        yield ",".join(headers) + "\n"

        for row in data:
            fila = [str(row.get(h, "")) for h in headers]
            yield ",".join(fila) + "\n"

    return Response(generate(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment;filename=reporte_{tipo}.csv"})

# -----------------------------
# Rutas de clientes y zonas
# -----------------------------
@app.route('/clientes', methods=['GET'])
@requiere_acceso("clientes")
def clientes():
    cursor = get_cursor()
    nombre_busqueda = request.args.get('nombre', '')
    if nombre_busqueda:
        cursor.execute("SELECT id, nombre, departamento, municipio FROM clientes WHERE nombre LIKE %s", (f"%{nombre_busqueda}%",))
    else:
        cursor.execute("SELECT id, nombre, departamento, municipio FROM clientes")
    clientes = cursor.fetchall()
    return render_template('clientes_list.html', clientes=clientes, nombre_busqueda=nombre_busqueda)


@app.route('/editar_cliente/<int:id>', methods=['POST'])
def editar_cliente(id):
    nuevo_departamento = request.form['departamento']
    nuevo_municipio = request.form['municipio']
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE clientes SET departamento=%s, municipio=%s WHERE id=%s",
        (nuevo_departamento, nuevo_municipio, id)
    )
    conn.commit()
    return redirect('/clientes')

@app.route('/eliminar_cliente/<int:id>', methods=['POST'])
def eliminar_cliente(id):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM clientes WHERE id=%s", (id,))
    conn.commit()
    return redirect('/clientes')

@app.route('/newzonas', methods=['GET', 'POST'])
@requiere_acceso("newzonas")
def newzonas():
    if request.method == 'POST':
        nombre_zona = request.form['nombre_zona']
        departamento = request.form['departamento']
        grupozona = request.form['grupozona']
        cursor = conn.cursor()
        cursor.execute("INSERT INTO zona_cliente (nombre_zona, departamento, grupozona) VALUES (%s, %s, %s)",
                       (nombre_zona, departamento, grupozona))
        conn.commit()

    cursor = get_cursor()
    cursor.execute("SELECT idzona, nombre_zona, departamento, grupozona FROM zona_cliente ORDER BY idzona DESC")
    zonas = cursor.fetchall()
    return render_template('zona_cliente_form.html', zonas=zonas)

@app.route('/zonas')
@requiere_acceso("zonas")
def listar_zonas():
    cursor = get_cursor()
    cursor.execute("SELECT idzona, nombre_zona, departamento, grupozona FROM zona_cliente ORDER BY nombre_zona")
    zonas = cursor.fetchall()

    consecutivos = {}
    for zona in zonas:
        nombre = zona['nombre_zona']
        consecutivos[nombre] = consecutivos.get(nombre, 0) + 1
        zona['consecutivo'] = consecutivos[nombre]

    return render_template('zona_cliente_list.html', zonas=zonas)

@app.route('/editar/<int:idzona>', methods=['GET', 'POST'])
def editar_zona(idzona):
    cursor = get_cursor()
    if request.method == 'POST':
        nombre_zona = request.form['nombre_zona']
        departamento = request.form['departamento']
        grupozona = request.form['grupozona']
        cursor.execute("UPDATE zona_cliente SET nombre_zona=%s, departamento=%s, grupozona=%s WHERE idzona=%s",
                       (nombre_zona, departamento, grupozona, idzona))
        conn.commit()
        return redirect('/zonas')
    else:
        cursor.execute("SELECT * FROM zona_cliente WHERE idzona=%s", (idzona,))
        zona = cursor.fetchone()
        return render_template('zona_cliente_edit.html', zona=zona)

@app.route('/eliminar/<int:idzona>')
def eliminar_zona(idzona):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM zona_cliente WHERE idzona=%s", (idzona,))
    conn.commit()
    return redirect('/zonas')

@app.route("/")
def inicio():
    if "usuario" in session:
        return render_template("inicio.html")
    # Si no está logueado → redirigir al login
    return redirect(url_for("login.login"))

# -----------------------------
# Punto de entrada
# -----------------------------
if __name__ == '__main__':
  app.run(host="0.0.0.0", port=8000,debug=True)

