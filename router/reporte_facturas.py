from flask import Blueprint, request, render_template
from utils.db import get_connection
from decimal import Decimal
from router.accesos import requiere_acceso

reporte_facturas_bp = Blueprint("reporte_facturas", __name__)

@reporte_facturas_bp.route("/reporte_facturas", methods=["GET"])
@requiere_acceso("reporte_facturas")
def reporte_facturas():
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT e.tipo AS Tipo_Vendedor,
            v.vendedor AS Vendedor,
            v.tipo_venta AS Venta,
            v.`tipo factura` AS TipoFactura,
            SUM(v.neto) AS Venta_Neta,
            0 AS DevTotal
        FROM ventas_historicas v
        LEFT JOIN (
            SELECT DISTINCT tipo, nombre 
            FROM cat_vendedor
        ) e ON v.vendedor = e.nombre
        WHERE v.fecha BETWEEN %s AND %s
        GROUP BY e.tipo, v.vendedor, v.tipo_venta, v.`tipo factura`

        UNION ALL

        SELECT e.tipo AS Tipo_Vendedor,
            r.vendedor AS Vendedor,
            r.tipofact AS Venta,
            'Bienes' AS TipoFactura,
            0 AS Venta_Neta,
            SUM(r.totaldv)/1.15 AS DevTotal
        FROM devoluciones r
        LEFT JOIN (
            SELECT DISTINCT tipo, nombre 
            FROM cat_vendedor
        ) e ON r.vendedor = e.nombre
        WHERE r.fechadv BETWEEN %s AND %s
        AND r.estado='Cerrado'
        GROUP BY e.tipo, r.vendedor, r.tipofact;

    """

    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    grupos = {}
    total_bienes = Decimal(0)
    total_mixta = Decimal(0)

    for r in rows:
        tipo = r["Tipo_Vendedor"]
        vendedor = r["Vendedor"]
        venta = r["Venta"].upper()
        tipo_factura = r["TipoFactura"]

        neto_ajustado = Decimal(r["Venta_Neta"]) - Decimal(r["DevTotal"])

        if tipo not in grupos:
            grupos[tipo] = {"vendedores": {}, "totales": {"Bienes": Decimal(0), "Mixtos": Decimal(0)}}

        if vendedor not in grupos[tipo]["vendedores"]:
            grupos[tipo]["vendedores"][vendedor] = {
                "ventas": {"CONTADO": {"Bienes": Decimal(0), "Mixtos": Decimal(0)},
                           "CREDITO": {"Bienes": Decimal(0), "Mixtos": Decimal(0)}},
                "totales": {"Bienes": Decimal(0), "Mixtos": Decimal(0)}
            }

        grupos[tipo]["vendedores"][vendedor]["ventas"][venta][tipo_factura] += neto_ajustado
        grupos[tipo]["vendedores"][vendedor]["totales"][tipo_factura] += neto_ajustado
        grupos[tipo]["totales"][tipo_factura] += neto_ajustado

        if tipo_factura == "Bienes":
            total_bienes += neto_ajustado
        elif tipo_factura == "Mixtos":
            total_mixta += neto_ajustado

    total_general = total_bienes + total_mixta

    # Calcular porcentajes
    for tipo, g in grupos.items():
        for vendedor, data in g["vendedores"].items():
            bienes = data["totales"]["Bienes"]
            mixta = data["totales"]["Mixtos"]
            total_vendedor = bienes + mixta
            if total_general > 0:
                data["porcentaje_bienes"] = (bienes / total_general * 100)
                data["porcentaje_mixta"] = (mixta / total_general * 100)
                data["porcentaje_total"] = (total_vendedor / total_general * 100)
            else:
                data["porcentaje_bienes"] = 0
                data["porcentaje_mixta"] = 0
                data["porcentaje_total"] = 0


    return render_template("reporte_facturas_totales.html",
                           grupos=grupos,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           total_bienes=total_bienes,
                           total_mixta=total_mixta,
                           total_general=total_general)

