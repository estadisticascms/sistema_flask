import pandas as pd
from flask import Blueprint, request, render_template
from utils.db import get_connection
from router.accesos import requiere_acceso

recibos_bp = Blueprint("recibos", __name__)

@recibos_bp.route("/cargar_recibos", methods=["GET", "POST"])
@requiere_acceso("cargar_recibos")
def cargar_recibos():
    mensaje = None
    ultimos = []

    if request.method == "POST":
        file = request.files["excel_file"]
        df = pd.read_excel(file)

        # Renombrar columnas
        df = df.rename(columns={
            "Fecha Recibo": "fecharc",
            "ROC #": "numerorc",
            "Beneficiario": "cliente",
            "Concepto": "concepto",
            "Monto Cordobas": "total"
        })

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        insertados = 0
        actualizados = 0

        for _, row in df.iterrows():
            sql = """
                INSERT INTO reciboscaja (fecharc, numerorc, cliente, concepto, total)
                VALUES (%s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    cliente = VALUES(cliente),
                    concepto = VALUES(concepto),
                    total = VALUES(total)
            """
            cursor.execute(sql, (
                pd.to_datetime(row["fecharc"]).date(),
                str(row["numerorc"]),
                str(row["cliente"]),
                str(row["concepto"]),
                float(row["total"])
            ))

            if cursor.rowcount == 1:
                insertados += 1
            elif cursor.rowcount == 2:
                actualizados += 1

        conn.commit()

        # Traer últimos 10 recibos cargados
        cursor.execute("""
            SELECT fecharc, numerorc, cliente, concepto, total
            FROM reciboscaja
            ORDER BY fecharc DESC, numerorc DESC
            LIMIT 10
        """)
        ultimos = cursor.fetchall()

        cursor.close()
        conn.close()

        mensaje = f"✅ Recibos cargados correctamente. Insertados: {insertados}, Actualizados: {actualizados}"

    return render_template("cargar_recibos.html", mensaje=mensaje, ultimos=ultimos)

@recibos_bp.route("/reporte_recibos", methods=["GET"])
@requiere_acceso("reporte_recibos")
def reporte_recibos():
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    tipo_cambio = float(request.args.get("tipo_cambio", 36.6243))  # valor ingresado por el usuario

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT fecharc AS Fecha,
               COUNT(numerorc) AS Cantidad,
               SUM(total) AS TotalCordobas
        FROM reciboscaja
        WHERE fecharc BETWEEN %s AND %s
        GROUP BY fecharc
        ORDER BY fecharc;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    # Calcular Subtotal, IVA y Total en Córdobas + Subtotal en USD
    for r in rows:
        subtotal_cordobas = float(r["TotalCordobas"]) / 1.15
        iva_cordobas = subtotal_cordobas * 0.15
        total_cordobas = subtotal_cordobas + iva_cordobas

        subtotal_usd = subtotal_cordobas / tipo_cambio

        r["SubtotalCordobas"] = subtotal_cordobas
        r["IVACordobas"] = iva_cordobas
        r["TotalCordobasCalc"] = total_cordobas
        r["SubtotalUSD"] = subtotal_usd

    # Totales generales
    total_general_subtotal = sum(r["SubtotalCordobas"] for r in rows)
    total_general_iva = sum(r["IVACordobas"] for r in rows)
    total_general_cordobas = sum(r["TotalCordobasCalc"] for r in rows)
    total_general_usd = sum(r["SubtotalUSD"] for r in rows)

    labels = [str(r["Fecha"]) for r in rows]
    valores = [r["TotalCordobasCalc"] for r in rows]

    return render_template("reporte_recibos.html",
                           rows=rows,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           tipo_cambio=tipo_cambio,
                           total_general_subtotal=total_general_subtotal,
                           total_general_iva=total_general_iva,
                           total_general_cordobas=total_general_cordobas,
                           total_general_usd=total_general_usd,
                           chart_labels=labels,
                           chart_values=valores)
