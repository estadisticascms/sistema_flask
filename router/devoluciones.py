import pandas as pd
from flask import Blueprint, request, render_template, redirect, url_for, flash
from utils.db import get_connection
from router.accesos import requiere_acceso

devoluciones_bp = Blueprint("devoluciones", __name__)

@devoluciones_bp.route("/devoluciones", methods=["GET", "POST"])
@requiere_acceso("devoluciones")
def devoluciones():
    ultima = None
    if request.method == "POST":
        file = request.files["archivo"]
        if not file:
            flash("No se seleccionó archivo", "danger")
            return redirect(url_for("devoluciones.devoluciones"))

        df = pd.read_excel(file)
        df = df.rename(columns={
            "# Devolucion Interna": "numerodv",
            "# Ref. Devolucion": "referencia",
            "Fecha Devolucion": "fechadv",
            "# Factura": "factura",
            "Cred./Cont.": "tipofact",
            "Nombre de Cliente": "cliente",
            "Facturador": "facturador",
            "Vendedor Asignado": "vendedor",
            "Motivo Devolucion": "motivo",
            "Total": "totaldv",
            "Saldo Factura": "saldofact",
            "Estado": "estado"
        })

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        for _, row in df.iterrows():
        # 👇 Condición: si vendedor está vacío, usar facturador
            vendedor = row.get("vendedor")
            if not vendedor or str(vendedor).strip() == "":
                vendedor = row.get("facturador")

            sql = """
                INSERT INTO devoluciones
                (numerodv, referencia, fechadv, factura, tipofact,
                 cliente, facturador, vendedor, motivo, totaldv,
                 saldofact, estado)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON DUPLICATE KEY UPDATE
                    cliente = VALUES(cliente),
                    facturador = VALUES(facturador),
                    vendedor = VALUES(vendedor),
                    motivo = VALUES(motivo),
                    totaldv = VALUES(totaldv),
                    saldofact = VALUES(saldofact),
                    estado = VALUES(estado)
            """
            cursor.execute(sql, (
                str(row["numerodv"]),
                str(row.get("referencia")),
                pd.to_datetime(row.get("fechadv")).date() if pd.notnull(row.get("fechadv")) else None,
                str(row.get("factura")),
                str(row.get("tipofact")),
                str(row.get("cliente")),
                str(row.get("facturador")),
                str(vendedor),
                str(row.get("motivo")),
                float(row.get("totaldv")) if pd.notnull(row.get("totaldv")) else None,
                float(row.get("saldofact")) if pd.notnull(row.get("saldofact")) else None,
                str(row.get("estado"))
            ))

        conn.commit()

        # Traer la última devolución cargada
        cursor.execute("""
            SELECT * FROM devoluciones
            ORDER BY fechadv DESC, numerodv DESC
            LIMIT 1
        """)
        ultima = cursor.fetchone()

        cursor.close()
        conn.close()

        flash("✅ Devoluciones cargadas correctamente", "success")

    return render_template("devoluciones.html", ultima=ultima)





