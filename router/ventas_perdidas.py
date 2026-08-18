from flask import Blueprint, render_template, request, redirect, url_for, flash
from datetime import date
from utils.db import get_connection
from werkzeug.utils import secure_filename
import os
import pandas as pd
from router.accesos import requiere_acceso

ventas_perdidas_bp = Blueprint("ventas_perdidas", __name__)

# -----------------------------
# Ventas Perdidas
# -----------------------------
@ventas_perdidas_bp.route("/ventas_perdidas", methods=["GET", "POST"])
@requiere_acceso("ventas_perdidas")
def ventas_perdidas():
    fecha_hoy = date.today().strftime("%Y-%m-%d")

    if request.method == "POST":
        nomb_vendedor = request.form.get("nomb_vendedor")
        nomb_cliente = request.form.get("nomb_cliente")
        fecha = request.form.get("fecha")
        numero_parte = request.form.get("numero_parte")
        descripcion = request.form.get("descripcion")
        cantidad = request.form.get("cantidad")
        nota = request.form.get("nota")

        if not all([nomb_vendedor, nomb_cliente, fecha, numero_parte, descripcion, cantidad, nota]):
            flash("⚠️ Todos los campos son obligatorios", "warning")
        else:
            conn = get_connection()
            cursor = conn.cursor()
            sql = """
            INSERT INTO ventas_perdidas
            (nomb_vendedor, nomb_cliente, fecha, numero_parte, descripcion, cantidad, nota)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """
            cursor.execute(sql, (nomb_vendedor, nomb_cliente, fecha, numero_parte, descripcion, cantidad, nota))
            conn.commit()
            cursor.close()
            conn.close()
            flash("✅ Venta perdida registrada correctamente", "success")

        return redirect(url_for("ventas_perdidas.ventas_perdidas"))

    # --- Filtros por vendedor y fechas ---
    filtro_vendedor = request.args.get("filtro_vendedor", "")
    fecha_inicio = request.args.get("fecha_inicio", "")
    fecha_fin = request.args.get("fecha_fin", "")

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = "SELECT * FROM ventas_perdidas WHERE 1=1"
    params = []

    if filtro_vendedor:
        query += " AND nomb_vendedor=%s"
        params.append(filtro_vendedor)

    if fecha_inicio:
        query += " AND fecha >= %s"
        params.append(fecha_inicio)

    if fecha_fin:
        query += " AND fecha <= %s"
        params.append(fecha_fin)

    query += " ORDER BY fecha DESC"
    cursor.execute(query, tuple(params))
    registros = cursor.fetchall()

    cursor.execute("SELECT DISTINCT nomb_vendedor FROM ventas_perdidas ORDER BY nomb_vendedor")
    vendedores = [row["nomb_vendedor"] for row in cursor.fetchall()]

    cursor.close()
    conn.close()

    return render_template("ventas_perdidas.html",
                           registros=registros,
                           fecha_hoy=fecha_hoy,
                           vendedores=vendedores,
                           filtro_vendedor=filtro_vendedor,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin)


# -----------------------------
# Editar registro
# -----------------------------
@ventas_perdidas_bp.route("/editar_venta_perdida/<int:id>", methods=["GET", "POST"])
@requiere_acceso("editar_venta_perdida")
def editar_venta_perdida(id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if request.method == "POST":
        fecha = request.form.get("fecha")
        vendedor = request.form.get("nomb_vendedor")
        cliente = request.form.get("nomb_cliente")
        numero_parte = request.form.get("numero_parte")
        descripcion = request.form.get("descripcion")
        cantidad = request.form.get("cantidad")
        nota = request.form.get("nota")

        sql = """
        UPDATE ventas_perdidas
        SET fecha=%s, nomb_vendedor=%s, nomb_cliente=%s, numero_parte=%s,
            descripcion=%s, cantidad=%s, nota=%s
        WHERE id=%s
        """
        cursor.execute(sql, (fecha, vendedor, cliente, numero_parte, descripcion, cantidad, nota, id))
        conn.commit()
        cursor.close()
        conn.close()
        flash("✅ Registro actualizado correctamente", "success")
        return redirect(url_for("ventas_perdidas.ventas_perdidas"))
    else:
        cursor.execute("SELECT * FROM ventas_perdidas WHERE id=%s", (id,))
        registro = cursor.fetchone()
        cursor.close()
        conn.close()
        return render_template("ventas_perdidas_edit.html", registro=registro)


# -----------------------------
# Eliminar registro
# -----------------------------
@ventas_perdidas_bp.route("/eliminar_venta_perdida/<int:id>")
@requiere_acceso("eliminar_venta_perdida")
def eliminar_venta_perdida(id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ventas_perdidas WHERE id=%s", (id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("🗑️ Registro eliminado correctamente", "success")
    return redirect(url_for("ventas_perdidas.ventas_perdidas"))


# -----------------------------
# Importar desde Excel
# -----------------------------
@ventas_perdidas_bp.route("/importar_ventas_perdidas", methods=["POST"])
@requiere_acceso("importar_ventas_perdidas")
def importar_ventas_perdidas():
    if "archivo_excel" not in request.files:
        flash("No se seleccionó archivo", "danger")
        return redirect(url_for("ventas_perdidas.ventas_perdidas"))

    archivo = request.files["archivo_excel"]
    if archivo.filename == "":
        flash("Nombre de archivo vacío", "danger")
        return redirect(url_for("ventas_perdidas.ventas_perdidas"))

    filename = secure_filename(archivo.filename)
    filepath = os.path.join("uploads", filename)  # asegúrate que exista la carpeta
    archivo.save(filepath)

    try:
        df = pd.read_excel(filepath)
        columnas = ["nomb_vendedor", "nomb_cliente", "fecha",
                    "numero_parte", "descripcion", "cantidad", "nota"]

        for col in columnas:
            if col not in df.columns:
                flash(f"Columna faltante en Excel: {col}", "danger")
                os.remove(filepath)
                return redirect(url_for("ventas_perdidas.ventas_perdidas"))

        conn = get_connection()
        cursor = conn.cursor()
        for _, row in df.iterrows():
            sql = """
            INSERT INTO ventas_perdidas
            (nomb_vendedor, nomb_cliente, fecha, numero_parte, descripcion, cantidad, nota)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """
            cursor.execute(sql, (
                row["nomb_vendedor"], row["nomb_cliente"], row["fecha"],
                row["numero_parte"], row["descripcion"], row["cantidad"], row["nota"]
            ))
        conn.commit()
        cursor.close()
        conn.close()
        flash("✅ Archivo importado correctamente", "success")

    except Exception as e:
        flash(f"Error al procesar archivo: {e}", "danger")

    finally:
        os.remove(filepath)

    return redirect(url_for("ventas_perdidas.ventas_perdidas"))

