from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from utils.db import get_connection
from functools import wraps


accesos_bp = Blueprint("accesos", __name__)

# Decorador para proteger vistas
def requiere_acceso(vista):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            usuario = session.get("usuario")
            if not usuario:
                return redirect(url_for("login"))

            conn = get_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT permitido FROM accesos_vistas WHERE usuario=%s AND vista=%s", (usuario, vista))
            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if row and row["permitido"]:
                return f(*args, **kwargs)
            else:
                flash("⚠️ No tienes acceso", "warning")
                return redirect(url_for("inicio"))        
        return wrapped
    return decorator

# Pantalla de gestión de accesos
@accesos_bp.route("/accesos", methods=["GET", "POST"])
@requiere_acceso("accesos")
def gestionar_accesos():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if request.method == "POST":
        usuario = request.form["usuario"]
        vista = request.form["vista"]
        permitido = request.form.get("permitido") == "on"

        cursor.execute("""
            INSERT INTO accesos_vistas (usuario, vista, permitido)
            VALUES (%s, %s, %s)
            ON DUPLICATE KEY UPDATE permitido=%s
        """, (usuario, vista, permitido, permitido))
        conn.commit()

    cursor.execute("SELECT * FROM accesos_vistas ORDER BY usuario, vista")
    accesos = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template("accesos.html", accesos=accesos)
