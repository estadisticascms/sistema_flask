from flask import Blueprint, render_template, request, redirect, url_for, session
from werkzeug.security import generate_password_hash, check_password_hash
from utils.db import get_connection
from router.accesos import requiere_acceso

usuarios_bp = Blueprint("usuarios", __name__)

# -----------------------------
# Registro de usuarios
# -----------------------------
@usuarios_bp.route("/registro", methods=["GET", "POST"])
@requiere_acceso("registro")
def registro():
    if request.method == "POST":
        usuario = request.form["usuario"]
        password = request.form["password"]
        rol = request.form.get("rol", "usuario")

        hash_pass = generate_password_hash(password)

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO usuarios (usuario, password, rol) VALUES (%s, %s, %s)", (usuario, hash_pass, rol))
        # Insertar accesos iniciales
        cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s)", (usuario, "inicio", 1))
        cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s)", (usuario, "ventas_perdidas", 1))

        # Todas las demás vistas en denegado
        vistas_denegadas = [
            "clientes", "zonas", "newzonas", "resumen", "cargar_recibos", "devoluciones",
            "reporteZona", "reporte_ventas_dia", "reporte_ventas", "rep_vend_totales", "rep_vend_netocreco",
            "reporte_facturas", "reporte_ventas_mes", "ventas_zonas_mes", "top_marcas", "reporte_recibos",
            "accesos", "usuarios"
        ]

        for v in vistas_denegadas:
            cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s)", (usuario, v, 0))

        conn.commit()
        cursor.close()
        conn.close()

        return redirect(url_for("usuarios.listar_usuarios"))

    return render_template("registro.html")

# -----------------------------
# Login
# -----------------------------
@usuarios_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"]
        password = request.form["password"]

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM usuarios WHERE usuario=%s", (usuario,))
        user = cursor.fetchone()
        cursor.close()
        conn.close()

        if user and check_password_hash(user["password"], password):
            session["usuario"] = usuario
            session["rol"] = user["rol"]
            return redirect(url_for("inicio"))
        else:
            return "❌ Usuario o contraseña incorrectos"

    return render_template("login.html")

# -----------------------------
# Logout
# -----------------------------
@usuarios_bp.route("/logout")
def logout():
    session.pop("usuario", None)
    session.pop("rol", None)
    return redirect(url_for("usuarios.login"))

# -----------------------------
# Listar usuarios
# -----------------------------
@usuarios_bp.route("/usuarios")
@requiere_acceso("usuarios")
def listar_usuarios():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, usuario, rol, creado FROM usuarios ORDER BY creado DESC")
    usuarios = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template("usuarios.html", usuarios=usuarios)

# -----------------------------
# Editar usuario
# -----------------------------
@usuarios_bp.route("/usuarios/editar/<int:id>", methods=["GET", "POST"])
def editar_usuario(id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if request.method == "POST":
        usuario = request.form["usuario"]
        rol = request.form["rol"]
        password = request.form.get("password")

        if password:  # Si el campo no está vacío, actualizar contraseña
            hash_pass = generate_password_hash(password)
            cursor.execute("UPDATE usuarios SET usuario=%s, rol=%s, password=%s WHERE id=%s",
                        (usuario, rol, hash_pass, id))
        else:
            cursor.execute("UPDATE usuarios SET usuario=%s, rol=%s WHERE id=%s",
                        (usuario, rol, id))

        conn.commit()
        cursor.close()
        conn.close()
        return redirect(url_for("usuarios.listar_usuarios"))

    cursor.execute("SELECT * FROM usuarios WHERE id=%s", (id,))
    user = cursor.fetchone()
    cursor.close()
    conn.close()
    return render_template("editar_usuario.html", user=user)

# -----------------------------
# Editar contraseña de usuario
# -----------------------------
@usuarios_bp.route("/usuarios/password/<int:id>", methods=["GET", "POST"])
@requiere_acceso("usuarios")
def editar_password(id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if request.method == "POST":
        nueva_password = request.form["password"]
        hash_pass = generate_password_hash(nueva_password)

        cursor.execute("UPDATE usuarios SET password=%s WHERE id=%s", (hash_pass, id))
        conn.commit()
        cursor.close()
        conn.close()
        return redirect(url_for("usuarios.listar_usuarios"))

    cursor.execute("SELECT id, usuario FROM usuarios WHERE id=%s", (id,))
    user = cursor.fetchone()
    cursor.close()
    conn.close()
    return render_template("editar_password.html", user=user)

# -----------------------------
# Eliminar usuario
# -----------------------------
@usuarios_bp.route("/usuarios/eliminar/<int:id>")
def eliminar_usuario(id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM usuarios WHERE id=%s", (id,))
    conn.commit()
    cursor.close()
    conn.close()
    return redirect(url_for("usuarios.listar_usuarios"))



