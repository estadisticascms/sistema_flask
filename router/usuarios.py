from flask import Blueprint, flash, redirect, render_template, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

from utils.db import get_connection
from router.accesos import requiere_acceso, vistas_validas

usuarios_bp = Blueprint("usuarios", __name__)

ROLES = {"usuario": "Usuario", "admin": "Administrador"}
INICIALES = ["ventas_perdidas"]   # lo que recibe un usuario nuevo si no se copian permisos de otro


def limpiar_nombre(valor):
    return " ".join((valor or "").split())[:100]


# -----------------------------
# Registro de usuarios
# -----------------------------
@usuarios_bp.route("/registro", methods=["GET", "POST"])
@requiere_acceso("usuarios")
def registro():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT usuario FROM usuarios ORDER BY usuario")
        existentes = [r["usuario"] for r in cursor.fetchall()]
        if request.method == "POST":
            usuario = limpiar_nombre(request.form.get("usuario"))
            password, confirmar = request.form.get("password", ""), request.form.get("confirm_password", "")
            rol = request.form.get("rol") if request.form.get("rol") in ROLES else "usuario"
            copiar = request.form.get("copiar_de", "")
            error = None
            if not usuario:
                error = "Escriba el nombre de usuario."
            elif usuario.lower() in {e.lower() for e in existentes}:
                error = f"Ya existe el usuario «{usuario}»."
            elif len(password) < 6:
                error = "La contraseña debe tener al menos 6 caracteres."
            elif password != confirmar:
                error = "Las contraseñas no coinciden."
            if error:
                flash(error, "warning")
                return render_template("usuario_form.html", user={"usuario": usuario, "rol": rol}, nuevo=True, roles=ROLES,
                                       existentes=existentes, copiar_de=copiar)
            cursor.execute("INSERT INTO usuarios (usuario, password, rol) VALUES (%s, %s, %s)", (usuario, generate_password_hash(password), rol))
            if copiar:
                cursor.execute("SELECT vista FROM accesos_vistas WHERE usuario = %s AND permitido = 1", (copiar,))
                vistas = [r["vista"] for r in cursor.fetchall()]
            else:
                vistas = INICIALES
            # Sin fila = sin acceso: solo se guardan los permisos habilitados
            for v in vistas:
                cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, 1) "
                               "ON DUPLICATE KEY UPDATE permitido = 1", (usuario, v))
            conn.commit()
            flash(f"✅ Usuario {usuario} creado" + (f" con los permisos de {copiar}." if copiar else ". Revise sus permisos."), "success")
            return redirect(url_for("accesos.gestionar_accesos", usuario=usuario))
    finally:
        cursor.close()
        conn.close()
    return render_template("usuario_form.html", user={"usuario": "", "rol": "usuario"}, nuevo=True, roles=ROLES, existentes=existentes, copiar_de="")

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
            session["usuario"] = user["usuario"]
            session["rol"] = user["rol"]
            return redirect(url_for("inicio"))
        return render_template("login.html", error="Usuario o contraseña incorrectos.")

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
    validas = vistas_validas()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, usuario, rol, creado FROM usuarios ORDER BY rol = 'admin' DESC, usuario")
    usuarios = cursor.fetchall()
    cursor.execute("SELECT usuario, vista FROM accesos_vistas WHERE permitido = 1")
    habilitadas = {}
    for r in cursor.fetchall():
        if r["vista"] in validas:
            habilitadas[r["usuario"].lower()] = habilitadas.get(r["usuario"].lower(), 0) + 1
    cursor.close()
    conn.close()
    for u in usuarios:
        u["habilitadas"] = habilitadas.get(u["usuario"].lower(), 0)
    return render_template("usuarios.html", usuarios=usuarios, total=len(validas), roles=ROLES, yo=(session.get("usuario") or "").lower())

# -----------------------------
# Editar usuario (nombre, rol y, si se escribe, contraseña)
# -----------------------------
@usuarios_bp.route("/usuarios/editar/<int:id>", methods=["GET", "POST"])
@requiere_acceso("usuarios")
def editar_usuario(id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM usuarios WHERE id=%s", (id,))
        user = cursor.fetchone()
        if not user:
            flash("El usuario no existe.", "warning")
            return redirect(url_for("usuarios.listar_usuarios"))
        if request.method == "POST":
            nuevo = limpiar_nombre(request.form.get("usuario"))
            rol = request.form.get("rol") if request.form.get("rol") in ROLES else "usuario"
            password, confirmar = request.form.get("password", ""), request.form.get("confirm_password", "")
            es_yo = user["usuario"].lower() == (session.get("usuario") or "").lower()
            error = None
            cursor.execute("SELECT id FROM usuarios WHERE usuario = %s AND id <> %s", (nuevo, id))
            if not nuevo:
                error = "Escriba el nombre de usuario."
            elif cursor.fetchone():
                error = f"Ya existe otro usuario «{nuevo}»."
            elif password and len(password) < 6:
                error = "La contraseña debe tener al menos 6 caracteres."
            elif password != confirmar:
                error = "Las contraseñas no coinciden."
            elif es_yo and rol != "admin" and user["rol"] == "admin":
                error = "No puede quitarse a sí mismo el rol de administrador."
            if error:
                flash(error, "warning")
                return render_template("usuario_form.html", user={**user, "usuario": nuevo, "rol": rol}, nuevo=False, roles=ROLES)
            if password:
                cursor.execute("UPDATE usuarios SET usuario=%s, rol=%s, password=%s WHERE id=%s", (nuevo, rol, generate_password_hash(password), id))
            else:
                cursor.execute("UPDATE usuarios SET usuario=%s, rol=%s WHERE id=%s", (nuevo, rol, id))
            if nuevo != user["usuario"]:   # los permisos están por nombre: se renombran también
                cursor.execute("UPDATE accesos_vistas SET usuario=%s WHERE usuario=%s", (nuevo, user["usuario"]))
                if es_yo:
                    session["usuario"] = nuevo
            conn.commit()
            flash(f"✅ Usuario {nuevo} actualizado" + (" (contraseña cambiada)." if password else "."), "success")
            return redirect(url_for("usuarios.listar_usuarios"))
    finally:
        cursor.close()
        conn.close()
    return render_template("usuario_form.html", user=user, nuevo=False, roles=ROLES)

# -----------------------------
# Cambiar contraseña (se mantiene la ruta anterior; usa el mismo formulario de edición)
# -----------------------------
@usuarios_bp.route("/usuarios/password/<int:id>", methods=["GET", "POST"])
@requiere_acceso("usuarios")
def editar_password(id):
    return redirect(url_for("usuarios.editar_usuario", id=id))

# -----------------------------
# Eliminar usuario
# -----------------------------
@usuarios_bp.route("/usuarios/eliminar/<int:id>", methods=["POST"])
@requiere_acceso("usuarios")
def eliminar_usuario(id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT usuario, rol FROM usuarios WHERE id=%s", (id,))
        user = cursor.fetchone()
        if not user:
            flash("El usuario no existe.", "warning")
        elif user["usuario"].lower() == (session.get("usuario") or "").lower():
            flash("No puede eliminar su propio usuario.", "warning")
        else:
            cursor.execute("SELECT COUNT(*) AS n FROM usuarios WHERE rol = 'admin'")
            admins = cursor.fetchone()["n"]
            if user["rol"] == "admin" and admins <= 1:
                flash("No se puede eliminar el único administrador.", "warning")
            else:
                cursor.execute("DELETE FROM accesos_vistas WHERE usuario=%s", (user["usuario"],))
                cursor.execute("DELETE FROM usuarios WHERE id=%s", (id,))
                conn.commit()
                flash(f"Usuario {user['usuario']} eliminado junto con sus permisos.", "info")
    finally:
        cursor.close()
        conn.close()
    return redirect(url_for("usuarios.listar_usuarios"))
