"""Permisos por pantalla (tabla accesos_vistas: usuario, vista, permitido).

Cada ruta protegida usa @requiere_acceso("vista"). El decorador deja el nombre de la vista en la función, así la pantalla de
accesos arma el catálogo sola: recorre las rutas de la app, las cruza con el menú (router/menu.py) para mostrar nombres legibles
agrupados por sección, y lo que no está en el menú va a "Otros permisos"."""
from functools import wraps

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from utils.db import get_connection

accesos_bp = Blueprint("accesos", __name__)

# Permisos que no son una pantalla del menú: nombre legible
OTROS = {
    "cheques_config": ("Cheques: bancos y cuentas contables", "Agregar o editar bancos, consecutivos y cuentas"),
    "editar_venta_perdida": ("Ventas perdidas: editar", "Corregir registros de ventas perdidas"),
    "eliminar_venta_perdida": ("Ventas perdidas: eliminar", "Borrar registros de ventas perdidas"),
    "importar_ventas_perdidas": ("Ventas perdidas: importar", "Subir ventas perdidas desde Excel"),
    "registro": ("Crear usuarios (antiguo)", "Ya no se usa: crear usuarios depende de «Usuarios»"),
    "reportes": ("Reportes (antiguo)", "Pantalla de reportes anterior"),
}
# Permisos que siempre conserva quien administra accesos (para no quedarse fuera)
PROPIOS = {"accesos", "usuarios"}


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
        wrapped.vista_acceso = vista   # para armar el catálogo de permisos
        return wrapped
    return decorator


def catalogo_vistas():
    """Lista de grupos [{titulo, icono, color, items: [{vista, nombre, descripcion, icono, subgrupo}]}] según el menú y las rutas.
    Cada permiso aparece una sola vez, en su área de trabajo (los atajos de Gerencia no se repiten)."""
    from router.menu import SECCIONES, vista_por_url
    por_url = vista_por_url()
    todas = set(por_url.values())
    grupos, vistos = [], set()
    for s in SECCIONES:
        if s.get("acceso_rapido"):
            continue
        items = {}
        for sub in s["grupos"]:
            for nombre, url, icono, descripcion, *_ in sub["items"]:
                vista = por_url.get(url)
                if not vista or vista in vistos:
                    continue
                if vista in items:   # varias opciones del menú con el mismo permiso (p. ej. Cheques y Nuevo cheque)
                    items[vista]["nombre"] += f" · {nombre}"
                    continue
                items[vista] = {"vista": vista, "nombre": nombre, "descripcion": descripcion, "icono": icono, "subgrupo": sub["titulo"]}
        vistos |= set(items)
        if items:
            grupos.append({"id": s["id"], "titulo": s["titulo"], "icono": s["icono"], "color": s["color"], "items": list(items.values())})
    otros = [{"vista": v, "nombre": OTROS.get(v, (v.replace("_", " ").capitalize(), ""))[0], "descripcion": OTROS.get(v, ("", ""))[1],
              "icono": "bi-key", "subgrupo": ""} for v in sorted(todas - vistos)]
    if otros:
        grupos.append({"id": "otros", "titulo": "Otros permisos", "icono": "bi-key-fill", "color": "gris", "items": otros})
    return grupos


def vistas_validas():
    return {i["vista"] for g in catalogo_vistas() for i in g["items"]}


def leer_usuarios(cursor):
    cursor.execute("SELECT id, usuario, rol, creado FROM usuarios ORDER BY usuario")
    usuarios = cursor.fetchall()
    cursor.execute("SELECT usuario, vista, permitido FROM accesos_vistas")
    permisos = {}
    for r in cursor.fetchall():
        permisos.setdefault(r["usuario"].lower(), {})[r["vista"]] = bool(r["permitido"])
    return usuarios, permisos


# Pantalla de gestión de accesos
@accesos_bp.route("/accesos", methods=["GET", "POST"])
@requiere_acceso("accesos")
def gestionar_accesos():
    validas = vistas_validas()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if request.method == "POST":
            usuario = request.form.get("usuario", "")
            cursor.execute("SELECT usuario FROM usuarios WHERE usuario = %s", (usuario,))
            fila = cursor.fetchone()
            if not fila:
                flash("El usuario no existe.", "warning")
                return redirect(url_for("accesos.gestionar_accesos"))
            usuario = fila["usuario"]
            marcadas = set(request.form.getlist("vista")) & validas
            if usuario.lower() == (session.get("usuario") or "").lower():
                marcadas |= PROPIOS   # no se puede quitar a sí mismo la administración
            for v in validas:
                cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s) "
                               "ON DUPLICATE KEY UPDATE permitido = VALUES(permitido)", (usuario, v, 1 if v in marcadas else 0))
            conn.commit()
            flash(f"✅ Permisos de {usuario} guardados: {len(marcadas)} pantalla(s) habilitada(s).", "success")
            return redirect(url_for("accesos.gestionar_accesos", usuario=usuario))

        usuarios, permisos = leer_usuarios(cursor)
        nombres = {u["usuario"].lower() for u in usuarios}
        huerfanos = sorted(u for u in permisos if u not in nombres)
    finally:
        cursor.close()
        conn.close()

    elegido = request.args.get("usuario") or (usuarios[0]["usuario"] if usuarios else "")
    actual = next((u for u in usuarios if u["usuario"].lower() == elegido.lower()), None)
    grupos = catalogo_vistas()
    for u in usuarios:
        u["habilitadas"] = sum(1 for v, p in permisos.get(u["usuario"].lower(), {}).items() if p and v in validas)
    return render_template("accesos.html", usuarios=usuarios, actual=actual, grupos=grupos, total=len(validas),
                           permisos=permisos.get(actual["usuario"].lower(), {}) if actual else {},
                           huerfanos=huerfanos, propio=bool(actual and actual["usuario"].lower() == (session.get("usuario") or "").lower()),
                           propios=PROPIOS)


@accesos_bp.route("/accesos/copiar", methods=["POST"])
@requiere_acceso("accesos")
def copiar_accesos():
    """Deja al usuario destino con exactamente los mismos permisos que el de origen."""
    origen, destino = request.form.get("origen", ""), request.form.get("destino", "")
    if not origen or origen.lower() == destino.lower():
        flash("Elija otro usuario para copiar sus permisos.", "warning")
        return redirect(url_for("accesos.gestionar_accesos", usuario=destino))
    validas = vistas_validas()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT vista FROM accesos_vistas WHERE usuario = %s AND permitido = 1", (origen,))
        marcadas = {r["vista"] for r in cursor.fetchall()} & validas
        if destino.lower() == (session.get("usuario") or "").lower():
            marcadas |= PROPIOS
        for v in validas:
            cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s) "
                           "ON DUPLICATE KEY UPDATE permitido = VALUES(permitido)", (destino, v, 1 if v in marcadas else 0))
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash(f"✅ {destino} ahora tiene los mismos permisos que {origen} ({len(marcadas)} pantallas).", "success")
    return redirect(url_for("accesos.gestionar_accesos", usuario=destino))


@accesos_bp.route("/accesos/limpiar", methods=["POST"])
@requiere_acceso("accesos")
def limpiar_huerfanos():
    """Borra permisos que quedaron de usuarios que ya no existen."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE a FROM accesos_vistas a LEFT JOIN usuarios u ON u.usuario = a.usuario WHERE u.id IS NULL")
        borrados = cursor.rowcount
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash(f"Se borraron {borrados} permisos de usuarios que ya no existen.", "info")
    return redirect(url_for("accesos.gestionar_accesos"))
