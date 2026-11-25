"""Módulo de seguimiento de entregas: roles, permisos y pantallas (especificación: docs/entregas/ESPECIFICACION_MODULO_ENTREGAS.md).

Usuarios: los mismos de la tabla `usuarios` (el administrador crea usuario y contraseña en Admin › Usuarios). No hay otro login.
Roles del módulo: tabla ent_usuario_rol (docs/sql/ent_entregas.sql), una fila por rol; un usuario puede tener varios.

Permisos en dos niveles:
1. Pantalla (accesos_vistas, igual que el resto del sistema): decide qué opciones ve en el menú. Al guardar los roles de un
   usuario se le habilitan exactamente las pantallas de Entregas de sus roles (sincronizar_accesos).
2. Rol (ent_usuario_rol): cada pantalla vuelve a validar en el servidor que el usuario tenga uno de los roles que la usan
   (requiere_rol). Una pantalla habilitada a mano en Accesos sin el rol no se abre.
"""
from functools import wraps

from flask import Blueprint, flash, g, redirect, render_template, request, session, url_for
from mysql.connector import errorcode
from mysql.connector.errors import ProgrammingError

from router.accesos import requiere_acceso
from utils.db import get_connection

entregas_bp = Blueprint("entregas", __name__)

ROLES = {
    "vendedor": ("Vendedor", "Sube la orden, agrega transporte y hora, decide ante faltantes y registra las facturas"),
    "bodega": ("Bodeguero", "Ve solo las líneas de su bodega: confirma existencias y registra la preparación"),
    "coordinador_bodega": ("Coordinador de bodega", "Bodegas con traslado: existencias, aplicar traslado, traer y recibir"),
    "paqueteria": ("Paquetería", "Seguimiento de todas las órdenes, bodegas con traslado, retiro, empaque y rutas"),
    "mensajero": ("Mensajero", "Hace la ruta y registra la entrega, fotos, recibo y flete"),
    "admin": ("Administrador de Entregas", "Catálogos, configuración, usuarios del módulo, tablero y reportes"),
}

# Pantallas del módulo: vista (permiso), título, url, icono, descripción, roles que la usan, parte en que se construye.
# Las mismas opciones están en router/menu.py (sección «Entregas»).
PANTALLAS = [
    ("entregas_subir_orden", "Subir orden", "/entregas/subir_orden", "bi-file-earmark-arrow-up",
     "PDF de la orden → vista previa → transporte y hora → enviar", ("vendedor",), 4),
    ("entregas_mis_envios", "Mis envíos", "/entregas/mis_envios", "bi-send-check",
     "Estado de mis envíos, faltantes por decidir y listos para facturar", ("vendedor",), 7),
    ("entregas_bodega", "Mi bodega", "/entregas/bodega", "bi-box-seam",
     "Líneas de mi bodega: existencias y preparación", ("bodega",), 6),
    ("entregas_coordinador", "Bodegas con traslado", "/entregas/coordinador", "bi-arrow-left-right",
     "Existencias de bodegas con traslado; traslados por aplicar, en camino y por recibir", ("coordinador_bodega",), 8),
    ("entregas_seguimiento", "Seguimiento de órdenes", "/entregas/seguimiento", "bi-kanban",
     "Tablero de todas las órdenes abiertas y la preparación por bodega", ("paqueteria",), 6),
    ("entregas_bodegas_traslado", "Bodegas que aplican traslado", "/entregas/bodegas_traslado", "bi-signpost-split",
     "Qué bodegas requieren traslado y hacia qué bodega", ("paqueteria", "admin"), 3),
    ("entregas_empaque", "Retiro y empaque", "/entregas/empaque", "bi-box2-heart",
     "Retiro por bodega, bultos, peso y etiqueta", ("paqueteria",), 10),
    ("entregas_rutas", "Rutas", "/entregas/rutas", "bi-signpost-2",
     "Armar las rutas del día y asignarlas al mensajero", ("paqueteria",), 11),
    ("entregas_mis_rutas", "Mis rutas", "/entregas/mis_rutas", "bi-geo-alt",
     "Paradas del día, «Salí» y registro de la entrega con fotos", ("mensajero",), 11),
    ("entregas_catalogos", "Catálogos y configuración", "/entregas/catalogos", "bi-sliders",
     "Bodegas, transportistas, horarios y tiempos de la alerta", ("admin",), 3),
    ("entregas_usuarios", "Usuarios y roles", "/entregas/usuarios", "bi-people",
     "Roles de cada usuario en Entregas y la bodega del bodeguero", ("admin",), 2),
    ("entregas_reportes", "Tablero y reportes", "/entregas/reportes", "bi-bar-chart-line",
     "Tiempos por etapa, entregas a tiempo, faltantes y fletes", ("admin",), 12),
]
VISTAS = {p[0]: p for p in PANTALLAS}
VISTAS_DE_ROL = {rol: {p[0] for p in PANTALLAS if rol in p[5]} for rol in ROLES}


# -----------------------------
# Roles del usuario
# -----------------------------
def tabla_faltante(e):
    return isinstance(e, ProgrammingError) and e.errno == errorcode.ER_NO_SUCH_TABLE


def usuario_actual():
    """(id, usuario, rol del sistema) de la sesión, o None."""
    if "ent_usuario" not in g:
        g.ent_usuario = None
        if session.get("usuario"):
            conn = get_connection()
            cursor = conn.cursor(dictionary=True)
            try:
                cursor.execute("SELECT id, usuario, rol FROM usuarios WHERE usuario = %s", (session["usuario"],))
                g.ent_usuario = cursor.fetchone()
            finally:
                cursor.close()
                conn.close()
    return g.ent_usuario


def leer_roles(cursor, usuario_id, solo_activos=True):
    """{rol: bodega_id} del usuario. Vacío si la migración todavía no está aplicada."""
    try:
        cursor.execute("SELECT rol, bodega_id FROM ent_usuario_rol WHERE usuario_id = %s" + (" AND activo = 1" if solo_activos else ""),
                       (usuario_id,))
        return {r[0]: r[1] for r in cursor.fetchall()}
    except ProgrammingError as e:
        if tabla_faltante(e):
            return {}
        raise


def roles_actuales():
    """{rol: bodega_id} activos del usuario de la sesión (una consulta por petición)."""
    if "ent_roles" not in g:
        u = usuario_actual()
        g.ent_roles = {}
        if u:
            conn = get_connection()
            cursor = conn.cursor()
            try:
                g.ent_roles = leer_roles(cursor, u["id"])
            finally:
                cursor.close()
                conn.close()
    return g.ent_roles


def requiere_rol(vista, *roles, admin_sistema=False):
    """Pantalla del módulo: permiso de la pantalla (accesos_vistas) + uno de los roles. admin_sistema: también la abre el
    administrador del sistema (usuarios.rol = 'admin'), para poder asignar los primeros roles."""
    def decorador(f):
        @wraps(f)
        def interna(*args, **kwargs):
            u = usuario_actual()
            if set(roles_actuales()) & set(roles) or (admin_sistema and u and u["rol"] == "admin"):
                return f(*args, **kwargs)
            from router.bitacora import registrar
            registrar("denegado")
            nombres = " o ".join(ROLES[r][0] for r in roles)
            flash(f"⚠️ Esta pantalla de Entregas es para el rol {nombres}. Pida al administrador que se lo asigne.", "warning")
            return redirect(url_for("inicio"))
        return requiere_acceso(vista)(interna)
    return decorador


def sincronizar_accesos(cursor, usuario, roles):
    """Deja habilitadas en accesos_vistas exactamente las pantallas de Entregas de los roles dados (las demás pantallas del
    sistema no se tocan)."""
    habilitar = set().union(*(VISTAS_DE_ROL[r] for r in roles)) if roles else set()
    for vista in VISTAS:
        cursor.execute("INSERT INTO accesos_vistas (usuario, vista, permitido) VALUES (%s, %s, %s) "
                       "ON DUPLICATE KEY UPDATE permitido = VALUES(permitido)", (usuario, vista, 1 if vista in habilitar else 0))


def historial_en_entregas(cursor, usuario_id):
    """Cuántos registros del módulo hacen referencia al usuario (envíos, rutas, facturas, eventos, líneas). Si hay, no se puede
    borrar el usuario: se le quitan los roles."""
    consultas = ["SELECT COUNT(*) FROM ent_envios WHERE vendedor_id = %s",
                 "SELECT COUNT(*) FROM ent_rutas WHERE mensajero_id = %s",
                 "SELECT COUNT(*) FROM ent_envio_facturas WHERE registrada_por = %s",
                 "SELECT COUNT(*) FROM ent_envio_eventos WHERE usuario_id = %s",
                 "SELECT COUNT(*) FROM ent_envio_lineas WHERE existencia_por = %s OR actualizado_por = %s"]
    total = 0
    try:
        for sql in consultas:
            cursor.execute(sql, (usuario_id,) * sql.count("%s"))
            total += _primero(cursor)
    except ProgrammingError as e:
        if tabla_faltante(e):
            return 0
        raise
    return total


def _primero(cursor):
    fila = cursor.fetchone()
    return list(fila.values())[0] if isinstance(fila, dict) else fila[0]


def borrar_roles(cursor, usuario_id):
    try:
        cursor.execute("DELETE FROM ent_usuario_rol WHERE usuario_id = %s", (usuario_id,))
    except ProgrammingError as e:
        if not tabla_faltante(e):
            raise


def registrar_evento(cursor, envio_id, usuario_id, a_estado, de_estado=None, campo="estado", linea_id=None, comentario=None):
    """Una fila del historial (ent_envio_eventos). La usa la función única de transiciones (regla 6.1, partes 6 en adelante)
    y la creación del envío."""
    cursor.execute("INSERT INTO ent_envio_eventos (envio_id, linea_id, usuario_id, campo, de_estado, a_estado, comentario) "
                   "VALUES (%s, %s, %s, %s, %s, %s, %s)",
                   (envio_id, linea_id, usuario_id, campo, de_estado, a_estado, (comentario or "")[:255] or None))


def requiere_modulo(f):
    """Pantallas del módulo que no son una opción del menú (detalle del envío): sesión iniciada y algún rol activo de Entregas.
    Lo que puede ver cada rol dentro de la pantalla lo decide la propia pantalla."""
    @wraps(f)
    def interna(*args, **kwargs):
        if not session.get("usuario"):
            return redirect(url_for("usuarios.login"))
        if not roles_actuales():
            flash("⚠️ No tiene un rol en Entregas.", "warning")
            return redirect(url_for("inicio"))
        return f(*args, **kwargs)
    return interna


@entregas_bp.context_processor
def contexto_entregas():
    """Pestañas del módulo: solo las pantallas de los roles del usuario."""
    mis = set(roles_actuales()) if session.get("usuario") else set()
    pendientes = {}
    if mis:
        try:
            from router.entregas_operacion import pendientes_actuales
            pendientes = pendientes_actuales()
        except Exception:   # los contadores nunca impiden abrir una pantalla
            pendientes = {}
    return {"ent_pestanas": [p for p in PANTALLAS if mis & set(p[5])], "ent_roles": ROLES, "ent_mis_roles": mis,
            "ent_pendientes": pendientes}


# -----------------------------
# Usuarios y roles (parte 2)
# -----------------------------
def migracion_aplicada(cursor):
    cursor.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = 'ent_usuario_rol'")
    return _primero(cursor) > 0


def leer_bodegas(cursor):
    try:
        cursor.execute("SELECT id, nombre, nombre_en_orden, requiere_traslado FROM ent_bodegas WHERE activo = 1 ORDER BY nombre")
        return cursor.fetchall()
    except ProgrammingError as e:
        if tabla_faltante(e):
            return []
        raise


@entregas_bp.route("/entregas/usuarios")
@requiere_rol("entregas_usuarios", "admin", admin_sistema=True)
def usuarios():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not migracion_aplicada(cursor):
            return render_template("entregas_usuarios.html", sin_migracion=True, usuarios=[], bodegas={})
        cursor.execute("SELECT id, usuario, rol FROM usuarios ORDER BY usuario")
        lista = cursor.fetchall()
        cursor.execute("SELECT usuario_id, rol, bodega_id, activo FROM ent_usuario_rol")
        filas = cursor.fetchall()
        bodegas = {b["id"]: b for b in leer_bodegas(cursor)}
    finally:
        cursor.close()
        conn.close()
    for u in lista:
        mias = [f for f in filas if f["usuario_id"] == u["id"]]
        u["roles"] = {f["rol"]: f["bodega_id"] for f in mias}
        u["activo"] = all(f["activo"] for f in mias) if mias else True
    con_roles = [u for u in lista if u["roles"]]
    return render_template("entregas_usuarios.html", sin_migracion=False, usuarios=lista, con_roles=con_roles, bodegas=bodegas)


@entregas_bp.route("/entregas/usuarios/<int:id_>", methods=["GET", "POST"])
@requiere_rol("entregas_usuarios", "admin", admin_sistema=True)
def usuario_roles(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if not migracion_aplicada(cursor):
            flash("Falta aplicar la migración docs/sql/ent_entregas.sql.", "warning")
            return redirect(url_for("entregas.usuarios"))
        cursor.execute("SELECT id, usuario, rol FROM usuarios WHERE id = %s", (id_,))
        u = cursor.fetchone()
        if not u:
            flash("El usuario no existe.", "warning")
            return redirect(url_for("entregas.usuarios"))
        bodegas = leer_bodegas(cursor)
        cursor.execute("SELECT rol, bodega_id, activo FROM ent_usuario_rol WHERE usuario_id = %s", (id_,))
        actuales = cursor.fetchall()
        roles = {f["rol"]: f["bodega_id"] for f in actuales}
        activo = all(f["activo"] for f in actuales) if actuales else True

        if request.method == "POST":
            marcados = [r for r in ROLES if r in request.form.getlist("rol")]
            activo = request.form.get("activo") == "1"
            bodega = request.form.get("bodega_id", type=int)
            ids_bodegas = {b["id"] for b in bodegas}
            avisos = []
            if "bodega" in marcados:
                if bodega not in ids_bodegas:
                    if bodegas:
                        flash("Escoja la bodega del bodeguero.", "warning")
                        return render_template("entregas_usuario_roles.html", u=u, roles={r: (bodega if r == "bodega" else None)
                                               for r in marcados}, activo=activo, bodegas=bodegas)
                    bodega = None
                    avisos.append("todavía no hay bodegas en el catálogo: asígnele su bodega cuando se carguen (parte 3); "
                                  "mientras tanto no verá líneas")
            es_yo = u["usuario"].lower() == (session.get("usuario") or "").lower()
            if es_yo and "admin" in roles and "admin" not in marcados and u["rol"] != "admin":
                flash("No puede quitarse a sí mismo el rol de administrador de Entregas.", "warning")
                return redirect(url_for("entregas.usuario_roles", id_=id_))
            cursor.execute("DELETE FROM ent_usuario_rol WHERE usuario_id = %s", (id_,))
            for r in marcados:
                cursor.execute("INSERT INTO ent_usuario_rol (usuario_id, rol, bodega_id, activo) VALUES (%s, %s, %s, %s)",
                               (id_, r, bodega if r == "bodega" else None, 1 if activo else 0))
            sincronizar_accesos(cursor, u["usuario"], marcados if activo else [])
            conn.commit()
            nombres = ", ".join(ROLES[r][0] for r in marcados) or "ninguno"
            flash(f"✅ Roles de {u['usuario']} en Entregas: {nombres}{'' if activo or not marcados else ' (suspendido)'}."
                  + (" Aviso: " + "; ".join(avisos) + "." if avisos else ""), "warning" if avisos else "success")
            return redirect(url_for("entregas.usuarios"))
    finally:
        cursor.close()
        conn.close()
    return render_template("entregas_usuario_roles.html", u=u, roles=roles, activo=activo, bodegas=bodegas)


# -----------------------------
# Pantallas de cada rol (se construyen en las partes siguientes)
# -----------------------------
def _registrar_pendiente(vista):
    p = VISTAS[vista]

    @requiere_rol(vista, *p[5])
    def ver():
        return render_template("entregas_pendiente.html", p=p)
    entregas_bp.add_url_rule(p[2], endpoint=vista, view_func=ver)


# Pantallas que ya tienen su propio código (no se registran como «en construcción»)
CONSTRUIDAS = {"entregas_usuarios",                                   # parte 2, aquí
               "entregas_catalogos", "entregas_bodegas_traslado",     # parte 3, router/entregas_catalogos.py
               "entregas_subir_orden",                                # parte 4, router/entregas_orden.py
               "entregas_bodega", "entregas_coordinador", "entregas_seguimiento",   # parte 6, router/entregas_operacion.py
               "entregas_mis_envios",                                 # parte 7, router/entregas_vendedor.py
               "entregas_empaque",                                    # partes 9 y 10, router/entregas_paqueteria.py
               "entregas_rutas", "entregas_mis_rutas"}                # parte 11, router/entregas_rutas.py
for _vista in VISTAS:
    if _vista not in CONSTRUIDAS:
        _registrar_pendiente(_vista)
