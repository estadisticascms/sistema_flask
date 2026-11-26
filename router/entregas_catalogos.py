"""Entregas · parte 3: catálogos de bodegas, transportistas y horarios, y configuración de la alerta de tiempo.

- Catálogos y configuración (/entregas/catalogos): rol Administrador de Entregas.
- Bodegas que aplican traslado (/entregas/bodegas_traslado): roles Paquetería y Administrador. Es lo único que paquetería
  cambia de las bodegas (regla 6.3: requiere_traslado solo paquetería o admin). El cambio aplica a las órdenes que se suban
  después; las líneas ya creadas conservan su responsable.

La bodega destino del traslado no es fija: depende del tipo de producto, y la escoge el coordinador en cada línea al aplicar el
traslado (parte 8). Por eso el catálogo solo dice si la bodega requiere traslado.

Bodega de vendedor («solo factura»): la mercadería ya se le entregó al cliente; de sus líneas solo se envía la factura junto con el
resto del paquete. No lleva traslado ni bodeguero.

Las bodegas no se borran si ya tienen líneas o bodegueros: se desactivan. Quién cambió qué y cuándo queda en la bitácora de
accesos (cada POST se registra con usuario, fecha y ruta).
"""
import re
from datetime import time, timedelta

from flask import Blueprint, flash, redirect, render_template, request, url_for

from router.entregas import contexto_entregas, requiere_rol
from utils.db import get_connection

entregas_catalogos_bp = Blueprint("entregas_catalogos", __name__)
entregas_catalogos_bp.context_processor(contexto_entregas)

TIPOS_TRANSPORTISTA = {"empresa": "Empresa de envíos", "bus": "Bus", "retira_cliente": "Cliente retira", "reparto_propio": "Reparto propio"}
DIAS = [("L", "Lunes"), ("M", "Martes"), ("X", "Miércoles"), ("J", "Jueves"), ("V", "Viernes"), ("S", "Sábado"), ("D", "Domingo")]

# Configuración de la alerta de tiempo (sección 4.8). Los [PENDIENTE] los da la empresa: sin valor quedan vacíos.
CONFIGURACION = [
    ("hora_corte", "Hora de corte", "hora", "Órdenes enviadas después de esta hora con poco tiempo se marcan en rojo", False),
    ("minutos_urgencia", "Minutos de urgencia", "minutos", "Después de la hora de corte: rojo si faltan menos de estos minutos para la hora requerida", False),
    ("minutos_preparacion", "Minutos de preparación", "minutos", "Tiempo estimado para preparar una orden en bodega", True),
    ("minutos_traslado_bodega", "Minutos de traslado entre bodegas", "minutos", "Se suma si alguna línea requiere traslado", True),
    ("minutos_facturacion", "Minutos de facturación", "minutos", "Tiempo estimado para modificar la orden y facturar", True),
    ("minutos_empaque", "Minutos de empaque", "minutos", "Tiempo estimado para empacar y etiquetar", True),
    ("margen_amarillo", "Margen para amarillo (minutos)", "minutos", "Amarillo (justo) si el margen es menor que esto", True),
    ("bodega_factura_traslado", "Factura de líneas con traslado", "opcion",
     "De qué bodega sale la factura después del traslado. Sin definir se usa la destino, como dice la especificación", True),
]
OPCIONES_CONFIG = {"bodega_factura_traslado": [("destino", "Bodega destino (la que escoge el coordinador)"), ("origen", "Bodega de origen")]}
TABS = ("bodegas", "transportistas", "configuracion")


# -----------------------------
# Utilidades
# -----------------------------
def texto(valor, largo):
    return " ".join((valor or "").split())[:largo]


def hora_texto(valor):
    """TIME de MySQL (llega como timedelta) o time -> 'HH:MM'."""
    if valor is None:
        return ""
    if isinstance(valor, timedelta):
        minutos = int(valor.total_seconds()) // 60
        return f"{minutos // 60:02d}:{minutos % 60:02d}"
    if isinstance(valor, time):
        return valor.strftime("%H:%M")
    return str(valor)[:5]


def leer_hora(valor):
    """'HH:MM' -> 'HH:MM:00', '' -> None, inválida -> 'error'."""
    valor = (valor or "").strip()
    if not valor:
        return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", valor)
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        return "error"
    return f"{int(m.group(1)):02d}:{m.group(2)}:00"


def leer_entero(valor):
    """'' -> None, entero >= 0 -> int, otro -> 'error'."""
    valor = (valor or "").strip()
    if not valor:
        return None
    return int(valor) if valor.isdigit() else "error"


def con_cursor(funcion):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        resultado = funcion(cursor)
        conn.commit()
        return resultado
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()


# -----------------------------
# Datos
# -----------------------------
def leer_bodegas(cursor):
    cursor.execute("""
        SELECT b.*, (SELECT COUNT(*) FROM ent_envio_lineas l WHERE l.bodega_id = b.id OR l.bodega_prepara_id = b.id) AS lineas
        FROM ent_bodegas b
        ORDER BY b.activo DESC, b.nombre""")
    bodegas = cursor.fetchall()
    cursor.execute("""SELECT r.bodega_id, u.usuario, r.activo FROM ent_usuario_rol r JOIN usuarios u ON u.id = r.usuario_id
                      WHERE r.rol = 'bodega' ORDER BY u.usuario""")
    bodegueros = cursor.fetchall()
    for b in bodegas:
        b["bodegueros"] = [x["usuario"] + ("" if x["activo"] else " (suspendido)") for x in bodegueros if x["bodega_id"] == b["id"]]
    sin_bodega = [x["usuario"] for x in bodegueros if x["bodega_id"] is None]
    return bodegas, sin_bodega


def bodegas_de_existencias(cursor):
    """Nombres de bodega de las existencias cargadas que todavía no están en el catálogo (para agregarlas sin escribirlas)."""
    cursor.execute("""SELECT DISTINCT e.Bodega AS nombre FROM existencias_bodega e
                      WHERE e.Bodega IS NOT NULL AND e.Bodega <> ''
                        AND NOT EXISTS (SELECT 1 FROM ent_bodegas b WHERE b.nombre_en_orden = e.Bodega)
                      ORDER BY e.Bodega""")
    return [r["nombre"] for r in cursor.fetchall()]


def leer_transportistas(cursor):
    cursor.execute("SELECT * FROM ent_transportistas ORDER BY activo DESC, nombre")
    lista = cursor.fetchall()
    cursor.execute("SELECT * FROM ent_transportista_horarios ORDER BY transportista_id, hora_salida, destino")
    horarios = cursor.fetchall()
    cursor.execute("SELECT transportista_id, COUNT(*) AS n FROM ent_envios GROUP BY transportista_id")
    usados = {r["transportista_id"]: r["n"] for r in cursor.fetchall()}
    cursor.execute("SELECT transportista_id, COUNT(*) AS n FROM ent_ruta_paradas GROUP BY transportista_id")
    for r in cursor.fetchall():
        usados[r["transportista_id"]] = usados.get(r["transportista_id"], 0) + r["n"]
    for t in lista:
        t["limite"] = hora_texto(t["hora_limite_recepcion"])
        t["horarios"] = [{**h, "salida": hora_texto(h["hora_salida"])} for h in horarios if h["transportista_id"] == t["id"]]
        t["usado"] = usados.get(t["id"], 0)
    return lista


def leer_configuracion(cursor):
    cursor.execute("SELECT clave, valor FROM ent_configuracion")
    return {r["clave"]: r["valor"] for r in cursor.fetchall()}


# -----------------------------
# Catálogos y configuración (administrador)
# -----------------------------
@entregas_catalogos_bp.route("/entregas/catalogos")
@requiere_rol("entregas_catalogos", "admin")
def catalogos():
    tab = request.args.get("tab") if request.args.get("tab") in TABS else "bodegas"

    def leer(cursor):
        bodegas, sin_bodega = leer_bodegas(cursor)
        return {"bodegas": bodegas, "sin_bodega": sin_bodega, "de_existencias": bodegas_de_existencias(cursor),
                "transportistas": leer_transportistas(cursor), "config": leer_configuracion(cursor)}
    d = con_cursor(leer)
    from router.entregas_horario import CLAVES, valores as valores_horario
    hv = valores_horario(d["config"])
    d["horario_laboral"] = {k: (hv[k].split("-") + ["", ""])[:2] if hv[k] else ["", ""] for k in CLAVES[:3]}
    d["horario_laboral"]["feriados"] = hv["feriados"]
    editar_b = next((b for b in d["bodegas"] if str(b["id"]) == request.args.get("bodega")), None)
    editar_t = next((t for t in d["transportistas"] if str(t["id"]) == request.args.get("transportista")), None)
    return render_template("entregas_catalogos.html", tab=tab, d=d, editar_b=editar_b, editar_t=editar_t,
                           nueva_b=request.args.get("bodega") == "nueva", nuevo_t=request.args.get("transportista") == "nuevo",
                           tipos=TIPOS_TRANSPORTISTA, dias=DIAS, configuracion=CONFIGURACION, opciones=OPCIONES_CONFIG)


@entregas_catalogos_bp.route("/entregas/catalogos/bodega", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def guardar_bodega():
    id_ = request.form.get("id", type=int)
    nombre_orden = texto(request.form.get("nombre_en_orden"), 60).upper()
    nombre = texto(request.form.get("nombre"), 60) or nombre_orden
    activo = request.form.get("activo") == "1"
    solo_factura = request.form.get("solo_factura") == "1"
    requiere = request.form.get("requiere_traslado") == "1" and not solo_factura
    volver = url_for("entregas_catalogos.catalogos", tab="bodegas", bodega=id_ or "nueva")

    def guardar(cursor):
        if not nombre_orden:
            return "Escriba el nombre de la bodega tal como sale en la orden (por ejemplo BODEGA #1)."
        cursor.execute("SELECT id FROM ent_bodegas WHERE nombre_en_orden = %s AND id <> %s", (nombre_orden, id_ or 0))
        if cursor.fetchone():
            return f"Ya existe la bodega «{nombre_orden}»."
        if id_:
            cursor.execute("UPDATE ent_bodegas SET nombre_en_orden = %s, nombre = %s, requiere_traslado = %s, solo_factura = %s, activo = %s "
                           "WHERE id = %s", (nombre_orden, nombre, int(requiere), int(solo_factura), int(activo), id_))
        else:
            cursor.execute("INSERT INTO ent_bodegas (nombre_en_orden, nombre, requiere_traslado, solo_factura, activo) VALUES (%s, %s, %s, %s, %s)",
                           (nombre_orden, nombre, int(requiere), int(solo_factura), int(activo)))
        return None
    error = con_cursor(guardar)
    if error:
        flash(error, "warning")
        return redirect(volver)
    flash(f"✅ Bodega {nombre_orden} guardada.", "success")
    return redirect(url_for("entregas_catalogos.catalogos", tab="bodegas"))


@entregas_catalogos_bp.route("/entregas/catalogos/bodegas_existencias", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def agregar_de_existencias():
    elegidas = [texto(n, 60).upper() for n in request.form.getlist("nombre") if texto(n, 60)]

    def agregar(cursor):
        disponibles = {n.upper() for n in bodegas_de_existencias(cursor)}
        nuevas = [n for n in elegidas if n in disponibles]
        for n in nuevas:
            cursor.execute("INSERT INTO ent_bodegas (nombre_en_orden, nombre) VALUES (%s, %s)", (n, n))
        return nuevas
    nuevas = con_cursor(agregar)
    flash(f"✅ Se agregaron {len(nuevas)} bodega(s). Revise el nombre corto y cuáles aplican traslado." if nuevas
          else "No se agregó ninguna bodega.", "success" if nuevas else "warning")
    return redirect(url_for("entregas_catalogos.catalogos", tab="bodegas"))


@entregas_catalogos_bp.route("/entregas/catalogos/bodega/<int:id_>/eliminar", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def eliminar_bodega(id_):
    def eliminar(cursor):
        bodegas, _ = leer_bodegas(cursor)
        b = next((x for x in bodegas if x["id"] == id_), None)
        if not b:
            return "La bodega no existe."
        if b["lineas"] or b["bodegueros"]:
            return f"No se puede borrar {b['nombre_en_orden']}: tiene órdenes o bodegueros. Desactívela."
        cursor.execute("DELETE FROM ent_bodegas WHERE id = %s", (id_,))
        return None
    error = con_cursor(eliminar)
    flash(error or "Bodega borrada.", "warning" if error else "info")
    return redirect(url_for("entregas_catalogos.catalogos", tab="bodegas"))


@entregas_catalogos_bp.route("/entregas/catalogos/transportista", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def guardar_transportista():
    f = request.form
    id_ = f.get("id", type=int)
    datos = {"nombre": texto(f.get("nombre"), 100), "tipo": f.get("tipo") if f.get("tipo") in TIPOS_TRANSPORTISTA else None,
             "direccion": texto(f.get("direccion"), 255) or None, "mapa_url": texto(f.get("mapa_url"), 255) or None,
             "telefono": texto(f.get("telefono"), 30) or None, "minutos_traslado": leer_entero(f.get("minutos_traslado")),
             "hora_limite_recepcion": leer_hora(f.get("hora_limite_recepcion")), "activo": int(f.get("activo") == "1")}
    error = None
    if not datos["nombre"]:
        error = "Escriba el nombre del transportista."
    elif not datos["tipo"]:
        error = "Escoja el tipo de transportista."
    elif datos["minutos_traslado"] == "error":
        error = "Los minutos de traslado deben ser un número entero."
    elif datos["hora_limite_recepcion"] == "error":
        error = "La hora límite de recepción debe ser HH:MM (por ejemplo 16:30)."
    elif datos["mapa_url"] and not re.match(r"https?://", datos["mapa_url"]):
        error = "El enlace del mapa debe empezar con http:// o https://."
    if error:
        flash(error, "warning")
        return redirect(url_for("entregas_catalogos.catalogos", tab="transportistas", transportista=id_ or "nuevo"))

    def guardar(cursor):
        columnas = list(datos)
        if id_:
            cursor.execute(f"UPDATE ent_transportistas SET {', '.join(c + ' = %s' for c in columnas)} WHERE id = %s",
                           [datos[c] for c in columnas] + [id_])
        else:
            cursor.execute(f"INSERT INTO ent_transportistas ({', '.join(columnas)}) VALUES ({', '.join(['%s'] * len(columnas))})",
                           [datos[c] for c in columnas])
    con_cursor(guardar)
    flash(f"✅ Transportista {datos['nombre']} guardado.", "success")
    return redirect(url_for("entregas_catalogos.catalogos", tab="transportistas"))


@entregas_catalogos_bp.route("/entregas/catalogos/transportista/<int:id_>/eliminar", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def eliminar_transportista(id_):
    def eliminar(cursor):
        t = next((x for x in leer_transportistas(cursor) if x["id"] == id_), None)
        if not t:
            return "El transportista no existe."
        if t["usado"]:
            return f"No se puede borrar {t['nombre']}: ya tiene envíos o paradas. Desactívelo."
        cursor.execute("DELETE FROM ent_transportista_horarios WHERE transportista_id = %s", (id_,))
        cursor.execute("DELETE FROM ent_transportistas WHERE id = %s", (id_,))
        return None
    error = con_cursor(eliminar)
    flash(error or "Transportista borrado con sus horarios.", "warning" if error else "info")
    return redirect(url_for("entregas_catalogos.catalogos", tab="transportistas"))


@entregas_catalogos_bp.route("/entregas/catalogos/transportista/<int:id_>/horario", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def agregar_horario(id_):
    destino = texto(request.form.get("destino"), 100)
    salida = leer_hora(request.form.get("hora_salida"))
    dias = ",".join(d for d, _ in DIAS if d in request.form.getlist("dias")) or None
    volver = url_for("entregas_catalogos.catalogos", tab="transportistas", transportista=id_)
    if not destino or salida == "error":       # sin hora: sale a cualquier hora hacia ese destino
        flash("Escriba el destino y, si tiene horario fijo, la hora de salida (HH:MM).", "warning")
        return redirect(volver)

    def agregar(cursor):
        cursor.execute("SELECT id FROM ent_transportistas WHERE id = %s", (id_,))
        if not cursor.fetchone():
            return False
        cursor.execute("INSERT INTO ent_transportista_horarios (transportista_id, destino, hora_salida, dias) VALUES (%s, %s, %s, %s)",
                       (id_, destino, salida, dias))
        return True
    ok = con_cursor(agregar)
    flash(("✅ Horario agregado." if salida else f"✅ {destino}: sale a cualquier hora.") if ok else "El transportista no existe.", "success")
    return redirect(volver)


@entregas_catalogos_bp.route("/entregas/catalogos/horario/<int:id_>/eliminar", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def eliminar_horario(id_):
    def eliminar(cursor):
        cursor.execute("SELECT transportista_id FROM ent_transportista_horarios WHERE id = %s", (id_,))
        fila = cursor.fetchone()
        cursor.execute("DELETE FROM ent_transportista_horarios WHERE id = %s", (id_,))
        return fila["transportista_id"] if fila else None
    tid = con_cursor(eliminar)
    flash("Horario borrado.", "info")
    return redirect(url_for("entregas_catalogos.catalogos", tab="transportistas", transportista=tid))


@entregas_catalogos_bp.route("/entregas/catalogos/horario_laboral", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def guardar_horario_laboral():
    """Horario laboral y feriados con que se miden los tiempos en los reportes (router/entregas_horario.py)."""
    from router.entregas_horario import CLAVES, leer_feriados, leer_rango
    valores, errores = {}, []
    for clave, nombre in zip(CLAVES[:3], ("Lunes a viernes", "Sábado", "Domingo")):
        desde, hasta = request.form.get(clave + "_desde", "").strip(), request.form.get(clave + "_hasta", "").strip()
        rango = leer_rango(f"{desde}-{hasta}" if desde or hasta else "")
        if rango == "error":
            errores.append(f"{nombre}: escriba la hora de entrada y la de salida (la salida después de la entrada).")
        valores[clave] = f"{rango[0].strftime('%H:%M')}-{rango[1].strftime('%H:%M')}" if rango and rango != "error" else ""
    if not valores[CLAVES[0]] and not any(e.startswith("Lunes") for e in errores):
        errores.append("Lunes a viernes es obligatorio.")
    _, _, feriados = leer_feriados(request.form.get("feriados", ""))
    if feriados == "error":
        errores.append("Feriados: use dd/mm (todos los años) o dd/mm/aaaa (solo esa fecha), separados por comas.")
    valores["feriados"] = feriados
    if errores:
        flash("No se guardó el horario. " + " ".join(errores), "warning")
        return redirect(url_for("entregas_catalogos.catalogos", tab="configuracion"))

    def guardar(cursor):
        for clave, valor in valores.items():
            cursor.execute("INSERT INTO ent_configuracion (clave, valor) VALUES (%s, %s) ON DUPLICATE KEY UPDATE valor = VALUES(valor)",
                           (clave, valor))
    con_cursor(guardar)
    flash("✅ Horario laboral guardado.", "success")
    return redirect(url_for("entregas_catalogos.catalogos", tab="configuracion"))


@entregas_catalogos_bp.route("/entregas/catalogos/configuracion", methods=["POST"])
@requiere_rol("entregas_catalogos", "admin")
def guardar_configuracion():
    valores, errores = {}, []
    for clave, nombre, tipo, _, _ in CONFIGURACION:
        crudo = request.form.get(clave, "")
        if tipo == "hora":
            v = leer_hora(crudo)
            valores[clave] = v[:5] if v and v != "error" else v
        elif tipo == "opcion":
            valores[clave] = crudo if crudo in dict(OPCIONES_CONFIG[clave]) else None
        else:
            valores[clave] = leer_entero(crudo)
        if valores[clave] == "error":
            errores.append(f"{nombre}: {'use HH:MM' if tipo == 'hora' else 'escriba un número entero de minutos'}.")
        elif valores[clave] is None and clave in ("hora_corte", "minutos_urgencia"):
            errores.append(f"{nombre} es obligatorio.")
    if errores:
        flash("No se guardó. " + " ".join(errores), "warning")
        return redirect(url_for("entregas_catalogos.catalogos", tab="configuracion"))

    def guardar(cursor):
        for clave, valor in valores.items():
            if valor is None:   # vacío = todavía no definido (pendiente)
                cursor.execute("DELETE FROM ent_configuracion WHERE clave = %s", (clave,))
            else:
                cursor.execute("INSERT INTO ent_configuracion (clave, valor) VALUES (%s, %s) ON DUPLICATE KEY UPDATE valor = VALUES(valor)",
                               (clave, str(valor)))
    con_cursor(guardar)
    flash("✅ Configuración guardada.", "success")
    return redirect(url_for("entregas_catalogos.catalogos", tab="configuracion"))


# -----------------------------
# Bodegas que aplican traslado (paquetería y administrador)
# -----------------------------
@entregas_catalogos_bp.route("/entregas/bodegas_traslado", methods=["GET", "POST"])
@requiere_rol("entregas_bodegas_traslado", "paqueteria", "admin")
def bodegas_traslado():
    if request.method == "POST":
        id_ = request.form.get("id", type=int)
        requiere = request.form.get("requiere_traslado") == "1"

        def guardar(cursor):
            bodegas, _ = leer_bodegas(cursor)
            b = next((x for x in bodegas if x["id"] == id_ and x["activo"]), None)
            if not b:
                return None, "La bodega no existe o está desactivada."
            if b["solo_factura"] and requiere:
                return None, "Es una bodega de vendedor (solo factura): no lleva traslado."
            cursor.execute("UPDATE ent_bodegas SET requiere_traslado = %s WHERE id = %s", (int(requiere), id_))
            return b, None
        b, error = con_cursor(guardar)
        if error:
            flash(error, "warning")
        else:
            flash(f"✅ {b['nombre']}: {'requiere traslado' if requiere else 'sin traslado'}. Aplica a las órdenes que se suban desde ahora.",
                  "success")
        return redirect(url_for("entregas_catalogos.bodegas_traslado"))
    bodegas, _ = con_cursor(leer_bodegas)
    return render_template("entregas_bodegas_traslado.html", bodegas=[b for b in bodegas if b["activo"] and not b["solo_factura"]])
