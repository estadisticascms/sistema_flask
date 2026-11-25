"""Entregas · parte 4: subir la orden (PDF → lectura → control de cuadre → vista previa → transporte y hora → guardar) y
detalle del envío con su historial.

Flujo de «Subir orden» (rol Vendedor):
1. El vendedor sube el PDF. Se lee (router/entregas_pdf.py); si no cuadra o no se reconoce, se rechaza con el motivo.
2. Vista previa: lo leído, la bodega de cada línea según el catálogo (y si requiere traslado), y lo que impide enviarla
   (bodega desconocida o desactivada, pedido ya subido). El PDF queda en una carpeta temporal identificada por un código.
3. El vendedor agrega transporte, destino, hora requerida, prioridad y notas (forma de pago y flete: opcionales, [PENDIENTE])
   y envía. El servidor vuelve a leer el PDF temporal (no confía en lo que venga del navegador), valida todo otra vez, guarda el
   envío en ORDEN_ENVIADA con sus líneas y guarda el PDF original.

Cada línea se crea con su responsable de la existencia (coordinador si su bodega requiere traslado; si no, bodeguero) y su estado
de traslado (PENDIENTE o NO_REQUIERE), según el catálogo en ese momento. La bodega que prepara es la misma de la línea; en las
líneas con traslado queda vacía hasta que el coordinador escoge el destino al aplicar el traslado (depende del tipo de producto).

Alerta de tiempo (parte 5, router/entregas_alerta.py): se calcula en vivo mientras el vendedor escoge transporte y hora, otra vez
en el servidor al enviar, y se recalcula al abrir el detalle. Avisa pero no bloquea: en rojo el vendedor debe marcar que confirma la
hora (queda riesgo_confirmado = 1) o cambiarla.

Bodegas de vendedor («solo factura» en el catálogo): la mercadería ya se entregó al cliente. Sus líneas se graban con la existencia
confirmada y como ya entregadas (no aparecen en ninguna cola de bodega ni de retiro), pero su factura se pide en Facturar y sale en la
etiqueta junto con las demás.

Bodegas fuera del catálogo (desconocidas o desactivadas): suelen ser productos que no se envían (por ejemplo, bodegas de
vendedores). El vendedor marca «estas líneas no se envían» y se graba el resto; las excluidas quedan anotadas en el historial y
como ajuste en Facturar. El cuadre se sigue verificando contra el PDF completo. No se excluyen solas: una bodega real recién creada
o mal escrita dejaría producto sin enviar sin que nadie lo note.
"""
import os
import re
import time
import uuid
from datetime import datetime

from flask import Blueprint, abort, current_app, flash, jsonify, redirect, render_template, request, send_file, url_for
from mysql.connector.errors import IntegrityError

from router.entregas import (ROLES, contexto_entregas, registrar_evento, requiere_modulo, requiere_rol, roles_actuales,
                             usuario_actual)
from router.entregas_avisos import avisar, avisar_sin_repetir, destinatarios, responsables_linea
from router.entregas_flujo import recalcular_estado
from router.entregas_alerta import ahora_nicaragua, alerta_de_envio, calcular as calcular_alerta, leer_configuracion
from router.entregas_catalogos import TIPOS_TRANSPORTISTA, hora_texto, texto
from router.entregas_pdf import ErrorLectura, leer_orden
from utils.db import get_connection

entregas_orden_bp = Blueprint("entregas_orden", __name__)
entregas_orden_bp.context_processor(contexto_entregas)

TAMANO_MAXIMO = 10 * 1024 * 1024        # 10 MB
HORAS_TEMPORAL = 24                     # los PDF leídos y no enviados se borran después de esto
ESTADOS = ["ORDEN_ENVIADA", "EN_REVISION", "POR_FACTURAR", "FACTURADA", "LISTA_PARA_RETIRO", "RETIRADA", "EMPACADA", "EN_RUTA",
           "ENTREGADA", "ANULADA"]
NOMBRE_ESTADO = {"ORDEN_ENVIADA": "Orden enviada", "EN_REVISION": "En revisión", "POR_FACTURAR": "Por facturar", "FACTURADA": "Facturada",
                 "LISTA_PARA_RETIRO": "Lista para retiro", "RETIRADA": "Retirada", "EMPACADA": "Empacada", "EN_RUTA": "En ruta",
                 "ENTREGADA": "Entregada", "ANULADA": "Anulada"}


# -----------------------------
# Archivos
# -----------------------------
def carpeta(*partes):
    ruta = os.path.join(current_app.config["UPLOAD_FOLDER"], "entregas", *partes)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def ruta_temporal(token):
    if not re.fullmatch(r"[0-9a-f]{32}", token or ""):
        return None
    return os.path.join(carpeta("tmp"), f"{token}.pdf")


def limpiar_temporales():
    limite = time.time() - HORAS_TEMPORAL * 3600
    for nombre in os.listdir(carpeta("tmp")):
        ruta = os.path.join(carpeta("tmp"), nombre)
        try:
            if os.path.getmtime(ruta) < limite:
                os.remove(ruta)
        except OSError:
            pass


def ruta_absoluta(relativa):
    """pdf_url guardado (relativo a la carpeta de subidas) -> ruta en disco, sin salir de uploads/entregas."""
    base = os.path.realpath(os.path.join(current_app.config["UPLOAD_FOLDER"], "entregas"))
    ruta = os.path.realpath(os.path.join(current_app.config["UPLOAD_FOLDER"], relativa or ""))
    return ruta if ruta.startswith(base + os.sep) and os.path.isfile(ruta) else None


# -----------------------------
# Revisión de la orden contra el catálogo
# -----------------------------
def lineas_que_van(orden):
    """Líneas con bodega válida del catálogo (las que se graban en el envío)."""
    return [l for l in orden["lineas"] if l.get("valida")]


def texto_excluidas(orden):
    """«MICHAEL LOPEZ BLANCO: líneas 3, 4, 5 (C$ 3,986.50)» por cada bodega fuera del catálogo."""
    grupos = {}
    for l in orden["lineas"]:
        if not l.get("valida"):
            g = grupos.setdefault(l["bodega"], [[], 0.0])
            g[0].append(str(l["linea"])); g[1] += l["subtotal"]
    return "; ".join(f"{b}: línea{'s' if len(n) > 1 else ''} {', '.join(n)} (C$ {m:,.2f})" for b, (n, m) in grupos.items())


def analizar(cursor, orden):
    """Agrega a cada línea su bodega del catálogo y devuelve (errores que impiden enviar, avisos). Las bodegas fuera del catálogo no
    van en «errores»: quedan en orden["errores_bodega"] y el vendedor decide si esas líneas no se envían."""
    errores, avisos = [], []
    cursor.execute("SELECT * FROM ent_bodegas")
    bodegas = {b["nombre_en_orden"]: b for b in cursor.fetchall()}
    cursor.execute("SELECT r.rol, r.bodega_id, u.usuario FROM ent_usuario_rol r JOIN usuarios u ON u.id = r.usuario_id "
                   "WHERE r.activo = 1 AND r.rol IN ('bodega', 'coordinador_bodega')")
    personas = cursor.fetchall()
    bodegueros = {}
    for p in personas:
        if p["rol"] == "bodega" and p["bodega_id"]:
            bodegueros.setdefault(p["bodega_id"], []).append(p["usuario"])
    hay_coordinador = any(p["rol"] == "coordinador_bodega" for p in personas)

    desconocidas, inactivas, sin_bodeguero = [], [], set()
    for l in orden["lineas"]:
        b = bodegas.get(l["bodega"].upper())
        l["b"] = b
        l["valida"] = bool(b and b["activo"])
        if not b:
            desconocidas.append(l["bodega"])
            continue
        if not b["activo"]:
            inactivas.append(l["bodega"])
            continue
        l["solo_factura"] = bool(b["solo_factura"])
        l["prepara_id"] = None if b["requiere_traslado"] and not b["solo_factura"] else b["id"]    # con traslado: lo define el coordinador
        l["responsable"] = "coordinador" if b["requiere_traslado"] and not b["solo_factura"] else "bodeguero"
        if b["solo_factura"]:
            continue
        if not b["requiere_traslado"] and not bodegueros.get(b["id"]):
            sin_bodeguero.add(b["nombre"])
    orden["errores_bodega"] = []
    if desconocidas:
        orden["errores_bodega"].append(f"Bodega que no está en el catálogo: {', '.join(sorted(set(desconocidas)))}.")
    if inactivas:
        orden["errores_bodega"].append(f"Bodega desactivada en el catálogo: {', '.join(sorted(set(inactivas)))}.")
    if (desconocidas or inactivas) and not lineas_que_van(orden):
        errores.append("Ninguna línea de la orden es de una bodega del catálogo: no hay nada que enviar. Si debe enviarse, avise al "
                       "administrador de Entregas para que agregue la bodega con ese nombre exacto.")
    cursor.execute("SELECT id, estado FROM ent_envios WHERE pedido_num = %s AND estado <> 'ANULADA'", (orden["pedido"],))
    previo = cursor.fetchone()
    if previo:
        errores.append(f"El pedido {orden['pedido']} ya se subió (envío #{previo['id']}, {NOMBRE_ESTADO[previo['estado']]}). "
                       f"Solo se puede volver a subir si el anterior se anula.")
    if sin_bodeguero:
        avisos.append(f"Sin bodeguero asignado: {', '.join(sorted(sin_bodeguero))}. Nadie verá esas líneas para prepararlas hasta "
                      f"que el administrador asigne uno.")
    if any(l.get("responsable") == "coordinador" for l in orden["lineas"]) and not hay_coordinador:
        avisos.append("La orden tiene líneas con traslado y no hay ningún coordinador de bodega activo.")
    if desconocidas or inactivas:
        nombres = sorted(set(desconocidas + inactivas))
        current_app.logger.warning("Entregas: pedido %s con bodegas no válidas: %s", orden["pedido"], nombres)
        orden["bodegas_no_validas"] = nombres
    return errores, avisos


def transportistas_activos(cursor):
    cursor.execute("SELECT id, nombre, tipo, minutos_traslado, hora_limite_recepcion FROM ent_transportistas WHERE activo = 1 ORDER BY nombre")
    lista = cursor.fetchall()
    cursor.execute("SELECT transportista_id, destino, hora_salida, dias FROM ent_transportista_horarios ORDER BY hora_salida")
    horarios = cursor.fetchall()
    for t in lista:
        t["limite"] = hora_texto(t["hora_limite_recepcion"])
        t["horarios"] = [h for h in horarios if h["transportista_id"] == t["id"]]
        t["destinos"] = sorted({h["destino"] for h in t["horarios"]})
    return lista


def hay_traslado(orden):
    return any(l.get("responsable") == "coordinador" for l in orden["lineas"])


def alerta_para(cursor, transportistas, transportista_id, hora_requerida, destino, traslado):
    t = next((x for x in transportistas if x["id"] == transportista_id), None)
    return calcular_alerta(leer_configuracion(cursor), hora_requerida, traslado, t, t["horarios"] if t else [], destino)


def leer_datos_envio(form, orden, transportistas):
    """Datos que agrega el vendedor -> (datos, errores)."""
    errores = []
    tid = form.get("transportista_id", type=int)
    if tid not in {t["id"] for t in transportistas}:
        errores.append("Escoja el transporte.")
    try:
        hora = datetime.strptime(form.get("hora_requerida", ""), "%Y-%m-%dT%H:%M")
    except ValueError:
        hora = None
        errores.append("Escriba la fecha y hora requerida de salida.")
    contado = orden["terminos"].upper().startswith("CONTADO")
    forma_pago = form.get("forma_pago") if contado and form.get("forma_pago") in ("transferencia", "contra_entrega") else None
    datos = {"transportista_id": tid, "destino": texto(form.get("destino"), 100) or None, "hora_requerida": hora,
             "prioridad": "urgente" if form.get("prioridad") == "urgente" else "normal",
             "forma_pago": forma_pago, "flete_paga": "empresa" if form.get("flete_paga") == "empresa" else "cliente",
             "notas": (form.get("notas") or "").strip()[:2000] or None}
    return datos, errores


def guardar_envio(cursor, orden, datos, usuario_id, token):
    """Inserta el envío y sus líneas, mueve el PDF y registra el evento. Devuelve el id."""
    cursor.execute("""INSERT INTO ent_envios (pedido_num, cliente, terminos, fecha_pedido, ingresado_por_orden, subtotal_orden, vendedor_id,
                          transportista_id, destino, hora_requerida, prioridad, riesgo, riesgo_confirmado, forma_pago, flete_paga, notas, estado)
                      VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'ORDEN_ENVIADA')""",
                   (orden["pedido"], orden["cliente"][:150], orden["terminos"][:30], orden["fecha"], (orden["ingresado_por"] or "")[:60] or None,
                    orden["subtotal"], usuario_id, datos["transportista_id"], datos["destino"], datos["hora_requerida"], datos["prioridad"],
                    datos["riesgo"], int(datos["riesgo_confirmado"]), datos["forma_pago"], datos["flete_paga"], datos["notas"]))
    envio_id = cursor.lastrowid
    ahora_ni = ahora_nicaragua()
    for l in lineas_que_van(orden):
        b = l["b"]
        solo = l["solo_factura"]      # bodega de vendedor: ya entregada al cliente, solo va la factura
        cursor.execute("""INSERT INTO ent_envio_lineas (envio_id, linea, bodega_id, codigo, producto, medida, marca, cantidad, subtotal,
                              existencia, existencia_por, existencia_en, preparacion, traslado_estado, responsable, bodega_prepara_id)
                          VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                       (envio_id, l["linea"], b["id"], l["codigo"][:50], l["producto"][:150] or None, l["medida"][:60] or None,
                        l["marca"][:40] or None, l["cantidad"], l["subtotal"],
                        "CONFIRMADA" if solo else "PENDIENTE", usuario_id if solo else None, ahora_ni if solo else None,
                        "RETIRADA" if solo else "PENDIENTE",
                        "PENDIENTE" if b["requiere_traslado"] and not solo else "NO_REQUIERE", l["responsable"], l["prepara_id"]))
    anio = (orden["fecha"] or datetime.now()).year
    relativa = os.path.join("entregas", "ordenes", str(anio), f"{orden['pedido']}_{envio_id}.pdf")
    carpeta("ordenes", str(anio))
    os.replace(ruta_temporal(token), os.path.join(current_app.config["UPLOAD_FOLDER"], relativa))
    cursor.execute("UPDATE ent_envios SET pdf_url = %s WHERE id = %s", (relativa.replace(os.sep, "/"), envio_id))
    alerta = {"verde": "verde", "amarillo": "amarilla", "rojo": "roja"}.get(datos["riesgo"], "sin calcular")
    # Aviso a paquetería y al responsable de cada línea (sección 4.1)
    a_quien = destinatarios(cursor, "paqueteria")
    for l in lineas_que_van(orden):
        if not l["solo_factura"]:
            a_quien += responsables_linea(cursor, {"responsable": l["responsable"], "bodega_id": l["b"]["id"]})
    avisar(cursor, a_quien, "orden_nueva", f"Orden nueva {orden['pedido']} ({orden['cliente'][:60]}): sale "
           f"{datos['hora_requerida'].strftime('%d/%m %H:%M')}" + (" · EN RIESGO" if datos["riesgo"] == "rojo" else ""),
           envio_id=envio_id, creado_por=usuario_id)
    registrar_evento(cursor, envio_id, usuario_id, "ORDEN_ENVIADA",
                     comentario=f"Orden {orden['pedido']} subida desde el PDF ({len(lineas_que_van(orden))} líneas); alerta {alerta}"
                                + (", hora confirmada por el vendedor" if datos["riesgo_confirmado"] else ""))
    solo = [l for l in lineas_que_van(orden) if l["solo_factura"]]
    if solo:
        grupos = {}
        for l in solo:
            grupos.setdefault(l["b"]["nombre"], []).append(str(l["linea"]))
        registrar_evento(cursor, envio_id, usuario_id, "SOLO_FACTURA", campo="existencia",
                         comentario="Mercadería ya entregada al cliente, solo va la factura: "
                                    + "; ".join(f"{b} (línea{'s' if len(n) > 1 else ''} {', '.join(n)})" for b, n in grupos.items()))
    if len(lineas_que_van(orden)) < len(orden["lineas"]):
        registrar_evento(cursor, envio_id, usuario_id, "NO_SE_ENVIAN", campo="existencia",
                         comentario=f"No se envían (bodega fuera del catálogo): {texto_excluidas(orden)}")
    if solo:
        cursor.execute("SELECT * FROM ent_envios WHERE id = %s", (envio_id,))
        recalcular_estado(cursor, cursor.fetchone(), usuario_id)     # si todas eran «solo factura», queda lista para facturar
    return envio_id


# -----------------------------
# Subir orden
# -----------------------------
@entregas_orden_bp.route("/entregas/subir_orden", methods=["GET", "POST"])
@requiere_rol("entregas_subir_orden", "vendedor")
def subir_orden():
    if request.method == "GET":
        return render_template("entregas_subir_orden.html", paso=1)
    accion = request.form.get("accion")

    if accion == "leer":
        archivo = request.files.get("pdf")
        contenido = archivo.read(TAMANO_MAXIMO + 1) if archivo else b""
        if not contenido:
            return render_template("entregas_subir_orden.html", paso=1, error="Seleccione el PDF de la orden.")
        if len(contenido) > TAMANO_MAXIMO:
            return render_template("entregas_subir_orden.html", paso=1, error="El archivo pesa más de 10 MB.")
        try:
            orden = leer_orden(contenido)
        except ErrorLectura as e:
            return render_template("entregas_subir_orden.html", paso=1, error=f"No se aceptó la orden: {e}.", archivo=archivo.filename)
        limpiar_temporales()
        avisar_bodegas_no_validas(orden)
        token = uuid.uuid4().hex
        with open(ruta_temporal(token), "wb") as f:
            f.write(contenido)
        return vista_previa(orden, token, archivo.filename)

    if accion == "enviar":
        token = request.form.get("token", "")
        ruta = ruta_temporal(token)
        if not ruta or not os.path.isfile(ruta):
            flash("La orden leída ya no está disponible (pasaron más de 24 horas o ya se envió). Súbala otra vez.", "warning")
            return redirect(url_for("entregas_orden.subir_orden"))
        with open(ruta, "rb") as f:
            try:
                orden = leer_orden(f.read())
            except ErrorLectura as e:
                return render_template("entregas_subir_orden.html", paso=1, error=f"No se aceptó la orden: {e}.")
        return enviar(orden, token)
    return redirect(url_for("entregas_orden.subir_orden"))


def avisar_bodegas_no_validas(orden):
    """Avisa al administrador de Entregas que una orden trae bodegas que no están (o están desactivadas) en el catálogo."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        analizar(cursor, orden)
        for nombre in orden.get("bodegas_no_validas") or []:
            avisar_sin_repetir(cursor, destinatarios(cursor, "admin"), "bodega_desconocida",
                               f"La bodega «{nombre}» aparece en órdenes y no está activa en el catálogo. Agréguela en Catálogos con ese "
                               f"nombre exacto; si es bodega de un vendedor (mercadería ya entregada), márquela «solo factura» para que su "
                               f"factura vaya en el paquete.",
                               creado_por=usuario_actual()["id"])
        conn.commit()
    finally:
        cursor.close()
        conn.close()


def vista_previa(orden, token, archivo, datos=None, errores_form=None, alerta=None, pedir_confirmacion=False):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        errores, avisos = analizar(cursor, orden)
        transportistas = transportistas_activos(cursor)
    finally:
        cursor.close()
        conn.close()
    if not transportistas:
        errores.append("No hay transportistas en el catálogo. Avise al administrador de Entregas.")
    u = usuario_actual()
    return render_template("entregas_subir_orden.html", paso=2, orden=orden, token=token, archivo=archivo, errores=errores, avisos=avisos,
                           errores_form=errores_form or [], datos=datos or {}, transportistas=transportistas, tipos=TIPOS_TRANSPORTISTA,
                           contado=orden["terminos"].upper().startswith("CONTADO"), usuario=u, traslado=hay_traslado(orden),
                           alerta=alerta, pedir_confirmacion=pedir_confirmacion, errores_bodega=orden.get("errores_bodega", []),
                           excluidas=texto_excluidas(orden), van=len(lineas_que_van(orden)))


def enviar(orden, token):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        errores, _ = analizar(cursor, orden)
        transportistas = transportistas_activos(cursor)
        datos, errores_form = leer_datos_envio(request.form, orden, transportistas)
        if orden.get("errores_bodega") and request.form.get("excluir_bodegas") != "1":
            errores_form.append("La orden trae líneas de bodegas fuera del catálogo: marque que no se envían para grabar el resto.")
        alerta = None
        if not errores_form:
            alerta = alerta_para(cursor, transportistas, datos["transportista_id"], datos["hora_requerida"], datos["destino"],
                                 hay_traslado(orden))
        confirmada = request.form.get("confirmar_riesgo") == "1"
        falta_confirmar = bool(alerta and alerta["nivel"] == "rojo" and not confirmada)
        if errores or errores_form or falta_confirmar:
            cursor.close()
            conn.close()
            conn = None
            return vista_previa(orden, token, request.form.get("archivo"), datos=request.form, errores_form=errores_form, alerta=alerta,
                                pedir_confirmacion=falta_confirmar)
        datos["riesgo"] = alerta["nivel"]
        datos["riesgo_confirmado"] = alerta["nivel"] == "rojo" and confirmada
        u = usuario_actual()
        try:
            envio_id = guardar_envio(cursor, orden, datos, u["id"], token)
            conn.commit()
        except IntegrityError as e:
            conn.rollback()
            if e.errno == 1062:
                flash(f"El pedido {orden['pedido']} ya se subió mientras usted revisaba la orden.", "warning")
                return redirect(url_for("entregas_orden.subir_orden"))
            raise
        except Exception:
            conn.rollback()
            raise
    finally:
        if conn:
            cursor.close()
            conn.close()
    flash(f"✅ Orden {orden['pedido']} enviada con {len(orden['lineas'])} líneas. Queda en «Orden enviada»."
          + (" Va marcada EN RIESGO para bodega y paquetería." if datos["riesgo_confirmado"] else ""), "success")
    return redirect(url_for("entregas_orden.envio", id_=envio_id))


@entregas_orden_bp.route("/entregas/subir_orden/alerta", methods=["POST"])
@requiere_rol("entregas_subir_orden", "vendedor")
def alerta_en_vivo():
    """Alerta mientras el vendedor escoge transporte y hora (la decisión final se vuelve a calcular al enviar)."""
    try:
        hora = datetime.strptime(request.form.get("hora_requerida", ""), "%Y-%m-%dT%H:%M")
    except ValueError:
        hora = None
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        transportistas = transportistas_activos(cursor)
        alerta = alerta_para(cursor, transportistas, request.form.get("transportista_id", type=int), hora,
                             texto(request.form.get("destino"), 100), request.form.get("traslado") == "1")
    finally:
        cursor.close()
        conn.close()
    return jsonify({"nivel": alerta["nivel"], "html": render_template("_entregas_alerta.html", alerta=alerta, en_formulario=True),
                    "sugerencia": alerta["sugerencia"]["hora"].strftime("%Y-%m-%dT%H:%M") if alerta["sugerencia"] else None})


# -----------------------------
# Detalle del envío
# -----------------------------
def lineas_visibles(roles, usuario_id, envio, lineas):
    """Qué líneas del envío puede ver el usuario (None: no puede ver el envío). Paquetería, administrador y el vendedor dueño ven
    todo; el coordinador, las líneas con traslado; el bodeguero, las de su bodega o las que prepara."""
    if "admin" in roles or "paqueteria" in roles or ("vendedor" in roles and envio["vendedor_id"] == usuario_id):
        return lineas
    mias = []
    for l in lineas:
        if "coordinador_bodega" in roles and (l["responsable"] == "coordinador" or l["traslado_estado"] != "NO_REQUIERE"):
            mias.append(l)
        elif "bodega" in roles and roles["bodega"] and roles["bodega"] in (l["bodega_id"], l["bodega_prepara_id"]):
            mias.append(l)
    return mias or None


def leer_envio(cursor, id_):
    cursor.execute("""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista, t.tipo AS transportista_tipo
                      FROM ent_envios e JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                      WHERE e.id = %s""", (id_,))
    envio = cursor.fetchone()
    if not envio:
        return None, [], []
    cursor.execute("""SELECT l.*, b.nombre AS bodega, b.nombre_en_orden, b.solo_factura, p.nombre AS bodega_prepara
                      FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id LEFT JOIN ent_bodegas p ON p.id = l.bodega_prepara_id
                      WHERE l.envio_id = %s ORDER BY l.linea""", (id_,))
    lineas = cursor.fetchall()
    cursor.execute("""SELECT ev.*, u.usuario FROM ent_envio_eventos ev JOIN usuarios u ON u.id = ev.usuario_id
                      WHERE ev.envio_id = %s ORDER BY ev.creado_en, ev.id""", (id_,))
    return envio, lineas, cursor.fetchall()


@entregas_orden_bp.route("/entregas/envio/<int:id_>")
@requiere_modulo
def envio(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        envio, lineas, eventos = leer_envio(cursor, id_)
        alerta = alerta_de_envio(cursor, envio, lineas) if envio else None
        facturas = []
        from router.entregas_rutas import entrega_de_envio
        entrega = entrega_de_envio(cursor, envio)
        if envio:
            cursor.execute("""SELECT f.*, b.nombre AS bodega, u.usuario FROM ent_envio_facturas f JOIN ent_bodegas b ON b.id = f.bodega_id
                              JOIN usuarios u ON u.id = f.registrada_por WHERE f.envio_id = %s ORDER BY b.nombre""", (id_,))
            facturas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    visibles = lineas_visibles(roles_actuales(), usuario_actual()["id"], envio, lineas) if envio else None
    if visibles is None:
        flash("El envío no existe o no le corresponde.", "warning")
        return redirect(url_for("inicio"))
    eventos = [e for e in eventos if e["linea_id"] is None or any(l["id"] == e["linea_id"] for l in visibles)]
    return render_template("entregas_envio.html", e=envio, lineas=visibles, todas=len(visibles) == len(lineas), eventos=eventos,
                           nombre_estado=NOMBRE_ESTADO, estados=ESTADOS, tipos=TIPOS_TRANSPORTISTA, roles=ROLES, alerta=alerta,
                           facturas=facturas, entrega=entrega, es_dueno=envio["vendedor_id"] == usuario_actual()["id"], mis_roles=roles_actuales())


@entregas_orden_bp.route("/entregas/envio/<int:id_>/pdf")
@requiere_modulo
def envio_pdf(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        envio, lineas, _ = leer_envio(cursor, id_)
    finally:
        cursor.close()
        conn.close()
    if not envio or lineas_visibles(roles_actuales(), usuario_actual()["id"], envio, lineas) is None:
        abort(404)
    ruta = ruta_absoluta(envio["pdf_url"])
    if not ruta:
        abort(404)
    return send_file(ruta, mimetype="application/pdf", download_name=f"orden_{envio['pedido_num']}.pdf")
