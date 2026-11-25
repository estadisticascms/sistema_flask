"""Entregas · parte 11: rutas del mensajero (sección 4.10 de la especificación).

Paquetería (/entregas/rutas):
- Agrupa envíos EMPACADOS en una ruta del día y la asigna a un mensajero. Se crea una parada por transportista, ordenadas por la
  hora requerida más temprana de sus envíos. Si un envío se agrega a una ruta que ya tiene una parada sin entregar de su mismo
  transportista, va en esa parada.
- Mientras la ruta está ARMADA puede agregar o quitar envíos, cambiar el mensajero o deshacer la ruta.
- Puede reordenar las paradas que faltan por entregar (también con la ruta EN_CURSO), por ejemplo si una hora límite está en riesgo.

Mensajero (/entregas/mis_rutas, pensada para celular):
- Ve sus paradas en orden con dirección, enlace al mapa, teléfono y hora límite.
- «Salí»: la ruta pasa a EN_CURSO y sus envíos a EN_RUTA.
- En cada parada registra la hora de entrega, la guía o encomienda, una foto de lo entregado, una foto del recibo y el monto del
  recibo; si el recibo cubre varios envíos puede detallar el monto por envío (ent_envios.flete_monto). Con un solo envío, el monto
  del recibo es su flete. Al registrar la parada, sus envíos pasan a ENTREGADA y se avisa al vendedor.
- «Regresé»: con todas las paradas entregadas, la ruta pasa a CERRADA con la hora de regreso.

Fotos (regla 6.4): se comprimen en el celular antes de subir y otra vez en el servidor (máximo 1600 px, JPEG); se guardan en
uploads/entregas/fotos/<año>/<mes>/ y en MySQL solo va la ruta.

[PENDIENTE] 2: quién cobra cuando el pago es en la entrega y cómo se registra. NO hay registro de cobros: la parada solo muestra
«COBRAR» como aviso al mensajero, igual que la etiqueta.

Los cambios de estado del envío pasan por entregas_flujo.poner_estado() (EMPACADA → EN_RUTA → ENTREGADA).
"""
import io
import os
import secrets
from collections import defaultdict
from datetime import datetime, timedelta

from flask import Blueprint, abort, current_app, flash, redirect, render_template, request, send_file, url_for

from router.entregas import contexto_entregas, registrar_evento, requiere_modulo, requiere_rol, roles_actuales, usuario_actual
from router.entregas_avisos import avisar, destinatarios
from router.entregas_catalogos import TIPOS_TRANSPORTISTA
from router.entregas_flujo import ErrorFlujo, ahora, poner_estado
from router.entregas_formato import hora12
from router.entregas_orden import NOMBRE_ESTADO, carpeta, lineas_visibles, ruta_absoluta
from utils.db import get_connection

entregas_rutas_bp = Blueprint("entregas_rutas", __name__)
entregas_rutas_bp.context_processor(contexto_entregas)

LADO_MAXIMO = 1600          # px del lado mayor de la foto guardada
CALIDAD_JPEG = 72
BYTES_MAXIMOS = 25 * 1024 * 1024
FOTOS = {"entrega": "foto_entrega_url", "recibo": "foto_recibo_url"}


# -----------------------------
# Lectura
# -----------------------------
def lista_sql(valores):
    return ", ".join(["%s"] * len(valores))


def mensajeros(cursor):
    cursor.execute("""SELECT u.id, u.usuario FROM ent_usuario_rol r JOIN usuarios u ON u.id = r.usuario_id
                      WHERE r.rol = 'mensajero' AND r.activo = 1 ORDER BY u.usuario""")
    return cursor.fetchall()


def es_mensajero(cursor, usuario_id):
    cursor.execute("SELECT 1 FROM ent_usuario_rol WHERE usuario_id = %s AND rol = 'mensajero' AND activo = 1", (usuario_id,))
    return cursor.fetchone() is not None


def envios_sin_ruta(cursor):
    """Envíos EMPACADOS que todavía no están en una ruta, agrupados por transportista."""
    cursor.execute("""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista, t.tipo AS transportista_tipo
                      FROM ent_envios e JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                      WHERE e.estado = 'EMPACADA' AND e.parada_id IS NULL ORDER BY t.nombre, e.hora_requerida""")
    grupos = {}
    for e in cursor.fetchall():
        grupos.setdefault(e["transportista_id"], {"transportista": e["transportista"], "tipo": e["transportista_tipo"], "envios": []})
        grupos[e["transportista_id"]]["envios"].append(e)
    return sorted(grupos.values(), key=lambda g: min(e["hora_requerida"] for e in g["envios"]))


def leer_rutas(cursor, where, params):
    """Rutas con sus paradas (en orden) y los envíos de cada parada."""
    cursor.execute(f"""SELECT r.*, u.usuario AS mensajero FROM ent_rutas r JOIN usuarios u ON u.id = r.mensajero_id
                       WHERE {where} ORDER BY r.fecha DESC, r.id DESC""", params)
    rutas = cursor.fetchall()
    if not rutas:
        return []
    ids = [r["id"] for r in rutas]
    cursor.execute(f"""SELECT p.*, t.nombre AS transportista, t.tipo AS transportista_tipo, t.direccion, t.mapa_url, t.telefono,
                              t.hora_limite_recepcion
                       FROM ent_ruta_paradas p JOIN ent_transportistas t ON t.id = p.transportista_id
                       WHERE p.ruta_id IN ({lista_sql(ids)}) ORDER BY p.ruta_id, p.orden""", ids)
    paradas = cursor.fetchall()
    envios = defaultdict(list)
    if paradas:
        pids = [p["id"] for p in paradas]
        cursor.execute(f"""SELECT e.*, u.usuario AS vendedor FROM ent_envios e JOIN usuarios u ON u.id = e.vendedor_id
                           WHERE e.parada_id IN ({lista_sql(pids)}) ORDER BY e.hora_requerida, e.id""", pids)
        for e in cursor.fetchall():
            envios[e["parada_id"]].append(e)
    momento = ahora()
    por_ruta = defaultdict(list)
    for p in paradas:
        p["envios"] = envios[p["id"]]
        vigentes = [e for e in p["envios"] if e["estado"] != "ANULADA"]
        p["hora_requerida"] = min((e["hora_requerida"] for e in vigentes), default=None)
        p["atrasada"] = not p["entregado_en"] and p["hora_requerida"] is not None and p["hora_requerida"] < momento
        p["mapa"] = mapa_de(p)
        p["cobrar"] = any(e["forma_pago"] == "contra_entrega" for e in vigentes)
        p["bultos"] = sum(e["bultos"] or 0 for e in vigentes)
        por_ruta[p["ruta_id"]].append(p)
    for r in rutas:
        r["paradas"] = por_ruta[r["id"]]
        r["entregadas"] = sum(1 for p in r["paradas"] if p["entregado_en"])
        r["envios"] = sum(len(p["envios"]) for p in r["paradas"])
        r["completa"] = bool(r["paradas"]) and r["entregadas"] == len(r["paradas"])
    return rutas


def mapa_de(parada):
    if parada["mapa_url"]:
        return parada["mapa_url"]
    if parada["direccion"]:
        from urllib.parse import quote
        return "https://www.google.com/maps/search/?api=1&query=" + quote(parada["direccion"])
    return None


def bloquear_ruta(cursor, ruta_id):
    cursor.execute("SELECT r.*, u.usuario AS mensajero FROM ent_rutas r JOIN usuarios u ON u.id = r.mensajero_id WHERE r.id = %s FOR UPDATE",
                   (ruta_id,))
    ruta = cursor.fetchone()
    if not ruta:
        raise ErrorFlujo("La ruta no existe.")
    return ruta


def bloquear_envios(cursor, envio_ids):
    ids = sorted({int(i) for i in envio_ids})
    if not ids:
        return []
    cursor.execute(f"SELECT * FROM ent_envios WHERE id IN ({lista_sql(ids)}) ORDER BY id FOR UPDATE", ids)
    return cursor.fetchall()


def paradas_de(cursor, ruta_id):
    cursor.execute("SELECT * FROM ent_ruta_paradas WHERE ruta_id = %s ORDER BY orden, id", (ruta_id,))
    return cursor.fetchall()


def renumerar(cursor, ruta_id):
    for i, p in enumerate(paradas_de(cursor, ruta_id), start=1):
        if p["orden"] != i:
            cursor.execute("UPDATE ent_ruta_paradas SET orden = %s WHERE id = %s", (i, p["id"]))


def quitar_paradas_vacias(cursor, ruta_id):
    cursor.execute("""DELETE p FROM ent_ruta_paradas p LEFT JOIN ent_envios e ON e.parada_id = p.id
                      WHERE p.ruta_id = %s AND p.entregado_en IS NULL AND e.id IS NULL""", (ruta_id,))
    renumerar(cursor, ruta_id)


# -----------------------------
# Paquetería: armar y ajustar rutas
# -----------------------------
def exigir_paqueteria(roles):
    if "paqueteria" not in roles:
        raise ErrorFlujo("Las rutas las arma paquetería.")


def validar_envios_para_ruta(envios, pedidos):
    if len(envios) != len(set(pedidos)):
        raise ErrorFlujo("Algún envío escogido ya no existe.")
    for e in envios:
        if e["estado"] != "EMPACADA":
            raise ErrorFlujo(f"Pedido {e['pedido_num']}: solo van a ruta los envíos empacados (está {NOMBRE_ESTADO.get(e['estado'], e['estado']).lower()}).")
        if e["parada_id"]:
            raise ErrorFlujo(f"Pedido {e['pedido_num']}: ya está en una ruta.")


def poner_en_paradas(cursor, ruta, envios, usuario_id):
    """Pone cada envío en la parada sin entregar de su transportista, o crea paradas nuevas al final (ordenadas por hora requerida)."""
    existentes = {p["transportista_id"]: p for p in paradas_de(cursor, ruta["id"]) if not p["entregado_en"]}
    siguiente = len(paradas_de(cursor, ruta["id"])) + 1
    nuevos = defaultdict(list)
    for e in envios:
        nuevos[e["transportista_id"]].append(e)
    for transportista_id in sorted(nuevos, key=lambda t: min(e["hora_requerida"] for e in nuevos[t])):
        parada = existentes.get(transportista_id)
        if parada:
            parada_id, orden = parada["id"], parada["orden"]
        else:
            cursor.execute("INSERT INTO ent_ruta_paradas (ruta_id, transportista_id, orden) VALUES (%s, %s, %s)",
                           (ruta["id"], transportista_id, siguiente))
            parada_id, orden = cursor.lastrowid, siguiente
            siguiente += 1
        for e in nuevos[transportista_id]:
            cursor.execute("UPDATE ent_envios SET parada_id = %s WHERE id = %s", (parada_id, e["id"]))
            registrar_evento(cursor, e["id"], usuario_id, e["estado"], de_estado=e["estado"],
                             comentario=f"En la ruta #{ruta['id']} del {ruta['fecha'].strftime('%d/%m')} ({ruta['mensajero']}), parada {orden}")


def armar_ruta(cursor, usuario_id, roles, mensajero_id, fecha, envio_ids):
    """Crea la ruta con sus paradas. Devuelve el id de la ruta."""
    exigir_paqueteria(roles)
    if not mensajero_id or not es_mensajero(cursor, mensajero_id):
        raise ErrorFlujo("Escoja el mensajero de la ruta.")
    if not fecha or fecha < ahora().date():
        raise ErrorFlujo("La fecha de la ruta no puede ser anterior a hoy.")
    if not envio_ids:
        raise ErrorFlujo("Marque los envíos que van en la ruta.")
    envios = bloquear_envios(cursor, envio_ids)
    validar_envios_para_ruta(envios, envio_ids)
    cursor.execute("INSERT INTO ent_rutas (mensajero_id, fecha) VALUES (%s, %s)", (mensajero_id, fecha))
    ruta = bloquear_ruta(cursor, cursor.lastrowid)
    poner_en_paradas(cursor, ruta, envios, usuario_id)
    avisar(cursor, [mensajero_id], "ruta_asignada",
           f"Ruta #{ruta['id']} del {fecha.strftime('%d/%m')}: {len(envios)} envío(s) en {len(paradas_de(cursor, ruta['id']))} parada(s)",
           creado_por=usuario_id)
    return ruta["id"]


def agregar_a_ruta(cursor, usuario_id, roles, ruta_id, envio_ids):
    exigir_paqueteria(roles)
    ruta = bloquear_ruta(cursor, ruta_id)
    if ruta["estado"] != "ARMADA":
        raise ErrorFlujo(f"La ruta #{ruta_id} ya salió: arme otra ruta para esos envíos.")
    if not envio_ids:
        raise ErrorFlujo("Marque los envíos que quiere agregar.")
    envios = bloquear_envios(cursor, envio_ids)
    validar_envios_para_ruta(envios, envio_ids)
    poner_en_paradas(cursor, ruta, envios, usuario_id)
    avisar(cursor, [ruta["mensajero_id"]], "ruta_asignada", f"Ruta #{ruta_id}: se agregaron {len(envios)} envío(s)", creado_por=usuario_id)


def quitar_de_ruta(cursor, usuario_id, roles, envio_id):
    exigir_paqueteria(roles)
    cursor.execute("SELECT p.ruta_id FROM ent_envios e JOIN ent_ruta_paradas p ON p.id = e.parada_id WHERE e.id = %s", (envio_id,))
    fila = cursor.fetchone()
    if not fila:
        raise ErrorFlujo("Ese envío no está en una ruta.")
    ruta = bloquear_ruta(cursor, fila["ruta_id"])
    if ruta["estado"] != "ARMADA":
        raise ErrorFlujo(f"La ruta #{ruta['id']} ya salió: no se pueden quitar envíos.")
    envio = bloquear_envios(cursor, [envio_id])[0]
    cursor.execute("UPDATE ent_envios SET parada_id = NULL WHERE id = %s", (envio_id,))
    registrar_evento(cursor, envio_id, usuario_id, envio["estado"], de_estado=envio["estado"], comentario=f"Quitado de la ruta #{ruta['id']}")
    quitar_paradas_vacias(cursor, ruta["id"])
    if not paradas_de(cursor, ruta["id"]):
        cursor.execute("DELETE FROM ent_rutas WHERE id = %s", (ruta["id"],))
        return True
    return False


def deshacer_ruta(cursor, usuario_id, roles, ruta_id):
    exigir_paqueteria(roles)
    ruta = bloquear_ruta(cursor, ruta_id)
    if ruta["estado"] != "ARMADA":
        raise ErrorFlujo(f"La ruta #{ruta_id} ya salió: no se puede deshacer.")
    cursor.execute("""SELECT e.* FROM ent_envios e JOIN ent_ruta_paradas p ON p.id = e.parada_id WHERE p.ruta_id = %s
                      ORDER BY e.id FOR UPDATE""", (ruta_id,))
    for e in cursor.fetchall():
        cursor.execute("UPDATE ent_envios SET parada_id = NULL WHERE id = %s", (e["id"],))
        registrar_evento(cursor, e["id"], usuario_id, e["estado"], de_estado=e["estado"], comentario=f"Se deshizo la ruta #{ruta_id}")
    cursor.execute("DELETE FROM ent_ruta_paradas WHERE ruta_id = %s", (ruta_id,))
    cursor.execute("DELETE FROM ent_rutas WHERE id = %s", (ruta_id,))
    avisar(cursor, [ruta["mensajero_id"]], "ruta_asignada", f"La ruta #{ruta_id} del {ruta['fecha'].strftime('%d/%m')} se deshizo",
           creado_por=usuario_id)


def cambiar_mensajero(cursor, usuario_id, roles, ruta_id, mensajero_id):
    exigir_paqueteria(roles)
    ruta = bloquear_ruta(cursor, ruta_id)
    if ruta["estado"] != "ARMADA":
        raise ErrorFlujo(f"La ruta #{ruta_id} ya salió: no se puede cambiar el mensajero.")
    if not mensajero_id or not es_mensajero(cursor, mensajero_id):
        raise ErrorFlujo("Escoja un mensajero.")
    if mensajero_id == ruta["mensajero_id"]:
        return
    cursor.execute("UPDATE ent_rutas SET mensajero_id = %s WHERE id = %s", (mensajero_id, ruta_id))
    avisar(cursor, [mensajero_id], "ruta_asignada", f"Se le asignó la ruta #{ruta_id} del {ruta['fecha'].strftime('%d/%m')}", creado_por=usuario_id)
    avisar(cursor, [ruta["mensajero_id"]], "ruta_asignada", f"La ruta #{ruta_id} se asignó a otro mensajero", creado_por=usuario_id)


def mover_parada(cursor, roles, parada_id, direccion):
    """Sube (-1) o baja (+1) una parada sin entregar, intercambiándola con la parada sin entregar vecina."""
    exigir_paqueteria(roles)
    cursor.execute("SELECT ruta_id FROM ent_ruta_paradas WHERE id = %s", (parada_id,))
    fila = cursor.fetchone()
    if not fila:
        raise ErrorFlujo("La parada no existe.")
    ruta = bloquear_ruta(cursor, fila["ruta_id"])
    if ruta["estado"] == "CERRADA":
        raise ErrorFlujo("La ruta ya está cerrada.")
    pendientes = [p for p in paradas_de(cursor, ruta["id"]) if not p["entregado_en"]]
    pos = next((i for i, p in enumerate(pendientes) if p["id"] == parada_id), None)
    if pos is None:
        raise ErrorFlujo("Esa parada ya se entregó.")
    otra = pos + direccion
    if not 0 <= otra < len(pendientes):
        return
    a, b = pendientes[pos], pendientes[otra]
    cursor.execute("UPDATE ent_ruta_paradas SET orden = %s WHERE id = %s", (b["orden"], a["id"]))
    cursor.execute("UPDATE ent_ruta_paradas SET orden = %s WHERE id = %s", (a["orden"], b["id"]))


# -----------------------------
# Mensajero: salir, entregar, regresar
# -----------------------------
def exigir_su_ruta(roles, usuario_id, ruta):
    if "mensajero" not in roles or ruta["mensajero_id"] != usuario_id:
        raise ErrorFlujo("Esa ruta es de otro mensajero.")


def salir(cursor, usuario_id, roles, ruta_id):
    """«Salí»: la ruta pasa a EN_CURSO y sus envíos EMPACADOS a EN_RUTA. Los anulados se quedan fuera. Devuelve cuántos salieron."""
    ruta = bloquear_ruta(cursor, ruta_id)
    exigir_su_ruta(roles, usuario_id, ruta)
    if ruta["estado"] != "ARMADA":
        raise ErrorFlujo(f"La ruta #{ruta_id} ya había salido.")
    cursor.execute("""SELECT e.* FROM ent_envios e JOIN ent_ruta_paradas p ON p.id = e.parada_id WHERE p.ruta_id = %s
                      ORDER BY e.id FOR UPDATE""", (ruta_id,))
    envios = cursor.fetchall()
    salen = [e for e in envios if e["estado"] == "EMPACADA"]
    if not salen:
        raise ErrorFlujo("La ruta no tiene envíos para entregar.")
    for e in envios:
        if e["estado"] != "EMPACADA":       # anulado mientras la ruta estaba armada: no sale
            cursor.execute("UPDATE ent_envios SET parada_id = NULL WHERE id = %s", (e["id"],))
    quitar_paradas_vacias(cursor, ruta_id)
    momento = ahora()
    cursor.execute("UPDATE ent_rutas SET estado = 'EN_CURSO', salida = %s WHERE id = %s", (momento, ruta_id))
    for e in salen:
        poner_estado(cursor, e, "EN_RUTA", usuario_id, f"Salió en la ruta #{ruta_id} ({ruta['mensajero']})")
    return len(salen)


def registrar_parada(cursor, usuario_id, roles, parada_id, entregado_en, guia, monto, fotos, montos_envio):
    """Registra la entrega de una parada: hora, guía, fotos (rutas ya guardadas) y monto del recibo; sus envíos pasan a ENTREGADA.
    montos_envio: {envio_id: monto} (vacío si no se detalló). Devuelve la lista de pedidos entregados."""
    cursor.execute("SELECT ruta_id FROM ent_ruta_paradas WHERE id = %s", (parada_id,))
    fila = cursor.fetchone()
    if not fila:
        raise ErrorFlujo("La parada no existe.")
    ruta = bloquear_ruta(cursor, fila["ruta_id"])
    exigir_su_ruta(roles, usuario_id, ruta)
    if ruta["estado"] != "EN_CURSO":
        raise ErrorFlujo("Marque «Salí» antes de registrar las entregas." if ruta["estado"] == "ARMADA" else "La ruta ya está cerrada.")
    cursor.execute("SELECT * FROM ent_ruta_paradas WHERE id = %s FOR UPDATE", (parada_id,))
    parada = cursor.fetchone()
    if parada["entregado_en"]:
        raise ErrorFlujo("Esa parada ya se registró.")
    cursor.execute("SELECT * FROM ent_envios WHERE parada_id = %s ORDER BY id FOR UPDATE", (parada_id,))
    envios = [e for e in cursor.fetchall() if e["estado"] == "EN_RUTA"]
    if not envios:
        raise ErrorFlujo("Esa parada no tiene envíos en ruta.")
    if entregado_en is None:
        raise ErrorFlujo("Escriba la hora de entrega.")
    if entregado_en > ahora() + timedelta(minutes=5):
        raise ErrorFlujo("La hora de entrega no puede ser futura.")
    if ruta["salida"] and entregado_en < ruta["salida"] - timedelta(minutes=1):
        raise ErrorFlujo(f"La hora de entrega no puede ser antes de la salida ({hora12(ruta['salida'])}).")
    if monto is None or monto < 0 or monto >= 10000000:
        raise ErrorFlujo("Escriba el monto del recibo (0 si no hubo cobro de flete).")
    if not fotos.get("entrega") or not fotos.get("recibo"):
        raise ErrorFlujo("Tome la foto de lo entregado y la foto del recibo.")
    monto = round(monto, 2)
    ids = {e["id"] for e in envios}
    detalle = {k: v for k, v in (montos_envio or {}).items() if k in ids}
    if len(envios) == 1:
        detalle = {envios[0]["id"]: monto}
    elif detalle:
        if len(detalle) != len(envios) or any(v is None or v < 0 for v in detalle.values()):
            raise ErrorFlujo("Si detalla el monto por envío, escríbalo para todos los envíos de la parada.")
        if abs(sum(detalle.values()) - monto) > 0.01:
            raise ErrorFlujo(f"La suma por envío (C$ {sum(detalle.values()):,.2f}) no da el monto del recibo (C$ {monto:,.2f}).")
    guia = " ".join((guia or "").split())[:50] or None
    cursor.execute("""UPDATE ent_ruta_paradas SET entregado_en = %s, guia = %s, monto_recibo = %s, foto_entrega_url = %s, foto_recibo_url = %s
                      WHERE id = %s""", (entregado_en, guia, monto, fotos["entrega"], fotos["recibo"], parada_id))
    cursor.execute("SELECT nombre FROM ent_transportistas WHERE id = %s", (parada["transportista_id"],))
    transportista = cursor.fetchone()["nombre"]
    for e in envios:
        flete = detalle.get(e["id"])
        if flete is not None:
            cursor.execute("UPDATE ent_envios SET flete_monto = %s WHERE id = %s", (round(flete, 2), e["id"]))
        texto = f"Entregada en {transportista} a las {hora12(entregado_en)}" + (f", guía {guia}" if guia else "")
        poner_estado(cursor, e, "ENTREGADA", usuario_id, texto)
        avisar(cursor, [e["vendedor_id"]], "entregada", f"Pedido {e['pedido_num']} entregado en {transportista}" + (f" (guía {guia})" if guia else ""),
               envio_id=e["id"], creado_por=usuario_id)
    return [e["pedido_num"] for e in envios]


def cerrar_ruta(cursor, usuario_id, roles, ruta_id):
    """«Regresé»: con todas las paradas registradas, la ruta pasa a CERRADA con la hora de regreso."""
    ruta = bloquear_ruta(cursor, ruta_id)
    if not ("paqueteria" in roles or ("mensajero" in roles and ruta["mensajero_id"] == usuario_id)):
        raise ErrorFlujo("Esa ruta es de otro mensajero.")
    if ruta["estado"] != "EN_CURSO":
        raise ErrorFlujo("La ruta no está en curso.")
    if any(not p["entregado_en"] for p in paradas_de(cursor, ruta_id)):
        raise ErrorFlujo("Faltan paradas por registrar.")
    cursor.execute("UPDATE ent_rutas SET estado = 'CERRADA', regreso = %s WHERE id = %s", (ahora(), ruta_id))
    avisar(cursor, destinatarios(cursor, "paqueteria"), "ruta_cerrada", f"Ruta #{ruta_id} ({ruta['mensajero']}) cerrada: todas las paradas entregadas",
           creado_por=usuario_id)


# -----------------------------
# Fotos
# -----------------------------
def comprimir_foto(archivo):
    """Archivo subido -> bytes JPEG (máximo LADO_MAXIMO px). Sin Pillow en el servidor, guarda la foto tal cual si es JPEG o PNG."""
    datos = archivo.read(BYTES_MAXIMOS + 1)
    if not datos:
        return None, None
    if len(datos) > BYTES_MAXIMOS:
        raise ErrorFlujo("La foto es demasiado grande (máximo 25 MB).")
    try:
        from PIL import Image, ImageOps
    except ImportError:
        if datos[:3] == b"\xff\xd8\xff":
            return datos, "jpg"
        if datos[:8] == b"\x89PNG\r\n\x1a\n":
            return datos, "png"
        raise ErrorFlujo("La foto debe ser JPG o PNG.")
    try:
        img = Image.open(io.BytesIO(datos))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((LADO_MAXIMO, LADO_MAXIMO))
        salida = io.BytesIO()
        img.save(salida, "JPEG", quality=CALIDAD_JPEG, optimize=True)
        return salida.getvalue(), "jpg"
    except ErrorFlujo:
        raise
    except Exception:
        raise ErrorFlujo("No se pudo leer la foto. Tómela otra vez.")


def guardar_fotos(parada_id, comprimidas):
    """{tipo: (bytes, extensión)} -> {tipo: ruta relativa a la carpeta de subidas}. Devuelve también las rutas absolutas (para borrar)."""
    momento = ahora()
    destino = carpeta("fotos", momento.strftime("%Y"), momento.strftime("%m"))
    relativas, absolutas = {}, []
    for tipo, (datos, ext) in comprimidas.items():
        nombre = f"parada{parada_id}_{tipo}_{secrets.token_hex(4)}.{ext}"
        ruta = os.path.join(destino, nombre)
        with open(ruta, "wb") as f:
            f.write(datos)
        absolutas.append(ruta)
        relativas[tipo] = os.path.relpath(ruta, current_app.config["UPLOAD_FOLDER"]).replace(os.sep, "/")
    return relativas, absolutas


def puede_ver_parada(cursor, roles, usuario_id, parada_id):
    cursor.execute("SELECT p.*, r.mensajero_id FROM ent_ruta_paradas p JOIN ent_rutas r ON r.id = p.ruta_id WHERE p.id = %s", (parada_id,))
    parada = cursor.fetchone()
    if not parada:
        return None
    if "paqueteria" in roles or "admin" in roles or parada["mensajero_id"] == usuario_id:
        return parada
    cursor.execute("SELECT * FROM ent_envios WHERE parada_id = %s", (parada_id,))
    for e in cursor.fetchall():
        cursor.execute("SELECT * FROM ent_envio_lineas WHERE envio_id = %s", (e["id"],))
        if lineas_visibles(roles, usuario_id, e, cursor.fetchall()) is not None:
            return parada
    return None


# -----------------------------
# Lectura de formularios
# -----------------------------
def leer_fecha(texto):
    try:
        return datetime.strptime(texto or "", "%Y-%m-%d").date()
    except ValueError:
        return None


def leer_monto(texto):
    texto = (texto or "").strip().replace(",", "")
    if not texto:
        return None
    try:
        return float(texto)
    except ValueError:
        return None


def leer_ids(nombre):
    ids = []
    for v in request.form.getlist(nombre):
        try:
            ids.append(int(v))
        except ValueError:
            pass
    return ids


def ejecutar(funcion, *args):
    """Corre una función de la lógica en una transacción. Devuelve (resultado, True) o (None, False) si no se permitió."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        resultado = funcion(cursor, *args)
        conn.commit()
        return resultado, True
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return None, False
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()


# -----------------------------
# Pantallas de paquetería
# -----------------------------
@entregas_rutas_bp.route("/entregas/rutas")
@requiere_rol("entregas_rutas", "paqueteria")
def rutas():
    hoy = ahora().date()
    fecha = leer_fecha(request.args.get("fecha")) or hoy
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        sin_ruta = envios_sin_ruta(cursor)
        lista = leer_rutas(cursor, "r.fecha = %s OR r.estado <> 'CERRADA'", (fecha,))
        lista_mensajeros = mensajeros(cursor)
    finally:
        cursor.close()
        conn.close()
    armadas = [r for r in lista if r["estado"] == "ARMADA"]
    return render_template("entregas_rutas.html", sin_ruta=sin_ruta, rutas=lista, armadas=armadas, mensajeros=lista_mensajeros,
                           fecha=fecha, hoy=hoy, nombre_estado=NOMBRE_ESTADO, tipos=TIPOS_TRANSPORTISTA, activo="entregas_rutas")


def volver_a_rutas():
    return redirect(url_for("entregas_rutas.rutas", fecha=request.form.get("fecha_vista") or None))


@entregas_rutas_bp.route("/entregas/rutas/armar", methods=["POST"])
@requiere_rol("entregas_rutas", "paqueteria")
def armar():
    u = usuario_actual()
    ids = leer_ids("envio_id")
    if request.form.get("accion") == "agregar":
        _, ok = ejecutar(agregar_a_ruta, u["id"], roles_actuales(), request.form.get("ruta_id", type=int), ids)
        if ok:
            flash(f"✅ {len(ids)} envío(s) agregado(s) a la ruta #{request.form.get('ruta_id')}.", "success")
    else:
        ruta_id, ok = ejecutar(armar_ruta, u["id"], roles_actuales(), request.form.get("mensajero_id", type=int),
                               leer_fecha(request.form.get("fecha")), ids)
        if ok:
            flash(f"✅ Ruta #{ruta_id} armada con {len(ids)} envío(s). El mensajero ya la ve en «Mis rutas».", "success")
    return volver_a_rutas()


@entregas_rutas_bp.route("/entregas/rutas/quitar", methods=["POST"])
@requiere_rol("entregas_rutas", "paqueteria")
def quitar():
    borrada, ok = ejecutar(quitar_de_ruta, usuario_actual()["id"], roles_actuales(), request.form.get("envio_id", type=int))
    if ok:
        flash("✅ Envío quitado de la ruta." + (" La ruta quedó vacía y se eliminó." if borrada else ""), "success")
    return volver_a_rutas()


@entregas_rutas_bp.route("/entregas/rutas/deshacer", methods=["POST"])
@requiere_rol("entregas_rutas", "paqueteria")
def deshacer():
    _, ok = ejecutar(deshacer_ruta, usuario_actual()["id"], roles_actuales(), request.form.get("ruta_id", type=int))
    if ok:
        flash("✅ Ruta deshecha: sus envíos volvieron a «sin ruta».", "success")
    return volver_a_rutas()


@entregas_rutas_bp.route("/entregas/rutas/mensajero", methods=["POST"])
@requiere_rol("entregas_rutas", "paqueteria")
def mensajero():
    _, ok = ejecutar(cambiar_mensajero, usuario_actual()["id"], roles_actuales(), request.form.get("ruta_id", type=int),
                     request.form.get("mensajero_id", type=int))
    if ok:
        flash("✅ Mensajero cambiado.", "success")
    return volver_a_rutas()


@entregas_rutas_bp.route("/entregas/rutas/mover", methods=["POST"])
@requiere_rol("entregas_rutas", "paqueteria")
def mover():
    direccion = -1 if request.form.get("direccion") == "subir" else 1
    ejecutar(mover_parada, roles_actuales(), request.form.get("parada_id", type=int), direccion)
    return volver_a_rutas()


@entregas_rutas_bp.route("/entregas/rutas/cerrar", methods=["POST"])
@requiere_modulo
def cerrar():
    _, ok = ejecutar(cerrar_ruta, usuario_actual()["id"], roles_actuales(), request.form.get("ruta_id", type=int))
    if ok:
        flash("✅ Ruta cerrada.", "success")
    if request.form.get("desde") == "mensajero":
        return redirect(url_for("entregas_rutas.mis_rutas"))
    return volver_a_rutas()


# -----------------------------
# Pantallas del mensajero
# -----------------------------
@entregas_rutas_bp.route("/entregas/mis_rutas")
@requiere_rol("entregas_mis_rutas", "mensajero")
def mis_rutas():
    u = usuario_actual()
    hoy = ahora().date()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        lista = leer_rutas(cursor, "r.mensajero_id = %s AND (r.estado <> 'CERRADA' OR r.fecha = %s)", (u["id"], hoy))
    finally:
        cursor.close()
        conn.close()
    lista.sort(key=lambda r: ({"EN_CURSO": 0, "ARMADA": 1, "CERRADA": 2}[r["estado"]], r["fecha"], r["id"]))
    return render_template("entregas_mis_rutas.html", rutas=lista, hoy=hoy, ahora=ahora(), nombre_estado=NOMBRE_ESTADO,
                           tipos=TIPOS_TRANSPORTISTA, activo="entregas_mis_rutas")


@entregas_rutas_bp.route("/entregas/mis_rutas/salir", methods=["POST"])
@requiere_rol("entregas_mis_rutas", "mensajero")
def salir_ruta():
    n, ok = ejecutar(salir, usuario_actual()["id"], roles_actuales(), request.form.get("ruta_id", type=int))
    if ok:
        flash(f"✅ Ruta en curso: {n} envío(s) en ruta. Registre cada parada al entregar.", "success")
    return redirect(url_for("entregas_rutas.mis_rutas"))


@entregas_rutas_bp.route("/entregas/mis_rutas/entregar", methods=["POST"])
@requiere_rol("entregas_mis_rutas", "mensajero")
def entregar():
    parada_id = request.form.get("parada_id", type=int)
    volver = redirect(url_for("entregas_rutas.mis_rutas") + f"#parada-{parada_id}")
    try:
        entregado_en = datetime.strptime(request.form.get("entregado_en") or "", "%Y-%m-%dT%H:%M")
    except ValueError:
        entregado_en = None
    montos = {}
    for clave, valor in request.form.items():
        if clave.startswith("flete_") and valor.strip():
            try:
                montos[int(clave[6:])] = leer_monto(valor)
            except ValueError:
                pass
    try:
        comprimidas = {}
        for tipo in FOTOS:
            archivo = request.files.get(f"foto_{tipo}")
            datos, ext = comprimir_foto(archivo) if archivo else (None, None)
            if datos:
                comprimidas[tipo] = (datos, ext)
    except ErrorFlujo as e:
        flash(str(e), "warning")
        return volver
    if len(comprimidas) < len(FOTOS):
        flash("Tome la foto de lo entregado y la foto del recibo.", "warning")
        return volver
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    absolutas = []
    try:
        relativas, absolutas = guardar_fotos(parada_id, comprimidas)
        pedidos = registrar_parada(cursor, u["id"], roles_actuales(), parada_id, entregado_en, request.form.get("guia"),
                                   leer_monto(request.form.get("monto_recibo")), relativas, montos)
        conn.commit()
        absolutas = []
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return volver
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
        for ruta in absolutas:          # la parada no se guardó: no dejar fotos huérfanas
            try:
                os.remove(ruta)
            except OSError:
                pass
    flash(f"✅ Parada registrada: pedido(s) {', '.join(pedidos)} entregado(s).", "success")
    return redirect(url_for("entregas_rutas.mis_rutas"))


@entregas_rutas_bp.route("/entregas/parada/<int:parada_id>/foto/<tipo>")
@requiere_modulo
def foto(parada_id, tipo):
    if tipo not in FOTOS:
        abort(404)
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        parada = puede_ver_parada(cursor, roles_actuales(), usuario_actual()["id"], parada_id)
    finally:
        cursor.close()
        conn.close()
    ruta = ruta_absoluta(parada[FOTOS[tipo]]) if parada else None
    if not ruta:
        abort(404)
    return send_file(ruta, mimetype="image/png" if ruta.endswith(".png") else "image/jpeg", max_age=3600)


def entrega_de_envio(cursor, envio):
    """Datos de la ruta y la parada de un envío (para su detalle), o None."""
    if not envio or not envio.get("parada_id"):
        return None
    cursor.execute("""SELECT p.*, r.id AS ruta_id, r.fecha, r.estado AS ruta_estado, r.salida, u.usuario AS mensajero,
                             t.nombre AS transportista
                      FROM ent_ruta_paradas p JOIN ent_rutas r ON r.id = p.ruta_id JOIN usuarios u ON u.id = r.mensajero_id
                      JOIN ent_transportistas t ON t.id = p.transportista_id WHERE p.id = %s""", (envio["parada_id"],))
    return cursor.fetchone()


def contar_rutas(cursor, roles, usuario_id):
    """Pendientes de rutas: paquetería, envíos empacados sin ruta; mensajero, paradas sin registrar de sus rutas abiertas."""
    r = {}
    if "paqueteria" in roles:
        cursor.execute("SELECT COUNT(*) AS n FROM ent_envios WHERE estado = 'EMPACADA' AND parada_id IS NULL")
        r["entregas_rutas"] = cursor.fetchone()["n"]
    if "mensajero" in roles and usuario_id:
        cursor.execute("""SELECT COUNT(*) AS n FROM ent_ruta_paradas p JOIN ent_rutas r ON r.id = p.ruta_id
                          WHERE r.mensajero_id = %s AND r.estado <> 'CERRADA' AND p.entregado_en IS NULL""", (usuario_id,))
        r["entregas_mis_rutas"] = cursor.fetchone()["n"]
    return r


