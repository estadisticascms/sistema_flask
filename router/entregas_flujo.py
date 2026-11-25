"""Entregas: flujo de estados (reglas técnicas 6.1 a 6.3 de la especificación).

TODO cambio de estado del envío y de cada línea pasa por este archivo:
- poner_estado(): única función que escribe ent_envios.estado; valida el estado de origen y registra el evento.
- recalcular_estado(): calcula los estados automáticos del envío desde sus líneas y sus facturas y llama a poner_estado() si
  cambian. Se llama después de cada cambio de existencia, decisión, factura, traslado o preparación.
- registrar_existencia(): existencia de líneas por su responsable (bodeguero de esa bodega o coordinador).
- decidir_faltante(): decisión del vendedor ante un faltante (quitar, va lo disponible, cambiar de bodega o esperar).
- registrar_facturas(): una factura por bodega de facturación, solo con el envío POR_FACTURAR.
- anular(): vendedor dueño o administrador, con motivo.

Estados automáticos implementados (los siguientes se agregan con sus partes):
- EN_REVISION: se registró la primera existencia (confirmada o faltante) de alguna línea.
- POR_FACTURAR: todas las líneas CONFIRMADA o QUITADA (al menos una confirmada), ninguna FALTANTE y ningún traslado PENDIENTE.
  Si una línea vuelve a FALTANTE antes de facturar, el envío regresa a EN_REVISION.
- FACTURADA: está POR_FACTURAR y tiene una factura por cada bodega de facturación con líneas no quitadas.
- LISTA_PARA_RETIRO: está FACTURADA, todas las líneas no quitadas preparadas (LISTA) y lo trasladado ya llegó (RECIBIDO).
- RETIRADA: paquetería retiró (por bodega) todas las líneas no quitadas.

Estado manual de paquetería (parte 10): EMPACADA, al registrar bultos y peso (empacar()); después se pueden corregir sin cambiar
el estado, para reimprimir la etiqueta.

Preparación (parte 9): la registra el bodeguero de la bodega que prepara (bodega_prepara_id): en preparación y lista. Requiere la
existencia confirmada y, en líneas con traslado, que la mercadería haya llegado (RECIBIDO). Se puede preparar antes de facturar:
facturar no depende de la preparación. El retiro (RETIRADA) lo marca paquetería por bodega, solo con la orden LISTA_PARA_RETIRO.

Bodega de facturación de una línea: su bodega; en líneas con traslado, según la configuración «bodega_factura_traslado»
([PENDIENTE] 6): «destino» (la que escogió el coordinador; es lo que dice la especificación y es el valor si no se configura) u
«origen». Se factura toda la orden junta ([PENDIENTE] 7: por ahora así lo indica la especificación).

Ventas perdidas: al quitar una línea, o al aceptar solo la parte disponible, el resto se registra en la tabla ventas_perdidas
que ya usa el sistema (motivo «Sin existencia»), como pide la sección 5.
"""
from datetime import datetime

from router.entregas import registrar_evento
from router.entregas_avisos import avisar, destinatarios, responsables_linea

ABIERTOS = ("ORDEN_ENVIADA", "EN_REVISION", "POR_FACTURAR", "FACTURADA", "LISTA_PARA_RETIRO", "RETIRADA", "EMPACADA", "EN_RUTA")
ANTES_DE_FACTURAR = ("ORDEN_ENVIADA", "EN_REVISION", "POR_FACTURAR")     # mientras tanto se puede confirmar o corregir existencias
AUTOMATICOS = {"EN_REVISION", "POR_FACTURAR", "FACTURADA", "LISTA_PARA_RETIRO", "RETIRADA"}
# Estado de origen permitido para cada estado de destino
ORIGENES = {
    "EN_REVISION": {"ORDEN_ENVIADA", "POR_FACTURAR"},
    "POR_FACTURAR": {"ORDEN_ENVIADA", "EN_REVISION"},
    "FACTURADA": {"POR_FACTURAR"},
    "LISTA_PARA_RETIRO": {"FACTURADA"},
    "RETIRADA": {"LISTA_PARA_RETIRO"},
    "EMPACADA": {"RETIRADA"},
    "ANULADA": set(ABIERTOS),          # «desde cualquier estado», con motivo; solo el vendedor dueño o el administrador
}
EXISTENCIA_DESDE = {        # existencia nueva -> existencias desde las que el responsable puede llegar
    "CONFIRMADA": {"PENDIENTE", "FALTANTE"},
    "FALTANTE": {"PENDIENTE", "CONFIRMADA"},
}


class ErrorFlujo(Exception):
    """El cambio no está permitido (rol, estado de origen o dato inválido). El mensaje se muestra tal cual."""


def ahora():
    from router.entregas_alerta import ahora_nicaragua
    return ahora_nicaragua()


def bloquear_envio(cursor, envio_id):
    """Lee el envío con bloqueo de fila: dos personas cambiando a la vez no pisan el recálculo."""
    cursor.execute("SELECT * FROM ent_envios WHERE id = %s FOR UPDATE", (envio_id,))
    return cursor.fetchone()


def poner_estado(cursor, envio, nuevo, usuario_id, comentario=None):
    """Única función que cambia ent_envios.estado. Valida el estado de origen y registra el evento."""
    actual = envio["estado"]
    if nuevo == actual:
        return False
    if actual not in ORIGENES.get(nuevo, set()):
        raise ErrorFlujo(f"El envío {envio['pedido_num']} no puede pasar de {actual} a {nuevo}.")
    cursor.execute("UPDATE ent_envios SET estado = %s WHERE id = %s", (nuevo, envio["id"]))
    registrar_evento(cursor, envio["id"], usuario_id, nuevo, de_estado=actual, comentario=comentario)
    envio["estado"] = nuevo
    return True


# -----------------------------
# Facturación
# -----------------------------
def modo_factura_traslado(cursor):
    cursor.execute("SELECT valor FROM ent_configuracion WHERE clave = 'bodega_factura_traslado'")
    fila = cursor.fetchone()
    valor = (fila["valor"] if isinstance(fila, dict) else fila[0]) if fila else None
    return valor if valor in ("destino", "origen") else "destino"


def bodega_de_factura(linea, modo):
    """Bodega de facturación de una línea (None si es con traslado, se factura en destino y aún no tiene destino)."""
    if linea["traslado_estado"] != "NO_REQUIERE" and modo == "destino":
        return linea["bodega_prepara_id"]
    return linea["bodega_id"]


def bodegas_de_facturacion(lineas, modo):
    """Conjunto de bodegas que deben tener factura: las de las líneas no quitadas."""
    return {bodega_de_factura(l, modo) for l in lineas if l["existencia"] != "QUITADA"}


def estado_calculado(envio, lineas, facturadas=frozenset(), modo="destino"):
    """Estado automático que corresponde según las líneas y las facturas (de la orden enviada hasta retirada)."""
    activas = [l for l in lineas if l["existencia"] != "QUITADA"]
    if envio["estado"] == "FACTURADA":
        listas = activas and all(l["preparacion"] in ("LISTA", "RETIRADA") for l in activas) \
            and all(l["traslado_estado"] in ("NO_REQUIERE", "RECIBIDO") for l in activas)
        return "LISTA_PARA_RETIRO" if listas else "FACTURADA"
    if envio["estado"] == "LISTA_PARA_RETIRO":
        return "RETIRADA" if activas and all(l["preparacion"] == "RETIRADA" for l in activas) else "LISTA_PARA_RETIRO"
    if envio["estado"] not in ANTES_DE_FACTURAR:
        return envio["estado"]
    listas = (activas and all(l["existencia"] == "CONFIRMADA" for l in activas)
              and not any(l["traslado_estado"] == "PENDIENTE" for l in activas))
    if listas:
        requeridas = bodegas_de_facturacion(lineas, modo)
        if None not in requeridas and requeridas <= set(facturadas):
            return "FACTURADA" if envio["estado"] == "POR_FACTURAR" else "POR_FACTURAR"
        return "POR_FACTURAR"
    if any(l["existencia"] != "PENDIENTE" and not l.get("solo_factura") for l in lineas):
        return "EN_REVISION"
    return "ORDEN_ENVIADA" if envio["estado"] == "ORDEN_ENVIADA" else "EN_REVISION"


def recalcular_estado(cursor, envio, usuario_id):
    """Pone el estado automático que corresponde (puede avanzar más de un paso) y avisa. Devuelve el último estado nuevo o None."""
    cursor.execute("SELECT l.existencia, l.traslado_estado, l.preparacion, l.bodega_id, l.bodega_prepara_id, b.solo_factura "
                   "FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id WHERE l.envio_id = %s", (envio["id"],))
    lineas = cursor.fetchall()
    cursor.execute("SELECT bodega_id FROM ent_envio_facturas WHERE envio_id = %s", (envio["id"],))
    facturadas = {f["bodega_id"] if isinstance(f, dict) else f[0] for f in cursor.fetchall()}
    modo = modo_factura_traslado(cursor)
    cambio = None
    for _ in range(6):
        nuevo = estado_calculado(envio, lineas, facturadas, modo)
        if nuevo == envio["estado"]:
            break
        anterior = envio["estado"]
        comentario = {"POR_FACTURAR": "Todas las existencias confirmadas: el vendedor ya puede facturar",
                      "FACTURADA": "Registradas las facturas de todas las bodegas",
                      "LISTA_PARA_RETIRO": "Todo preparado y lo trasladado recibido: lista para retiro",
                      "RETIRADA": "Paquetería retiró la mercadería de todas las bodegas",
                      "EN_REVISION": ("Primera existencia registrada" if anterior == "ORDEN_ENVIADA"
                                      else "Volvió a revisión: una línea quedó faltante")}.get(nuevo)
        poner_estado(cursor, envio, nuevo, usuario_id, comentario)
        cambio = nuevo
        if nuevo == "POR_FACTURAR":
            avisar(cursor, [envio["vendedor_id"]], "listo_facturar",
                   f"Pedido {envio['pedido_num']}: listo para facturar. Revise los ajustes y registre una factura por bodega.",
                   envio_id=envio["id"], creado_por=usuario_id)
        elif nuevo == "FACTURADA":
            avisar(cursor, destinatarios(cursor, "paqueteria"), "facturada", f"Pedido {envio['pedido_num']}: facturado.",
                   envio_id=envio["id"], creado_por=usuario_id)
        elif nuevo == "LISTA_PARA_RETIRO":
            avisar(cursor, destinatarios(cursor, "paqueteria"), "lista_retiro",
                   f"Pedido {envio['pedido_num']}: listo para retiro en bodega.", envio_id=envio["id"], creado_por=usuario_id)
        elif nuevo == "EN_REVISION" and anterior == "POR_FACTURAR":
            avisar(cursor, [envio["vendedor_id"]], "faltante",
                   f"Pedido {envio['pedido_num']}: volvió a revisión, una línea quedó faltante. No facture todavía.",
                   envio_id=envio["id"], creado_por=usuario_id)
    return cambio


# -----------------------------
# Existencias (responsable de la línea)
# -----------------------------
def puede_registrar_existencia(roles, linea):
    """Regla 6.3: la existencia la cambia solo el responsable de la línea."""
    if linea["responsable"] == "coordinador":
        return "coordinador_bodega" in roles
    return "bodega" in roles and roles["bodega"] is not None and roles["bodega"] == linea["bodega_id"]


def leer_lineas_bloqueando(cursor, ids):
    cursor.execute(f"SELECT l.*, b.nombre AS bodega FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id "
                   f"WHERE l.id IN ({', '.join(['%s'] * len(ids))})", ids)
    lineas = {l["id"]: l for l in cursor.fetchall()}
    envios = {}
    for envio_id in sorted({l["envio_id"] for l in lineas.values()}):    # bloqueo en orden fijo: sin abrazos mortales
        envios[envio_id] = bloquear_envio(cursor, envio_id)
    return lineas, envios


def registrar_existencia(cursor, usuario_id, roles, cambios, comentario=None):
    """cambios: [(linea_id, 'CONFIRMADA' | 'FALTANTE', cant_disponible o None)]. Todo o nada: si una línea no se puede cambiar,
    lanza ErrorFlujo y el que llama hace rollback. Devuelve {envio_id: estado nuevo o None}."""
    if not cambios:
        raise ErrorFlujo("No se marcó ninguna línea.")
    lineas, envios = leer_lineas_bloqueando(cursor, [c[0] for c in cambios])
    momento = ahora()
    for linea_id, nueva, disponible in cambios:
        l = lineas.get(linea_id)
        if not l:
            raise ErrorFlujo("Una de las líneas ya no existe.")
        envio = envios[l["envio_id"]]
        nombre = f"Pedido {envio['pedido_num']}, línea {l['linea']} ({l['codigo']})"
        if envio["estado"] not in ANTES_DE_FACTURAR:
            raise ErrorFlujo(f"{nombre}: el envío está {envio['estado'].replace('_', ' ').lower()}; ya no se cambian existencias.")
        if not puede_registrar_existencia(roles, l):
            raise ErrorFlujo(f"{nombre}: la existencia la registra "
                             f"{'el coordinador de bodega' if l['responsable'] == 'coordinador' else 'el bodeguero de esa bodega'}.")
        if nueva not in EXISTENCIA_DESDE:
            raise ErrorFlujo(f"{nombre}: cambio no válido.")
        if l["existencia"] == nueva and nueva == "CONFIRMADA":
            continue
        if l["existencia"] not in EXISTENCIA_DESDE[nueva] and not (nueva == "FALTANTE" and l["existencia"] == "FALTANTE"):
            raise ErrorFlujo(f"{nombre}: está {l['existencia'].lower()}; no se puede marcar {nueva.lower()}.")
        if nueva == "FALTANTE":
            if disponible is None or disponible < 0 or disponible >= float(l["cantidad"]):
                raise ErrorFlujo(f"{nombre}: indique cuánto hay disponible (de 0 a menos de {float(l['cantidad']):g}).")
        else:
            disponible = None
        cursor.execute("UPDATE ent_envio_lineas SET existencia = %s, cant_disponible = %s, decision = NULL, existencia_por = %s, "
                       "existencia_en = %s, actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                       (nueva, disponible, usuario_id, momento, usuario_id, momento, linea_id))
        detalle = comentario
        if nueva == "FALTANTE":
            detalle = f"Hay {disponible:g} de {float(l['cantidad']):g}" + (f" · {comentario}" if comentario else "")
            avisar(cursor, [envio["vendedor_id"]] + destinatarios(cursor, "paqueteria"), "faltante",
                   f"Pedido {envio['pedido_num']}: falta {l['codigo']} en {l['bodega']} (hay {disponible:g} de {float(l['cantidad']):g}). "
                   f"El vendedor decide.", envio_id=envio["id"], linea_id=linea_id, creado_por=usuario_id)
        registrar_evento(cursor, envio["id"], usuario_id, nueva, de_estado=l["existencia"], campo="existencia", linea_id=linea_id,
                         comentario=detalle)
        l["existencia"] = nueva
    return {envio_id: recalcular_estado(cursor, envio, usuario_id) for envio_id, envio in envios.items()}


# -----------------------------
# Decisión del vendedor ante un faltante
# -----------------------------
def codigo_catalogo(cursor, numero_parte, marca):
    """Código del catálogo por número de parte y marca (en el PDF, «Cod./#Parte» es el número de parte). None si no es único."""
    cursor.execute("SELECT DISTINCT `Codigo Producto` AS c FROM catalogo_productos WHERE `Numero Parte` = %s "
                   "AND REPLACE(UPPER(COALESCE(Marca, '')), ' ', '') = REPLACE(UPPER(%s), ' ', '')", (numero_parte, marca or ""))
    filas = cursor.fetchall()
    return filas[0]["c"] if len(filas) == 1 else None


def registrar_venta_perdida(cursor, envio, linea, cantidad, vendedor, detalle):
    cursor.execute("""INSERT INTO ventas_perdidas (nomb_vendedor, nomb_cliente, fecha, numero_parte, codigo_producto, descripcion, marca,
                          cantidad, motivo, nota) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Sin existencia', %s)""",
                   (vendedor[:50], envio["cliente"][:100], ahora().date(), linea["codigo"][:100], codigo_catalogo(cursor, linea["codigo"], linea["marca"]),
                    (linea["producto"] or "")[:100] or None, (linea["marca"] or "")[:100] or None, cantidad,
                    f"Entregas: pedido {envio['pedido_num']}, línea {linea['linea']} en {linea['bodega']} ({detalle})"))


def decidir_faltante(cursor, usuario, roles, linea_id, decision, bodega_nueva=None):
    """decision: 'quitar' | 'parcial' (va la parte disponible) | 'cambiar_bodega' | 'esperar'. Solo el vendedor dueño del envío,
    con la línea FALTANTE y el envío antes de facturar. Devuelve el texto de lo hecho."""
    lineas, envios = leer_lineas_bloqueando(cursor, [linea_id])
    l = lineas.get(linea_id)
    if not l:
        raise ErrorFlujo("La línea ya no existe.")
    envio = envios[l["envio_id"]]
    nombre = f"Pedido {envio['pedido_num']}, línea {l['linea']} ({l['codigo']})"
    if "vendedor" not in roles or envio["vendedor_id"] != usuario["id"]:
        raise ErrorFlujo(f"{nombre}: solo el vendedor de la orden decide sus faltantes.")
    if envio["estado"] not in ANTES_DE_FACTURAR:
        raise ErrorFlujo(f"{nombre}: la orden ya no está en revisión.")
    if l["existencia"] != "FALTANTE":
        raise ErrorFlujo(f"{nombre}: ya no está faltante.")
    cantidad, disponible = float(l["cantidad"]), float(l["cant_disponible"] or 0)
    momento, uid = ahora(), usuario["id"]
    if decision == "quitar":
        cursor.execute("UPDATE ent_envio_lineas SET existencia = 'QUITADA', decision = 'quitar', actualizado_por = %s, actualizado_en = %s "
                       "WHERE id = %s", (uid, momento, linea_id))
        registrar_venta_perdida(cursor, envio, l, cantidad, usuario["usuario"], "quitada de la orden")
        registrar_evento(cursor, envio["id"], uid, "QUITADA", "FALTANTE", "existencia", linea_id,
                         f"El vendedor la quita de la orden; {cantidad:g} como venta perdida")
        hecho = f"Línea {l['linea']} quitada; {cantidad:g} quedan como venta perdida. Quítela de la orden al facturar."
    elif decision == "parcial":
        if disponible <= 0:
            raise ErrorFlujo(f"{nombre}: no hay parte disponible; quítela, cámbiela de bodega o espere.")
        cursor.execute("UPDATE ent_envio_lineas SET existencia = 'CONFIRMADA', decision = NULL, actualizado_por = %s, actualizado_en = %s "
                       "WHERE id = %s", (uid, momento, linea_id))
        registrar_venta_perdida(cursor, envio, l, cantidad - disponible, usuario["usuario"], f"va {disponible:g} de {cantidad:g}")
        registrar_evento(cursor, envio["id"], uid, "CONFIRMADA", "FALTANTE", "existencia", linea_id,
                         f"Va la parte disponible: {disponible:g} de {cantidad:g}; {cantidad - disponible:g} como venta perdida")
        hecho = f"Línea {l['linea']}: van {disponible:g} de {cantidad:g}. Ajuste la cantidad en la orden al facturar."
    elif decision == "cambiar_bodega":
        cursor.execute("SELECT * FROM ent_bodegas WHERE id = %s AND activo = 1 AND solo_factura = 0", (bodega_nueva,))
        b = cursor.fetchone()
        if not b or b["id"] == l["bodega_id"]:
            raise ErrorFlujo(f"{nombre}: escoja otra bodega activa.")
        responsable = "coordinador" if b["requiere_traslado"] else "bodeguero"
        cursor.execute("""UPDATE ent_envio_lineas SET bodega_id = %s, responsable = %s, traslado_estado = %s, bodega_prepara_id = %s,
                              existencia = 'PENDIENTE', cant_disponible = NULL, decision = 'cambiar_bodega', existencia_por = NULL,
                              existencia_en = NULL, traslado_num = NULL, actualizado_por = %s, actualizado_en = %s WHERE id = %s""",
                       (b["id"], responsable, "PENDIENTE" if b["requiere_traslado"] else "NO_REQUIERE",
                        None if b["requiere_traslado"] else b["id"], uid, momento, linea_id))
        registrar_evento(cursor, envio["id"], uid, "PENDIENTE", "FALTANTE", "existencia", linea_id,
                         f"Cambio de bodega: {l['bodega']} → {b['nombre']}")
        avisar(cursor, responsables_linea(cursor, {"responsable": responsable, "bodega_id": b["id"]}), "cambio_bodega",
               f"Pedido {envio['pedido_num']}: nueva línea para confirmar, {l['codigo']} ({cantidad:g}) pasó de {l['bodega']} a {b['nombre']}.",
               envio_id=envio["id"], linea_id=linea_id, creado_por=uid)
        hecho = f"Línea {l['linea']} pasó a {b['nombre']}; queda pendiente de confirmar."
    elif decision == "esperar":
        cursor.execute("UPDATE ent_envio_lineas SET decision = 'esperar', actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                       (uid, momento, linea_id))
        registrar_evento(cursor, envio["id"], uid, "FALTANTE", "FALTANTE", "existencia", linea_id, "El vendedor decide esperar")
        hecho = f"Línea {l['linea']}: se espera a que la bodega la confirme."
    else:
        raise ErrorFlujo("Decisión no válida.")
    if recalcular_estado(cursor, envio, uid) == "POR_FACTURAR":
        hecho += " La orden quedó lista para facturar."
    return hecho


# -----------------------------
# Facturas (una por bodega de facturación)
# -----------------------------
def registrar_facturas(cursor, usuario, roles, envio_id, facturas):
    """facturas: {bodega_id: número}. Solo el vendedor dueño y con el envío POR_FACTURAR. Guarda las que vengan con número (puede
    registrarlas por partes) y recalcula: con todas, el envío pasa a FACTURADA. Devuelve (guardadas, faltan)."""
    envio = bloquear_envio(cursor, envio_id)
    if not envio:
        raise ErrorFlujo("El envío no existe.")
    if "vendedor" not in roles or envio["vendedor_id"] != usuario["id"]:
        raise ErrorFlujo("Solo el vendedor de la orden registra sus facturas.")
    if envio["estado"] != "POR_FACTURAR":
        raise ErrorFlujo(f"Pedido {envio['pedido_num']}: no se pueden registrar facturas mientras falte confirmar existencias o aplicar "
                         f"un traslado (está {envio['estado'].replace('_', ' ').lower()}).")
    cursor.execute("SELECT l.*, b.nombre AS bodega FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id WHERE l.envio_id = %s",
                   (envio_id,))
    lineas = cursor.fetchall()
    requeridas = bodegas_de_facturacion(lineas, modo_factura_traslado(cursor))
    if None in requeridas:
        raise ErrorFlujo("Hay líneas con traslado sin bodega destino: el coordinador debe escogerla antes de facturar.")
    guardadas = 0
    for bodega_id, numero in facturas.items():
        numero = " ".join((numero or "").split())[:30]
        if not numero:
            continue
        if bodega_id not in requeridas:
            raise ErrorFlujo("Esa bodega no tiene líneas para facturar en esta orden.")
        cursor.execute("SELECT factura_num FROM ent_envio_facturas WHERE envio_id = %s AND bodega_id = %s", (envio_id, bodega_id))
        previa = cursor.fetchone()
        if previa and previa["factura_num"] == numero:
            continue
        cursor.execute("INSERT INTO ent_envio_facturas (envio_id, bodega_id, factura_num, registrada_por) VALUES (%s, %s, %s, %s) "
                       "ON DUPLICATE KEY UPDATE factura_num = VALUES(factura_num), registrada_por = VALUES(registrada_por), "
                       "registrada_en = CURRENT_TIMESTAMP", (envio_id, bodega_id, numero, usuario["id"]))
        cursor.execute("SELECT nombre FROM ent_bodegas WHERE id = %s", (bodega_id,))
        bodega = cursor.fetchone()["nombre"]
        registrar_evento(cursor, envio_id, usuario["id"], "REGISTRADA" if not previa else "CORREGIDA", campo="factura",
                         comentario=f"{bodega}: factura {numero}" + (f" (antes {previa['factura_num']})" if previa else ""))
        guardadas += 1
    recalcular_estado(cursor, envio, usuario["id"])
    cursor.execute("SELECT bodega_id FROM ent_envio_facturas WHERE envio_id = %s", (envio_id,))
    faltan = requeridas - {f["bodega_id"] for f in cursor.fetchall()}
    return guardadas, len(faltan)


# -----------------------------
# Anular
# -----------------------------
def anular(cursor, usuario, roles, envio_id, motivo):
    envio = bloquear_envio(cursor, envio_id)
    if not envio:
        raise ErrorFlujo("El envío no existe.")
    if not ("admin" in roles or ("vendedor" in roles and envio["vendedor_id"] == usuario["id"])):
        raise ErrorFlujo("Solo el vendedor de la orden o el administrador pueden anularla.")
    motivo = " ".join((motivo or "").split())[:255]
    if not motivo:
        raise ErrorFlujo("Escriba el motivo de la anulación.")
    poner_estado(cursor, envio, "ANULADA", usuario["id"], f"Anulada: {motivo}")
    cursor.execute("UPDATE ent_envios SET motivo_anulacion = %s WHERE id = %s", (motivo, envio_id))
    cursor.execute("SELECT responsable, bodega_id FROM ent_envio_lineas WHERE envio_id = %s AND existencia <> 'QUITADA'", (envio_id,))
    lineas = cursor.fetchall()
    afectados = destinatarios(cursor, "paqueteria")
    for l in lineas:
        afectados += responsables_linea(cursor, l)
    avisar(cursor, afectados, "anulada", f"Pedido {envio['pedido_num']} anulado: {motivo}", envio_id=envio_id, creado_por=usuario["id"])


# -----------------------------
# Traslados (coordinador de bodega)
# -----------------------------
# acción -> (estado de traslado de origen, estado nuevo, estados del envío en que se permite)
TRASLADO = {
    "aplicar": ("PENDIENTE", "APLICADO", ANTES_DE_FACTURAR),
    "en_camino": ("APLICADO", "EN_CAMINO", ANTES_DE_FACTURAR + ("FACTURADA",)),
    "recibido": ("EN_CAMINO", "RECIBIDO", ANTES_DE_FACTURAR + ("FACTURADA",)),
}


def cantidad_que_va(linea):
    return float(linea["cant_disponible"]) if linea["existencia"] == "CONFIRMADA" and linea["cant_disponible"] is not None else float(linea["cantidad"])


def registrar_traslado(cursor, usuario_id, roles, accion, linea_ids, destino_id=None, numero=None):
    """Regla 6.3: el traslado lo cambia solo el coordinador. «aplicar» exige la existencia confirmada, el número del traslado del
    sistema de ventas y la bodega destino (la escoge el coordinador según el tipo de producto); con eso el vendedor ya puede
    facturar. «en_camino»: mandó a traer la mercadería. «recibido»: llegó; la línea pasa a la cola del bodeguero de la bodega
    destino, que la prepara (parte 9). Todo o nada. Devuelve {envio_id: estado nuevo o None}."""
    if "coordinador_bodega" not in roles:
        raise ErrorFlujo("Los traslados los registra el coordinador de bodega.")
    if accion not in TRASLADO:
        raise ErrorFlujo("Acción de traslado no válida.")
    if not linea_ids:
        raise ErrorFlujo("No se marcó ninguna línea.")
    desde, hacia, estados_envio = TRASLADO[accion]
    destino = None
    if accion == "aplicar":
        numero = " ".join((numero or "").split())[:30]
        if not numero:
            raise ErrorFlujo("Escriba el número del traslado aplicado en el sistema de ventas.")
        cursor.execute("SELECT * FROM ent_bodegas WHERE id = %s AND activo = 1 AND solo_factura = 0", (destino_id,))
        destino = cursor.fetchone()
        if not destino:
            raise ErrorFlujo("Escoja la bodega a la que se traslada.")
    lineas, envios = leer_lineas_bloqueando(cursor, list(linea_ids))
    momento = ahora()
    for linea_id in linea_ids:
        l = lineas.get(linea_id)
        if not l:
            raise ErrorFlujo("Una de las líneas ya no existe.")
        envio = envios[l["envio_id"]]
        nombre = f"Pedido {envio['pedido_num']}, línea {l['linea']} ({l['codigo']})"
        if l["traslado_estado"] == "NO_REQUIERE":
            raise ErrorFlujo(f"{nombre}: no requiere traslado.")
        if envio["estado"] not in estados_envio:
            raise ErrorFlujo(f"{nombre}: el envío está {envio['estado'].replace('_', ' ').lower()}.")
        if l["existencia"] == "QUITADA":
            raise ErrorFlujo(f"{nombre}: la línea se quitó de la orden.")
        if accion == "aplicar" and l["existencia"] != "CONFIRMADA":
            raise ErrorFlujo(f"{nombre}: confirme primero la existencia.")
        if l["traslado_estado"] != desde:
            raise ErrorFlujo(f"{nombre}: el traslado está {l['traslado_estado'].replace('_', ' ').lower()}.")
        if accion == "aplicar":
            if destino["id"] == l["bodega_id"]:
                raise ErrorFlujo(f"{nombre}: el destino no puede ser la misma bodega.")
            cursor.execute("UPDATE ent_envio_lineas SET traslado_estado = 'APLICADO', traslado_num = %s, traslado_aplicado_en = %s, "
                           "bodega_prepara_id = %s, actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                           (numero, momento, destino["id"], usuario_id, momento, linea_id))
            comentario = f"Traslado N° {numero}: {l['bodega']} → {destino['nombre']}"
        elif accion == "en_camino":
            cursor.execute("UPDATE ent_envio_lineas SET traslado_estado = 'EN_CAMINO', actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                           (usuario_id, momento, linea_id))
            comentario = "Se mandó a traer la mercadería"
        else:
            cursor.execute("UPDATE ent_envio_lineas SET traslado_estado = 'RECIBIDO', traslado_recibido_en = %s, actualizado_por = %s, "
                           "actualizado_en = %s WHERE id = %s", (momento, usuario_id, momento, linea_id))
            cursor.execute("SELECT nombre FROM ent_bodegas WHERE id = %s", (l["bodega_prepara_id"],))
            destino_nombre = cursor.fetchone()["nombre"]
            comentario = f"Llegó a {destino_nombre}"
            avisar(cursor, destinatarios(cursor, "bodega", l["bodega_prepara_id"]), "traslado_recibido",
                   f"Pedido {envio['pedido_num']}: llegó por traslado {l['codigo']} ({cantidad_que_va(l):g}) a {destino_nombre}. Ya se puede preparar.",
                   envio_id=envio["id"], linea_id=linea_id, creado_por=usuario_id)
        registrar_evento(cursor, envio["id"], usuario_id, hacia, de_estado=desde, campo="traslado", linea_id=linea_id, comentario=comentario)
        l["traslado_estado"] = hacia
    return {envio_id: recalcular_estado(cursor, envio, usuario_id) for envio_id, envio in envios.items()}


# -----------------------------
# Preparación (bodeguero de la bodega que prepara) y retiro (paquetería)
# -----------------------------
PREPARACION = {"empezar": ({"PENDIENTE"}, "EN_PREPARACION"), "lista": ({"PENDIENTE", "EN_PREPARACION"}, "LISTA")}
SE_PREPARA_EN = ("EN_REVISION", "POR_FACTURAR", "FACTURADA")


def puede_preparar(roles, linea):
    """Regla 6.3: la preparación la registra solo el bodeguero de bodega_prepara_id."""
    return "bodega" in roles and roles["bodega"] is not None and roles["bodega"] == linea["bodega_prepara_id"]


def registrar_preparacion(cursor, usuario_id, roles, accion, linea_ids):
    """«empezar» (EN_PREPARACION) o «lista» (LISTA) de una o varias líneas. Todo o nada. Devuelve {envio_id: estado nuevo o None}."""
    if accion not in PREPARACION:
        raise ErrorFlujo("Acción de preparación no válida.")
    if not linea_ids:
        raise ErrorFlujo("No se marcó ninguna línea.")
    desde, hacia = PREPARACION[accion]
    lineas, envios = leer_lineas_bloqueando(cursor, list(linea_ids))
    momento = ahora()
    for linea_id in linea_ids:
        l = lineas.get(linea_id)
        if not l:
            raise ErrorFlujo("Una de las líneas ya no existe.")
        envio = envios[l["envio_id"]]
        nombre = f"Pedido {envio['pedido_num']}, línea {l['linea']} ({l['codigo']})"
        if not puede_preparar(roles, l):
            raise ErrorFlujo(f"{nombre}: la prepara el bodeguero de la bodega {'destino del traslado' if l['traslado_estado'] != 'NO_REQUIERE' else 'de la línea'}.")
        if l["existencia"] != "CONFIRMADA":
            raise ErrorFlujo(f"{nombre}: la existencia no está confirmada.")
        if envio["estado"] not in SE_PREPARA_EN:
            raise ErrorFlujo(f"{nombre}: el envío está {envio['estado'].replace('_', ' ').lower()}.")
        if l["traslado_estado"] not in ("NO_REQUIERE", "RECIBIDO"):
            raise ErrorFlujo(f"{nombre}: la mercadería trasladada todavía no llegó; se prepara cuando el coordinador la marque recibida.")
        if l["preparacion"] == hacia:
            continue
        if l["preparacion"] not in desde:
            raise ErrorFlujo(f"{nombre}: ya está {l['preparacion'].replace('_', ' ').lower()}.")
        cursor.execute("UPDATE ent_envio_lineas SET preparacion = %s, actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                       (hacia, usuario_id, momento, linea_id))
        registrar_evento(cursor, envio["id"], usuario_id, hacia, de_estado=l["preparacion"], campo="preparacion", linea_id=linea_id)
        l["preparacion"] = hacia
    return {envio_id: recalcular_estado(cursor, envio, usuario_id) for envio_id, envio in envios.items()}


def retirar_bodega(cursor, usuario_id, roles, envio_id, bodega_id):
    """Paquetería retira de una bodega todas las líneas listas de esa orden a la vez (sección 4.2). Solo con la orden
    LISTA_PARA_RETIRO (todo preparado y lo trasladado recibido). Devuelve (líneas retiradas, estado nuevo o None)."""
    if "paqueteria" not in roles:
        raise ErrorFlujo("El retiro lo registra paquetería.")
    envio = bloquear_envio(cursor, envio_id)
    if not envio:
        raise ErrorFlujo("El envío no existe.")
    if envio["estado"] != "LISTA_PARA_RETIRO":
        raise ErrorFlujo(f"Pedido {envio['pedido_num']}: el retiro se habilita cuando la orden está facturada, todo preparado y lo "
                         f"trasladado recibido (está {envio['estado'].replace('_', ' ').lower()}).")
    cursor.execute("SELECT * FROM ent_envio_lineas WHERE envio_id = %s AND bodega_prepara_id = %s AND existencia <> 'QUITADA'",
                   (envio_id, bodega_id))
    lineas = [l for l in cursor.fetchall() if l["preparacion"] == "LISTA"]
    if not lineas:
        raise ErrorFlujo(f"Pedido {envio['pedido_num']}: no hay líneas listas para retirar en esa bodega.")
    cursor.execute("SELECT nombre FROM ent_bodegas WHERE id = %s", (bodega_id,))
    bodega = cursor.fetchone()["nombre"]
    momento = ahora()
    for l in lineas:
        cursor.execute("UPDATE ent_envio_lineas SET preparacion = 'RETIRADA', actualizado_por = %s, actualizado_en = %s WHERE id = %s",
                       (usuario_id, momento, l["id"]))
        registrar_evento(cursor, envio_id, usuario_id, "RETIRADA", de_estado="LISTA", campo="preparacion", linea_id=l["id"],
                         comentario=f"Retirada de {bodega}")
    return len(lineas), recalcular_estado(cursor, envio, usuario_id)


# -----------------------------
# Empaque (paquetería)
# -----------------------------
def empacar(cursor, usuario_id, roles, envio_id, bultos, peso):
    """Registra bultos y peso. Con la orden RETIRADA la pasa a EMPACADA; con la orden ya EMPACADA solo corrige los datos (para
    reimprimir la etiqueta). Devuelve True si cambió el estado."""
    if "paqueteria" not in roles:
        raise ErrorFlujo("El empaque lo registra paquetería.")
    envio = bloquear_envio(cursor, envio_id)
    if not envio:
        raise ErrorFlujo("El envío no existe.")
    if envio["estado"] not in ("RETIRADA", "EMPACADA"):
        raise ErrorFlujo(f"Pedido {envio['pedido_num']}: se empaca cuando paquetería ya retiró toda la mercadería "
                         f"(está {envio['estado'].replace('_', ' ').lower()}).")
    if not isinstance(bultos, int) or bultos < 1 or bultos > 999:
        raise ErrorFlujo("Escriba cuántos bultos son (1 o más).")
    if peso is None or peso <= 0 or peso >= 1000000:
        raise ErrorFlujo("Escriba el peso en kg (mayor que 0).")
    peso = round(peso, 2)
    cursor.execute("UPDATE ent_envios SET bultos = %s, peso_kg = %s WHERE id = %s", (bultos, peso, envio_id))
    texto = f"{bultos} bulto(s), {peso:g} kg"
    if envio["estado"] == "RETIRADA":
        poner_estado(cursor, envio, "EMPACADA", usuario_id, f"Empacada: {texto}")
        return True
    if (envio["bultos"], float(envio["peso_kg"] or 0)) != (bultos, peso):
        registrar_evento(cursor, envio_id, usuario_id, "EMPACADA", de_estado="EMPACADA",
                         comentario=f"Corrigió el empaque: {texto} (antes {envio['bultos']} bulto(s), {float(envio['peso_kg'] or 0):g} kg)")
    return False
