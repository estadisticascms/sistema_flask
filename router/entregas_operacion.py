"""Entregas · parte 6: existencias y faltantes (bodeguero y coordinador) y tablero de seguimiento de paquetería.

- Mi bodega (/entregas/bodega, rol Bodeguero): líneas de su bodega cuya existencia le toca (bodegas sin traslado), por prioridad
  y riesgo. Confirma o marca faltante por línea o toda su parte de una orden. Preparación (parte 9): las líneas que prepara su
  bodega, incluida la mercadería trasladada que ya llegó; lo que viene en camino se muestra aparte y no se puede marcar.
- Bodegas con traslado (/entregas/coordinador, rol Coordinador): lo mismo para las líneas de bodegas con traslado, y los traslados
  (parte 8): aplicar con número y bodega destino (según el tipo de producto, que se muestra), mandar a traer y marcar recibido.
- Seguimiento de órdenes (/entregas/seguimiento, rol Paquetería): todas las órdenes abiertas, avance por bodega, tiempo en el paso
  actual y alerta de tiempo. Solo ve; no confirma en nombre de otros.
- Pendientes (/entregas/pendientes): contadores por rol para las pestañas, que se actualizan solos cada minuto.

Los cambios pasan por router/entregas_flujo.py (permiso del responsable, evento y recálculo del estado del envío).
"""
from collections import defaultdict

from flask import Blueprint, flash, jsonify, redirect, render_template, request, url_for

from router.entregas import contexto_entregas, requiere_modulo, requiere_rol, roles_actuales, usuario_actual
from router.entregas_avisos import avisar, contar_no_leidos, responsables_linea
from router.entregas_alerta import ahora_nicaragua, alertas_de_envios, duracion
from router.entregas_flujo import (ABIERTOS, ANTES_DE_FACTURAR, SE_PREPARA_EN, ErrorFlujo, cantidad_que_va, registrar_existencia,
                                   registrar_preparacion, registrar_traslado)
from router.entregas_orden import NOMBRE_ESTADO
from utils.db import get_connection

entregas_operacion_bp = Blueprint("entregas_operacion", __name__)
entregas_operacion_bp.context_processor(contexto_entregas)

ORDEN_RIESGO = {"rojo": 0, "amarillo": 1, "verde": 2, None: 3}


def a_numero(valor):
    try:
        return float(str(valor).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def nivel(envio, alerta):
    """Nivel que se muestra: rojo si el vendedor confirmó la hora en riesgo, si no el recalculado."""
    if envio.get("riesgo_confirmado"):
        return "rojo"
    return alerta["nivel"] if alerta else None


def orden_cola(envio):
    return (ORDEN_RIESGO.get(envio["nivel"], 3), envio["prioridad"] != "urgente", envio["hora_requerida"])


# -----------------------------
# Colas de existencias (bodeguero y coordinador)
# -----------------------------
def cola_existencias(cursor, responsable, bodega_id=None):
    """Envíos antes de facturar con líneas cuya existencia le toca a este responsable. -> lista de envíos con sus líneas."""
    filtro, params = "l.responsable = %s", [responsable]
    if responsable == "bodeguero":
        filtro += " AND l.bodega_id = %s"
        params.append(bodega_id)
    cursor.execute(f"""
        SELECT l.*, b.nombre AS bodega FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id JOIN ent_bodegas b ON b.id = l.bodega_id
        WHERE {filtro} AND e.estado IN ({', '.join(['%s'] * len(ANTES_DE_FACTURAR))})
        ORDER BY l.envio_id, l.linea""", params + list(ANTES_DE_FACTURAR))
    mias = cursor.fetchall()
    if not mias:
        return []
    ids = sorted({l["envio_id"] for l in mias})
    marcas = ", ".join(["%s"] * len(ids))
    cursor.execute(f"""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista FROM ent_envios e
                       JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                       WHERE e.id IN ({marcas})""", ids)
    envios = cursor.fetchall()
    cursor.execute(f"SELECT envio_id, existencia, traslado_estado, preparacion FROM ent_envio_lineas WHERE envio_id IN ({marcas})", ids)
    todas = defaultdict(list)
    for l in cursor.fetchall():
        todas[l["envio_id"]].append(l)
    alertas = alertas_de_envios(cursor, envios, todas)
    for e in envios:
        e["alerta"] = alertas[e["id"]]
        e["nivel"] = nivel(e, e["alerta"])
        e["lineas"] = [l for l in mias if l["envio_id"] == e["id"]]
        e["pendientes"] = sum(1 for l in e["lineas"] if l["existencia"] == "PENDIENTE")
        e["faltantes"] = sum(1 for l in e["lineas"] if l["existencia"] == "FALTANTE")
        e["otras"] = len(todas[e["id"]]) - len(e["lineas"])
    envios.sort(key=lambda e: (e["pendientes"] == 0 and e["faltantes"] == 0,) + orden_cola(e))
    return envios


def tipo_producto(cursor, numero_parte, marca):
    """Tipo de producto para ayudar al coordinador a escoger el destino: «Tipo Producto» del catálogo y familia de la clasificación,
    buscando por número de parte y marca (en el PDF, «Cod./#Parte» es el número de parte)."""
    cursor.execute("""SELECT c.`Codigo Producto` AS codigo, c.`Tipo Producto` AS tipo, k.familia, k.subfamilia FROM catalogo_productos c
                      LEFT JOIN clase_prod k ON k.codigo = c.`Codigo Producto`
                      WHERE c.`Numero Parte` = %s AND REPLACE(UPPER(COALESCE(c.Marca, '')), ' ', '') = REPLACE(UPPER(%s), ' ', '') LIMIT 2""",
                   (numero_parte, marca or ""))
    filas = cursor.fetchall()
    if len(filas) != 1:
        return None
    f = filas[0]
    partes = [p for p in (f["tipo"], f["familia"], f["subfamilia"]) if p and p not in ("Por clasificar", "Sin definir")]
    return {"codigo": f["codigo"], "texto": " · ".join(dict.fromkeys(partes))}


TRASLADO_ABIERTO = ("PENDIENTE", "APLICADO", "EN_CAMINO")


def cola_traslados(cursor):
    """Líneas con traslado sin recibir de envíos abiertos (hasta facturada), agrupadas por envío."""
    estados = ANTES_DE_FACTURAR + ("FACTURADA",)
    cursor.execute(f"""
        SELECT l.*, b.nombre AS bodega, d.nombre AS destino, e.pedido_num, e.cliente, e.hora_requerida, e.prioridad, e.estado AS estado_envio
        FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id JOIN ent_bodegas b ON b.id = l.bodega_id
        LEFT JOIN ent_bodegas d ON d.id = l.bodega_prepara_id
        WHERE l.traslado_estado IN ('PENDIENTE', 'APLICADO', 'EN_CAMINO') AND l.existencia <> 'QUITADA'
          AND e.estado IN ({', '.join(['%s'] * len(estados))})
        ORDER BY e.hora_requerida, l.envio_id, l.linea""", list(estados))
    lineas = cursor.fetchall()
    grupos = {"aplicar": [], "en_camino": [], "recibir": []}
    for l in lineas:
        l["va"] = cantidad_que_va(l)
        if l["traslado_estado"] == "PENDIENTE":
            l["tipo"] = tipo_producto(cursor, l["codigo"], l["marca"])
            grupos["aplicar"].append(l)
        else:
            grupos["en_camino" if l["traslado_estado"] == "APLICADO" else "recibir"].append(l)
    return grupos


def cola_preparacion(cursor, bodega_id):
    """Lo que prepara esta bodega: envíos con líneas confirmadas por preparar (y lo trasladado ya recibido), más lo que viene en
    camino por traslado. -> (envíos con sus líneas a preparar, líneas en camino)."""
    estados = ", ".join(["%s"] * len(SE_PREPARA_EN))
    cursor.execute(f"""
        SELECT l.*, b.nombre AS bodega FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id JOIN ent_bodegas b ON b.id = l.bodega_id
        WHERE l.bodega_prepara_id = %s AND l.existencia = 'CONFIRMADA' AND l.preparacion IN ('PENDIENTE', 'EN_PREPARACION')
          AND l.traslado_estado IN ('NO_REQUIERE', 'RECIBIDO') AND e.estado IN ({estados})
        ORDER BY l.envio_id, l.linea""", [bodega_id] + list(SE_PREPARA_EN))
    lineas = cursor.fetchall()
    cursor.execute(f"""
        SELECT l.*, b.nombre AS bodega, e.pedido_num, e.hora_requerida FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
        JOIN ent_bodegas b ON b.id = l.bodega_id
        WHERE l.bodega_prepara_id = %s AND l.traslado_estado IN ('APLICADO', 'EN_CAMINO') AND l.existencia <> 'QUITADA' AND e.estado IN ({estados})
        ORDER BY e.hora_requerida""", [bodega_id] + list(SE_PREPARA_EN))
    en_camino = cursor.fetchall()
    for l in lineas + en_camino:
        l["va"] = cantidad_que_va(l)
    if not lineas:
        return [], en_camino
    ids = sorted({l["envio_id"] for l in lineas})
    marcas = ", ".join(["%s"] * len(ids))
    cursor.execute(f"""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista FROM ent_envios e
                       JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                       WHERE e.id IN ({marcas})""", ids)
    envios = cursor.fetchall()
    cursor.execute(f"SELECT envio_id, existencia, traslado_estado, preparacion FROM ent_envio_lineas WHERE envio_id IN ({marcas})", ids)
    todas = defaultdict(list)
    for l in cursor.fetchall():
        todas[l["envio_id"]].append(l)
    alertas = alertas_de_envios(cursor, envios, todas)
    for e in envios:
        e["alerta"] = alertas[e["id"]]
        e["nivel"] = nivel(e, e["alerta"])
        e["lineas"] = [l for l in lineas if l["envio_id"] == e["id"]]
    envios.sort(key=orden_cola)
    return envios, en_camino


def pantalla_existencias(modo):
    roles = roles_actuales()
    bodega_id = roles.get("bodega") if modo == "bodega" else None
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        bodega = None
        if modo == "bodega" and bodega_id:
            cursor.execute("SELECT * FROM ent_bodegas WHERE id = %s", (bodega_id,))
            bodega = cursor.fetchone()
        envios = cola_existencias(cursor, "bodeguero" if modo == "bodega" else "coordinador", bodega_id) \
            if (modo != "bodega" or bodega_id) else []
        traslados, bodegas = None, []
        preparar, en_camino = cola_preparacion(cursor, bodega_id) if modo == "bodega" and bodega_id else ([], [])
        if modo == "coordinador":
            traslados = cola_traslados(cursor)
            cursor.execute("SELECT id, nombre FROM ent_bodegas WHERE activo = 1 AND solo_factura = 0 ORDER BY nombre")
            bodegas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    return render_template("entregas_existencias.html", modo=modo, bodega=bodega, sin_bodega=modo == "bodega" and not bodega_id,
                           envios=envios, nombre_estado=NOMBRE_ESTADO, ahora=ahora_nicaragua(), traslados=traslados, bodegas=bodegas, preparar=preparar, en_camino=en_camino,
                           activo="entregas_bodega" if modo == "bodega" else "entregas_coordinador",
                           totales={"pendientes": sum(e["pendientes"] for e in envios), "faltantes": sum(e["faltantes"] for e in envios),
                                    "rojo": sum(1 for e in envios if e["nivel"] == "rojo")})


@entregas_operacion_bp.route("/entregas/bodega")
@requiere_rol("entregas_bodega", "bodega")
def bodega():
    return pantalla_existencias("bodega")


@entregas_operacion_bp.route("/entregas/coordinador")
@requiere_rol("entregas_coordinador", "coordinador_bodega")
def coordinador():
    return pantalla_existencias("coordinador")


@entregas_operacion_bp.route("/entregas/bodega/existencia", methods=["POST"])
@requiere_rol("entregas_bodega", "bodega")
def existencia_bodega():
    return guardar_existencia(url_for("entregas_operacion.bodega"))


@entregas_operacion_bp.route("/entregas/coordinador/existencia", methods=["POST"])
@requiere_rol("entregas_coordinador", "coordinador_bodega")
def existencia_coordinador():
    return guardar_existencia(url_for("entregas_operacion.coordinador"))


@entregas_operacion_bp.route("/entregas/coordinador/traslado", methods=["POST"])
@requiere_rol("entregas_coordinador", "coordinador_bodega")
def traslado():
    """Aplicar (una línea, con número y destino), mandar a traer o marcar recibido (una o varias líneas)."""
    accion = request.form.get("accion")
    ids = [int(x) for x in request.form.getlist("linea_id") if x.isdigit()]
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cambios = registrar_traslado(cursor, u["id"], roles_actuales(), accion, ids, request.form.get("destino_id", type=int),
                                     request.form.get("numero"))
        conn.commit()
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return redirect(url_for("entregas_operacion.coordinador") + "#traslados")
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    mensaje = {"aplicar": "✅ Traslado aplicado.", "en_camino": f"✅ {len(ids)} línea(s) en camino.",
               "recibido": f"✅ {len(ids)} línea(s) recibida(s): se avisó al bodeguero que las prepara."}[accion]
    if "POR_FACTURAR" in cambios.values():
        mensaje += " La orden quedó lista para facturar."
    flash(mensaje, "success")
    return redirect(url_for("entregas_operacion.coordinador") + "#traslados")


@entregas_operacion_bp.route("/entregas/bodega/preparacion", methods=["POST"])
@requiere_rol("entregas_bodega", "bodega")
def preparacion():
    """Empezar o marcar lista una línea o toda la parte de una orden que prepara esta bodega."""
    accion = request.form.get("accion")
    ids = [int(x) for x in request.form.getlist("linea_id") if x.isdigit()]
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cambios = registrar_preparacion(cursor, u["id"], roles_actuales(), accion, ids)
        conn.commit()
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return redirect(url_for("entregas_operacion.bodega") + "#preparacion")
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    mensaje = f"✅ {len(ids)} línea(s) {'en preparación' if accion == 'empezar' else 'lista(s)'}."
    if "LISTA_PARA_RETIRO" in cambios.values():
        mensaje += " La orden quedó lista para retiro: se avisó a paquetería."
    flash(mensaje, "success")
    return redirect(url_for("entregas_operacion.bodega") + "#preparacion")


def guardar_existencia(volver):
    """Confirmar o marcar faltante: una línea (linea_id) o toda la parte de una orden (varias linea_id). El permiso de cada línea
    se valida en entregas_flujo (solo su responsable)."""
    accion = request.form.get("accion")
    ids = [int(x) for x in request.form.getlist("linea_id") if x.isdigit()]
    nueva = {"confirmar": "CONFIRMADA", "faltante": "FALTANTE"}.get(accion)
    if not nueva or not ids:
        flash("No se marcó ninguna línea.", "warning")
        return redirect(volver)
    disponible = a_numero(request.form.get("cant_disponible")) if nueva == "FALTANTE" else None
    if nueva == "FALTANTE" and len(ids) != 1:
        flash("El faltante se marca línea por línea, con la cantidad disponible.", "warning")
        return redirect(volver)
    comentario = " ".join((request.form.get("comentario") or "").split())[:200] or None
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cambios = registrar_existencia(cursor, u["id"], roles_actuales(), [(i, nueva, disponible) for i in ids], comentario)
        conn.commit()
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return redirect(volver)
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    mensaje = (f"✅ {len(ids)} línea(s) confirmada(s)." if nueva == "CONFIRMADA" else "Línea marcada como faltante: se avisa al vendedor y a paquetería.")
    if "POR_FACTURAR" in cambios.values():
        mensaje += " La orden quedó completa: el vendedor ya puede facturar."
    flash(mensaje, "success" if nueva == "CONFIRMADA" else "warning")
    return redirect(volver)


# -----------------------------
# Seguimiento de paquetería
# -----------------------------
def resumen_bodegas(lineas):
    """Avance por bodega de un envío: [{bodega, total, confirmadas, faltantes, pendientes, quitadas, traslado_pend, responsable}]."""
    grupos = {}
    for l in lineas:
        g = grupos.setdefault(l["bodega"], {"bodega": l["bodega"], "total": 0, "confirmadas": 0, "faltantes": 0, "pendientes": 0,
                                            "quitadas": 0, "traslado_pend": 0, "traslado_camino": 0, "listas": 0, "retiradas": 0,
                                            "responsable": l["responsable"], "solo_factura": l.get("solo_factura")})
        g["total"] += 1
        g[{"CONFIRMADA": "confirmadas", "FALTANTE": "faltantes", "PENDIENTE": "pendientes", "QUITADA": "quitadas"}[l["existencia"]]] += 1
        if l["traslado_estado"] == "PENDIENTE" and l["existencia"] != "QUITADA":
            g["traslado_pend"] += 1
        elif l["traslado_estado"] in ("APLICADO", "EN_CAMINO") and l["existencia"] != "QUITADA":
            g["traslado_camino"] += 1
        if l["preparacion"] == "LISTA":
            g["listas"] += 1
        elif l["preparacion"] == "RETIRADA":
            g["retiradas"] += 1
    return sorted(grupos.values(), key=lambda g: g["bodega"])


def leer_seguimiento(cursor, f):
    where, params = [f"e.estado IN ({', '.join(['%s'] * len(ABIERTOS))})"], list(ABIERTOS)
    if f["estado"] in ABIERTOS:
        where.append("e.estado = %s"); params.append(f["estado"])
    if f["q"]:
        where.append("(e.pedido_num LIKE %s OR e.cliente LIKE %s)"); params += [f"%{f['q']}%"] * 2
    if f["bodega"]:
        where.append("EXISTS (SELECT 1 FROM ent_envio_lineas x WHERE x.envio_id = e.id AND x.bodega_id = %s)"); params.append(f["bodega"])
    cursor.execute(f"""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista FROM ent_envios e
                       JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                       WHERE {' AND '.join(where)}""", params)
    envios = cursor.fetchall()
    if not envios:
        return []
    ids = [e["id"] for e in envios]
    marcas = ", ".join(["%s"] * len(ids))
    cursor.execute(f"""SELECT l.*, b.nombre AS bodega, b.solo_factura FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id
                       WHERE l.envio_id IN ({marcas}) ORDER BY l.linea""", ids)
    lineas = defaultdict(list)
    for l in cursor.fetchall():
        lineas[l["envio_id"]].append(l)
    # Desde cuándo está en su estado actual (último evento de estado) y última actividad
    cursor.execute(f"""SELECT envio_id, MAX(CASE WHEN campo = 'estado' THEN creado_en END) AS desde_estado, MAX(creado_en) AS ultima
                       FROM ent_envio_eventos WHERE envio_id IN ({marcas}) GROUP BY envio_id""", ids)
    tiempos = {r["envio_id"]: r for r in cursor.fetchall()}
    alertas = alertas_de_envios(cursor, envios, lineas)
    ahora = ahora_nicaragua()
    for e in envios:
        e["alerta"] = alertas[e["id"]]
        e["nivel"] = nivel(e, e["alerta"])
        e["bodegas"] = resumen_bodegas(lineas[e["id"]])
        e["faltantes"] = sum(g["faltantes"] for g in e["bodegas"])
        e["pendientes"] = sum(g["pendientes"] for g in e["bodegas"])
        t = tiempos.get(e["id"], {})
        desde = t.get("desde_estado") or e["creado_en"]
        e["en_estado"] = duracion(max(0, int((ahora - desde).total_seconds() // 60)))
        e["ultima"] = t.get("ultima") or e["creado_en"]
    if f["riesgo"]:
        envios = [e for e in envios if e["nivel"] in ("rojo", "amarillo")]
    envios.sort(key=lambda e: (e["faltantes"] == 0,) + orden_cola(e))
    return envios


@entregas_operacion_bp.route("/entregas/seguimiento")
@requiere_rol("entregas_seguimiento", "paqueteria")
def seguimiento():
    f = {"estado": request.args.get("estado", ""), "q": " ".join((request.args.get("q") or "").split())[:40],
         "bodega": request.args.get("bodega", type=int), "riesgo": request.args.get("riesgo") == "1"}
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        envios = leer_seguimiento(cursor, f)
        cursor.execute("SELECT id, nombre FROM ent_bodegas WHERE activo = 1 ORDER BY nombre")
        bodegas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    resumen = {"abiertas": len(envios), "por_confirmar": sum(1 for e in envios if e["pendientes"]),
               "faltantes": sum(1 for e in envios if e["faltantes"]), "riesgo": sum(1 for e in envios if e["nivel"] == "rojo"),
               "por_facturar": sum(1 for e in envios if e["estado"] == "POR_FACTURAR")}
    return render_template("entregas_seguimiento.html", envios=envios, f=f, bodegas=bodegas, resumen=resumen, nombre_estado=NOMBRE_ESTADO,
                           estados=[s for s in ABIERTOS], ahora=ahora_nicaragua())


# -----------------------------
# Pendientes por rol (contadores de las pestañas)
# -----------------------------
def contar_pendientes(cursor, roles, usuario_id=None):
    """{vista: cantidad} de lo que espera a este usuario. Bodeguero y coordinador: líneas por confirmar; paquetería: órdenes nuevas o
    con faltantes."""
    r = {}
    estados = ", ".join(["%s"] * len(ANTES_DE_FACTURAR))
    if roles.get("bodega"):
        cursor.execute(f"""SELECT COUNT(*) AS n FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
                           WHERE l.responsable = 'bodeguero' AND l.bodega_id = %s AND l.existencia = 'PENDIENTE' AND e.estado IN ({estados})""",
                       [roles["bodega"]] + list(ANTES_DE_FACTURAR))
        n = cursor.fetchone()["n"]
        cursor.execute("""SELECT COUNT(*) AS n FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
                          WHERE l.bodega_prepara_id = %s AND l.existencia = 'CONFIRMADA' AND l.preparacion IN ('PENDIENTE', 'EN_PREPARACION')
                            AND l.traslado_estado IN ('NO_REQUIERE', 'RECIBIDO') AND e.estado IN ('EN_REVISION', 'POR_FACTURAR', 'FACTURADA')""",
                       (roles["bodega"],))
        r["entregas_bodega"] = n + cursor.fetchone()["n"]
    if "coordinador_bodega" in roles:
        cursor.execute(f"""SELECT COUNT(*) AS n FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
                           WHERE l.responsable = 'coordinador' AND l.existencia = 'PENDIENTE' AND e.estado IN ({estados})""",
                       list(ANTES_DE_FACTURAR))
        n = cursor.fetchone()["n"]
        cursor.execute("""SELECT COUNT(*) AS n FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
                          WHERE l.existencia <> 'QUITADA' AND e.estado IN ('ORDEN_ENVIADA', 'EN_REVISION', 'POR_FACTURAR', 'FACTURADA')
                            AND ((l.traslado_estado = 'PENDIENTE' AND l.existencia = 'CONFIRMADA') OR l.traslado_estado IN ('APLICADO', 'EN_CAMINO'))""")
        r["entregas_coordinador"] = n + cursor.fetchone()["n"]
    if "paqueteria" in roles:
        cursor.execute("""SELECT COUNT(*) AS n FROM ent_envios e WHERE e.estado = 'ORDEN_ENVIADA'
                          OR (e.estado IN ('EN_REVISION', 'POR_FACTURAR') AND EXISTS
                              (SELECT 1 FROM ent_envio_lineas l WHERE l.envio_id = e.id AND l.existencia = 'FALTANTE'))""")
        r["entregas_seguimiento"] = cursor.fetchone()["n"]
        cursor.execute("SELECT COUNT(*) AS n FROM ent_envios WHERE estado IN ('LISTA_PARA_RETIRO', 'RETIRADA')")
        r["entregas_empaque"] = cursor.fetchone()["n"]
    if "vendedor" in roles and usuario_id:
        cursor.execute("""SELECT COUNT(*) AS n FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id
                          WHERE e.vendedor_id = %s AND l.existencia = 'FALTANTE' AND (l.decision IS NULL OR l.decision <> 'esperar')
                            AND e.estado IN ('ORDEN_ENVIADA', 'EN_REVISION', 'POR_FACTURAR')""", (usuario_id,))
        n = cursor.fetchone()["n"]
        cursor.execute("SELECT COUNT(*) AS n FROM ent_envios WHERE vendedor_id = %s AND estado = 'POR_FACTURAR'", (usuario_id,))
        r["entregas_mis_envios"] = n + cursor.fetchone()["n"]
    from router.entregas_rutas import contar_rutas
    r.update(contar_rutas(cursor, roles, usuario_id))
    if usuario_id:
        r["avisos"] = contar_no_leidos(cursor, usuario_id)
    return r


def pendientes_actuales():
    roles = roles_actuales()
    if not roles:
        return {}
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        return contar_pendientes(cursor, roles, usuario_actual()["id"])
    finally:
        cursor.close()
        conn.close()


# -----------------------------
# Avisos y recordatorios
# -----------------------------
@entregas_operacion_bp.route("/entregas/avisos")
@requiere_modulo
def avisos():
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""SELECT a.*, e.pedido_num, c.usuario AS de FROM ent_avisos a LEFT JOIN ent_envios e ON e.id = a.envio_id
                          LEFT JOIN usuarios c ON c.id = a.creado_por WHERE a.usuario_id = %s
                          AND (a.leido_en IS NULL OR a.creado_en >= NOW() - INTERVAL 7 DAY) ORDER BY a.leido_en IS NOT NULL, a.creado_en DESC
                          LIMIT 200""", (u["id"],))
        lista = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    return render_template("entregas_avisos.html", avisos=lista, activo="")


@entregas_operacion_bp.route("/entregas/avisos/leer", methods=["POST"])
@requiere_modulo
def leer_avisos():
    """Marca leídos: uno (id, y abre su envío) o todos."""
    u = usuario_actual()
    id_ = request.form.get("id", type=int)
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if id_:
            cursor.execute("UPDATE ent_avisos SET leido_en = NOW() WHERE id = %s AND usuario_id = %s AND leido_en IS NULL", (id_, u["id"]))
            cursor.execute("SELECT envio_id FROM ent_avisos WHERE id = %s AND usuario_id = %s", (id_, u["id"]))
            fila = cursor.fetchone()
        else:
            cursor.execute("UPDATE ent_avisos SET leido_en = NOW() WHERE usuario_id = %s AND leido_en IS NULL", (u["id"],))
            fila = None
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    if fila and fila["envio_id"] and request.form.get("abrir"):
        return redirect(url_for("entregas_orden.envio", id_=fila["envio_id"]))
    return redirect(url_for("entregas_operacion.avisos"))


@entregas_operacion_bp.route("/entregas/recordar/<int:envio_id>", methods=["POST"])
@requiere_rol("entregas_seguimiento", "paqueteria")
def recordar(envio_id):
    """Recordatorio de paquetería a quien tiene pendiente el paso actual: responsables de las existencias por confirmar y, si hay
    faltantes sin decidir o la orden está lista para facturar, el vendedor."""
    u = usuario_actual()
    nota = " ".join((request.form.get("nota") or "").split())[:120]
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM ent_envios WHERE id = %s", (envio_id,))
        envio = cursor.fetchone()
        cursor.execute("SELECT * FROM ent_envio_lineas WHERE envio_id = %s", (envio_id,))
        lineas = cursor.fetchall()
        a_quien, que = [], []
        if envio and envio["estado"] in ANTES_DE_FACTURAR:
            for l in lineas:
                if l["existencia"] == "PENDIENTE":
                    a_quien += responsables_linea(cursor, l)
            if any(l["existencia"] == "PENDIENTE" for l in lineas):
                que.append("confirmar existencias")
            if any(l["existencia"] == "FALTANTE" and l["decision"] != "esperar" for l in lineas):
                a_quien.append(envio["vendedor_id"]); que.append("decidir faltantes")
            if envio["estado"] == "POR_FACTURAR":
                a_quien.append(envio["vendedor_id"]); que.append("facturar")
        if envio and envio["estado"] in ANTES_DE_FACTURAR + ("FACTURADA",):
            mover = [l for l in lineas if l["existencia"] != "QUITADA" and
                     ((l["traslado_estado"] == "PENDIENTE" and l["existencia"] == "CONFIRMADA") or l["traslado_estado"] in ("APLICADO", "EN_CAMINO"))]
            preparar = [l for l in lineas if l["existencia"] == "CONFIRMADA" and l["preparacion"] in ("PENDIENTE", "EN_PREPARACION")
                        and l["traslado_estado"] in ("NO_REQUIERE", "RECIBIDO")]
            if preparar and envio["estado"] == "FACTURADA":
                for l in preparar:
                    a_quien += responsables_linea(cursor, {"responsable": "bodeguero", "bodega_id": l["bodega_prepara_id"]})
                que.append("preparar")
            if mover:
                a_quien += [x for x in responsables_linea(cursor, {"responsable": "coordinador", "bodega_id": None})]
                que.append("aplicar o traer traslados")
        enviados = avisar(cursor, a_quien, "recordatorio",
                          f"Recordatorio de paquetería, pedido {envio['pedido_num']}: {', '.join(que)}" + (f". {nota}" if nota else ""),
                          envio_id=envio_id, creado_por=u["id"]) if a_quien else 0
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash(f"Recordatorio enviado a {enviados} persona(s)." if enviados else "No hay a quién recordarle: nada pendiente que dependa de otra persona.",
          "success" if enviados else "info")
    return redirect(url_for("entregas_operacion.seguimiento"))


@entregas_operacion_bp.route("/entregas/pendientes")
def pendientes():
    if not roles_actuales():
        return jsonify({})
    return jsonify(pendientes_actuales())
