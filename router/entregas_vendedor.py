"""Entregas · parte 7: pantallas del vendedor.

- Mis envíos (/entregas/mis_envios): sus órdenes con estado y alerta; faltantes por decidir (quitar, va lo disponible, cambiar de
  bodega o esperar); listas para facturar; anular con motivo.
- Facturar (/entregas/envio/<id>/facturar): qué ajustar en la orden del sistema de ventas (líneas quitadas, cantidades parciales,
  cambios de bodega) y una factura por bodega de facturación. Solo con la orden POR_FACTURAR (regla 4.5).

Todos los cambios pasan por router/entregas_flujo.py.
"""
from collections import defaultdict

from flask import Blueprint, flash, redirect, render_template, request, url_for

from router.entregas import contexto_entregas, requiere_modulo, requiere_rol, roles_actuales, usuario_actual
from router.entregas_alerta import alertas_de_envios
from router.entregas_flujo import (ABIERTOS, ANTES_DE_FACTURAR, ErrorFlujo, anular, bodega_de_factura, bodegas_de_facturacion,
                                   decidir_faltante, modo_factura_traslado, registrar_facturas)
from router.entregas_orden import NOMBRE_ESTADO
from utils.db import get_connection

entregas_vendedor_bp = Blueprint("entregas_vendedor", __name__)
entregas_vendedor_bp.context_processor(contexto_entregas)

DECISIONES = {"quitar": "Quitar", "parcial": "Va lo disponible", "cambiar_bodega": "Cambiar de bodega", "esperar": "Esperar"}


def ejecutar(funcion):
    """Corre un cambio del flujo en una transacción: ErrorFlujo -> rollback y mensaje."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        resultado = funcion(cursor)
        conn.commit()
        return resultado, None
    except ErrorFlujo as e:
        conn.rollback()
        return None, str(e)
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()


def leer_mis_envios(cursor, vendedor_id):
    cursor.execute(f"""SELECT e.*, t.nombre AS transportista FROM ent_envios e JOIN ent_transportistas t ON t.id = e.transportista_id
                       WHERE e.vendedor_id = %s AND (e.estado IN ({', '.join(['%s'] * len(ABIERTOS))})
                             OR e.actualizado_en >= NOW() - INTERVAL 7 DAY)
                       ORDER BY e.hora_requerida""", [vendedor_id] + list(ABIERTOS))
    envios = cursor.fetchall()
    if not envios:
        return []
    ids = [e["id"] for e in envios]
    marcas = ", ".join(["%s"] * len(ids))
    cursor.execute(f"""SELECT l.*, b.nombre AS bodega FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id
                       WHERE l.envio_id IN ({marcas}) ORDER BY l.linea""", ids)
    lineas = defaultdict(list)
    for l in cursor.fetchall():
        lineas[l["envio_id"]].append(l)
    cursor.execute(f"SELECT envio_id, COUNT(*) AS n FROM ent_envio_facturas WHERE envio_id IN ({marcas}) GROUP BY envio_id", ids)
    facturas = {r["envio_id"]: r["n"] for r in cursor.fetchall()}
    alertas = alertas_de_envios(cursor, envios, lineas)
    modo = modo_factura_traslado(cursor)
    for e in envios:
        e["lineas"] = lineas[e["id"]]
        e["alerta"] = alertas[e["id"]]
        e["nivel"] = "rojo" if e["riesgo_confirmado"] else (e["alerta"]["nivel"] if e["alerta"] else None)
        e["faltantes"] = [l for l in e["lineas"] if l["existencia"] == "FALTANTE"] if e["estado"] in ANTES_DE_FACTURAR else []
        e["confirmadas"] = sum(1 for l in e["lineas"] if l["existencia"] in ("CONFIRMADA", "QUITADA"))
        e["facturas"] = facturas.get(e["id"], 0)
        e["por_facturar"] = len(bodegas_de_facturacion(e["lineas"], modo)) if e["estado"] == "POR_FACTURAR" else 0
    return envios


@entregas_vendedor_bp.route("/entregas/mis_envios")
@requiere_rol("entregas_mis_envios", "vendedor")
def mis_envios():
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        envios = leer_mis_envios(cursor, u["id"])
        cursor.execute("SELECT id, nombre, requiere_traslado FROM ent_bodegas WHERE activo = 1 AND solo_factura = 0 ORDER BY nombre")
        bodegas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    grupos = {"faltantes": [e for e in envios if e["faltantes"]],
              "facturar": [e for e in envios if e["estado"] == "POR_FACTURAR"],
              "proceso": [e for e in envios if e["estado"] in ABIERTOS and not e["faltantes"] and e["estado"] != "POR_FACTURAR"],
              "cerrados": [e for e in envios if e["estado"] not in ABIERTOS]}
    return render_template("entregas_mis_envios.html", g=grupos, bodegas=bodegas, nombre_estado=NOMBRE_ESTADO, decisiones=DECISIONES)


@entregas_vendedor_bp.route("/entregas/faltante/<int:linea_id>", methods=["POST"])
@requiere_rol("entregas_mis_envios", "vendedor")
def decidir(linea_id):
    decision = request.form.get("decision")
    bodega = request.form.get("bodega_id", type=int)
    u, roles = usuario_actual(), roles_actuales()
    hecho, error = ejecutar(lambda c: decidir_faltante(c, u, roles, linea_id, decision, bodega))
    flash(error or f"✅ {hecho}", "warning" if error else "success")
    return redirect(url_for("entregas_vendedor.mis_envios"))


@entregas_vendedor_bp.route("/entregas/envio/<int:id_>/anular", methods=["POST"])
@requiere_modulo      # el permiso (vendedor dueño o administrador) lo valida entregas_flujo.anular
def anular_envio(id_):
    u, roles = usuario_actual(), roles_actuales()
    _, error = ejecutar(lambda c: anular(c, u, roles, id_, request.form.get("motivo")))
    flash(error or "Orden anulada. Se avisó a paquetería y a las bodegas.", "warning" if error else "info")
    return redirect(request.referrer or url_for("entregas_vendedor.mis_envios"))


# -----------------------------
# Facturar
# -----------------------------
def datos_factura(cursor, envio_id):
    cursor.execute("""SELECT e.*, t.nombre AS transportista FROM ent_envios e JOIN ent_transportistas t ON t.id = e.transportista_id
                      WHERE e.id = %s""", (envio_id,))
    envio = cursor.fetchone()
    if not envio:
        return None
    cursor.execute("""SELECT l.*, b.nombre AS bodega, b.solo_factura, p.nombre AS bodega_prepara FROM ent_envio_lineas l JOIN ent_bodegas b ON b.id = l.bodega_id
                      LEFT JOIN ent_bodegas p ON p.id = l.bodega_prepara_id WHERE l.envio_id = %s ORDER BY l.linea""", (envio_id,))
    lineas = cursor.fetchall()
    cursor.execute("""SELECT f.*, b.nombre AS bodega, u.usuario FROM ent_envio_facturas f JOIN ent_bodegas b ON b.id = f.bodega_id
                      JOIN usuarios u ON u.id = f.registrada_por WHERE f.envio_id = %s""", (envio_id,))
    facturas = {f["bodega_id"]: f for f in cursor.fetchall()}
    cursor.execute("""SELECT linea_id, comentario FROM ent_envio_eventos WHERE envio_id = %s AND campo = 'existencia'
                      AND comentario LIKE 'Cambio de bodega:%%' ORDER BY id""", (envio_id,))
    cambios = {r["linea_id"]: r["comentario"] for r in cursor.fetchall()}
    cursor.execute("SELECT comentario FROM ent_envio_eventos WHERE envio_id = %s AND a_estado = 'NO_SE_ENVIAN' ORDER BY id", (envio_id,))
    no_van = [r["comentario"] for r in cursor.fetchall()]
    modo = modo_factura_traslado(cursor)
    cursor.execute("SELECT id, nombre FROM ent_bodegas")
    nombres = {b["id"]: b["nombre"] for b in cursor.fetchall()}
    # Qué ajustar en la orden del sistema de ventas para que quede igual a lo que sale
    ajustes = []
    for l in lineas:
        if l["existencia"] == "QUITADA":
            ajustes.append({"linea": l, "texto": f"Quitar la línea {l['linea']} ({l['codigo']}, {float(l['cantidad']):g})"})
        elif l["existencia"] == "CONFIRMADA" and l["cant_disponible"] is not None:
            ajustes.append({"linea": l, "texto": f"Línea {l['linea']} ({l['codigo']}): cambiar la cantidad de {float(l['cantidad']):g} a "
                                                 f"{float(l['cant_disponible']):g}"})
        if l["id"] in cambios and l["existencia"] != "QUITADA":
            ajustes.append({"linea": l, "texto": f"Línea {l['linea']} ({l['codigo']}): {cambios[l['id']].replace('Cambio de bodega: ', 'cambiar la bodega de ')}"})
    for texto in no_van:
        ajustes.append({"linea": None, "texto": texto.replace("No se envían (bodega fuera del catálogo): ",
                                                              "No van en este envío (bodega fuera del catálogo): ")})
    # Una factura por bodega de facturación
    grupos = {}
    for l in lineas:
        if l["existencia"] == "QUITADA":
            continue
        b = bodega_de_factura(l, modo)
        g = grupos.setdefault(b, {"bodega_id": b, "bodega": nombres.get(b, "destino por definir"), "lineas": [], "factura": facturas.get(b)})
        g["lineas"].append(l)
    sobrantes = [f for b, f in facturas.items() if b not in grupos]
    return {"envio": envio, "lineas": lineas, "ajustes": ajustes, "grupos": sorted(grupos.values(), key=lambda g: g["bodega"]),
            "sobrantes": sobrantes, "modo": modo, "requeridas": bodegas_de_facturacion(lineas, modo)}


@entregas_vendedor_bp.route("/entregas/envio/<int:id_>/facturar", methods=["GET", "POST"])
@requiere_rol("entregas_mis_envios", "vendedor")
def facturar(id_):
    u, roles = usuario_actual(), roles_actuales()
    if request.method == "POST":
        facturas = {int(k[8:]): v for k, v in request.form.items() if k.startswith("factura_") and k[8:].isdigit()}
        resultado, error = ejecutar(lambda c: registrar_facturas(c, u, roles, id_, facturas))
        if error:
            flash(error, "warning")
        else:
            guardadas, faltan = resultado
            flash(("✅ Facturas registradas: la orden quedó FACTURADA." if not faltan else
                   f"Se guardaron {guardadas} factura(s); faltan {faltan} bodega(s)."), "success" if not faltan else "info")
            if not faltan:
                return redirect(url_for("entregas_orden.envio", id_=id_))
        return redirect(url_for("entregas_vendedor.facturar", id_=id_))
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        d = datos_factura(cursor, id_)
    finally:
        cursor.close()
        conn.close()
    if not d or d["envio"]["vendedor_id"] != u["id"]:
        flash("El envío no existe o no es suyo.", "warning")
        return redirect(url_for("entregas_vendedor.mis_envios"))
    return render_template("entregas_facturar.html", d=d, e=d["envio"], nombre_estado=NOMBRE_ESTADO)
