"""Entregas · partes 9 y 10: retiro y empaque (rol Paquetería).

Retiro y empaque (/entregas/empaque):
- Listas para retiro: órdenes LISTA_PARA_RETIRO (facturadas, todo preparado y lo trasladado recibido), agrupadas por la bodega donde
  se recoge (bodega que prepara). «Retiré de esta bodega» marca RETIRADA todas sus líneas listas a la vez; con todas las bodegas
  retiradas, la orden pasa sola a RETIRADA.
- En preparación: órdenes facturadas que todavía no están listas (avance por bodega); el retiro no se habilita.
- Retiradas: por empacar. Paquetería registra bultos y peso: la orden pasa a EMPACADA.
- Empacadas: listas para armar la ruta (parte 11). Se puede corregir bultos y peso y reimprimir la etiqueta.

Etiqueta (/entregas/envio/<id>/etiqueta): número de pedido, números de factura, cliente, destino, transporte y bultos; «COBRAR»
destacado si el vendedor indicó pago en la entrega ([PENDIENTE] 1: mientras no se confirme de dónde sale la forma de pago, la
indica el vendedor al subir la orden). Se imprime una etiqueta por bulto («bulto 1 de N»), o una sola con ?una=1. Tamaño 100 x 150 mm.
"""
from collections import defaultdict

from flask import Blueprint, abort, flash, redirect, render_template, request, url_for

from router.entregas import contexto_entregas, requiere_modulo, requiere_rol, roles_actuales, usuario_actual
from router.entregas_alerta import ahora_nicaragua, alertas_de_envios
from router.entregas_catalogos import TIPOS_TRANSPORTISTA
from router.entregas_flujo import ErrorFlujo, cantidad_que_va, empacar, retirar_bodega
from router.entregas_orden import NOMBRE_ESTADO
from utils.db import get_connection

entregas_paqueteria_bp = Blueprint("entregas_paqueteria", __name__)
entregas_paqueteria_bp.context_processor(contexto_entregas)

ORDEN_RIESGO = {"rojo": 0, "amarillo": 1, "verde": 2, None: 3}


def leer_retiro(cursor):
    estados = ("FACTURADA", "LISTA_PARA_RETIRO", "RETIRADA", "EMPACADA")
    cursor.execute(f"""SELECT e.*, u.usuario AS vendedor, t.nombre AS transportista FROM ent_envios e
                       JOIN usuarios u ON u.id = e.vendedor_id JOIN ent_transportistas t ON t.id = e.transportista_id
                       WHERE e.estado IN ({', '.join(['%s'] * len(estados))})""", estados)
    envios = cursor.fetchall()
    if not envios:
        return {"listas": [], "preparando": [], "retiradas": [], "empacadas": []}
    ids = [e["id"] for e in envios]
    marcas = ", ".join(["%s"] * len(ids))
    cursor.execute(f"""SELECT l.*, b.nombre AS bodega, b.solo_factura, p.nombre AS bodega_prepara FROM ent_envio_lineas l
                       JOIN ent_bodegas b ON b.id = l.bodega_id LEFT JOIN ent_bodegas p ON p.id = l.bodega_prepara_id
                       WHERE l.envio_id IN ({marcas}) ORDER BY l.linea""", ids)
    lineas = defaultdict(list)
    for l in cursor.fetchall():
        l["va"] = cantidad_que_va(l)
        lineas[l["envio_id"]].append(l)
    cursor.execute(f"""SELECT f.envio_id, f.factura_num, b.nombre AS bodega FROM ent_envio_facturas f JOIN ent_bodegas b ON b.id = f.bodega_id
                       WHERE f.envio_id IN ({marcas}) ORDER BY b.nombre""", ids)
    facturas = defaultdict(list)
    for f in cursor.fetchall():
        facturas[f["envio_id"]].append(f)
    alertas = alertas_de_envios(cursor, envios, lineas)
    for e in envios:
        e["facturas"] = facturas[e["id"]]
        e["alerta"] = alertas[e["id"]]
        e["nivel"] = "rojo" if e["riesgo_confirmado"] else (e["alerta"]["nivel"] if e["alerta"] else None)
        activas = [l for l in lineas[e["id"]] if l["existencia"] != "QUITADA"]
        grupos = {}
        for l in activas:
            g = grupos.setdefault(l["bodega_prepara_id"], {"bodega_id": l["bodega_prepara_id"], "bodega": l["bodega_prepara"] or "destino por definir",
                                                           "lineas": [], "listas": 0, "retiradas": 0, "camino": 0,
                                                           "solo_factura": bool(l["solo_factura"])})
            g["lineas"].append(l)
            g["listas"] += l["preparacion"] == "LISTA"
            g["retiradas"] += l["preparacion"] == "RETIRADA"
            g["camino"] += l["traslado_estado"] in ("PENDIENTE", "APLICADO", "EN_CAMINO")
        e["bodegas"] = sorted(grupos.values(), key=lambda g: g["bodega"])
    clave = lambda e: (ORDEN_RIESGO.get(e["nivel"], 3), e["prioridad"] != "urgente", e["hora_requerida"])  # noqa: E731
    return {"listas": sorted([e for e in envios if e["estado"] == "LISTA_PARA_RETIRO"], key=clave),
            "preparando": sorted([e for e in envios if e["estado"] == "FACTURADA"], key=clave),
            "retiradas": sorted([e for e in envios if e["estado"] == "RETIRADA"], key=clave),
            "empacadas": sorted([e for e in envios if e["estado"] == "EMPACADA"], key=clave)}


@entregas_paqueteria_bp.route("/entregas/empaque")
@requiere_rol("entregas_empaque", "paqueteria")
def empaque():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        d = leer_retiro(cursor)
    finally:
        cursor.close()
        conn.close()
    return render_template("entregas_empaque.html", d=d, nombre_estado=NOMBRE_ESTADO, ahora=ahora_nicaragua(), activo="entregas_empaque")


@entregas_paqueteria_bp.route("/entregas/empaque/retirar", methods=["POST"])
@requiere_rol("entregas_empaque", "paqueteria")
def retirar():
    envio_id, bodega_id = request.form.get("envio_id", type=int), request.form.get("bodega_id", type=int)
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        n, nuevo = retirar_bodega(cursor, u["id"], roles_actuales(), envio_id, bodega_id)
        conn.commit()
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return redirect(url_for("entregas_paqueteria.empaque"))
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    flash(f"✅ {n} línea(s) retirada(s)." + (" La orden quedó RETIRADA: pasa a empaque." if nuevo == "RETIRADA" else ""), "success")
    return redirect(url_for("entregas_paqueteria.empaque"))


@entregas_paqueteria_bp.route("/entregas/empaque/empacar", methods=["POST"])
@requiere_rol("entregas_empaque", "paqueteria")
def empacar_envio():
    envio_id = request.form.get("envio_id", type=int)
    try:
        bultos = int(request.form.get("bultos", ""))
    except ValueError:
        bultos = None
    try:
        peso = float((request.form.get("peso_kg") or "").replace(",", "."))
    except ValueError:
        peso = None
    u = usuario_actual()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cambio = empacar(cursor, u["id"], roles_actuales(), envio_id, bultos, peso)
        conn.commit()
    except ErrorFlujo as e:
        conn.rollback()
        flash(str(e), "warning")
        return redirect(url_for("entregas_paqueteria.empaque"))
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    flash("✅ Empacada. Imprima la etiqueta." if cambio else "✅ Bultos y peso corregidos. Reimprima la etiqueta.", "success")
    return redirect(url_for("entregas_paqueteria.empaque", imprimir=envio_id))


@entregas_paqueteria_bp.route("/entregas/envio/<int:id_>/etiqueta")
@requiere_modulo
def etiqueta(id_):
    roles = roles_actuales()
    if not ({"paqueteria", "admin"} & set(roles)):
        abort(403)
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("""SELECT e.*, t.nombre AS transportista, t.tipo AS transportista_tipo, t.direccion AS transportista_direccion,
                                 u.usuario AS vendedor
                          FROM ent_envios e JOIN ent_transportistas t ON t.id = e.transportista_id JOIN usuarios u ON u.id = e.vendedor_id
                          WHERE e.id = %s""", (id_,))
        e = cursor.fetchone()
        cursor.execute("""SELECT f.factura_num, b.nombre AS bodega FROM ent_envio_facturas f JOIN ent_bodegas b ON b.id = f.bodega_id
                          WHERE f.envio_id = %s ORDER BY b.nombre""", (id_,))
        facturas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    if not e:
        abort(404)
    if e["estado"] not in ("EMPACADA", "EN_RUTA", "ENTREGADA") or not e["bultos"]:
        flash("La etiqueta se imprime después de registrar bultos y peso.", "warning")
        return redirect(url_for("entregas_paqueteria.empaque"))
    copias = 1 if request.args.get("una") else e["bultos"]
    return render_template("entregas_etiqueta.html", e=e, facturas=facturas, copias=copias, tipos=TIPOS_TRANSPORTISTA)
