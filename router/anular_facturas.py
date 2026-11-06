"""Anular facturas ya cargadas (se anularon en el sistema después de subir las ventas).

Una factura se identifica por número + bodega + fecha (el número se repite entre bodegas). Al anular:
  - ventas_historicas: se borra la fila (así sale de todos los reportes de ventas),
  - ventas_por_producto: sus líneas quedan con Estatus 'Anulado' (los reportes ya las excluyen),
  - ventas_armados: se borran sus líneas (número + fecha).
Todo lo que había se guarda en facturas_anuladas.respaldo para poder restaurarla, y mientras esté anulada las cargas no la
vuelven a meter (ver anuladas_activas)."""
import json
from datetime import date, datetime
from decimal import Decimal

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from utils.db import get_connection
from router.accesos import requiere_acceso

anular_facturas_bp = Blueprint("anular_facturas", __name__)


def _num(factura):
    """`No. Factura` de ventas_por_producto es numérico."""
    try:
        return float(factura)
    except (TypeError, ValueError):
        return None


def _json(v):
    if isinstance(v, Decimal):
        return str(v)
    if isinstance(v, (date, datetime)):
        return v.isoformat()
    return v


def clave_txt(factura, bodega, fecha):
    return f"{factura}|{bodega}|{fecha.isoformat() if isinstance(fecha, date) else fecha}"


def leer_clave(texto):
    try:
        factura, bodega, fecha = (texto or "").split("|", 2)
        return factura.strip(), bodega, date.fromisoformat(fecha)
    except ValueError:
        return None


def anuladas_activas(cursor=None):
    """Facturas anuladas a mano y no restauradas: {"identificadores": set, "producto": {(num, bodega, 'aaaa-mm-dd')}, "armados": {(factura, fecha)}}.
    Las cargas las usan para no volver a meterlas."""
    propio = cursor is None
    if propio:
        conn = get_connection()
        cursor = conn.cursor()
    try:
        cursor.execute("SELECT factura, bodega, fecha, identificador FROM facturas_anuladas WHERE restaurado IS NULL")
        filas = cursor.fetchall()
    finally:
        if propio:
            cursor.close()
            conn.close()
    filas = [tuple(f.values()) if isinstance(f, dict) else f for f in filas]
    return {
        "identificadores": {i for _, _, _, i in filas if i},
        "producto": {(int(_num(f)), b, d.isoformat()) for f, b, d, _ in filas if _num(f) is not None},
        "armados": {(f, d) for f, _, d, _ in filas},
    }


# -----------------------------
# Lectura de una factura
# -----------------------------
def buscar(cursor, factura):
    """Todas las facturas con ese número (cada bodega/fecha por separado), estén en ventas o solo en ventas por producto."""
    encontradas = {}
    cursor.execute("SELECT * FROM ventas_historicas WHERE factura = %s", (factura,))
    for v in cursor.fetchall():
        k = (v["factura"], v["Bodega"], v["Fecha"])
        encontradas.setdefault(k, {"ventas": [], "lineas": 0, "lineas_vigentes": 0})["ventas"].append(v)
    n = _num(factura)
    if n is not None:
        cursor.execute("SELECT Bodega, Fecha, COUNT(*) AS lineas, SUM(Estatus <> 'Anulado') AS vigentes, SUM(`Sub-Total 2`) AS venta, "
                       "MAX(Cliente) AS cliente, MAX(Vendedor) AS vendedor, MAX(`Contado/Credito`) AS tipo "
                       "FROM ventas_por_producto WHERE `No. Factura` = %s GROUP BY Bodega, Fecha", (n,))
        for p in cursor.fetchall():
            k = (factura, p["Bodega"], date.fromisoformat(str(p["Fecha"])[:10]))
            e = encontradas.setdefault(k, {"ventas": [], "lineas": 0, "lineas_vigentes": 0})
            e.update(lineas=int(p["lineas"]), lineas_vigentes=int(p["vigentes"] or 0), venta_productos=float(p["venta"] or 0),
                     cliente_p=p["cliente"], vendedor_p=p["vendedor"], tipo_p=p["tipo"])
    cursor.execute("SELECT fecha, anulado, anulado_por FROM facturas_anuladas WHERE factura = %s AND restaurado IS NULL", (factura,))
    ya = {f["fecha"]: f for f in cursor.fetchall()}
    salida = []
    for (fac, bodega, fecha), e in sorted(encontradas.items(), key=lambda kv: kv[0][2], reverse=True):
        v = e["ventas"][0] if e["ventas"] else None
        salida.append({
            "clave": clave_txt(fac, bodega, fecha), "factura": fac, "bodega": bodega, "fecha": fecha,
            "cliente": v["cliente"] if v else e.get("cliente_p"), "vendedor": v["vendedor"] if v else e.get("vendedor_p"),
            "tipo": v["tipo_venta"] if v else (e.get("tipo_p") or "").upper(), "total": float(v["Total"]) if v else None,
            "estado": v["estado_factura"] if v else None, "en_ventas": bool(v), "lineas": e["lineas"], "lineas_vigentes": e["lineas_vigentes"],
            "anulable": bool(v) or e["lineas_vigentes"] > 0, "ya_anulada": ya.get(fecha),
        })
    return salida


def detalle(cursor, factura, bodega, fecha):
    """Todo lo que se va a tocar al anular, con avisos (abonos, devoluciones)."""
    cursor.execute("SELECT * FROM ventas_historicas WHERE factura = %s AND Bodega = %s AND Fecha = %s", (factura, bodega, fecha))
    ventas = cursor.fetchall()
    productos = []
    n = _num(factura)
    if n is not None:
        cursor.execute("SELECT * FROM ventas_por_producto WHERE `No. Factura` = %s AND Bodega = %s AND LEFT(Fecha, 10) = %s", (n, bodega, fecha.isoformat()))
        productos = cursor.fetchall()
    cursor.execute("SELECT * FROM ventas_armados WHERE nfactura = %s AND fecha = %s ORDER BY id", (factura, fecha))
    armados = cursor.fetchall()

    avisos = []
    cursor.execute("SELECT nrecibo, fechapago, mpagado FROM recuperacion WHERE nfactura = %s AND fechafact = %s", (factura, fecha))
    abonos = cursor.fetchall()
    if abonos:
        avisos.append(f"Tiene {len(abonos)} abono(s) registrados en Recuperación (recibos {', '.join(str(a['nrecibo']) for a in abonos[:8])}, "
                      f"C$ {sum(float(a['mpagado'] or 0) for a in abonos):,.2f}). Esos abonos no se borran.")
    cursor.execute("SELECT numerodv, fechadv, totaldv FROM devoluciones WHERE factura = %s AND fechadv >= %s", (factura, fecha))
    devs = cursor.fetchall()
    if devs:
        avisos.append(f"Hay {len(devs)} devolución(es) con este número de factura (C$ {sum(float(d['totaldv'] or 0) for d in devs):,.2f}). "
                      "No se borran: revise si son de esta factura.")
    if armados and not ventas and not productos:
        avisos.append("Solo está en Armado de mangueras.")
    cursor.execute("SELECT id, anulado, anulado_por FROM facturas_anuladas WHERE factura = %s AND bodega = %s AND fecha = %s AND restaurado IS NULL",
                   (factura, bodega, fecha))
    ya = cursor.fetchone()
    return {"ventas": ventas, "productos": productos, "armados": armados, "avisos": avisos, "ya_anulada": ya,
            "vigentes": sum(1 for p in productos if p["Estatus"] != "Anulado")}


# -----------------------------
# Vistas
# -----------------------------
@anular_facturas_bp.route("/anular_facturas")
@requiere_acceso("anular_facturas")
def inicio():
    factura = "".join((request.args.get("factura") or "").split())
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        resultados = buscar(cursor, factura) if factura else None
        cursor.execute("SELECT id, factura, bodega, fecha, tipo_venta, cliente, vendedor, total, motivo, anulado_por, anulado, restaurado_por, restaurado "
                       "FROM facturas_anuladas ORDER BY anulado DESC LIMIT 60")
        historial = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    if resultados and len(resultados) == 1 and resultados[0]["anulable"] and not resultados[0]["ya_anulada"]:
        return redirect(url_for("anular_facturas.revisar", clave=resultados[0]["clave"]))
    return render_template("anular_facturas.html", factura=factura, resultados=resultados, historial=historial)


@anular_facturas_bp.route("/anular_facturas/revisar")
@requiere_acceso("anular_facturas")
def revisar():
    clave = leer_clave(request.args.get("clave"))
    if not clave:
        flash("⚠️ Factura no válida.", "warning")
        return redirect(url_for("anular_facturas.inicio"))
    factura, bodega, fecha = clave
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        d = detalle(cursor, factura, bodega, fecha)
    finally:
        cursor.close()
        conn.close()
    if not d["ventas"] and not d["vigentes"] and not d["armados"]:
        flash(f"⚠️ La factura {factura} ({bodega}, {fecha:%d/%m/%Y}) no tiene nada vigente que anular.", "warning")
        return redirect(url_for("anular_facturas.inicio", factura=factura))
    return render_template("anular_facturas_revisar.html", factura=factura, bodega=bodega, fecha=fecha, clave=clave_txt(factura, bodega, fecha),
                           d=d, v=d["ventas"][0] if d["ventas"] else None)


@anular_facturas_bp.route("/anular_facturas/anular", methods=["POST"])
@requiere_acceso("anular_facturas")
def anular():
    clave = leer_clave(request.form.get("clave"))
    motivo = " ".join((request.form.get("motivo") or "").split())[:255]
    if not clave:
        flash("⚠️ Factura no válida.", "warning")
        return redirect(url_for("anular_facturas.inicio"))
    factura, bodega, fecha = clave
    if not motivo or request.form.get("confirmo") != "si":
        flash("⚠️ Escriba el motivo y marque la casilla de confirmación.", "warning")
        return redirect(url_for("anular_facturas.revisar", clave=clave_txt(*clave)))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        conn.start_transaction()
        d = detalle(cursor, factura, bodega, fecha)   # se vuelve a leer dentro de la transacción
        if d["ya_anulada"]:
            raise ValueError(f"ya fue anulada el {d['ya_anulada']['anulado']:%d/%m/%Y %H:%M} por {d['ya_anulada']['anulado_por']}")
        if not d["ventas"] and not d["vigentes"] and not d["armados"]:
            raise ValueError("no tiene nada vigente que anular")
        v = d["ventas"][0] if d["ventas"] else None
        p = d["productos"][0] if d["productos"] else {}
        respaldo = {t: [{k: _json(x) for k, x in fila.items()} for fila in d[n]]
                    for t, n in (("ventas_historicas", "ventas"), ("ventas_por_producto", "productos"), ("ventas_armados", "armados"))}
        cursor.execute("""INSERT INTO facturas_anuladas (factura, bodega, fecha, identificador, tipo_venta, cliente, vendedor, total, motivo, respaldo, anulado_por)
                          VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                       (factura, bodega, fecha, v["identificador"] if v else None,
                        v["tipo_venta"] if v else (p.get("Contado/Credito") or "").upper()[:10] or None,
                        (v["cliente"] if v else p.get("Cliente")) or None, (v["vendedor"] if v else p.get("Vendedor")) or None,
                        v["Total"] if v else None, motivo, json.dumps(respaldo, ensure_ascii=False), session.get("usuario")))
        cursor.execute("DELETE FROM ventas_historicas WHERE factura = %s AND Bodega = %s AND Fecha = %s", (factura, bodega, fecha))
        borradas_v = cursor.rowcount
        marcadas = 0
        if _num(factura) is not None:
            cursor.execute("UPDATE ventas_por_producto SET Estatus = 'Anulado' WHERE `No. Factura` = %s AND Bodega = %s AND LEFT(Fecha, 10) = %s AND Estatus <> 'Anulado'",
                           (_num(factura), bodega, fecha.isoformat()))
            marcadas = cursor.rowcount
        cursor.execute("DELETE FROM ventas_armados WHERE nfactura = %s AND fecha = %s", (factura, fecha))
        borradas_a = cursor.rowcount
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se anuló la factura {factura}: {e}", "danger")
        return redirect(url_for("anular_facturas.revisar", clave=clave_txt(*clave)))
    finally:
        cursor.close()
        conn.close()

    partes = [f"{borradas_v} de ventas" if borradas_v else None, f"{marcadas} líneas de productos" if marcadas else None,
              f"{borradas_a} líneas de armados" if borradas_a else None]
    flash(f"✅ Factura {factura} ({bodega}, {fecha:%d/%m/%Y}) anulada: " + ", ".join(x for x in partes if x) + ". Se puede restaurar desde el historial.",
          "success")
    return redirect(url_for("anular_facturas.inicio"))


@anular_facturas_bp.route("/anular_facturas/restaurar/<int:id_anulada>", methods=["POST"])
@requiere_acceso("anular_facturas")
def restaurar(id_anulada):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        conn.start_transaction()
        cursor.execute("SELECT * FROM facturas_anuladas WHERE id = %s AND restaurado IS NULL FOR UPDATE", (id_anulada,))
        a = cursor.fetchone()
        if not a:
            raise ValueError("no existe o ya fue restaurada")
        r = json.loads(a["respaldo"])

        def insertar(tabla, filas, quitar=()):
            for f in filas:
                cols = [c for c in f if c not in quitar]
                cursor.execute(f"INSERT INTO {tabla} ({', '.join(f'`{c}`' for c in cols)}) VALUES ({', '.join(['%s'] * len(cols))})",
                               [f[c] for c in cols])

        for f in r["ventas_historicas"]:
            cursor.execute("SELECT 1 FROM ventas_historicas WHERE identificador = %s", (f["identificador"],))
            if cursor.fetchone():
                raise ValueError(f"la factura ya está otra vez en ventas ({f['identificador']})")
        insertar("ventas_historicas", r["ventas_historicas"])
        if r["ventas_por_producto"] and _num(a["factura"]) is not None:
            # Se devuelven las líneas tal como estaban (con su estatus original)
            cursor.execute("DELETE FROM ventas_por_producto WHERE `No. Factura` = %s AND Bodega = %s AND LEFT(Fecha, 10) = %s",
                           (_num(a["factura"]), a["bodega"], a["fecha"].isoformat()))
            insertar("ventas_por_producto", r["ventas_por_producto"])
        insertar("ventas_armados", r["ventas_armados"])   # con su id original (los autoincrementales no se reutilizan)
        cursor.execute("UPDATE facturas_anuladas SET restaurado = NOW(), restaurado_por = %s WHERE id = %s", (session.get("usuario"), id_anulada))
        conn.commit()
        flash(f"✅ Factura {a['factura']} ({a['bodega']}, {a['fecha']:%d/%m/%Y}) restaurada.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se restauró la factura: {e}", "danger")
    finally:
        cursor.close()
        conn.close()
    return redirect(url_for("anular_facturas.inicio"))
