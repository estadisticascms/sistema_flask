"""Carga de la recuperación (reporte "Comisiones Crédito" del sistema) a la tabla recuperacion.

El reporte viene agrupado por vendedor: una fila con el nombre del vendedor, sus facturas cobradas
(cada una con su facturador) y una fila "TOTAL ...".

Un mismo recibo (ROC) puede abonar dos veces a la misma factura, así que los pagos se comparan por grupo
factura + recibo: si el archivo trae las mismas líneas que la base de datos no se toca nada; si cambian,
se reemplaza el grupo. Lo que sí es error: que lo pagado a una factura (este archivo + otros recibos ya
cargados) pase del total de la factura; esa factura no se carga."""
import json
import os
import uuid
from datetime import date
from decimal import Decimal

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import iguales
from router.cargar_ventas import ErrorArchivo, a_decimal, a_fecha, a_texto, leer_filas, limpiar_temporales, ruta_temporal

cargar_recuperacion_bp = Blueprint("cargar_recuperacion", __name__)

PREFIJO = "carga_recuperacion_"
MAX_DETALLE = 300
# columna del reporte -> columna de la tabla
COLUMNAS = {"FACTURADOR": "facturador", "# FACTURA/ND": "nfactura", "# ROC": "nrecibo", "CLIENTE": "cliente",
            "FECHA FACTURA": "fechafact", "FECHA PAGO": "fechapago", "DIAS RECUPER.": "drecuperado",
            "SUBTOTAL": "subtotal", "DESCUENTO": "descuento", "IMPUESTO": "impuesto", "TOTAL": "total",
            "MONTO PAGADO": "mpagado", "SUBTOTAL-DESCUENTO": "sub_desc", "REAL ABONADO": "realabono"}
CAMPOS = ["vendedor", "facturador", "nfactura", "nrecibo", "cliente", "fechafact", "fechapago", "drecuperado",
          "subtotal", "descuento", "impuesto", "total", "mpagado", "sub_desc", "realabono"]
MONTOS = {"subtotal", "descuento", "impuesto", "total", "mpagado", "sub_desc", "realabono"}
NUMERICAS = MONTOS | {"drecuperado"}
LARGOS = {"vendedor": 50, "facturador": 45, "nfactura": 15, "nrecibo": 15, "cliente": 45}
NOMBRES = {"vendedor": "Vendedor", "facturador": "Facturador", "nfactura": "Factura", "nrecibo": "Recibo (ROC)",
           "cliente": "Cliente", "fechafact": "Fecha factura", "fechapago": "Fecha pago", "drecuperado": "Días",
           "subtotal": "Subtotal", "descuento": "Descuento", "impuesto": "Impuesto", "total": "Total",
           "mpagado": "Monto pagado", "sub_desc": "Subtotal − desc.", "realabono": "Real abonado", "clave": "Factura / recibo",
           "lineas": "Líneas del recibo en la factura"}


def leer_recuperacion(archivo):
    filas = leer_filas(archivo)
    inicio, pos = None, None
    for i, fila in enumerate(filas[:40]):
        nombres = {a_texto(v).upper(): j for j, v in enumerate(fila)}
        if all(c in nombres for c in COLUMNAS):
            inicio, pos = i, {c: nombres[c] for c in COLUMNAS}
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas del reporte Comisiones Crédito "
                           "(FACTURADOR, # FACTURA/ND, # ROC, FECHA PAGO, MONTO PAGADO, REAL ABONADO, ...)")

    def celda(fila, col):
        j = pos[col]
        return fila[j] if j < len(fila) else None

    registros, errores, advertencias, vistos = [], [], [], {}
    vendedor, recortados = None, 0
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        primera = a_texto(celda(fila, "FACTURADOR"))
        factura = a_texto(celda(fila, "# FACTURA/ND"))
        if not factura:
            if primera.upper().startswith("TOTAL"):
                continue
            if primera:           # fila de grupo: nombre del vendedor
                vendedor = " ".join(primera.split())
            continue
        r = {"vendedor": vendedor, "facturador": primera or None, "nfactura": factura.lstrip("#").strip(),
             "nrecibo": a_texto(celda(fila, "# ROC")) or None, "cliente": a_texto(celda(fila, "CLIENTE")) or None,
             "fechafact": a_fecha(celda(fila, "FECHA FACTURA")), "fechapago": a_fecha(celda(fila, "FECHA PAGO"))}
        if not vendedor:
            errores.append(f"Fila {num}: factura {factura} sin vendedor (no está debajo de una fila de vendedor); no se carga")
            continue
        if not r["nrecibo"]:
            errores.append(f"Fila {num}: factura {factura} sin número de recibo (ROC); no se carga")
            continue
        if r["fechapago"] is None:
            errores.append(f"Fila {num}: factura {factura} sin fecha de pago válida; no se carga")
            continue
        try:
            for col, campo in COLUMNAS.items():
                if campo in MONTOS:
                    r[campo] = a_decimal(celda(fila, col))
            dias = a_texto(celda(fila, "DIAS RECUPER.")).replace(",", "")
            r["drecuperado"] = int(float(dias)) if dias else None
        except ValueError as e:
            errores.append(f"Fila {num}: factura {factura}: {e}; no se carga")
            continue
        for campo, largo in LARGOS.items():
            if r[campo] and len(r[campo]) > largo:
                if campo in ("nfactura", "nrecibo"):
                    errores.append(f"Fila {num}: {NOMBRES[campo]} '{r[campo]}' demasiado largo")
                    break
                r[campo] = r[campo][:largo]
                recortados += campo == "cliente"
        else:
            if r["fechafact"] and r["drecuperado"] is not None and abs((r["fechapago"] - r["fechafact"]).days - r["drecuperado"]) > 1:
                advertencias.append(f"Fila {num}: factura {factura}: {r['drecuperado']} días de recuperación, pero entre las fechas hay "
                                    f"{(r['fechapago'] - r['fechafact']).days}")
            r["clave"] = f"{r['nfactura']} / {r['nrecibo']}"
            r["_fila"] = num
            # Un recibo puede abonar dos veces a la misma factura: no se descartan líneas repetidas aquí,
            # se valida al comparar que lo pagado no pase del total de la factura
            vistos[r["clave"]] = vistos.get(r["clave"], 0) + 1
            registros.append(r)
    if not registros:
        raise ErrorArchivo("no contiene pagos")
    if recortados:
        advertencias.append(f"{recortados} nombre(s) de cliente tienen más de 45 caracteres y se guardan recortados (igual que los cargados antes)")
    dobles = sum(1 for n in vistos.values() if n > 1)
    if dobles:
        advertencias.append(f"{dobles} recibo(s) abonan más de una vez a la misma factura; se cargan todas sus líneas "
                            f"(se revisa que lo pagado no pase del total)")
    ceros = sum(1 for r in registros if r["mpagado"] == 0)
    if ceros:
        advertencias.append(f"{ceros} línea(s) con monto pagado 0 (recibos que no abonan a esa factura); se guardan igual que en cargas anteriores")
    return registros, errores, advertencias


TOLERANCIA = Decimal("1.00")   # C$ de diferencia aceptada por redondeo al comparar lo pagado con el total


def _valores(r):
    """Línea normalizada para comparar archivo y base de datos (números a 2 decimales, textos sin espacios extra)."""
    valores = []
    for c in CAMPOS:
        v = r[c]
        if c in NUMERICAS:
            v = str(Decimal(str(v)).quantize(Decimal("0.01"))) if v is not None else None
        elif isinstance(v, date):
            v = v.isoformat()
        elif isinstance(v, str):
            v = " ".join(v.split())
        valores.append(v)
    return tuple(valores)


def comparar(registros):
    """Compara por grupos factura + recibo y valida que lo pagado a cada factura no pase de su total."""
    facturas = sorted({r["nfactura"] for r in registros})
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    actuales = {}
    try:
        for i in range(0, len(facturas), 1000):
            lote = facturas[i:i + 1000]
            cursor.execute(f"SELECT * FROM recuperacion WHERE nfactura IN ({', '.join(['%s'] * len(lote))})", lote)
            for a in cursor.fetchall():
                actuales.setdefault(a["nfactura"], []).append(a)
    finally:
        cursor.close()
        conn.close()

    grupos = {}
    for r in registros:
        grupos.setdefault((r["nfactura"], r["nrecibo"]), []).append(r)

    # 1) Lo pagado por factura no puede pasar del total: archivo + recibos ya cargados que no vienen en el archivo
    errores, excedidas = [], set()
    for factura in facturas:
        del_archivo = [r for r in registros if r["nfactura"] == factura]
        recibos = {r["nrecibo"] for r in del_archivo}
        otros = [a for a in actuales.get(factura, []) if a["nrecibo"] not in recibos]
        pagado = sum(r["mpagado"] for r in del_archivo) + sum(Decimal(a["mpagado"] or 0) for a in otros)
        total = max([r["total"] for r in del_archivo] + [Decimal(a["total"] or 0) for a in otros])
        if pagado > total + TOLERANCIA:
            excedidas.add(factura)
            detalle = ", ".join(f"ROC {r['nrecibo']} C$ {r['mpagado']:,.2f}" for r in del_archivo)
            if otros:
                detalle += "; ya cargados: " + ", ".join(f"ROC {a['nrecibo']} C$ {Decimal(a['mpagado'] or 0):,.2f}" for a in otros)
            errores.append(f"Factura {factura}: lo pagado (C$ {pagado:,.2f}) pasa del total de la factura (C$ {total:,.2f}) "
                           f"[{detalle}]; no se carga esta factura")

    # 2) Comparar cada grupo factura + recibo con lo que hay en la base de datos
    nuevos, modificados, sin_cambio = [], [], 0
    for (factura, recibo), lineas in grupos.items():
        if factura in excedidas:
            continue
        en_bd = [a for a in actuales.get(factura, []) if a["nrecibo"] == recibo]
        if not en_bd:
            nuevos.extend(lineas)
        elif sorted(_valores(a) for a in en_bd) == sorted(_valores(r) for r in lineas):
            sin_cambio += len(lineas)
        else:
            if len(en_bd) != len(lineas):
                cambios = [("lineas", len(en_bd), len(lineas))]
            else:
                pares = zip(sorted(en_bd, key=lambda a: _valores(a)), sorted(lineas, key=lambda r: _valores(r)))
                cambios = [(c, a[c], r[c]) for a, r in pares for c in CAMPOS if not iguales(c, a[c], r[c], NUMERICAS)]
            modificados.append({"registro": lineas[0], "lineas": lineas, "cambios": cambios, "ids": [a["id"] for a in en_bd]})
    return {"nuevos": nuevos, "modificados": modificados, "sin_cambio": sin_cambio, "errores": errores,
            "excluidas": sum(len(l) for (f, _), l in grupos.items() if f in excedidas)}


def _json(v):
    return str(v) if isinstance(v, Decimal) else (v.isoformat() if isinstance(v, date) else v)


@cargar_recuperacion_bp.route("/cargar_recuperacion", methods=["GET", "POST"])
@requiere_acceso("cargar_recuperacion")
def cargar_recuperacion():
    if request.method == "GET":
        return render_template("cargar_recuperacion.html", ultimos=resumen_meses())
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return render_template("cargar_recuperacion.html", errores_archivo=["Falta el archivo de Comisiones Crédito."], ultimos=resumen_meses())
    try:
        registros, errores, advertencias = leer_recuperacion(archivo)
        dif = comparar(registros)
    except ErrorArchivo as e:
        return render_template("cargar_recuperacion.html", errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"], ultimos=resumen_meses())
    except Exception as e:
        return render_template("cargar_recuperacion.html", errores_archivo=[f"{archivo.filename}: error inesperado – {e}"], ultimos=resumen_meses())

    errores = dif["errores"] + errores
    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO), "w", encoding="utf-8") as f:
        json.dump({"nuevos": [[_json(r[c]) for c in CAMPOS] for r in dif["nuevos"]],
                   "modificados": [{"ids": m["ids"], "filas": [[_json(r[c]) for c in CAMPOS] for r in m["lineas"]]}
                                   for m in dif["modificados"]]},
                  f, ensure_ascii=False)
    session["carga_recuperacion_token"] = token
    fechas = [r["fechapago"] for r in registros]
    por_vendedor = {}
    for r in registros:
        v = por_vendedor.setdefault(r["vendedor"], {"vendedor": r["vendedor"], "lineas": 0, "pagado": Decimal(0), "abonado": Decimal(0)})
        v["lineas"] += 1
        v["pagado"] += r["mpagado"]
        v["abonado"] += r["realabono"]
    return render_template("cargar_recuperacion.html", validado=True, token=token, dif=dif, total=len(registros),
                           desde=min(fechas), hasta=max(fechas), pagado=sum(r["mpagado"] for r in registros),
                           abonado=sum(r["realabono"] for r in registros),
                           por_vendedor=sorted(por_vendedor.values(), key=lambda x: -x["abonado"]),
                           nombres=NOMBRES, errores=errores, advertencias=advertencias, max_detalle=MAX_DETALLE,
                           ultimos=resumen_meses())


@cargar_recuperacion_bp.route("/cargar_recuperacion/confirmar", methods=["POST"])
@requiere_acceso("cargar_recuperacion")
def confirmar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO)
    if not ruta or token != session.get("carga_recuperacion_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_recuperacion.cargar_recuperacion"))
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()   # todo o nada
        if datos["nuevos"]:
            cursor.executemany(f"INSERT INTO recuperacion ({', '.join(CAMPOS)}) VALUES ({', '.join(['%s'] * len(CAMPOS))})", datos["nuevos"])
        # Grupo factura + recibo que cambió: se reemplazan sus líneas (pueden ser más o menos que antes)
        for m in datos["modificados"]:
            cursor.execute(f"DELETE FROM recuperacion WHERE id IN ({', '.join(['%s'] * len(m['ids']))})", m["ids"])
            cursor.executemany(f"INSERT INTO recuperacion ({', '.join(CAMPOS)}) VALUES ({', '.join(['%s'] * len(CAMPOS))})", m["filas"])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_recuperacion.cargar_recuperacion"))
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop("carga_recuperacion_token", None)
    flash(f"✅ Recuperación guardada: {len(datos['nuevos']):,} pagos nuevos y {len(datos['modificados']):,} actualizados.", "success")
    return redirect(url_for("cargar_recuperacion.cargar_recuperacion"))


def resumen_meses():
    """Lo cargado por mes (para ver de un vistazo qué falta)."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT DATE_FORMAT(fechapago, '%Y-%m') AS mes, COUNT(*) AS lineas, SUM(mpagado) AS pagado, "
                       "SUM(realabono) AS abonado, MIN(fechapago) AS desde, MAX(fechapago) AS hasta "
                       "FROM recuperacion GROUP BY mes ORDER BY mes DESC LIMIT 6")
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
