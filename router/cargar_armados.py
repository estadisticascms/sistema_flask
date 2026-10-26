"""Carga del reporte "Factura de Armados" (armado de mangueras) a la tabla ventas_armados (docs/sql/ventas_armados.sql).

El reporte trae una línea por producto de cada factura de armado: la manguera armada, los trabajos (hacer adaptador, poner
terminal…) y los repuestos sueltos. MO es la mano de obra de la línea (ya multiplicada por la cantidad) y Otros es lo que se
cobra de soldadura. La columna Ganancia viene en 0: el costo y la utilidad se toman de ventas_historicas en el reporte.

Se compara por factura (número + fecha): si el archivo trae las mismas líneas que la base de datos no se toca nada; si
cambian (por ejemplo pasó de Journal a Pagado) se reemplazan las líneas de esa factura."""
import csv
import io
import json
import os
import re
import unicodedata
import uuid
from datetime import date
from decimal import Decimal

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import iguales
from router.cargar_ventas import (ErrorArchivo, a_decimal, a_factura, a_fecha, a_texto, leer_filas, limpiar_temporales,
                                  ruta_temporal)

cargar_armados_bp = Blueprint("cargar_armados", __name__)

PREFIJO = "carga_armados_"
MAX_DETALLE = 300
# encabezado del reporte (en mayúsculas, sin tildes ni símbolos) -> columna de la tabla
COLUMNAS = {"FACTURA": "nfactura", "FECHA": "fecha", "CONTADOCREDITO": "tipo_venta", "NOMBREDECLIENTE": "cliente",
            "FACTURO": "facturador", "VENDEDOR": "vendedor", "PRODUCTO": "producto", "CANTIDAD": "cantidad",
            "PRECIO": "precio", "SUBTOTAL": "subtotal", "GANANCIA": "ganancia", "MO": "mano_obra", "OTROS": "otros",
            "ESTATUS": "estatus"}
OPCIONALES = {"GANANCIA", "ESTATUS"}
CAMPOS = ["nfactura", "fecha", "tipo_venta", "cliente", "facturador", "vendedor", "producto", "cantidad", "precio",
          "subtotal", "ganancia", "mano_obra", "otros", "estatus"]
NUMERICAS = {"cantidad", "precio", "subtotal", "ganancia", "mano_obra", "otros"}
LARGOS = {"nfactura": 15, "cliente": 150, "facturador": 100, "vendedor": 100, "producto": 255, "estatus": 20}
NOMBRES = {"nfactura": "Factura", "fecha": "Fecha", "tipo_venta": "Contado/Crédito", "cliente": "Cliente",
           "facturador": "Facturó", "vendedor": "Vendedor", "producto": "Producto", "cantidad": "Cantidad",
           "precio": "Precio", "subtotal": "Subtotal", "ganancia": "Ganancia", "mano_obra": "Mano de obra",
           "otros": "Soldadura (otros)", "estatus": "Estatus", "clave": "Factura (fecha)", "lineas": "Líneas de la factura"}

# Trabajos que no son una manguera nueva (se cobra mano de obra sobre una pieza del cliente)
SERVICIO = re.compile(r"^(HACER|ELABORACION|REPARACION|PONER|CORTAR|CAMBIAR|SOLDAR)\b")


def tipo_linea(producto, mano_obra=0):
    """'manguera' (cuenta como manguera armada), 'servicio' (trabajo con mano de obra) o 'repuesto' (pieza suelta)."""
    p = " ".join((producto or "").upper().split())
    if p.startswith("MANGUERA"):
        return "manguera"
    if SERVICIO.match(p) or (mano_obra or 0) > 0:
        return "servicio"
    return "repuesto"


def _encabezado(valor):
    texto = unicodedata.normalize("NFKD", a_texto(valor).upper())
    texto = re.sub(r"[^A-Z0-9]", "", texto)
    return "FACTURO" if texto == "FACTUR" else texto   # "Facturó" llega como "Factur�" si el CSV viene en otra codificación


def _leer_csv(contenido):
    try:
        texto = contenido.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = contenido.decode("cp1252", errors="replace")
    primera = texto.split("\n", 1)[0]
    separador = ";" if primera.count(";") >= primera.count(",") else ","   # los montos traen coma de miles
    return list(csv.reader(io.StringIO(texto), delimiter=separador))


def leer_armados(archivo):
    if archivo.filename.lower().endswith((".csv", ".txt")):
        contenido = archivo.read()
        if not contenido:
            raise ErrorArchivo("el archivo está vacío")
        filas = _leer_csv(contenido)
    else:
        filas = leer_filas(archivo)
    inicio, pos = None, None
    for i, fila in enumerate(filas[:40]):
        nombres = {_encabezado(v): j for j, v in enumerate(fila)}
        if all(c in nombres for c in COLUMNAS if c not in OPCIONALES):
            inicio, pos = i, {c: nombres[c] for c in COLUMNAS if c in nombres}
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas del reporte Factura de Armados "
                           "(# Factura, Fecha, Contado/Credito, Producto, Cantidad, SubTotal, MO, Otros, ...)")

    def celda(fila, col):
        j = pos.get(col)
        return fila[j] if j is not None and j < len(fila) else None

    registros, errores, advertencias, recortados, descuadres = [], [], [], 0, []
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        factura = a_factura(celda(fila, "FACTURA"))
        producto = " ".join(a_texto(celda(fila, "PRODUCTO")).split())
        if not factura or factura.upper().startswith("TOTAL") or not producto:
            continue   # filas vacías o de totales
        r = {"nfactura": factura, "fecha": a_fecha(celda(fila, "FECHA")),
             "tipo_venta": a_texto(celda(fila, "CONTADOCREDITO")).upper(), "producto": producto,
             "cliente": a_texto(celda(fila, "NOMBREDECLIENTE")) or None, "facturador": a_texto(celda(fila, "FACTURO")) or None,
             "vendedor": a_texto(celda(fila, "VENDEDOR")) or None, "estatus": a_texto(celda(fila, "ESTATUS")) or None}
        if r["fecha"] is None:
            errores.append(f"Fila {num}: factura {factura} sin fecha válida; no se carga")
            continue
        if r["tipo_venta"] not in ("CONTADO", "CREDITO"):
            errores.append(f"Fila {num}: factura {factura}: Contado/Credito '{r['tipo_venta']}' no reconocido; no se carga")
            continue
        try:
            for col, campo in COLUMNAS.items():
                if campo in NUMERICAS:
                    r[campo] = a_decimal(celda(fila, col))
        except ValueError as e:
            errores.append(f"Fila {num}: factura {factura}: {e}; no se carga")
            continue
        if len(factura) > LARGOS["nfactura"]:
            errores.append(f"Fila {num}: factura '{factura}' demasiado larga; no se carga")
            continue
        for campo, largo in LARGOS.items():
            if r[campo] and len(r[campo]) > largo:
                r[campo] = r[campo][:largo]
                recortados += 1
        if abs(r["cantidad"] * r["precio"] - r["subtotal"]) > 1:
            descuadres.append(f"Fila {num}: factura {factura}: {r['cantidad']:,} × {r['precio']:,.2f} ≠ subtotal {r['subtotal']:,.2f}")
        r["clave"] = f"{factura} ({r['fecha']:%d/%m/%Y})"
        registros.append(r)
    if not registros:
        raise ErrorArchivo("no contiene líneas de facturas")
    if recortados:
        advertencias.append(f"{recortados} texto(s) muy largos se guardan recortados")
    advertencias += [f"{d} (se guarda el subtotal del archivo)" for d in descuadres]
    return registros, errores, advertencias


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
    """Compara por factura (número + fecha) con lo que ya hay en ventas_armados."""
    facturas = sorted({r["nfactura"] for r in registros})
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    actuales, en_ventas = {}, set()
    try:
        for i in range(0, len(facturas), 1000):
            lote = facturas[i:i + 1000]
            marcas = ", ".join(["%s"] * len(lote))
            cursor.execute(f"SELECT * FROM ventas_armados WHERE nfactura IN ({marcas})", lote)
            for a in cursor.fetchall():
                actuales.setdefault((a["nfactura"], a["fecha"]), []).append(a)
            cursor.execute(f"SELECT factura, Fecha FROM ventas_historicas WHERE factura IN ({marcas})", lote)
            en_ventas |= {(a["factura"], a["Fecha"]) for a in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    grupos = {}
    for r in registros:
        grupos.setdefault((r["nfactura"], r["fecha"]), []).append(r)

    nuevos, modificados, sin_cambio = [], [], 0
    for clave, lineas in grupos.items():
        en_bd = actuales.get(clave, [])
        if not en_bd:
            nuevos.extend(lineas)
        elif sorted(_valores(a) for a in en_bd) == sorted(_valores(r) for r in lineas):
            sin_cambio += len(lineas)
        else:
            if len(en_bd) != len(lineas):
                cambios = [("lineas", len(en_bd), len(lineas))]
            else:
                pares = zip(sorted(en_bd, key=_valores), sorted(lineas, key=_valores))
                cambios = [(c, a[c], r[c]) for a, r in pares for c in CAMPOS if not iguales(c, a[c], r[c], NUMERICAS)]
            modificados.append({"registro": lineas[0], "lineas": lineas, "cambios": cambios, "ids": [a["id"] for a in en_bd]})
    sin_venta = sorted(f"{f} ({d:%d/%m/%Y})" for f, d in grupos if (f, d) not in en_ventas)
    return {"nuevos": nuevos, "modificados": modificados, "sin_cambio": sin_cambio, "sin_venta": sin_venta}


def _json(v):
    return str(v) if isinstance(v, Decimal) else (v.isoformat() if isinstance(v, date) else v)


def _pantalla(**datos):
    return render_template("cargar_armados.html", ultimos=resumen_meses(), **datos)


@cargar_armados_bp.route("/cargar_armados", methods=["GET", "POST"])
@requiere_acceso("cargar_armados")
def cargar_armados():
    if request.method == "GET":
        return _pantalla()
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return _pantalla(errores_archivo=["Falta el archivo Factura de Armados."])
    try:
        registros, errores, advertencias = leer_armados(archivo)
        dif = comparar(registros)
    except ErrorArchivo as e:
        return _pantalla(errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return _pantalla(errores_archivo=[f"{archivo.filename}: error inesperado – {e}"])

    if dif["sin_venta"]:
        advertencias.insert(0, f"{len(dif['sin_venta'])} factura(s) aún no están en Ventas del día, así que el reporte no tendrá su "
                               f"costo ni ganancia hasta cargarlas: {', '.join(dif['sin_venta'][:20])}"
                               f"{' …' if len(dif['sin_venta']) > 20 else ''}")
    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO), "w", encoding="utf-8") as f:
        json.dump({"nuevos": [[_json(r[c]) for c in CAMPOS] for r in dif["nuevos"]],
                   "modificados": [{"ids": m["ids"], "filas": [[_json(r[c]) for c in CAMPOS] for r in m["lineas"]]}
                                   for m in dif["modificados"]]},
                  f, ensure_ascii=False)
    session["carga_armados_token"] = token
    fechas = [r["fecha"] for r in registros]
    mangueras = [r for r in registros if tipo_linea(r["producto"], r["mano_obra"]) == "manguera"]
    return _pantalla(validado=True, token=token, dif=dif, total=len(registros), desde=min(fechas), hasta=max(fechas),
                     facturas=len({(r["nfactura"], r["fecha"]) for r in registros}),
                     mangueras=sum(r["cantidad"] for r in mangueras), venta=sum(r["subtotal"] for r in registros),
                     mano_obra=sum(r["mano_obra"] for r in registros), otros=sum(r["otros"] for r in registros),
                     nombres=NOMBRES, errores=errores, advertencias=advertencias, max_detalle=MAX_DETALLE)


@cargar_armados_bp.route("/cargar_armados/confirmar", methods=["POST"])
@requiere_acceso("cargar_armados")
def confirmar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO)
    if not ruta or token != session.get("carga_armados_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_armados.cargar_armados"))
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    insertar = f"INSERT INTO ventas_armados ({', '.join(CAMPOS)}) VALUES ({', '.join(['%s'] * len(CAMPOS))})"
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()   # todo o nada
        if datos["nuevos"]:
            cursor.executemany(insertar, datos["nuevos"])
        # Factura que cambió: se reemplazan sus líneas (pueden ser más o menos que antes)
        for m in datos["modificados"]:
            cursor.execute(f"DELETE FROM ventas_armados WHERE id IN ({', '.join(['%s'] * len(m['ids']))})", m["ids"])
            cursor.executemany(insertar, m["filas"])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_armados.cargar_armados"))
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop("carga_armados_token", None)
    flash(f"✅ Armados guardados: {len(datos['nuevos']):,} líneas nuevas y {len(datos['modificados']):,} facturas actualizadas.", "success")
    return redirect(url_for("cargar_armados.cargar_armados"))


def resumen_meses():
    """Lo cargado por mes (para ver de un vistazo qué falta)."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT DATE_FORMAT(fecha, '%Y-%m') AS mes, COUNT(DISTINCT nfactura, fecha) AS facturas, COUNT(*) AS lineas, "
                       "SUM(subtotal) AS venta, SUM(mano_obra) AS mano_obra, SUM(otros) AS otros, MIN(fecha) AS desde, "
                       "MAX(fecha) AS hasta FROM ventas_armados GROUP BY mes ORDER BY mes DESC LIMIT 6")
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
