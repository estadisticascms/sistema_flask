import json
import os
import re
import uuid

from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, ruta_temporal, limpiar_temporales
from router.cargar_inventario import arreglar_texto
from router.cargar_catalogo import sin_tildes, a_numero, iguales, resumen_cambios, MAX_DETALLE, LIMITE_ELIMINADOS, LOTE_INSERT

cargar_existencias_bp = Blueprint("cargar_existencias", __name__)

PREFIJO_TEMP = "carga_existencias_"

# Columna de la tabla -> encabezado del Excel
COLUMNAS_EXCEL = {
    "Bodega": "Bodega", "Id Producto": "Id Producto", "Producto": "Producto",
    "Codigo_producto": "Código Producto", "Número Parte": "Número Parte", "Medida": "Medida",
    "Marca": "Marca", "Código SAC": "Código SAC", "UOM": "UOM",
    "Monto Inicial": "Monto Inicial", "Monto Entrada": "Monto Entrada",
    "Monto Ajuste Entrada": "Monto Ajuste Entrada", "Monto Ajuste Salida": "Monto Ajuste Salida",
    "Ajuste Costo Entrada": "Ajuste Costo Entrada", "Ajuste Costo Salida": "Ajuste Costo Salida",
    "Monto Traslados Enviado": "Monto Traslados Enviado", "Monto Traslado Recibido": "Monto Traslado Recibido",
    "Monto Salida": "Monto Salida", "Monto Final": "Monto Final",
    "Cantidad Inicial": "Cantidad Inicial", "Entradas": "Entradas", "Ajustes": "Ajustes",
    "Traslados Enviado": "Traslados Enviado", "Traslados Recibido": "Traslados Recibido",
    "Salidas": "Salidas", "Cant_Final": "Cantidad Final",
}
COLUMNAS = list(COLUMNAS_EXCEL)
TEXTO = {"Bodega", "Producto", "Codigo_producto", "Número Parte", "Medida", "Marca", "Código SAC", "UOM"}
NUMERICAS = set(COLUMNAS) - TEXTO
LARGOS = {"Bodega": 50, "Codigo_producto": 50}


# -----------------------------
# Lectura y validación
# -----------------------------
def leer_existencias(archivo):
    filas = leer_filas(archivo)

    periodo = None
    for fila in filas[:10]:
        m = re.search(r"Desde:\s*(\S+)\s+Al:\s*(\S+)", a_texto(fila[0]) if fila else "")
        if m:
            periodo = (m.group(1), m.group(2))

    objetivo = {sin_tildes(e): c for c, e in COLUMNAS_EXCEL.items()}
    inicio, pos = None, None
    for i, fila in enumerate(filas[:40]):
        encontrados = {objetivo[sin_tildes(a_texto(v))]: j for j, v in enumerate(fila)
                       if sin_tildes(a_texto(v)) in objetivo}
        if len(encontrados) == len(COLUMNAS):
            inicio, pos = i, encontrados
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas del Reporte de Inventario (Bodega, Código Producto, Cantidad Final, ...)")

    registros, errores, advertencias = [], [], []
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        if not any(a_texto(v) for v in fila):
            continue
        r, problemas = {}, []
        for col in COLUMNAS:
            valor = fila[pos[col]] if pos[col] < len(fila) else None
            if col in NUMERICAS:
                try:
                    r[col] = a_numero(valor)
                except ValueError:
                    r[col] = None
                    problemas.append(f"{COLUMNAS_EXCEL[col]} no es un número ('{a_texto(valor)}')")
            else:
                r[col] = arreglar_texto(a_texto(valor)) or None
        ref = f"Fila {num} ({r['Codigo_producto'] or '?'} / {r['Bodega'] or '?'})"
        if not r["Codigo_producto"] or not r["Bodega"]:
            errores.append(f"{ref}: no tiene Código Producto o Bodega; no se carga")
            continue
        largos = [c for c, m in LARGOS.items() if len(r[c]) > m]
        if largos:
            errores.append(f"{ref}: texto demasiado largo en {', '.join(largos)}; no se carga")
            continue
        if problemas:
            advertencias.append(f"{ref}: " + "; ".join(problemas) + "; se guarda vacío")
        r["_fila"] = num
        registros.append(r)
    if not registros:
        raise ErrorArchivo("no contiene existencias")
    return registros, errores, advertencias, periodo


def descuadres(registros):
    """Filas donde el propio reporte no cuadra: Final ≠ Inicial + Entradas + Ajustes − Enviados + Recibidos − Salidas."""
    avisos = []
    for r in registros:
        v = {c: r[c] or 0 for c in NUMERICAS}
        calculado = (v["Cantidad Inicial"] + v["Entradas"] + v["Ajustes"] - v["Traslados Enviado"]
                     + v["Traslados Recibido"] - v["Salidas"])
        if abs(calculado - v["Cant_Final"]) > 0.001:
            avisos.append(f"{r['Codigo_producto']} / {r['Bodega']}: el movimiento da {calculado:g} pero Cantidad Final es {v['Cant_Final']:g}")
    return avisos


# -----------------------------
# Comparación con la base de datos
# -----------------------------
def comparar(registros):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM existencias_bodega")
        actuales = cursor.fetchall()
        cursor.execute("SELECT `Codigo Producto` AS codigo, `Existencia Actual` AS existencia "
                       "FROM catalogo_productos WHERE `Codigo Producto` IS NOT NULL")
        catalogo = {f["codigo"]: f["existencia"] or 0 for f in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    por_clave = {(a["Codigo_producto"], a["Bodega"]): a for a in actuales}
    nuevos, modificados, sin_cambio = [], [], 0
    for r in registros:
        actual = por_clave.get((r["Codigo_producto"], r["Bodega"]))
        if actual is None:
            nuevos.append(r)
            continue
        cambios = [(c, actual[c], r[c]) for c in COLUMNAS if not iguales(c, actual[c], r[c], NUMERICAS)]
        if cambios:
            modificados.append({"codigo": r["Codigo_producto"], "bodega": r["Bodega"], "nombre": r["Producto"], "cambios": cambios})
        else:
            sin_cambio += 1
    en_archivo = {(r["Codigo_producto"], r["Bodega"]) for r in registros}
    eliminados = [a for k, a in por_clave.items() if k not in en_archivo]

    # Resumen por bodega: antes / después
    bodegas = {}
    for origen, filas in (("antes", actuales), ("despues", registros)):
        for f in filas:
            b = bodegas.setdefault(f["Bodega"], {"bodega": f["Bodega"], "antes": [0, 0, 0.0], "despues": [0, 0, 0.0]})
            b[origen][0] += 1
            b[origen][1] += f["Cant_Final"] or 0
            b[origen][2] += f["Monto Final"] or 0

    # Existencia total por producto vs catálogo
    totales = {}
    for r in registros:
        totales[r["Codigo_producto"]] = totales.get(r["Codigo_producto"], 0) + (r["Cant_Final"] or 0)
    vs_catalogo = [f"{c}: suma de bodegas {t:g} vs catálogo {catalogo[c]:g}"
                   for c, t in sorted(totales.items()) if c in catalogo and abs(t - catalogo[c]) > 0.001]
    sin_catalogo = sorted(c for c in totales if c not in catalogo)

    return {
        "actuales": len(actuales),
        "nuevos": nuevos,
        "modificados": modificados,
        "eliminados": eliminados,
        "sin_cambio": sin_cambio,
        "bodegas": sorted(bodegas.values(), key=lambda b: -b["despues"][2]),
        "vs_catalogo": vs_catalogo,
        "sin_catalogo": sin_catalogo,
    }


# -----------------------------
# Vistas
# -----------------------------
@cargar_existencias_bp.route("/cargar_existencias", methods=["GET", "POST"])
@requiere_acceso("cargar_existencias")
def cargar_existencias():
    if request.method == "GET":
        return render_template("cargar_existencias.html")

    archivo = request.files.get("archivo_existencias")
    if not archivo or not archivo.filename:
        return render_template("cargar_existencias.html", errores_archivo=["Falta el Reporte de Inventario."])
    try:
        registros, errores, advertencias, periodo = leer_existencias(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_existencias.html",
                               errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_existencias.html",
                               errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    # Producto + bodega repetido: se conserva la primera aparición
    vistos, unicos = {}, []
    for r in registros:
        clave = (r["Codigo_producto"], r["Bodega"])
        if clave in vistos:
            errores.append(f"Fila {r['_fila']}: {clave[0]} / {clave[1]} repetido (ya está en la fila {vistos[clave]}); se ignora esta fila")
            continue
        vistos[clave] = r["_fila"]
        unicos.append(r)
    registros = unicos

    try:
        dif = comparar(registros)
    except Exception as e:
        return render_template("cargar_existencias.html", errores_archivo=[f"Error al comparar con la base de datos: {e}"])

    pedir_confirmacion = dif["actuales"] and len(dif["eliminados"]) > dif["actuales"] * LIMITE_ELIMINADOS
    if pedir_confirmacion:
        advertencias.insert(0, f"Se eliminarían {len(dif['eliminados']):,} filas producto/bodega que hoy están en la tabla y no vienen en el archivo. "
                               "Verifique que el reporte no tenga filtros (bodega, marca, fechas) antes de confirmar.")
    if dif["sin_catalogo"]:
        advertencias.append(f"{len(dif['sin_catalogo'])} productos no existen en catalogo_productos: "
                            + ", ".join(dif["sin_catalogo"][:20]) + (" …" if len(dif["sin_catalogo"]) > 20 else ""))

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump([[r[c] for c in COLUMNAS] for r in registros], f, ensure_ascii=False)
    session["carga_existencias_token"] = token

    return render_template(
        "cargar_existencias.html",
        validado=True,
        token=token,
        total=len(registros),
        periodo=periodo,
        dif=dif,
        cambios_por_columna=resumen_cambios(dif["modificados"]),
        descuadres=descuadres(registros),
        max_detalle=MAX_DETALLE,
        errores=errores,
        advertencias=advertencias,
        pedir_confirmacion=pedir_confirmacion,
    )


@cargar_existencias_bp.route("/cargar_existencias/confirmar", methods=["POST"])
@requiere_acceso("cargar_existencias")
def confirmar_carga_existencias():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_existencias_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_existencias.cargar_existencias"))

    with open(ruta, encoding="utf-8") as f:
        filas = [tuple(fila) for fila in json.load(f)]

    columnas = ", ".join(f"`{c}`" for c in COLUMNAS)
    sql = f"INSERT INTO existencias_bodega ({columnas}) VALUES ({', '.join(['%s'] * len(COLUMNAS))})"

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Reemplazo completo en una sola transacción: si algo falla, la tabla queda como estaba
        conn.start_transaction()
        cursor.execute("DELETE FROM existencias_bodega")
        borrados = cursor.rowcount
        for i in range(0, len(filas), LOTE_INSERT):
            cursor.executemany(sql, filas[i:i + LOTE_INSERT])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se modificaron las existencias. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_existencias.cargar_existencias"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_existencias_token", None)
    flash(f"✅ Existencias por bodega actualizadas: se reemplazaron {borrados:,} registros por {len(filas):,} del archivo.", "success")
    return render_template("cargar_existencias.html", cargado=True)
