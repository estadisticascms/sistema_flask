import json
import os
import re
import unicodedata
import uuid

from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, ruta_temporal, limpiar_temporales
from router.cargar_inventario import arreglar_texto

cargar_catalogo_bp = Blueprint("cargar_catalogo", __name__)

PREFIJO_TEMP = "carga_catalogo_"
HOJA = "CATALOGO EXISTENCIAS"
MAX_DETALLE = 300           # filas a mostrar por lista en la vista previa
LIMITE_ELIMINADOS = 0.02    # si se eliminaría más del 2% del catálogo, se pide confirmación explícita
LOTE_INSERT = 2000

COLUMNAS = [
    "Id Producto", "Nombre", "Codigo Producto", "Numero Parte", "Medida", "Descripcion Corta",
    "Descripcion Larga", "Observaciones", "Equivalencias", "Codigo EAN", "Codigo SKU", "Codigo UPC",
    "Modelo", "Otros 1", "Otros 2", "Codigo SAC", "Unidad Medida", "Tipo Producto", "Categoria", "Marca",
    "Precio 1", "Precio 2", "Precio 3", "Precio 4", "Moneda", "Costo Promedio", "Ultimo Costo",
    "Existencia Actual",
]
NUMERICAS = {"Precio 1", "Precio 2", "Precio 3", "Precio 4", "Costo Promedio", "Ultimo Costo", "Existencia Actual"}
CAMPOS_VISIBLES = ["Existencia Actual", "Costo Promedio", "Ultimo Costo", "Precio 1", "Nombre", "Marca", "Numero Parte"]


def sin_tildes(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def a_numero(valor):
    texto = a_texto(valor).replace(",", "")
    if texto == "":
        return None
    return float(texto)  # ValueError si no es número


# -----------------------------
# Lectura del archivo
# -----------------------------
def leer_catalogo(archivo):
    try:
        filas = leer_filas(archivo, hoja=HOJA)
    except ErrorArchivo:
        archivo.seek(0)
        filas = leer_filas(archivo)  # por si la hoja viene con otro nombre

    objetivo = {sin_tildes(c): c for c in COLUMNAS}
    inicio, pos = None, None
    for i, fila in enumerate(filas[:40]):
        encontrados = {objetivo[sin_tildes(a_texto(v))]: j for j, v in enumerate(fila)
                       if sin_tildes(a_texto(v)) in objetivo}
        if len(encontrados) == len(COLUMNAS):
            inicio, pos = i, encontrados
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas del catálogo (Id Producto, Código Producto, Existencia Actual, ...)")

    productos, errores, advertencias, total_reportado = [], [], [], None
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        primera = a_texto(fila[pos["Id Producto"]])
        m = re.match(r"TOTAL DE REGISTROS:\s*(\d+)", primera, re.I)
        if m:
            total_reportado = int(m.group(1))
            continue
        if not any(a_texto(v) for v in fila):
            continue

        p = {}
        for col in COLUMNAS:
            valor = fila[pos[col]] if pos[col] < len(fila) else None
            if col in NUMERICAS:
                try:
                    p[col] = a_numero(valor)
                except ValueError:
                    p[col] = None
                    advertencias.append(f"Fila {num} ({a_texto(fila[pos['Codigo Producto']])}): {col} no es un número ('{a_texto(valor)}'); se guarda vacío")
            else:
                p[col] = arreglar_texto(a_texto(valor)) or None
        if not p["Codigo Producto"]:
            errores.append(f"Fila {num}: no tiene Código Producto ({p['Nombre'] or 'sin nombre'}); no se carga")
            continue
        p["_fila"] = num
        productos.append(p)

    if not productos:
        raise ErrorArchivo("no contiene productos")
    return productos, errores, advertencias, total_reportado


# -----------------------------
# Comparación con la base de datos
# -----------------------------
def iguales(col, a, b, numericas=NUMERICAS):
    if col in numericas:
        if a is None or b is None:
            return a is None and b is None
        return abs(float(a) - float(b)) <= 1e-6 * max(1.0, abs(float(a)))
    return _texto_comparable(a) == _texto_comparable(b)


def _texto_comparable(valor):
    """Ignora diferencias de espacios/saltos de línea y el texto 'null' que dejó la carga anterior."""
    if valor is None:
        return ""
    texto = " ".join(str(valor).split())
    return "" if texto.lower() == "null" else texto


def comparar(productos):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM catalogo_productos")
        actuales = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    por_codigo, basura = {}, []
    for a in actuales:
        if a["Codigo Producto"]:
            por_codigo[a["Codigo Producto"]] = a
        else:
            basura.append(a)

    nuevos, modificados, sin_cambio = [], [], 0
    for p in productos:
        actual = por_codigo.get(p["Codigo Producto"])
        if actual is None:
            nuevos.append(p)
            continue
        cambios = [(c, actual[c], p[c]) for c in COLUMNAS if not iguales(c, actual[c], p[c])]
        if cambios:
            modificados.append({"codigo": p["Codigo Producto"], "nombre": p["Nombre"], "cambios": cambios})
        else:
            sin_cambio += 1

    en_archivo = {p["Codigo Producto"] for p in productos}
    eliminados = [a for c, a in por_codigo.items() if c not in en_archivo]

    def existencias(filas):
        unidades = sum((f["Existencia Actual"] or 0) for f in filas)
        valor = sum((f["Existencia Actual"] or 0) * (f["Costo Promedio"] or 0) for f in filas)
        return unidades, valor

    return {
        "actuales": len(actuales),
        "nuevos": nuevos,
        "modificados": modificados,
        "eliminados": eliminados,
        "basura": basura,
        "sin_cambio": sin_cambio,
        "antes": existencias(list(por_codigo.values())),
        "despues": existencias(productos),
    }


def resumen_cambios(modificados):
    """Cuántos productos cambiaron en cada columna."""
    conteo = {}
    for m in modificados:
        for col, _, _ in m["cambios"]:
            conteo[col] = conteo.get(col, 0) + 1
    return sorted(conteo.items(), key=lambda x: -x[1])


# -----------------------------
# Vistas
# -----------------------------
@cargar_catalogo_bp.route("/cargar_catalogo", methods=["GET", "POST"])
@requiere_acceso("cargar_catalogo")
def cargar_catalogo():
    if request.method == "GET":
        return render_template("cargar_catalogo.html")

    archivo = request.files.get("archivo_catalogo")
    if not archivo or not archivo.filename:
        return render_template("cargar_catalogo.html", errores_archivo=["Falta el Reporte de Catálogo de Productos."])
    try:
        productos, errores, advertencias, total_reportado = leer_catalogo(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_catalogo.html",
                               errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_catalogo.html",
                               errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    # Códigos repetidos: se conserva la primera aparición
    vistos, unicos = {}, []
    for p in productos:
        if p["Codigo Producto"] in vistos:
            errores.append(f"Fila {p['_fila']}: Código {p['Codigo Producto']} repetido (ya está en la fila {vistos[p['Codigo Producto']]}); se ignora esta fila")
            continue
        vistos[p["Codigo Producto"]] = p["_fila"]
        unicos.append(p)
    productos = unicos

    if total_reportado is not None and total_reportado != len(productos):
        advertencias.insert(0, f"El reporte dice TOTAL DE REGISTROS: {total_reportado:,} pero se leyeron {len(productos):,} productos válidos")

    try:
        dif = comparar(productos)
    except Exception as e:
        return render_template("cargar_catalogo.html", errores_archivo=[f"Error al comparar con la base de datos: {e}"])

    pedir_confirmacion = dif["actuales"] and len(dif["eliminados"]) > dif["actuales"] * LIMITE_ELIMINADOS
    if pedir_confirmacion:
        advertencias.insert(0, f"Se eliminarían {len(dif['eliminados']):,} productos que hoy están en el catálogo y no vienen en el archivo. "
                               "Verifique que el reporte no tenga filtros antes de confirmar.")

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump([[p[c] for c in COLUMNAS] for p in productos], f, ensure_ascii=False)
    session["carga_catalogo_token"] = token

    return render_template(
        "cargar_catalogo.html",
        validado=True,
        token=token,
        total=len(productos),
        dif=dif,
        cambios_por_columna=resumen_cambios(dif["modificados"]),
        campos_visibles=CAMPOS_VISIBLES,
        max_detalle=MAX_DETALLE,
        errores=errores,
        advertencias=advertencias,
        pedir_confirmacion=pedir_confirmacion,
    )


@cargar_catalogo_bp.route("/cargar_catalogo/confirmar", methods=["POST"])
@requiere_acceso("cargar_catalogo")
def confirmar_carga_catalogo():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_catalogo_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_catalogo.cargar_catalogo"))

    with open(ruta, encoding="utf-8") as f:
        filas = [tuple(fila) for fila in json.load(f)]

    columnas = ", ".join(f"`{c}`" for c in COLUMNAS)
    sql = f"INSERT INTO catalogo_productos ({columnas}) VALUES ({', '.join(['%s'] * len(COLUMNAS))})"

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Reemplazo completo en una sola transacción: si algo falla, el catálogo queda como estaba
        conn.start_transaction()
        cursor.execute("DELETE FROM catalogo_productos")
        borrados = cursor.rowcount
        for i in range(0, len(filas), LOTE_INSERT):
            cursor.executemany(sql, filas[i:i + LOTE_INSERT])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se modificó el catálogo. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_catalogo.cargar_catalogo"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_catalogo_token", None)
    flash(f"✅ Catálogo actualizado: se reemplazaron {borrados:,} registros por {len(filas):,} productos del archivo.", "success")
    return render_template("cargar_catalogo.html", cargado=True)
