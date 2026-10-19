import json
import os
import uuid
from datetime import date, timedelta
from decimal import Decimal

from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import (
    ErrorArchivo, a_texto, a_decimal, a_fecha, leer_filas, buscar_encabezado,
    ruta_temporal, limpiar_temporales, TOLERANCIA_MONTO, DIAS_POSIBLE_DUPLICADO,
)

cargar_inventario_bp = Blueprint("cargar_inventario", __name__)

PREFIJO_TEMP = "carga_inventario_"
TIPOS_ENTRADA = {"ENTRADA", "AJUSTE ENTRADA", "TRASLADO ENTRADA"}
TIPOS_SALIDA = {"SALIDA", "AJUSTE SALIDA", "TRASLADO SALIDA"}
TIPOS_TRASLADO = {"TRASLADO ENTRADA", "TRASLADO SALIDA"}

# Columna de la tabla -> encabezado del Excel
COLUMNAS_EXCEL = {
    "id_transaccion": "Id Transaccion",
    "bodega": "Bodega",
    "fecha": "Fecha Transaccion",
    "tipo_transaccion": "Tipo Transaccion",
    "tipo_documento_origen": "Tipo Documento Origen",
    "id_documento_origen": "Id Documento Origen",
    "referencia": "Referencia",
    "cliente_proveedor": "Cliente/Proveedor/Empleado",
    "id_producto": "Id Producto",
    "producto": "Producto",
    "codigo_producto": "Codigo Producto",
    "numero_parte": "Numero Parte",
    "medida": "Medida",
    "marca": "Marca",
    "uom": "UOM",
    "cantidad": "Cantidad",
    "tipo_operacion": "Tipo Operacion",
    "costo_unitario_mn": "Costo Unitario Moneda Nacional",
    "total_mn": "Total Moneda Nacional",
    "tc": "T/C",
    "costo_unitario_me": "Costo Unitario Moneda Extranjera",
    "total_me": "Total Moneda Extranjera",
    "comentarios": "Descripcion/Comentarios",
}
COLUMNAS_SQL = list(COLUMNAS_EXCEL) + ["tipo_final"]
LARGOS = {
    "bodega": 100, "tipo_transaccion": 50, "tipo_documento_origen": 50, "id_documento_origen": 50,
    "referencia": 100, "cliente_proveedor": 150, "id_producto": 50, "producto": 150,
    "codigo_producto": 50, "numero_parte": 50, "medida": 50, "marca": 50, "uom": 20, "tipo_operacion": 50,
}
DECIMALES = {"cantidad": "0.01", "costo_unitario_mn": "0.01", "total_mn": "0.01", "tc": "0.0001",
             "costo_unitario_me": "0.01", "total_me": "0.01"}


def arreglar_texto(texto):
    """Corrige textos con doble codificación que trae el reporte (ej. 'CorrecciÃ³n' -> 'Corrección')."""
    if texto and ("Ã" in texto or "Â" in texto):
        try:
            return texto.encode("latin-1").decode("utf-8").strip()
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return texto


# -----------------------------
# Lectura y validación
# -----------------------------
def leer_reporte_inventario(archivo):
    filas = leer_filas(archivo)
    inicio, col = buscar_encabezado(filas, list(COLUMNAS_EXCEL.values()))
    lineas = []
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        if not a_texto(fila[col["Id Transaccion"]] if col["Id Transaccion"] < len(fila) else None):
            continue
        linea = {c: (fila[col[e]] if col[e] < len(fila) else None) for c, e in COLUMNAS_EXCEL.items()}
        linea["fila"] = num
        lineas.append(linea)
    if not lineas:
        raise ErrorArchivo("no contiene transacciones")
    return lineas


def convertir_linea(l):
    id_tx = a_texto(l["id_transaccion"])
    if not id_tx.isdigit():
        raise ValueError(f"Id Transaccion inválido '{id_tx}'")
    fecha = a_fecha(l["fecha"])
    if fecha is None:
        raise ValueError(f"fecha inválida '{a_texto(l['fecha'])}'")

    r = {c: (arreglar_texto(a_texto(l[c])) or None) for c in COLUMNAS_EXCEL if c not in DECIMALES}
    r["id_transaccion"] = int(id_tx)
    r["fecha"] = fecha.isoformat()
    exactos = {c: _decimal_exacto(l[c], c) for c in DECIMALES}
    for c, q in DECIMALES.items():
        r[c] = str(exactos[c].quantize(Decimal(q)))
    r["tipo_final"] = None
    # Se compara con los valores sin redondear (el costo unitario trae hasta 6 decimales)
    r["_calculado"] = exactos["cantidad"] * exactos["costo_unitario_mn"]
    r["_dif_total"] = abs(r["_calculado"] - exactos["total_mn"])

    if not r["bodega"]:
        raise ValueError("no tiene bodega")
    tipo, op = (r["tipo_transaccion"] or "").upper(), r["tipo_operacion"]
    if tipo == "AJUSTE":  # el histórico separa los ajustes por dirección
        if op not in ("+", "-"):
            raise ValueError(f"AJUSTE con Tipo Operacion desconocido '{op}'")
        tipo = "AJUSTE ENTRADA" if op == "+" else "AJUSTE SALIDA"
    if tipo not in TIPOS_ENTRADA | TIPOS_SALIDA:
        raise ValueError(f"Tipo Transaccion desconocido '{r['tipo_transaccion']}'")
    r["tipo_transaccion"] = tipo
    largos = [c for c, m in LARGOS.items() if r[c] and len(r[c]) > m]
    if largos:
        raise ValueError(f"texto demasiado largo en {', '.join(largos)}")
    return r


def _decimal_exacto(valor, columna):
    texto = a_texto(valor).replace(",", "")
    if not texto:
        return Decimal("0")
    try:
        return Decimal(texto)
    except Exception:
        raise ValueError(f"{COLUMNAS_EXCEL[columna]} inválido '{valor}'")


def es_anulada(r):
    return (r["comentarios"] or "").strip().upper() == "ANULADO"


def procesar(lineas, incluir_traslados):
    documentos, errores, advertencias = {}, [], []
    anuladas, traslados, lineas_traslado = set(), set(), 0

    for l in lineas:
        try:
            r = convertir_linea(l)
        except ValueError as e:
            clave = (a_texto(l["id_transaccion"]), a_texto(l["tipo_transaccion"]), a_texto(l["bodega"]), a_texto(l["fecha"])[:10])
            documentos.setdefault(clave, {"lineas": [], "errores": []})["errores"].append(f"fila {l['fila']}: {e}")
            continue

        ref = f"Transacción {r['id_transaccion']} ({r['tipo_transaccion']}, {r['bodega']}) fila {l['fila']}"
        if es_anulada(r):
            anuladas.add(f"{r['id_documento_origen'] or r['id_transaccion']} ({r['bodega']})")
            if Decimal(r["cantidad"]) != 0:
                advertencias.append(f"{ref}: marcada ANULADO pero con cantidad {r['cantidad']}; no se carga")
            continue
        if r["tipo_transaccion"] in TIPOS_TRASLADO and not incluir_traslados:
            traslados.add((r["id_transaccion"], r["tipo_transaccion"]))
            lineas_traslado += 1
            continue

        esperado = "+" if r["tipo_transaccion"] in TIPOS_ENTRADA else "-"
        if r["tipo_operacion"] != esperado:
            advertencias.append(f"{ref}: Tipo Operacion '{r['tipo_operacion']}' no corresponde a {r['tipo_transaccion']}")
        calculado = r.pop("_calculado")
        if r.pop("_dif_total") > TOLERANCIA_MONTO:
            advertencias.append(f"{ref}: Cantidad × Costo ({calculado:,.2f}) no coincide con Total MN ({Decimal(r['total_mn']):,.2f})")

        clave = (str(r["id_transaccion"]), r["tipo_transaccion"], r["bodega"], r["fecha"])
        documentos.setdefault(clave, {"lineas": [], "errores": []})["lineas"].append(r)

    validos = []
    for (id_tx, tipo, bodega, fecha), doc in documentos.items():
        if doc["errores"]:
            errores.append(f"Transacción {id_tx} ({tipo or '?'}, {bodega}) no se carga: " + "; ".join(doc["errores"]))
            continue
        if not doc["lineas"]:
            continue
        lin = doc["lineas"]
        validos.append({
            "id_transaccion": int(id_tx), "tipo": tipo, "bodega": bodega, "fecha": fecha,
            "documento": " ".join(filter(None, [lin[0]["tipo_documento_origen"], lin[0]["id_documento_origen"]
                                                if lin[0]["id_documento_origen"] != "0" else None])) or lin[0]["referencia"],
            "tercero": lin[0]["cliente_proveedor"],
            "total_mn": sum((Decimal(r["total_mn"]) for r in lin), Decimal("0")),
            "lineas": lin,
        })

    duplicados = []
    if validos:
        validos, duplicados, posibles = separar_duplicados(validos)
        advertencias.extend(posibles)
        advertencias.extend(comparar_con_ventas(validos))
    return validos, errores, advertencias, duplicados, sorted(anuladas), (len(traslados), lineas_traslado)


def separar_duplicados(documentos):
    desde = (min(date.fromisoformat(d["fecha"]) for d in documentos) - timedelta(days=DIAS_POSIBLE_DUPLICADO)).isoformat()
    ids = sorted({d["id_transaccion"] for d in documentos})
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            f"SELECT DISTINCT id_transaccion, tipo_transaccion, bodega, fecha FROM transacciones_inventario "
            f"WHERE fecha >= %s AND id_transaccion IN ({', '.join(['%s'] * len(ids))})",
            [desde] + ids,
        )
        existentes = {}
        for id_tx, tipo, bodega, fecha in cursor.fetchall():
            existentes.setdefault((id_tx, tipo, bodega), set()).add(fecha.isoformat())
    finally:
        cursor.close()
        conn.close()

    nuevos, duplicados, posibles = [], [], []
    for d in documentos:
        fechas_bd = existentes.get((d["id_transaccion"], d["tipo"], d["bodega"]))
        if not fechas_bd:
            nuevos.append(d)
        elif d["fecha"] in fechas_bd:
            d["motivo"] = "ya existe"
            duplicados.append(d)
        else:
            d["motivo"] = f"ya existe con fecha {', '.join(sorted(fechas_bd))}"
            duplicados.append(d)
            posibles.append(f"Transacción {d['id_transaccion']} ({d['tipo']}, {d['bodega']}): ya está en la base con fecha {', '.join(sorted(fechas_bd))}; no se carga")
    return nuevos, duplicados, posibles


def comparar_con_ventas(documentos):
    """El costo de las salidas por factura debe coincidir con el Costo Total de ventas_historicas."""
    salidas = {}
    for d in documentos:
        for r in d["lineas"]:
            if r["tipo_transaccion"] == "SALIDA" and (r["tipo_documento_origen"] or "").startswith("FACTURA"):
                clave = (r["id_documento_origen"], r["bodega"], r["fecha"])
                salidas[clave] = salidas.get(clave, Decimal("0")) + Decimal(r["total_mn"])
    if not salidas:
        return []
    fechas = sorted({f for _, _, f in salidas})
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT factura, Bodega, Fecha, `Costo Total` FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s",
                       (fechas[0], fechas[-1]))
        costos = {(f, b, fe.isoformat()): c for f, b, fe, c in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    fechas_con_ventas = {f for _, _, f in costos}
    avisos = []
    for (factura, bodega, fecha), costo in sorted(salidas.items()):
        if fecha not in fechas_con_ventas:
            continue  # ventas_historicas de ese día aún no se ha cargado
        costo_vh = costos.get((factura, bodega, fecha))
        if costo_vh is None:
            avisos.append(f"Factura {factura} ({bodega}, {fecha}): tiene salida de inventario pero no está en ventas_historicas")
        elif abs(costo - costo_vh) > Decimal("1.00"):
            avisos.append(f"Factura {factura} ({bodega}): costo de salida {costo:,.2f} no coincide con Costo Total de ventas_historicas {costo_vh:,.2f}")
    return avisos


def resumen_por_fecha(documentos):
    resumen = {}
    for d in documentos:
        fila = resumen.setdefault(d["fecha"], {"fecha": d["fecha"], "documentos": 0, "lineas": 0,
                                               "salidas": Decimal("0"), "entradas": Decimal("0")})
        fila["documentos"] += 1
        fila["lineas"] += len(d["lineas"])
        for r in d["lineas"]:
            fila["entradas" if r["tipo_transaccion"] in TIPOS_ENTRADA else "salidas"] += Decimal(r["total_mn"])
    return [resumen[k] for k in sorted(resumen)]


# -----------------------------
# Vistas
# -----------------------------
@cargar_inventario_bp.route("/cargar_inventario", methods=["GET", "POST"])
@requiere_acceso("cargar_inventario")
def cargar_inventario():
    if request.method == "GET":
        return render_template("cargar_inventario.html")

    incluir_traslados = request.form.get("incluir_traslados") == "on"
    archivo = request.files.get("archivo_inventario")
    if not archivo or not archivo.filename:
        return render_template("cargar_inventario.html", errores_archivo=["Falta el Reporte de Transacciones de Inventario."])
    try:
        lineas = leer_reporte_inventario(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_inventario.html",
                               errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_inventario.html",
                               errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    try:
        nuevos, errores, advertencias, duplicados, anuladas, n_traslados = procesar(lineas, incluir_traslados)
    except Exception as e:
        return render_template("cargar_inventario.html", errores_archivo=[f"Error al validar contra la base de datos: {e}"])

    token = None
    if nuevos:
        limpiar_temporales()
        token = uuid.uuid4().hex
        with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
            json.dump([r for d in nuevos for r in d["lineas"]], f, ensure_ascii=False)
        session["carga_inventario_token"] = token

    return render_template(
        "cargar_inventario.html",
        validado=True,
        token=token,
        incluir_traslados=incluir_traslados,
        total_lineas=len(lineas),
        nuevos=nuevos,
        lineas_nuevas=sum(len(d["lineas"]) for d in nuevos),
        resumen=resumen_por_fecha(nuevos),
        errores=errores,
        advertencias=advertencias,
        duplicados=duplicados,
        anuladas=anuladas,
        n_traslados=n_traslados,
    )


@cargar_inventario_bp.route("/cargar_inventario/confirmar", methods=["POST"])
@requiere_acceso("cargar_inventario")
def confirmar_carga_inventario():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_inventario_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_inventario.cargar_inventario"))

    with open(ruta, encoding="utf-8") as f:
        lineas = json.load(f)

    columnas = ", ".join(f"`{c}`" for c in COLUMNAS_SQL)
    sql = f"INSERT INTO transacciones_inventario ({columnas}) VALUES ({', '.join(['%s'] * len(COLUMNAS_SQL))})"

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Revalidar por si alguien cargó las mismas transacciones entre la validación y la confirmación
        ids = sorted({l["id_transaccion"] for l in lineas})
        fechas = sorted({l["fecha"] for l in lineas})
        cursor.execute(
            f"SELECT DISTINCT id_transaccion, tipo_transaccion, bodega, fecha FROM transacciones_inventario "
            f"WHERE fecha BETWEEN %s AND %s AND id_transaccion IN ({', '.join(['%s'] * len(ids))})",
            [fechas[0], fechas[-1]] + ids,
        )
        ya_existen = {(i, t, b, f.isoformat()) for i, t, b, f in cursor.fetchall()}
        pendientes = [l for l in lineas
                      if (l["id_transaccion"], l["tipo_transaccion"], l["bodega"], l["fecha"]) not in ya_existen]

        if pendientes:
            cursor.executemany(sql, [tuple(l[c] for c in COLUMNAS_SQL) for l in pendientes])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_inventario.cargar_inventario"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_inventario_token", None)

    documentos = {}
    for l in pendientes:
        documentos.setdefault((l["id_transaccion"], l["tipo_transaccion"], l["bodega"], l["fecha"]),
                              {"fecha": l["fecha"], "lineas": []})["lineas"].append(l)

    mensaje = f"✅ Se guardaron {len(pendientes)} líneas ({len(documentos)} transacciones) en transacciones_inventario."
    if ya_existen:
        mensaje += f" Se omitieron {len(ya_existen)} transacciones que ya existían."
    flash(mensaje, "success")
    return render_template("cargar_inventario.html", cargado=True, resumen=resumen_por_fecha(documentos.values()))
