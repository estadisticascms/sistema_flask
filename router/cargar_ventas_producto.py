import json
import os
import uuid
from datetime import date, timedelta
from decimal import Decimal

from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import (
    ErrorArchivo, a_texto, a_decimal, a_fecha, a_factura, leer_filas,
    ruta_temporal, limpiar_temporales, TOLERANCIA_MONTO, DIAS_POSIBLE_DUPLICADO,
)

cargar_ventas_producto_bp = Blueprint("cargar_ventas_producto", __name__)

PREFIJO_TEMP = "carga_productos_"
ESTATUS_VALIDOS = {"Pagado/Cerrado", "Pendiente", "Anulado"}

COLUMNAS_SQL = [
    "row_names", "No. Factura", "Tipo Factura", "Fecha", "Cod. CLiente", "Cliente", "Vendedor",
    "Facturador", "Contado/Credito", "Plazo Pago", "Codigo", "Numero_parte", "Num. Parte",
    "Descrip Prod.", "Marca", "Unidad Medida", "Cantidad Vendida", "Costo Unit.", "Precio Facturado",
    "Sub-Total", "% Desc", "Descuento", "Sub-Total 2", "IVA", "Total", "Bodega",
    "Proveedor Ult. Compra", "Estatus",
]
COLUMNAS_TEXTO = ["Tipo Factura", "Cliente", "Vendedor", "Facturador", "Contado/Credito", "Codigo",
                  "Numero_parte", "Num. Parte", "Descrip Prod.", "Marca", "Bodega", "Proveedor Ult. Compra"]
COLUMNAS_DOUBLE = ["Cantidad Vendida", "Costo Unit.", "Precio Facturado", "Sub-Total", "% Desc"]
COLUMNAS_DECIMAL = ["Descuento", "Sub-Total 2", "IVA", "Total"]


def a_float(valor, obligatorio=True):
    texto = a_texto(valor).replace(",", "")
    if texto == "":
        if obligatorio:
            raise ValueError("valor vacío")
        return None
    try:
        return float(texto)
    except ValueError:
        raise ValueError(f"número inválido '{valor}'")


# -----------------------------
# Lectura del archivo
# -----------------------------
def leer_reporte_productos(archivo):
    filas = leer_filas(archivo)
    requeridas = ["No. Factura", "Fecha", "Cliente", "Descrip Prod.", "Cantidad Vendida", "Total", "Bodega", "Estatus"]
    inicio = None
    for i, fila in enumerate(filas[:40]):
        if all(r in [a_texto(c) for c in fila] for r in requeridas):
            inicio = i
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas " + ", ".join(requeridas))

    nombres = [a_texto(c) for c in filas[inicio]]
    while nombres and not nombres[-1]:
        nombres.pop()
    datos = [f for f in filas[inicio + 1:] if any(a_texto(c) for c in f)]

    # El sistema exporta el encabezado "Num. Parte" sin datos: las columnas quedan corridas una posición
    ultima_usada = any(len(f) >= len(nombres) and a_texto(f[len(nombres) - 1]) for f in datos)
    if "Num. Parte" in nombres and not ultima_usada:
        nombres.remove("Num. Parte")
    # "Cod. Prod" viene dos veces: el primero es el código interno y el segundo el número de parte
    cod = [i for i, n in enumerate(nombres) if n == "Cod. Prod"]
    if len(cod) == 2:
        nombres[cod[0]], nombres[cod[1]] = "Codigo", "Numero_parte"

    faltan = [c for c in COLUMNAS_SQL if c not in ("row_names", "Num. Parte") and c not in nombres]
    if faltan:
        raise ErrorArchivo("faltan las columnas " + ", ".join(faltan))
    pos = {n: i for i, n in enumerate(nombres)}

    lineas = []
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        if not any(a_texto(c) for c in fila):
            continue
        linea = {n: (fila[i] if i < len(fila) else None) for n, i in pos.items()}
        linea["fila"] = num
        lineas.append(linea)
    if not lineas:
        raise ErrorArchivo("no contiene líneas de venta")

    # Si la columna Estatus no trae estados conocidos, las columnas están desalineadas
    conocidos = sum(1 for l in lineas if a_texto(l["Estatus"]) in ESTATUS_VALIDOS)
    if conocidos < len(lineas) * 0.9:
        raise ErrorArchivo("las columnas no están en el orden esperado (la columna Estatus no trae Pagado/Cerrado, Pendiente o Anulado)")
    return lineas


# -----------------------------
# Validación
# -----------------------------
def convertir_linea(l):
    factura = a_factura(l["No. Factura"])
    if not factura.isdigit():
        raise ValueError(f"No. Factura inválido '{a_texto(l['No. Factura'])}'")
    fecha = a_fecha(l["Fecha"])
    if fecha is None:
        raise ValueError(f"fecha inválida '{a_texto(l['Fecha'])}'")
    if not a_texto(l["Bodega"]):
        raise ValueError("no tiene bodega")
    estatus = a_texto(l["Estatus"])
    if not estatus:
        raise ValueError("no tiene Estatus")

    registro = {
        "row_names": None,
        "No. Factura": int(factura),
        "Fecha": fecha.isoformat(),
        "Cod. CLiente": a_float(l["Cod. CLiente"], obligatorio=False),
        "Plazo Pago": a_float(l["Plazo Pago"], obligatorio=False),
        "Unidad Medida": 0,  # la columna es numérica; el reporte trae texto (UNI, PIE, ...) como el histórico
        "Estatus": estatus,
    }
    for c in COLUMNAS_TEXTO:
        registro[c] = a_texto(l.get(c)) or None
    for c in COLUMNAS_DOUBLE:
        registro[c] = a_float(l[c])
    for c in COLUMNAS_DECIMAL:
        registro[c] = str(a_decimal(l[c]))
    return registro


def procesar(lineas):
    facturas, errores, advertencias = {}, [], []

    for l in lineas:
        clave = (a_factura(l["No. Factura"]), a_texto(l["Bodega"]), a_texto(l["Fecha"])[:10])
        fac = facturas.setdefault(clave, {"lineas": [], "errores": []})
        try:
            r = convertir_linea(l)
        except ValueError as e:
            fac["errores"].append(f"fila {l['fila']}: {e}")
            continue
        if r["Estatus"] not in ESTATUS_VALIDOS:
            advertencias.append(f"Factura {r['No. Factura']} fila {l['fila']}: Estatus poco común '{r['Estatus']}'")
        total_calc = Decimal(r["Sub-Total 2"]) + Decimal(r["IVA"])
        if abs(total_calc - Decimal(r["Total"])) > TOLERANCIA_MONTO:
            advertencias.append(f"Factura {r['No. Factura']} fila {l['fila']}: Sub-Total 2 + IVA ({total_calc}) no coincide con Total ({r['Total']})")
        fac["lineas"].append(r)

    validas = []
    for (factura, bodega, _), fac in facturas.items():
        if fac["errores"]:
            errores.append(f"Factura {factura} ({bodega}) no se carga: " + "; ".join(fac["errores"]))
            continue
        lin = fac["lineas"]
        estados = {r["Estatus"] for r in lin}
        if len({r["Fecha"] for r in lin}) > 1:
            errores.append(f"Factura {factura} ({bodega}) no se carga: sus líneas tienen fechas distintas")
            continue
        validas.append({
            "factura": int(factura),
            "bodega": bodega,
            "fecha": lin[0]["Fecha"],
            "cliente": lin[0]["Cliente"],
            "tipo": lin[0]["Contado/Credito"],
            "estatus": ", ".join(sorted(estados)),
            "anulada": estados == {"Anulado"},
            "total": sum((Decimal(r["Total"]) for r in lin), Decimal("0")),
            "lineas": lin,
        })

    duplicadas = []
    if validas:
        validas, duplicadas, posibles = separar_duplicadas(validas)
        advertencias.extend(posibles)
        advertencias.extend(comparar_con_ventas_historicas(validas))
    return validas, errores, advertencias, duplicadas


def separar_duplicadas(facturas):
    fechas = [date.fromisoformat(f["fecha"]) for f in facturas]
    desde = (min(fechas) - timedelta(days=DIAS_POSIBLE_DUPLICADO)).isoformat()
    numeros = sorted({f["factura"] for f in facturas})
    marcas = ", ".join(["%s"] * len(numeros))

    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            f"SELECT DISTINCT `No. Factura`, Bodega, Fecha FROM ventas_por_producto "
            f"WHERE Fecha >= %s AND `No. Factura` IN ({marcas})",
            [desde] + numeros,
        )
        existentes = {}
        for num, bodega, fecha in cursor.fetchall():
            existentes.setdefault((int(num), bodega), set()).add(str(fecha)[:10])
    finally:
        cursor.close()
        conn.close()

    nuevas, duplicadas, posibles = [], [], []
    for f in facturas:
        fechas_bd = existentes.get((f["factura"], f["bodega"]))
        if not fechas_bd:
            nuevas.append(f)
        elif f["fecha"] in fechas_bd:
            f["motivo"] = "ya existe"
            duplicadas.append(f)
        else:
            f["motivo"] = f"ya existe con fecha {', '.join(sorted(fechas_bd))}"
            duplicadas.append(f)
            posibles.append(f"Factura {f['factura']} ({f['bodega']}): ya está en la base con fecha {', '.join(sorted(fechas_bd))}; no se carga")
    return nuevas, duplicadas, posibles


def comparar_con_ventas_historicas(facturas):
    """Compara el total por factura contra ventas_historicas (si esas facturas ya se cargaron)."""
    vigentes = [f for f in facturas if not f["anulada"]]
    if not vigentes:
        return []
    numeros = sorted({str(f["factura"]) for f in vigentes})
    fechas = sorted({f["fecha"] for f in vigentes})
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            f"SELECT factura, Bodega, Fecha, Total FROM ventas_historicas "
            f"WHERE Fecha BETWEEN %s AND %s AND factura IN ({', '.join(['%s'] * len(numeros))})",
            [fechas[0], fechas[-1]] + numeros,
        )
        totales = {(str(n), b, str(fe)): t for n, b, fe, t in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    avisos = []
    for f in vigentes:
        lineas_vigentes = [r for r in f["lineas"] if r["Estatus"] != "Anulado"]
        total = sum((Decimal(r["Total"]) for r in lineas_vigentes), Decimal("0"))
        total_vh = totales.get((str(f["factura"]), f["bodega"], f["fecha"]))
        if total_vh is not None and abs(total - total_vh) > TOLERANCIA_MONTO:
            avisos.append(f"Factura {f['factura']} ({f['bodega']}): la suma por producto ({total:,.2f}) no coincide con ventas_historicas ({total_vh:,.2f})"
                          + (" — es factura mixta, puede incluir servicios" if (f["lineas"][0]["Tipo Factura"] or "").endswith("Mixta") else ""))
    return avisos


def resumen_por_fecha(facturas):
    resumen = {}
    for f in facturas:
        fila = resumen.setdefault(f["fecha"], {"fecha": f["fecha"], "facturas": 0, "anuladas": 0, "lineas": 0, "total": Decimal("0")})
        fila["facturas"] += 1
        fila["lineas"] += len(f["lineas"])
        if f["anulada"]:
            fila["anuladas"] += 1
        else:
            fila["total"] += sum((Decimal(r["Total"]) for r in f["lineas"] if r["Estatus"] != "Anulado"), Decimal("0"))
    return [resumen[k] for k in sorted(resumen)]


# -----------------------------
# Vistas
# -----------------------------
@cargar_ventas_producto_bp.route("/cargar_ventas_producto", methods=["GET", "POST"])
@requiere_acceso("cargar_ventas_producto")
def cargar_ventas_producto():
    if request.method == "GET":
        return render_template("cargar_ventas_producto.html")

    archivo = request.files.get("archivo_productos")
    if not archivo or not archivo.filename:
        return render_template("cargar_ventas_producto.html", errores_archivo=["Falta el Reporte de Ventas por Producto."])
    try:
        lineas = leer_reporte_productos(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_ventas_producto.html",
                               errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_ventas_producto.html",
                               errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    try:
        nuevas, errores, advertencias, duplicadas = procesar(lineas)
    except Exception as e:
        return render_template("cargar_ventas_producto.html",
                               errores_archivo=[f"Error al validar contra la base de datos: {e}"])

    token = None
    if nuevas:
        limpiar_temporales()
        token = uuid.uuid4().hex
        with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
            json.dump([r for fac in nuevas for r in fac["lineas"]], f, ensure_ascii=False)
        session["carga_productos_token"] = token

    return render_template(
        "cargar_ventas_producto.html",
        validado=True,
        token=token,
        total_lineas=len(lineas),
        nuevas=nuevas,
        lineas_nuevas=sum(len(f["lineas"]) for f in nuevas),
        resumen=resumen_por_fecha(nuevas),
        errores=errores,
        advertencias=advertencias,
        duplicadas=duplicadas,
    )


@cargar_ventas_producto_bp.route("/cargar_ventas_producto/confirmar", methods=["POST"])
@requiere_acceso("cargar_ventas_producto")
def confirmar_carga_ventas_producto():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_productos_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_ventas_producto.cargar_ventas_producto"))

    with open(ruta, encoding="utf-8") as f:
        lineas = json.load(f)

    columnas = ", ".join(f"`{c}`" for c in COLUMNAS_SQL)
    sql = f"INSERT INTO ventas_por_producto ({columnas}) VALUES ({', '.join(['%s'] * len(COLUMNAS_SQL))})"

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Revalidar por si alguien cargó las mismas facturas entre la validación y la confirmación
        numeros = sorted({l["No. Factura"] for l in lineas})
        fechas = sorted({l["Fecha"] for l in lineas})
        cursor.execute(
            f"SELECT DISTINCT `No. Factura`, Bodega, Fecha FROM ventas_por_producto "
            f"WHERE Fecha BETWEEN %s AND %s AND `No. Factura` IN ({', '.join(['%s'] * len(numeros))})",
            [fechas[0], fechas[-1]] + numeros,
        )
        ya_existen = {(int(n), b, str(fe)[:10]) for n, b, fe in cursor.fetchall()}
        pendientes = [l for l in lineas if (l["No. Factura"], l["Bodega"], l["Fecha"]) not in ya_existen]

        if pendientes:
            cursor.executemany(sql, [tuple(l[c] for c in COLUMNAS_SQL) for l in pendientes])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_ventas_producto.cargar_ventas_producto"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_productos_token", None)

    facturas = {}
    for l in pendientes:
        f = facturas.setdefault((l["No. Factura"], l["Bodega"], l["Fecha"]),
                                {"fecha": l["Fecha"], "lineas": [], "anulada": True})
        f["lineas"].append(l)
        f["anulada"] = f["anulada"] and l["Estatus"] == "Anulado"

    mensaje = f"✅ Se guardaron {len(pendientes)} líneas ({len(facturas)} facturas) en ventas_por_producto."
    if ya_existen:
        mensaje += f" Se omitieron {len(ya_existen)} facturas que ya existían."
    flash(mensaje, "success")
    return render_template("cargar_ventas_producto.html", cargado=True, resumen=resumen_por_fecha(facturas.values()))
