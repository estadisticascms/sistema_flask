from flask import Blueprint, request, render_template, redirect, url_for, flash
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, a_decimal, a_fecha, a_factura, leer_filas, buscar_encabezado
from router import carga_simple

devoluciones_bp = Blueprint("devoluciones", __name__)

# Reporte "Lista Devoluciones" del sistema -> tabla devoluciones (clave: numerodv)
COLUMNAS_DV = {
    "# Devolucion Interna": "numerodv", "# Ref. Devolucion": "referencia", "Fecha Devolucion": "fechadv",
    "# Factura": "factura", "Cred./Cont.": "tipofact", "Nombre de Cliente": "cliente", "Facturador": "facturador",
    "Vendedor Asignado": "vendedor", "Motivo Devolucion": "motivo", "Total": "totaldv",
    "Saldo Factura": "saldofact", "Estado": "estado",
}
CAMPOS_DV = list(COLUMNAS_DV.values())
NOMBRES_DV = {v: k for k, v in COLUMNAS_DV.items()}
LARGOS = {"numerodv": 15, "referencia": 15, "factura": 15, "tipofact": 45, "cliente": 100, "facturador": 100,
          "vendedor": 100, "motivo": 200, "estado": 15}
PREFIJO_DV = "carga_devoluciones_"
MAX_DETALLE = 300


def monto(valor):
    """Los montos vienen como texto 'C$ 2,954.52'."""
    texto = a_texto(valor)
    if texto.upper().startswith(("U$", "US$")):
        raise ValueError(f"monto en dólares '{texto}'")
    return a_decimal(texto)


def leer_devoluciones(archivo):
    filas = leer_filas(archivo)
    try:
        inicio, pos = buscar_encabezado(filas, list(COLUMNAS_DV))
    except ErrorArchivo:
        raise ErrorArchivo("no se encontraron las columnas de la Lista de Devoluciones "
                           "(# Devolucion Interna, Fecha Devolucion, # Factura, Total, Estado, ...)")

    def celda(fila, col):
        j = pos[col]
        return fila[j] if j < len(fila) else None

    registros, errores, advertencias, vistos = [], [], [], {}
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        numero = a_texto(celda(fila, "# Devolucion Interna"))
        if not numero:
            continue   # filas vacías o de totales al final del reporte
        r = {campo: (a_texto(celda(fila, col)) or None) for col, campo in COLUMNAS_DV.items()
             if campo not in ("fechadv", "totaldv", "saldofact")}
        r["factura"] = a_factura(r["factura"]) or None
        r["fechadv"] = a_fecha(celda(fila, "Fecha Devolucion"))
        if r["fechadv"] is None:
            errores.append(f"Fila {num}: devolución {numero} sin fecha válida; no se carga")
            continue
        try:
            r["totaldv"] = monto(celda(fila, "Total"))
            r["saldofact"] = monto(celda(fila, "Saldo Factura"))
        except ValueError as e:
            errores.append(f"Fila {num}: devolución {numero}: {e}; no se carga")
            continue
        if not r["vendedor"]:   # sin vendedor asignado: se toma el facturador (regla de siempre)
            r["vendedor"] = r["facturador"]
        if r["estado"] not in ("Cerrado", "Anulado"):
            advertencias.append(f"Fila {num}: devolución {numero} con estado '{r['estado'] or ''}' (normalmente Cerrado o Anulado)")
        if r["tipofact"] and r["tipofact"].upper() not in ("CREDITO", "CONTADO"):
            advertencias.append(f"Fila {num}: devolución {numero} con tipo '{r['tipofact']}' (se esperaba CREDITO o CONTADO)")
        demasiado_largo = [c for c, n in LARGOS.items() if r[c] and len(r[c]) > n]
        if "numerodv" in demasiado_largo:
            errores.append(f"Fila {num}: número de devolución '{numero}' demasiado largo; no se carga")
            continue
        for c in demasiado_largo:
            advertencias.append(f"Fila {num}: devolución {numero}: {NOMBRES_DV[c]} de más de {LARGOS[c]} caracteres; se recorta")
            r[c] = r[c][:LARGOS[c]]
        if numero in vistos:
            if vistos[numero] != r:
                errores.append(f"Fila {num}: devolución {numero} repetida con datos distintos; se conserva la primera fila")
            continue
        vistos[numero] = dict(r)
        registros.append(r)
    if not registros:
        raise ErrorArchivo("no contiene devoluciones")
    return registros, errores, advertencias


def revisar_facturas(registros, advertencias):
    """Avisa si la factura no está en las ventas cargadas o si la devolución supera la factura."""
    facturas = sorted({r["factura"] for r in registros if r["factura"]})
    if not facturas:
        return
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"SELECT factura, SUM(Total), MAX(cliente) FROM ventas_historicas "
                       f"WHERE factura IN ({', '.join(['%s'] * len(facturas))}) GROUP BY factura", facturas)
        ventas = {str(f): (total, cliente) for f, total, cliente in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()
    for r in registros:
        if r["estado"] == "Anulado" or not r["factura"]:
            continue
        v = ventas.get(r["factura"])
        if v is None:
            advertencias.append(f"Devolución {r['numerodv']}: la factura {r['factura']} no está en las ventas cargadas "
                                f"(¿falta cargar las ventas de ese día?)")
        elif v[0] is not None and r["totaldv"] > v[0] + 1:
            advertencias.append(f"Devolución {r['numerodv']}: C$ {r['totaldv']:,.2f} es mayor que el total de la "
                                f"factura {r['factura']} (C$ {v[0]:,.2f})")


def ultimas_devoluciones():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM devoluciones ORDER BY fechadv DESC, CAST(numerodv AS UNSIGNED) DESC LIMIT 10")
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()


@devoluciones_bp.route("/devoluciones", methods=["GET", "POST"])
@requiere_acceso("devoluciones")
def devoluciones():
    if request.method == "GET":
        return render_template("devoluciones.html", ultimas=ultimas_devoluciones())

    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return render_template("devoluciones.html", errores_archivo=["Falta el archivo Lista de Devoluciones."])
    try:
        registros, errores, advertencias = leer_devoluciones(archivo)
        revisar_facturas(registros, advertencias)
        dif = carga_simple.comparar("devoluciones", "numerodv", CAMPOS_DV, {"totaldv", "saldofact"}, registros)
    except ErrorArchivo as e:
        return render_template("devoluciones.html", errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("devoluciones.html", errores_archivo=[f"{archivo.filename}: error inesperado – {e}"])

    token = carga_simple.guardar_temporal(PREFIJO_DV, "carga_devoluciones_token", CAMPOS_DV, dif)
    fechas = [r["fechadv"] for r in registros]
    cerradas = [r for r in registros if r["estado"] != "Anulado"]
    return render_template("devoluciones.html", validado=True, token=token, dif=dif, total=len(registros),
                           monto=sum(r["totaldv"] for r in cerradas), anuladas=len(registros) - len(cerradas),
                           desde=min(fechas), hasta=max(fechas), nombres=NOMBRES_DV,
                           errores=errores, advertencias=advertencias, max_detalle=MAX_DETALLE)


@devoluciones_bp.route("/devoluciones/confirmar", methods=["POST"])
@requiere_acceso("devoluciones")
def confirmar_devoluciones():
    try:
        nuevos, modificados = carga_simple.confirmar("devoluciones", "numerodv", CAMPOS_DV, PREFIJO_DV,
                                                     "carga_devoluciones_token", request.form.get("token"))
    except ValueError as e:
        flash(f"⚠️ {e}", "warning")
        return redirect(url_for("devoluciones.devoluciones"))
    except Exception as e:
        flash(f"❌ No se guardó ninguna devolución. Error de base de datos: {e}", "danger")
        return redirect(url_for("devoluciones.devoluciones"))
    flash(f"✅ Devoluciones guardadas: {nuevos:,} nuevas y {modificados:,} actualizadas.", "success")
    return redirect(url_for("devoluciones.devoluciones"))
