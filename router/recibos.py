from flask import Blueprint, request, render_template, redirect, url_for, flash
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, a_decimal, a_fecha, leer_filas, buscar_encabezado
from router import carga_simple

recibos_bp = Blueprint("recibos", __name__)

# Reporte "Recibos de Caja" del sistema -> tabla reciboscaja (clave: numerorc)
COLUMNAS_RC = {"Fecha Recibo": "fecharc", "ROC #": "numerorc", "Beneficiario": "cliente",
               "Concepto": "concepto", "Monto Cordobas": "total"}
CAMPOS_RC = ["fecharc", "numerorc", "cliente", "concepto", "total"]
NOMBRES_RC = {v: k for k, v in COLUMNAS_RC.items()}
PREFIJO_RC = "carga_recibos_"
MAX_DETALLE = 300


def leer_recibos(archivo):
    filas = leer_filas(archivo)
    try:
        inicio, pos = buscar_encabezado(filas, list(COLUMNAS_RC))
    except ErrorArchivo:
        raise ErrorArchivo("no se encontraron las columnas del reporte de Recibos de Caja "
                           "(Fecha Recibo, ROC #, Beneficiario, Concepto, Monto Cordobas)")
    pos_usd = pos.get("Monto Dolares")

    def celda(fila, j):
        return fila[j] if j is not None and j < len(fila) else None

    recibos, errores, advertencias, vistos = [], [], [], {}
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        numero = a_texto(celda(fila, pos["ROC #"]))
        if not numero:
            continue   # filas vacías o de totales
        r = {"numerorc": numero, "cliente": a_texto(celda(fila, pos["Beneficiario"])) or None,
             "concepto": a_texto(celda(fila, pos["Concepto"])) or None,
             "fecharc": a_fecha(celda(fila, pos["Fecha Recibo"]))}
        if r["fecharc"] is None:
            errores.append(f"Fila {num}: ROC {numero} sin fecha válida ('{a_texto(celda(fila, pos['Fecha Recibo']))}'); no se carga")
            continue
        try:
            r["total"] = a_decimal(celda(fila, pos["Monto Cordobas"]))
        except ValueError as e:
            errores.append(f"Fila {num}: ROC {numero}: {e}; no se carga")
            continue
        if len(numero) > 10:
            errores.append(f"Fila {num}: ROC '{numero}' tiene más de 10 caracteres; no se carga")
            continue
        if r["cliente"] and len(r["cliente"]) > 100:
            advertencias.append(f"Fila {num}: ROC {numero}: beneficiario de más de 100 caracteres; se recorta")
            r["cliente"] = r["cliente"][:100]
        try:
            usd = a_decimal(celda(fila, pos_usd)) if pos_usd is not None else 0
        except ValueError:
            usd = 0
        if usd:
            advertencias.append(f"Fila {num}: ROC {numero} trae U$ {usd:,.2f} en Monto Dolares; "
                                f"la tabla solo guarda el monto en córdobas (C$ {r['total']:,.2f})")
        if r["total"] <= 0:
            advertencias.append(f"Fila {num}: ROC {numero} con monto C$ {r['total']:,.2f}")
        if numero in vistos:
            if vistos[numero] != r:
                errores.append(f"Fila {num}: ROC {numero} repetido con datos distintos; se conserva la primera fila")
            continue
        vistos[numero] = dict(r)
        recibos.append(r)
    if not recibos:
        raise ErrorArchivo("no contiene recibos")
    return recibos, errores, advertencias


def faltantes(recibos):
    """ROC consecutivos que faltan dentro del rango del archivo y tampoco están en la base de datos
    (anulados o un reporte exportado incompleto)."""
    numeros = sorted(int(r["numerorc"]) for r in recibos if r["numerorc"].isdigit())
    if len(numeros) < 2 or numeros[-1] - numeros[0] > 5000:
        return []
    presentes = set(numeros)
    candidatos = [str(n) for n in range(numeros[0], numeros[-1] + 1) if n not in presentes]
    if not candidatos:
        return []
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"SELECT numerorc FROM reciboscaja WHERE numerorc IN ({', '.join(['%s'] * len(candidatos))})", candidatos)
        en_bd = {r[0] for r in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()
    return [c for c in candidatos if c not in en_bd]


def ultimos_recibos():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT fecharc, numerorc, cliente, concepto, total FROM reciboscaja "
                       "ORDER BY fecharc DESC, CAST(numerorc AS UNSIGNED) DESC LIMIT 10")
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()


@recibos_bp.route("/cargar_recibos", methods=["GET", "POST"])
@requiere_acceso("cargar_recibos")
def cargar_recibos():
    if request.method == "GET":
        return render_template("cargar_recibos.html", ultimos=ultimos_recibos())

    archivo = request.files.get("excel_file")
    if not archivo or not archivo.filename:
        return render_template("cargar_recibos.html", errores_archivo=["Falta el archivo de Recibos de Caja."])
    try:
        recibos, errores, advertencias = leer_recibos(archivo)
        dif = carga_simple.comparar("reciboscaja", "numerorc", CAMPOS_RC, {"total"}, recibos)
        huecos = faltantes(recibos)
    except ErrorArchivo as e:
        return render_template("cargar_recibos.html", errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_recibos.html", errores_archivo=[f"{archivo.filename}: error inesperado – {e}"])

    if huecos:
        muestra = ", ".join(huecos[:30]) + (f" … ({len(huecos)} en total)" if len(huecos) > 30 else "")
        advertencias.insert(0, f"Faltan ROC consecutivos que no están en el archivo ni en la base de datos: {muestra}. "
                               "Verifique si fueron anulados o si el reporte se exportó incompleto.")
    token = carga_simple.guardar_temporal(PREFIJO_RC, "carga_recibos_token", CAMPOS_RC, dif)
    fechas = [r["fecharc"] for r in recibos]
    return render_template("cargar_recibos.html", validado=True, token=token, dif=dif, total=len(recibos),
                           monto=sum(r["total"] for r in recibos), monto_nuevos=sum(r["total"] for r in dif["nuevos"]),
                           desde=min(fechas), hasta=max(fechas), nombres=NOMBRES_RC,
                           errores=errores, advertencias=advertencias, max_detalle=MAX_DETALLE)


@recibos_bp.route("/cargar_recibos/confirmar", methods=["POST"])
@requiere_acceso("cargar_recibos")
def confirmar_recibos():
    try:
        nuevos, modificados = carga_simple.confirmar("reciboscaja", "numerorc", CAMPOS_RC, PREFIJO_RC,
                                                     "carga_recibos_token", request.form.get("token"))
    except ValueError as e:
        flash(f"⚠️ {e}", "warning")
        return redirect(url_for("recibos.cargar_recibos"))
    except Exception as e:
        flash(f"❌ No se guardó ningún recibo. Error de base de datos: {e}", "danger")
        return redirect(url_for("recibos.cargar_recibos"))
    flash(f"✅ Recibos guardados: {nuevos:,} nuevos y {modificados:,} actualizados.", "success")
    return redirect(url_for("recibos.cargar_recibos"))


@recibos_bp.route("/reporte_recibos", methods=["GET"])
@requiere_acceso("reporte_recibos")
def reporte_recibos():
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    tipo_cambio = float(request.args.get("tipo_cambio", 36.6243))  # valor ingresado por el usuario

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT fecharc AS Fecha,
               COUNT(numerorc) AS Cantidad,
               SUM(total) AS TotalCordobas
        FROM reciboscaja
        WHERE fecharc BETWEEN %s AND %s
        GROUP BY fecharc
        ORDER BY fecharc;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    # Calcular Subtotal, IVA y Total en Córdobas + Subtotal en USD
    for r in rows:
        subtotal_cordobas = float(r["TotalCordobas"]) / 1.15
        iva_cordobas = subtotal_cordobas * 0.15
        total_cordobas = subtotal_cordobas + iva_cordobas

        subtotal_usd = subtotal_cordobas / tipo_cambio

        r["SubtotalCordobas"] = subtotal_cordobas
        r["IVACordobas"] = iva_cordobas
        r["TotalCordobasCalc"] = total_cordobas
        r["SubtotalUSD"] = subtotal_usd

    # Totales generales
    total_general_subtotal = sum(r["SubtotalCordobas"] for r in rows)
    total_general_iva = sum(r["IVACordobas"] for r in rows)
    total_general_cordobas = sum(r["TotalCordobasCalc"] for r in rows)
    total_general_usd = sum(r["SubtotalUSD"] for r in rows)

    labels = [str(r["Fecha"]) for r in rows]
    valores = [r["TotalCordobasCalc"] for r in rows]

    return render_template("reporte_recibos.html",
                           rows=rows,
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           tipo_cambio=tipo_cambio,
                           total_general_subtotal=total_general_subtotal,
                           total_general_iva=total_general_iva,
                           total_general_cordobas=total_general_cordobas,
                           total_general_usd=total_general_usd,
                           chart_labels=labels,
                           chart_values=valores)
