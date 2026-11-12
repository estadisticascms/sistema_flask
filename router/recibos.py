from datetime import date, datetime, timedelta

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


DIAS_SEMANA = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]
TC_RECIBOS = 36.6243


def ultima_fecha_recibos():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MAX(fecharc) FROM reciboscaja")
        v = cursor.fetchone()[0]
        return v.date() if isinstance(v, datetime) else v
    finally:
        cursor.close()
        conn.close()


def recibos_por_dia(desde, hasta, tipo_cambio=TC_RECIBOS):
    """Recibos de caja por día: cantidad, subtotal (total ÷ 1.15), IVA 15%, total C$ y subtotal US$.
    La usan el reporte de recibos y el de ventas y recibos por día, así los números son los mismos."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT fecharc AS Fecha, COUNT(numerorc) AS Cantidad, SUM(total) AS TotalCordobas "
                       "FROM reciboscaja WHERE fecharc BETWEEN %s AND %s GROUP BY fecharc ORDER BY fecharc", (desde, hasta))
        rows = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    tipo_cambio = float(tipo_cambio) or TC_RECIBOS
    for r in rows:
        fecha = r["Fecha"].date() if isinstance(r["Fecha"], datetime) else r["Fecha"]
        r["Fecha"], r["Dia"] = fecha, DIAS_SEMANA[fecha.weekday()]
        subtotal = float(r["TotalCordobas"] or 0) / 1.15
        r["SubtotalCordobas"] = subtotal
        r["IVACordobas"] = subtotal * 0.15
        r["TotalCordobasCalc"] = subtotal * 1.15
        r["SubtotalUSD"] = subtotal / tipo_cambio
    total = {"Cantidad": sum(r["Cantidad"] for r in rows),
             **{k: sum(r[k] for r in rows) for k in ("SubtotalCordobas", "IVACordobas", "TotalCordobasCalc", "SubtotalUSD")}}
    total["dias"] = len(rows)
    total["promedio_dia"] = total["TotalCordobasCalc"] / len(rows) if rows else 0
    total["promedio_recibo"] = total["TotalCordobasCalc"] / total["Cantidad"] if total["Cantidad"] else 0
    total["max_dia"] = max((r["TotalCordobasCalc"] for r in rows), default=0)
    mejor = max(rows, key=lambda r: r["TotalCordobasCalc"], default=None)
    total["mejor"] = mejor
    return {"rows": rows, "total": total}


def _fecha(valor):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return None


@recibos_bp.route("/reporte_recibos", methods=["GET"])
@requiere_acceso("reporte_recibos")
def reporte_recibos():
    desde, hasta = _fecha(request.args.get("fecha_inicio")), _fecha(request.args.get("fecha_fin"))
    if not desde and not hasta:     # por defecto: el mes de la última fecha con recibos cargados
        hasta = ultima_fecha_recibos() or date.today()
        desde = hasta.replace(day=1)
    desde, hasta = desde or hasta, hasta or desde
    if hasta < desde:
        desde, hasta = hasta, desde
    try:
        tipo_cambio = float(str(request.args.get("tipo_cambio") or TC_RECIBOS).replace(",", ""))
        tipo_cambio = tipo_cambio if tipo_cambio > 0 else TC_RECIBOS
    except ValueError:
        tipo_cambio = TC_RECIBOS
    r = recibos_por_dia(desde, hasta, tipo_cambio)
    ultima = ultima_fecha_recibos() or date.today()
    rangos = [("Último día", ultima, ultima), ("Últimos 7 días", ultima - timedelta(days=6), ultima),
              ("Mes en curso", ultima.replace(day=1), ultima),
              ("Mes anterior", (ultima.replace(day=1) - timedelta(days=1)).replace(day=1), ultima.replace(day=1) - timedelta(days=1))]
    return render_template("reporte_recibos.html", rows=r["rows"], total=r["total"], desde=desde, hasta=hasta,
                           tipo_cambio=tipo_cambio, rangos=rangos,
                           chart_labels=[f"{x['Dia']} {x['Fecha']:%d/%m}" for x in r["rows"]],
                           chart_values=[round(x["TotalCordobasCalc"], 2) for x in r["rows"]])
