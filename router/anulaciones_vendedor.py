"""Facturas anuladas por vendedor y motivo, a partir del reporte «FACTURAS ANULADAS» del sistema.

- Carga (/cargar_anuladas_sistema): CSV (separado por ;) o Excel tal como lo da el sistema. Se valida, se muestra el resumen
  y al confirmar se guarda en facturas_anuladas_sistema (docs/sql/facturas_anuladas_sistema.sql). Si una factura ya estaba
  (número + fecha + total), se actualiza con lo que diga el archivo.
- Vendedor: el vendedor asignado; si viene en blanco, el facturador.
- Motivo: se clasifica la razón escrita al anular (NHE = no hay existencia, PMF = producto mal facturado, refacturación, …).
  La razón original siempre se guarda y se muestra.
- Reporte (/anulaciones_vendedor): por día o por mes, cuántas facturas anula cada vendedor y por qué motivo; Excel."""
import csv
import io
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta

from flask import Blueprint, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import sin_tildes
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, limpiar_temporales, ruta_temporal

anulaciones_vendedor_bp = Blueprint("anulaciones_vendedor", __name__)

PREFIJO_TEMP = "carga_anuladas_sis_"
VERDE, NARANJA = "027266", "FE6412"

# Motivos, en el orden en que se revisan (gana el primero que coincida). (clave, nombre, color, patrón sobre la razón en mayúsculas sin tildes)
MOTIVOS = [
    # NHE = no hay existencia; PNEB = producto no encontrado en bodega (también «bodega no encontró», «solo hay N físicas»)
    ("NHE", "No hay existencia (NHE)", "#FE6412",
     r"N+HE|\bPNEB\b|SIN EXISTENCIA|NO (SE )?ENCONTR|ENCONTRAD|F\S?SIC|FALTANTE|INCOMPLET|SOLO \d+ UNI|SOLO HAY|REMISI\S*N|NO ALCAN|UNIDAD\w* NO EST"),
    ("PMF", "Producto mal facturado (PMF)", "#027266",
     r"\bP ?M ?F|PFM|MAL FACTURAD|ERROR DE PRODUCTO|ERROR DEL PRODUCTO|NO IBA\b|CRUCE DE"),
    ("MAL_ESTADO", "Producto en mal estado", "#9C6644",
     r"MAL ESTADO|\bPME\b|DETERIOR|MOJADO|PEDACERIA|QUEBRAD|PINTURA|SARRO"),
    ("DUPLICADA", "Factura o ítem duplicado", "#6D597A", r"DU\W?P\W?LIC|DULPIC|TRIPLIC"),
    ("REFACTURACION", "Refacturación", "#4F6D8A", r"REFACT|FACTURAR DE NUEVO|SALGA OTRA FACTURA"),
    ("IMPRESION", "Error de impresión", "#8AB17D",
     r"IMPRES|IMPRIM|CINTA|MOVIDA|CONSECUTIVO|SE DA.O LA FACTURA|FACTURA SE DA|SE ROMPIO"),
    ("CLIENTE", "Cliente canceló o pidió mal", "#D9A441",
     r"CANCEL|NO LO QUISO|NO QUISO|NO LO QUIERE|NO QUIERE|NO LO LLEVO|NO LLEVO|NO VINO|NO RECIBE|NO ACEPTO|SOLICIT\w* +MAL|"
     r"CLIENTE SE EQUIVOCO|SE EQUIVOCO CON PEDIDO|CLIENTE CAMBIO|NO ANDABA|NO PUDO VENIR|ERROR DEL CLIENTE|COTIZACION|PROFORMA|"
     r"CLIENTE PIDIO|CLIENTE REGRES|PEDIDO CANCELADO|CLIENTE QUER|CLIENTE SOLICITA$"),
    ("PRODUCTO", "No era el producto (medida, aplicación, marca)", "#5E8C61",
     r"NO ERA\b(?! (EL CLIENTE|AL CLIENTE))|NO ERA EL (REPUESTO|PRODUCTO)|DIFERENTE|NO CORRESPONDE|APLICACI|MEDIDA|MUESTRA|MARCA|NO ES (LO|LA|EL|TUERCA)|"
     r"METALICO|LISO|TUERCA|PERNOS|MAS GRANDE|DESCRIPCI|CAMBIO POR (UN|OTRO)"),
    ("DATOS", "Cambio de precio, cantidad o datos de la factura", "#A8A29E",
     r"PRECIO|CANTIDAD|CR\S?DITO|CONTADO|\bRUC\b|RAZON SOCIAL|EXO?NER|EXHONER|NOMBRE DEL CLIENTE|RETENCION|\bIVA\b|DESCUENTO|MONTO|POR FECHA"),
    ("ERROR_FACTURA", "Error al facturar o del sistema", "#C9ADA7",
     r"ERROR (DE|EN|DEL) FACT|ERRO EN FACT|ERROR DE FACT|VENDEDOR|EQUIVOCAD|NO ERA (EL|AL) CLIENTE|NO ES AL CLIENTE|CALCULO|NUMERACION|"
     r"FACTUR\S* EN 0|EN 0 EL SISTEMA|GRABAD|GRABO|SISTEMA|MAL INGRESO|INGRESO DESDE IMPORT|CAJA|DE B\d\b"),
    ("OTRO", "Otro motivo", "#D6D3CE", None),
]
SIN_RAZON = ("SIN_RAZON", "Sin razón registrada", "#ECEBE8")
NOMBRE_MOTIVO = {m[0]: m[1] for m in MOTIVOS + [SIN_RAZON]}
COLOR_MOTIVO = {m[0]: m[2] for m in MOTIVOS + [SIN_RAZON]}
ORDEN_MOTIVOS = [m[0] for m in MOTIVOS] + [SIN_RAZON[0]]
_PATRONES = [(clave, re.compile(patron)) for clave, _, _, patron in MOTIVOS if patron]


def limpiar_razon(razon):
    """Quita el prefijo «FACTURA ANULADA:» que el sistema le pone a todas."""
    texto = " ".join((razon or "").split())
    return re.sub(r"^FACTURA ANULADA\s*:\s*", "", texto, flags=re.I).strip()


def motivo_de(razon):
    texto = sin_tildes(limpiar_razon(razon)).upper()
    if not texto:
        return SIN_RAZON[0]
    for clave, patron in _PATRONES:
        if patron.search(texto):
            return clave
    return "OTRO"


def limpiar_nombre(nombre):
    """«MOSTRADOR . . .» -> «MOSTRADOR» (y «MOSTRADOR . MONTOYA .» -> «MOSTRADOR MONTOYA»): el mostrador queda como un vendedor más."""
    nombre = " ".join((nombre or "").split())
    if nombre.upper().startswith("MOSTRADOR"):
        nombre = " ".join(p for p in nombre.split() if p != ".")
    return nombre


def nombre_vendedor(asignado, facturador):
    return limpiar_nombre(asignado) or limpiar_nombre(facturador) or "Sin vendedor"


# -----------------------------
# Lectura del archivo del sistema
# -----------------------------
COLUMNAS = {   # encabezado normalizado -> campo
    "FACTURA": "factura", "NOFACTURA": "factura", "FECHA": "fecha", "CLIENTE": "cliente", "FACTURADOR": "facturador",
    "VENDEDORASIGNADO": "vendedor_asignado", "VENDEDOR": "vendedor_asignado", "CREDCONT": "tipo", "CREDITOCONTADO": "tipo",
    "DESCRIPCION": "descripcion", "TOTAL": "total", "RAZONANULADO": "razon", "RAZON": "razon", "ANULADOPOR": "anulado_por",
    "FECHAANULADO": "fecha_anulado",
}
REQUERIDAS = ["factura", "fecha", "total", "facturador", "vendedor_asignado", "razon"]


def normal_columna(valor):
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(a_texto(valor)).upper())


def filas_del_archivo(archivo):
    nombre = (archivo.filename or "").lower()
    if not nombre.endswith(".csv"):
        return leer_filas(archivo)
    contenido = archivo.read()
    if not contenido:
        raise ErrorArchivo("el archivo está vacío")
    for codificacion in ("utf-8-sig", "cp1252"):
        try:
            texto = contenido.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    primeras = texto[:3000]
    separador = ";" if primeras.count(";") >= primeras.count(",") else ","
    return list(csv.reader(io.StringIO(texto), delimiter=separador))


def a_fecha(valor, con_hora=False):
    if isinstance(valor, datetime):
        return valor if con_hora else valor.date()
    if isinstance(valor, date):
        return datetime.combine(valor, datetime.min.time()) if con_hora else valor
    texto = a_texto(valor)
    if not texto or texto.startswith("0000"):
        return None
    for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            d = datetime.strptime(texto[:19], formato)
            return d if con_hora else d.date()
        except ValueError:
            continue
    return None


def a_monto(valor):
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)
    texto = a_texto(valor).replace(",", "").replace("C$", "").strip()
    try:
        return round(float(texto), 2) if texto else None
    except ValueError:
        return None


def leer_anuladas(archivo):
    """-> (registros, errores, advertencias). Cada registro es un dict listo para guardar."""
    filas = filas_del_archivo(archivo)
    inicio, posiciones = None, {}
    for i, fila in enumerate(filas[:30]):
        mapa = {}
        for j, celda in enumerate(fila):
            campo = COLUMNAS.get(normal_columna(celda))
            if campo and campo not in mapa:
                mapa[campo] = j
        if all(c in mapa for c in REQUERIDAS):
            inicio, posiciones = i + 1, mapa
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas # Factura, Fecha, Facturador, Vendedor Asignado, Total y Razon Anulado")

    registros, errores, advertencias = {}, [], []
    repetidas = 0
    for n, fila in enumerate(filas[inicio:], start=inicio + 1):
        valor = lambda campo: fila[posiciones[campo]] if campo in posiciones and posiciones[campo] < len(fila) else None
        factura = a_texto(valor("factura"))
        if not factura and not any(a_texto(c) for c in fila):
            continue
        fecha, total = a_fecha(valor("fecha")), a_monto(valor("total"))
        if not factura or not fecha or total is None:
            errores.append(f"Fila {n}: falta número, fecha o total ({factura or '—'}); no se carga.")
            continue
        asignado, facturador = a_texto(valor("vendedor_asignado")), a_texto(valor("facturador"))
        razon = limpiar_razon(a_texto(valor("razon")))
        r = {"factura": factura[:30], "fecha": fecha.isoformat(), "total": total,
             "cliente": " ".join(a_texto(valor("cliente")).split())[:150] or None,
             "facturador": " ".join(facturador.split())[:100] or None,
             "vendedor_asignado": " ".join(asignado.split())[:100] or None,
             "vendedor": nombre_vendedor(asignado, facturador)[:100],
             "tipo": a_texto(valor("tipo")).upper()[:10] or None,
             "descripcion": " ".join(a_texto(valor("descripcion")).split())[:255] or None,
             "razon": razon[:500] or None,
             "anulado_por": a_texto(valor("anulado_por"))[:50] or None}
        anulado = a_fecha(valor("fecha_anulado"), con_hora=True)
        r["fecha_anulado"] = anulado.strftime("%Y-%m-%d %H:%M:%S") if anulado else None
        llave = (r["factura"], r["fecha"], f"{total:.2f}")
        previo = registros.get(llave)
        if previo:
            repetidas += 1
            # la misma factura dos veces: queda la que tenga razón y la anulación más reciente
            if (bool(r["razon"]), r["fecha_anulado"] or "") <= (bool(previo["razon"]), previo["fecha_anulado"] or ""):
                continue
        registros[llave] = r
    if repetidas:
        advertencias.append(f"{repetidas} fila(s) repetidas en el archivo (misma factura, fecha y total): se cuentan una sola vez.")
    sin_razon = sum(1 for r in registros.values() if not r["razon"])
    if sin_razon:
        advertencias.append(f"{sin_razon} factura(s) sin razón de anulación (el sistema las trae sin razón ni fecha de anulado); "
                            "quedan como «Sin razón registrada».")
    return list(registros.values()), errores, advertencias


# -----------------------------
# Carga
# -----------------------------
CAMPOS_GUARDAR = ["factura", "fecha", "total", "cliente", "facturador", "vendedor_asignado", "vendedor", "tipo", "descripcion",
                  "razon", "anulado_por", "fecha_anulado"]
CAMPOS_COMPARAR = ["cliente", "facturador", "vendedor_asignado", "vendedor", "tipo", "descripcion", "razon", "anulado_por", "fecha_anulado"]


def comparar_con_base(cursor, registros):
    """Cuántas son nuevas, cuántas ya estaban igual y cuántas cambian."""
    if not registros:
        return 0, 0, 0
    fechas = [r["fecha"] for r in registros]
    cursor.execute(f"SELECT factura, fecha, total, {', '.join(CAMPOS_COMPARAR)} FROM facturas_anuladas_sistema WHERE fecha BETWEEN %s AND %s",
                   (min(fechas), max(fechas)))
    existentes = {}
    for f in cursor.fetchall():
        existentes[(f["factura"], f["fecha"].isoformat(), f"{float(f['total']):.2f}")] = f
    nuevas = iguales = cambian = 0
    for r in registros:
        e = existentes.get((r["factura"], r["fecha"], f"{r['total']:.2f}"))
        if e is None:
            nuevas += 1
            continue
        actual = {c: (e[c].strftime("%Y-%m-%d %H:%M:%S") if isinstance(e[c], datetime) else e[c]) for c in CAMPOS_COMPARAR}
        if all((actual[c] or None) == (r[c] or None) for c in CAMPOS_COMPARAR):
            iguales += 1
        else:
            cambian += 1
    return nuevas, iguales, cambian


def resumen_carga(registros):
    por_vendedor, por_motivo = {}, {}
    for r in registros:
        m = motivo_de(r["razon"])
        v = por_vendedor.setdefault(r["vendedor"], {"vendedor": r["vendedor"], "facturas": 0, "monto": 0.0})
        v["facturas"] += 1
        v["monto"] += r["total"]
        por_motivo[m] = por_motivo.get(m, 0) + 1
    return (sorted(por_vendedor.values(), key=lambda x: (-x["facturas"], -x["monto"])),
            [(m, NOMBRE_MOTIVO[m], por_motivo[m]) for m in ORDEN_MOTIVOS if m in por_motivo])


@anulaciones_vendedor_bp.route("/cargar_anuladas_sistema", methods=["GET", "POST"])
@requiere_acceso("cargar_anuladas_sistema")
def cargar():
    if request.method == "GET":
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga())
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga(), errores_archivo=["Seleccione el reporte de facturas anuladas."])
    if not archivo.filename.lower().endswith((".csv", ".xlsx", ".xls")):
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga(),
                               errores_archivo=[f"{archivo.filename}: debe ser CSV o Excel (.csv, .xlsx, .xls)."])
    try:
        registros, errores, advertencias = leer_anuladas(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga(),
                               errores_archivo=[f"{archivo.filename}: {e}. ¿Es el reporte «FACTURAS ANULADAS» del sistema?"])
    except Exception as e:
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga(),
                               errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])
    if not registros:
        return render_template("cargar_anuladas_sistema.html", ultima=ultima_carga(), errores=errores,
                               errores_archivo=[f"{archivo.filename}: no trae facturas anuladas."])

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        nuevas, iguales, cambian = comparar_con_base(cursor, registros)
    finally:
        cursor.close()
        conn.close()

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump({"archivo": archivo.filename[:255], "registros": registros}, f, ensure_ascii=False)
    session["carga_anuladas_sis_token"] = token
    vendedores, motivos = resumen_carga(registros)
    fechas = sorted(r["fecha"] for r in registros)
    return render_template("cargar_anuladas_sistema.html", validado=True, token=token, archivo=archivo.filename,
                           total=len(registros), monto=sum(r["total"] for r in registros), desde=fechas[0], hasta=fechas[-1],
                           nuevas=nuevas, iguales=iguales, cambian=cambian, vendedores=vendedores, motivos=motivos,
                           errores=errores, advertencias=advertencias, ultima=ultima_carga())


@anulaciones_vendedor_bp.route("/cargar_anuladas_sistema/confirmar", methods=["POST"])
@requiere_acceso("cargar_anuladas_sistema")
def confirmar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_anuladas_sis_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("anulaciones_vendedor.cargar"))
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    registros = datos["registros"]
    usuario = session.get("usuario")
    actualizar = ", ".join(f"{c} = VALUES({c})" for c in CAMPOS_GUARDAR if c not in ("factura", "fecha", "total"))
    sql = (f"INSERT INTO facturas_anuladas_sistema ({', '.join(CAMPOS_GUARDAR)}, archivo, cargado_por) "
           f"VALUES ({', '.join(['%s'] * (len(CAMPOS_GUARDAR) + 2))}) "
           f"ON DUPLICATE KEY UPDATE {actualizar}, archivo = VALUES(archivo), actualizado = NOW()")
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()
        for i in range(0, len(registros), 500):
            cursor.executemany(sql, [tuple(r[c] for c in CAMPOS_GUARDAR) + (datos["archivo"], usuario) for r in registros[i:i + 500]])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó. Error de base de datos: {e}", "danger")
        return redirect(url_for("anulaciones_vendedor.cargar"))
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop("carga_anuladas_sis_token", None)
    fechas = sorted(r["fecha"] for r in registros)
    flash(f"✅ {len(registros):,} facturas anuladas guardadas ({fechas[0]} al {fechas[-1]}).", "success")
    return redirect(url_for("anulaciones_vendedor.reporte", desde=fechas[0], hasta=fechas[-1]))


def ultima_carga():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT MAX(fecha) AS hasta, MIN(fecha) AS desde, COUNT(*) AS n, MAX(COALESCE(actualizado, cargado)) AS cuando "
                       "FROM facturas_anuladas_sistema")
        return cursor.fetchone()
    finally:
        cursor.close()
        conn.close()


# -----------------------------
# Reporte
# -----------------------------
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def leer_filtros(args):
    ultima = (ultima_carga() or {}).get("hasta") or date.today()
    hasta = fecha_param(args.get("hasta"), ultima)
    desde = min(fecha_param(args.get("desde"), hasta.replace(day=1)), hasta)
    ver = args.get("ver")
    if ver not in ("dia", "mes"):
        ver = "dia" if (hasta - desde).days <= 31 else "mes"
    motivo = args.get("motivo") if args.get("motivo") in NOMBRE_MOTIVO else ""
    return {"desde": desde, "hasta": hasta, "ultima": ultima, "ver": ver, "motivo": motivo,
            "por": "facturador" if args.get("por") == "facturador" else "vendedor",
            "vendedor": (args.get("vendedor") or "").strip(), "tipo": args.get("tipo") if args.get("tipo") in ("CREDITO", "CONTADO") else ""}


def periodos(ultima):
    inicio_mes = ultima.replace(day=1)
    fin_ant = inicio_mes - timedelta(days=1)
    return [("Día", ultima, ultima), ("Ayer", ultima - timedelta(days=1), ultima - timedelta(days=1)),
            ("Este mes", inicio_mes, ultima), ("Mes anterior", fin_ant.replace(day=1), fin_ant),
            ("Últimos 3 meses", (inicio_mes - timedelta(days=62)).replace(day=1), ultima),
            ("Este año", ultima.replace(month=1, day=1), ultima)]


def etiqueta_periodo(d, ver):
    return f"{d:%d/%m}" if ver == "dia" else f"{MESES[d.month - 1]} {d:%y}"


def calcular(f):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        where, params = ["fecha BETWEEN %s AND %s"], [f["desde"], f["hasta"]]
        if f["tipo"]:
            where.append("tipo = %s")
            params.append(f["tipo"])
        cursor.execute(f"SELECT * FROM facturas_anuladas_sistema WHERE {' AND '.join(where)} ORDER BY fecha DESC, fecha_anulado DESC", params)
        filas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    quien = (lambda r: limpiar_nombre(r["vendedor"])) if f["por"] == "vendedor" else (lambda r: limpiar_nombre(r["facturador"]) or "Sin facturador")
    personas_todas = sorted({quien(r) for r in filas})
    detalle, por_persona, por_periodo, por_motivo = [], {}, {}, {}
    total = monto = 0
    for r in filas:
        r["motivo"] = motivo_de(r["razon"])
        r["persona"] = quien(r)
        r["total"] = float(r["total"])
        if (f["motivo"] and r["motivo"] != f["motivo"]) or (f["vendedor"] and r["persona"] != f["vendedor"]):
            continue
        detalle.append(r)
        total += 1
        monto += r["total"]
        p = por_persona.setdefault(r["persona"], {"persona": r["persona"], "facturas": 0, "monto": 0.0, "motivos": {}})
        p["facturas"] += 1
        p["monto"] += r["total"]
        p["motivos"][r["motivo"]] = p["motivos"].get(r["motivo"], 0) + 1
        k = r["fecha"] if f["ver"] == "dia" else r["fecha"].replace(day=1)
        x = por_periodo.setdefault(k, {"periodo": k, "facturas": 0, "monto": 0.0, "motivos": {}})
        x["facturas"] += 1
        x["monto"] += r["total"]
        x["motivos"][r["motivo"]] = x["motivos"].get(r["motivo"], 0) + 1
        m = por_motivo.setdefault(r["motivo"], {"motivo": r["motivo"], "facturas": 0, "monto": 0.0, "razones": {}})
        m["facturas"] += 1
        m["monto"] += r["total"]
        m["razones"][r["razon"] or "—"] = m["razones"].get(r["razon"] or "—", 0) + 1

    motivos = [m for m in ORDEN_MOTIVOS if m in por_motivo]
    motivos.sort(key=lambda m: -por_motivo[m]["facturas"])
    personas = sorted(por_persona.values(), key=lambda x: (-x["facturas"], -x["monto"]))
    for p in personas:
        p["pct"] = p["facturas"] / total * 100 if total else 0
        p["principal"] = max(p["motivos"].items(), key=lambda x: x[1])[0] if p["motivos"] else None
    lista_motivos = []
    for m in motivos:
        x = por_motivo[m]
        lista_motivos.append({**x, "nombre": NOMBRE_MOTIVO[m], "color": COLOR_MOTIVO[m], "pct": x["facturas"] / total * 100 if total else 0,
                              "razones": sorted(x["razones"].items(), key=lambda y: -y[1])[:8]})
    # todos los días (o meses) del rango, también los que no tienen anulaciones
    lista_periodos, d = [], f["desde"] if f["ver"] == "dia" else f["desde"].replace(day=1)
    while d <= f["hasta"]:
        lista_periodos.append(por_periodo.get(d, {"periodo": d, "facturas": 0, "monto": 0.0, "motivos": {}}))
        d = d + timedelta(days=1) if f["ver"] == "dia" else (d.replace(day=28) + timedelta(days=4)).replace(day=1)
    for x in lista_periodos:
        x["etiqueta"] = etiqueta_periodo(x["periodo"], f["ver"])
    dias = (f["hasta"] - f["desde"]).days + 1
    return {"detalle": detalle, "total": total, "monto": monto, "personas": personas, "personas_todas": personas_todas,
            "motivos": motivos, "lista_motivos": lista_motivos, "periodos": lista_periodos,
            "maximo_periodo": max((x["facturas"] for x in lista_periodos), default=0),
            "promedio_dia": total / dias if dias else 0,
            "nhe": por_motivo.get("NHE", {}).get("facturas", 0), "pmf": por_motivo.get("PMF", {}).get("facturas", 0)}


@anulaciones_vendedor_bp.route("/anulaciones_vendedor")
@requiere_acceso("anulaciones_vendedor")
def reporte():
    f = leer_filtros(request.args)
    d = calcular(f)
    return render_template("anulaciones_vendedor.html", f=f, d=d, nombre_motivo=NOMBRE_MOTIVO, color_motivo=COLOR_MOTIVO,
                           orden_motivos=ORDEN_MOTIVOS, periodos=periodos(f["ultima"]), ultima=ultima_carga())


@anulaciones_vendedor_bp.route("/anulaciones_vendedor/excel")
@requiere_acceso("anulaciones_vendedor")
def excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    quien = "Vendedor" if f["por"] == "vendedor" else "Facturador"
    nota = (f"Del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y} · {d['total']:,} facturas anuladas · C$ {d['monto']:,.2f}"
            f"{' · ' + NOMBRE_MOTIVO[f['motivo']] if f['motivo'] else ''}{' · ' + f['vendedor'] if f['vendedor'] else ''}"
            f"{' · ' + f['tipo'] if f['tipo'] else ''} · Vendedor = asignado, o el facturador si viene en blanco")
    wb = Workbook()

    def hoja(ws, titulo, columnas):
        ws["A1"], ws["A2"] = titulo, nota
        ws["A1"].font = Font(bold=True, size=14, color=VERDE)
        ws["A2"].font = Font(size=9, color="52514E")
        for i, (nombre, ancho) in enumerate(columnas, start=1):
            c = ws.cell(row=4, column=i, value=nombre)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NARANJA if nombre in ("Total", "Facturas") else VERDE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.freeze_panes = "B5"

    def total_fila(ws, fila, columnas_suma):
        ws.cell(row=fila, column=1, value="TOTAL").font = Font(bold=True)
        for col, formato in columnas_suma:
            l = get_column_letter(col)
            c = ws.cell(row=fila, column=col, value=f"=SUM({l}5:{l}{fila - 1})")
            c.font = Font(bold=True)
            c.number_format = formato

    motivos = d["motivos"]
    ws = wb.active
    ws.title = f"Por {quien.lower()}"
    hoja(ws, f"Facturas anuladas por {quien.lower()} y motivo", [(quien, 30), ("Facturas", 10), ("% del total", 9), ("Monto C$", 14)]
         + [(NOMBRE_MOTIVO[m], 13) for m in motivos])
    for n, p in enumerate(d["personas"], start=5):
        valores = [p["persona"], p["facturas"], p["pct"] / 100, round(p["monto"], 2)] + [p["motivos"].get(m) for m in motivos]
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=n, column=i, value=v)
            c.number_format = "0.0%" if i == 3 else "#,##0.00" if i == 4 else c.number_format
    total_fila(ws, 5 + len(d["personas"]), [(2, "#,##0"), (4, "#,##0.00")] + [(5 + i, "#,##0") for i in range(len(motivos))])

    ws2 = wb.create_sheet("Por día" if f["ver"] == "dia" else "Por mes")
    hoja(ws2, "Facturas anuladas por " + ("día" if f["ver"] == "dia" else "mes"),
         [("Fecha" if f["ver"] == "dia" else "Mes", 12), ("Facturas", 10), ("Monto C$", 14)] + [(NOMBRE_MOTIVO[m], 13) for m in motivos])
    for n, x in enumerate(d["periodos"], start=5):
        primera = x["periodo"] if f["ver"] == "dia" else x["etiqueta"]
        valores = [primera, x["facturas"], round(x["monto"], 2)] + [x["motivos"].get(m) for m in motivos]
        for i, v in enumerate(valores, start=1):
            c = ws2.cell(row=n, column=i, value=v)
            c.number_format = "dd/mm/yyyy" if (i == 1 and f["ver"] == "dia") else "#,##0.00" if i == 3 else c.number_format
    total_fila(ws2, 5 + len(d["periodos"]), [(2, "#,##0"), (3, "#,##0.00")] + [(4 + i, "#,##0") for i in range(len(motivos))])

    ws3 = wb.create_sheet("Detalle")
    hoja(ws3, "Detalle de facturas anuladas", [("Fecha", 11), ("Factura", 11), ("Cliente", 38), (quien, 28), ("Facturador", 24),
                                               ("Vendedor asignado", 24), ("Tipo", 9), ("Total", 13), ("Motivo", 26), ("Razón (sistema)", 50),
                                               ("Anulado por", 13), ("Fecha anulado", 16)])
    for n, r in enumerate(d["detalle"], start=5):
        valores = [r["fecha"], r["factura"], r["cliente"], r["persona"], r["facturador"], r["vendedor_asignado"], r["tipo"], r["total"],
                   NOMBRE_MOTIVO[r["motivo"]], r["razon"], r["anulado_por"], r["fecha_anulado"]]
        for i, v in enumerate(valores, start=1):
            c = ws3.cell(row=n, column=i, value=v)
            c.number_format = {1: "dd/mm/yyyy", 8: "#,##0.00", 12: "dd/mm/yyyy hh:mm"}.get(i, c.number_format)
    ws3.auto_filter.ref = f"A4:L{4 + max(len(d['detalle']), 1)}"

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"anulaciones_{f['por']}_{f['desde']}_al_{f['hasta']}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
