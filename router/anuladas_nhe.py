"""Anulaciones por NHE (no hay existencia) con el detalle de productos.

Facturas (filtradas por la fecha de anulación):
- Reporte de anuladas del sistema (facturas_anuladas_sistema): la razón escrita al anular, clasificada con el mismo criterio de
  «Anulaciones por vendedor» (router/anulaciones_vendedor.motivo_de: NHE, PNEB, «sin existencia», «no se encontró», …).
- Módulo «Anular factura» (facturas_anuladas): el motivo escrito ahí, con el mismo criterio; las restauradas no cuentan.
Una factura que está en los dos (mismo número y fecha) sale una sola vez, con los dos orígenes.

Productos: las líneas de la factura en el detalle por producto (ventas_por_producto, número + fecha; con la bodega cuando el
módulo la tiene). Si no aparecen ahí, se usa el respaldo que guardó el módulo «Anular factura».
Producto señalado: la línea cuyo número de parte o código aparece en la razón de la anulación (por ejemplo «NHE US7145HN (B2)»),
comparando sin espacios, guiones ni tildes. Va primero en su factura; las demás líneas quedan como información. Los números de parte
solo con dígitos y de menos de 6 cifras deben aparecer como palabra aparte (para no confundirlos con otros números de la razón).

Por producto: último ingreso por orden de compra (transacciones_inventario, como Seguimiento de ingresos); si nunca hubo compra,
el último ajuste de entrada, y si ese ajuste es de 2021 (montaje de saldos al iniciar el sistema) se indica «Inventario inicial».
Existencia ACTUAL (existencias_bodega: la de la bodega de la factura, las OTRAS bodegas que tienen existencia y la total), es decir,
a la fecha del último inventario cargado. «En otra bodega» marca los casos sin existencia en la bodega de la factura pero con existencia
en otra (se pudo trasladar).
"""
import io
import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from router.accesos import requiere_acceso
from router.anulaciones_vendedor import limpiar_nombre, limpiar_razon, motivo_de
from router.cargar_catalogo import sin_tildes
from utils.db import get_connection

anuladas_nhe_bp = Blueprint("anuladas_nhe", __name__)
VERDE, NARANJA = "027266", "FE6412"
AÑO_INVENTARIO_INICIAL = 2021     # ajustes de entrada de ese año = montaje de saldos al empezar el sistema


def lista_sql(valores):
    return ", ".join(["%s"] * len(valores))


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def numero(factura):
    """El número de factura como en ventas_por_producto (`No. Factura` es numérico)."""
    try:
        return float(str(factura).strip())
    except (TypeError, ValueError):
        return None


# -----------------------------
# 1) Facturas anuladas por NHE
# -----------------------------
def facturas_nhe(cursor, desde, hasta):
    fin = hasta + timedelta(days=1)
    facturas = {}
    cursor.execute("""SELECT factura, fecha, fecha_anulado, razon, cliente, vendedor, facturador, total FROM facturas_anuladas_sistema
                      WHERE COALESCE(fecha_anulado, fecha) >= %s AND COALESCE(fecha_anulado, fecha) < %s""", (desde, fin))
    for r in cursor.fetchall():
        if motivo_de(r["razon"]) != "NHE":
            continue
        clave = (str(r["factura"]).strip(), a_fecha(r["fecha"]))
        facturas[clave] = {"factura": clave[0], "fecha": clave[1], "anulada": r["fecha_anulado"] or r["fecha"], "origenes": ["Sistema"],
                           "razon": limpiar_razon(r["razon"]), "cliente": r["cliente"], "vendedor": limpiar_nombre(r["vendedor"] or r["facturador"]),
                           "total": r["total"], "bodega": None, "respaldo": None}
    cursor.execute("""SELECT factura, bodega, fecha, anulado, motivo, cliente, vendedor, total, respaldo FROM facturas_anuladas
                      WHERE restaurado IS NULL AND anulado >= %s AND anulado < %s""", (desde, fin))
    for r in cursor.fetchall():
        if motivo_de(r["motivo"]) != "NHE":
            continue
        clave = (str(r["factura"]).strip(), a_fecha(r["fecha"]))
        if clave in facturas:
            f = facturas[clave]
            f["origenes"].append("Anular factura")
            f["bodega"] = r["bodega"]
            f["respaldo"] = r["respaldo"]
            if r["motivo"] and r["motivo"] not in f["razon"]:
                f["razon"] += f" / {r['motivo']}"
        else:
            facturas[clave] = {"factura": clave[0], "fecha": clave[1], "anulada": r["anulado"], "origenes": ["Anular factura"],
                               "razon": r["motivo"], "cliente": r["cliente"], "vendedor": limpiar_nombre(r["vendedor"]), "total": r["total"],
                               "bodega": r["bodega"], "respaldo": r["respaldo"]}
    return sorted(facturas.values(), key=lambda f: (f["anulada"], f["factura"]), reverse=True)


# -----------------------------
# 2) Productos de cada factura
# -----------------------------
def lineas_de_respaldo(f):
    """Líneas de ventas_por_producto que guardó el módulo «Anular factura» al anular."""
    try:
        respaldo = json.loads(f["respaldo"] or "{}")
    except (TypeError, ValueError):
        return []
    filas = respaldo.get("ventas_por_producto") if isinstance(respaldo, dict) else None
    return [{"codigo": x.get("Codigo"), "parte": x.get("Num. Parte") or x.get("Numero_parte"), "producto": x.get("Descrip Prod."),
             "marca": x.get("Marca"), "cantidad": float(x.get("Cantidad Vendida") or 0), "bodega": x.get("Bodega")}
            for x in (filas or []) if isinstance(x, dict)]


def compacto(texto):
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(texto or "").upper())


def mencionado(razon, valor):
    """¿El número de parte o código «valor» aparece en la razón?"""
    v = compacto(valor)
    if len(v) < 4:
        return False
    if v.isdigit() and len(v) < 6:      # número corto: debe ir como palabra aparte
        return v in {compacto(t) for t in re.split(r"[\s,;:()/]+", sin_tildes(razon or "").upper())}
    return v in compacto(razon)


def señalar(lineas, razon):
    """Marca las líneas mencionadas en la razón y las pone primero (orden estable)."""
    for l in lineas:
        l["señalado"] = mencionado(razon, l.get("parte")) or mencionado(razon, l.get("codigo"))
    return sorted(lineas, key=lambda l: not l["señalado"])


def armar(cursor, facturas):
    """Una fila por producto de cada factura, con último ingreso y existencia."""
    por_numero = defaultdict(list)
    numeros = sorted({n for n in (numero(f["factura"]) for f in facturas) if n is not None})
    for i in range(0, len(numeros), 500):
        lote = numeros[i:i + 500]
        cursor.execute(f"""SELECT `No. Factura` AS num, Fecha, Codigo, `Num. Parte` AS parte, Numero_parte, `Descrip Prod.` AS producto, Marca,
                                  `Cantidad Vendida` AS cantidad, Bodega FROM ventas_por_producto WHERE `No. Factura` IN ({lista_sql(lote)})""", lote)
        for x in cursor.fetchall():
            por_numero[(x["num"], a_fecha(x["Fecha"]))].append(x)
    filas = []
    for f in facturas:
        lineas = [{"codigo": x["Codigo"], "parte": x["parte"] or x["Numero_parte"], "producto": x["producto"], "marca": x["Marca"],
                   "cantidad": float(x["cantidad"] or 0), "bodega": x["Bodega"]}
                  for x in por_numero.get((numero(f["factura"]), f["fecha"]), [])
                  if not f["bodega"] or (x["Bodega"] or "").strip().upper() == f["bodega"].strip().upper()]
        if not lineas:
            lineas = lineas_de_respaldo(f)
        if not lineas:
            filas.append({**f, "codigo": None, "sin_detalle": True, "señalado": False})
        for l in señalar(lineas, f["razon"]):
            filas.append({**f, **l, "bodega": l["bodega"] or f["bodega"], "sin_detalle": False})
    codigos = sorted({r["codigo"] for r in filas if r["codigo"]})
    ingreso, ajuste, existencia, por_bodega, bodegas_de = {}, {}, defaultdict(float), {}, defaultdict(list)
    for i in range(0, len(codigos), 500):
        lote = codigos[i:i + 500]
        cursor.execute(f"""SELECT codigo_producto, MAX(fecha) AS f FROM transacciones_inventario WHERE tipo_transaccion = 'ENTRADA'
                           AND tipo_documento_origen = 'ORDEN COMPRA' AND codigo_producto IN ({lista_sql(lote)}) GROUP BY codigo_producto""", lote)
        for x in cursor.fetchall():
            ingreso[x["codigo_producto"]] = a_fecha(x["f"])
        cursor.execute(f"""SELECT codigo_producto, MAX(fecha) AS f FROM transacciones_inventario WHERE tipo_transaccion = 'AJUSTE ENTRADA'
                           AND codigo_producto IN ({lista_sql(lote)}) GROUP BY codigo_producto""", lote)
        for x in cursor.fetchall():
            ajuste[x["codigo_producto"]] = a_fecha(x["f"])
        cursor.execute(f"""SELECT Codigo_producto, Bodega, SUM(Cant_Final) AS cant FROM existencias_bodega
                           WHERE Codigo_producto IN ({lista_sql(lote)}) GROUP BY Codigo_producto, Bodega""", lote)
        for x in cursor.fetchall():
            cant = float(x["cant"] or 0)
            existencia[x["Codigo_producto"]] += cant
            por_bodega[(x["Codigo_producto"], (x["Bodega"] or "").strip().upper())] = cant
            if cant > 0:
                bodegas_de[x["Codigo_producto"]].append(((x["Bodega"] or "").strip(), cant))
    for r in filas:
        r["ultimo_ingreso"] = ingreso.get(r["codigo"])
        r["tipo_ingreso"] = "compra" if r["ultimo_ingreso"] else None
        if not r["ultimo_ingreso"] and ajuste.get(r["codigo"]):      # sin compras: el último ajuste de entrada
            r["ultimo_ingreso"] = ajuste[r["codigo"]]
            r["tipo_ingreso"] = "inicial" if r["ultimo_ingreso"].year == AÑO_INVENTARIO_INICIAL else "ajuste"
        r["texto_ingreso"] = {"compra": "", "ajuste": "Ajuste", "inicial": "Inventario inicial"}.get(r["tipo_ingreso"], "")
        r["existencia_total"] = existencia.get(r["codigo"], 0.0) if r["codigo"] else None
        r["existencia_bodega"] = por_bodega.get((r["codigo"], (r["bodega"] or "").strip().upper()), 0.0) if r["codigo"] else None
        propia = (r["bodega"] or "").strip().upper()
        r["otras_bodegas"] = sorted(((b, c) for b, c in bodegas_de.get(r["codigo"], []) if b.upper() != propia), key=lambda x: -x[1])
        r["en_otra_bodega"] = bool(r["codigo"] and not r["existencia_bodega"] and r["otras_bodegas"])
        r["ingreso_despues"] = bool(r["ultimo_ingreso"] and a_fecha(r["anulada"]) and r["ultimo_ingreso"] > a_fecha(r["anulada"]))
    return filas


def calcular(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        facturas = facturas_nhe(cursor, desde, hasta)
        filas = armar(cursor, facturas)
    finally:
        cursor.close()
        conn.close()
    productos = defaultdict(lambda: {"veces": 0, "cantidad": 0.0})
    hay_señalados = any(r["señalado"] for r in filas)
    for r in filas:
        if r["codigo"] and (r["señalado"] or not hay_señalados):
            p = productos[(r["codigo"], r["parte"], r["producto"], r["marca"])]
            p["veces"] += 1
            p["cantidad"] += r["cantidad"]
    resumen = {"facturas": len(facturas), "lineas": sum(1 for r in filas if r["codigo"]), "productos": len({r["codigo"] for r in filas if r["codigo"]}),
               "unidades": sum(r["cantidad"] for r in filas if r["codigo"]), "sin_detalle": sum(1 for r in filas if r["sin_detalle"]),
               "sistema": sum(1 for f in facturas if "Sistema" in f["origenes"]),
               "modulo": sum(1 for f in facturas if "Anular factura" in f["origenes"]),
               "identificadas": len({(r["factura"], r["fecha"]) for r in filas if r["señalado"]}), "top_señalados": hay_señalados}
    top = sorted(({"codigo": k[0], "parte": k[1], "producto": k[2], "marca": k[3], **v} for k, v in productos.items()),
                 key=lambda x: (-x["veces"], -x["cantidad"]))[:15]
    return filas, resumen, top


def leer_periodo(args):
    hoy = date.today()

    def fecha(nombre, defecto):
        try:
            return datetime.strptime(args.get(nombre) or "", "%Y-%m-%d").date()
        except ValueError:
            return defecto
    desde, hasta = fecha("desde", hoy.replace(day=1)), fecha("hasta", hoy)
    return (hasta, desde) if hasta < desde else (desde, hasta)


@anuladas_nhe_bp.route("/anuladas_nhe")
@requiere_acceso("anuladas_nhe")
def anuladas_nhe():
    desde, hasta = leer_periodo(request.args)
    filas, resumen, top = calcular(desde, hasta)
    return render_template("anuladas_nhe.html", filas=filas, r=resumen, top=top, desde=desde, hasta=hasta)


@anuladas_nhe_bp.route("/anuladas_nhe/excel")
@requiere_acceso("anuladas_nhe")
def excel():
    desde, hasta = leer_periodo(request.args)
    filas, resumen, _ = calcular(desde, hasta)
    wb = Workbook()
    ws = wb.active
    ws.title = "Anuladas por NHE"
    ws["A1"] = f"Anulaciones por NHE · anuladas del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}"
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"] = (f"{resumen['facturas']} factura(s), {resumen['lineas']} línea(s); en {resumen['identificadas']} se identificó el producto por la razón "
                f"(filas resaltadas). Existencia actual (último inventario cargado). "
                f"Generado {datetime.now():%d/%m/%Y %I:%M %p}")
    ws["A2"].font = Font(size=9, color="52514E")
    enc = ["Factura", "Fecha factura", "Fecha anulación", "Origen", "Razón", "Cliente", "Vendedor", "Señalado en la razón", "Bodega", "Código",
           "Nº parte", "Producto", "Marca", "Cantidad", "Último ingreso", "Tipo de ingreso", "Existencia en la bodega", "Existencia en otras bodegas", "Existencia total"]
    naranja = PatternFill("solid", fgColor="FFF1E8")
    for j, e in enumerate(enc, 1):
        c = ws.cell(row=4, column=j, value=e)
        c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=VERDE), Alignment(horizontal="center", wrap_text=True)
    n = 5
    for r in filas:
        valores = [r["factura"], r["fecha"], r["anulada"], " + ".join(r["origenes"]), r["razon"], r["cliente"], r["vendedor"],
                   "Sí" if r["señalado"] else None, r["bodega"], r["codigo"] or "(sin detalle de productos)", r.get("parte"), r.get("producto"),
                   r.get("marca"), r.get("cantidad"), r.get("ultimo_ingreso"),
                   {"compra": "Compra", "ajuste": "Ajuste", "inicial": "Inventario inicial"}.get(r.get("tipo_ingreso"), "Sin ingresos" if r["codigo"] else None),
                   r["existencia_bodega"], ", ".join(f"{b}: {c:g}" for b, c in r.get("otras_bodegas") or []) or None, r["existencia_total"]]
        for j, v in enumerate(valores, 1):
            c = ws.cell(row=n, column=j, value=v)
            if isinstance(v, datetime):
                c.number_format = "dd/mm/yyyy hh:mm AM/PM"
            elif isinstance(v, date):
                c.number_format = "dd/mm/yyyy"
            if r["señalado"]:
                c.fill = naranja
                c.font = Font(bold=True)
        n += 1
    for j, ancho in enumerate([10, 12, 18, 16, 40, 30, 22, 11, 20, 14, 18, 36, 14, 10, 13, 16, 13, 34, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.freeze_panes = "A5"
    if filas:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(enc))}{n - 1}"
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"anuladas_nhe_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
