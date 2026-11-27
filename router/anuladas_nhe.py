"""Anulaciones por NHE (no hay existencia) con el detalle de productos.

Facturas (filtradas por la fecha de anulación):
- Reporte de anuladas del sistema (facturas_anuladas_sistema): la razón escrita al anular, clasificada con el mismo criterio de
  «Anulaciones por vendedor» (router/anulaciones_vendedor.motivo_de: NHE, PNEB, «sin existencia», «no se encontró», …).
- Módulo «Anular factura» (facturas_anuladas): el motivo escrito ahí, con el mismo criterio; las restauradas no cuentan.
Una factura que está en los dos (mismo número y fecha) sale una sola vez, con los dos orígenes.

Productos: las líneas de la factura en el detalle por producto (ventas_por_producto, número + fecha; con la bodega cuando el
módulo la tiene). Si no aparecen ahí, se usa el respaldo que guardó el módulo «Anular factura».
Producto señalado: la línea cuyo número de parte o código aparece en la razón de la anulación (por ejemplo «NHE US7145HN (B2)»),
comparando sin espacios, guiones ni tildes. Va primero en su factura; las demás líneas quedan como información. Los códigos cortos
(3 caracteres sin contar guiones, como «A-10») y los de solo dígitos de menos de 6 cifras deben aparecer como palabra aparte (o en dos
palabras seguidas, «A 10»), para no confundirlos con partes de otras palabras o números; los de 2 caracteres o menos no se comparan.

Por producto: último ingreso por orden de compra (transacciones_inventario, como Seguimiento de ingresos); si nunca hubo compra,
el último ajuste de entrada, y si ese ajuste es de 2021 (montaje de saldos al iniciar el sistema) se indica «Inventario inicial».
Existencia ACTUAL (existencias_bodega: la de la bodega de la factura, las OTRAS bodegas que tienen existencia y la total), es decir,
a la fecha del último inventario cargado. «En otra bodega» marca los casos sin existencia en la bodega de la factura pero con existencia
en otra (se pudo trasladar).

Excel: hoja con todo el detalle y hoja «Para imprimir» en blanco y negro, agrupada por fecha de anulación y bodega; en las facturas
con el producto identificado solo va ese producto (marcado ►), en las demás van todas las líneas.
"""
import io
import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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


def palabras(razon):
    """Palabras de la razón (compactadas) y también cada par de palabras seguidas («A 10» -> «A10»)."""
    sueltas = [compacto(t) for t in re.split(r"[\s,;:()/]+", sin_tildes(razon or "").upper()) if compacto(t)]
    return set(sueltas) | {a + b for a, b in zip(sueltas, sueltas[1:])}


def mencionado(razon, valor):
    """¿El número de parte o código «valor» aparece en la razón?"""
    v = compacto(valor)
    if len(v) < 3:
        return False
    if len(v) < 4 or (v.isdigit() and len(v) < 6):     # código corto: debe ir como palabra aparte
        return v in palabras(razon)
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


def para_imprimir(filas):
    """Filas de la versión impresa: en facturas con producto identificado, solo ese; en las demás, todas sus líneas."""
    con_señalado = {(r["factura"], r["fecha"]) for r in filas if r["señalado"]}
    return [r for r in filas if r["señalado"] or (r["factura"], r["fecha"]) not in con_señalado]


def hoja_imprimir(wb, filas, desde, hasta):
    """Hoja «Para imprimir»: blanco y negro, agrupada por fecha de anulación y luego por bodega, con subtotales."""
    ws = wb.create_sheet("Para imprimir")
    fino, grueso = Side(style="thin", color="000000"), Side(style="medium", color="000000")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    gris_claro, gris = PatternFill("solid", fgColor="EDEDED"), PatternFill("solid", fgColor="D0D0D0")
    columnas = [("Factura", 9), ("Fecha fact.", 10), ("Cliente / vendedor", 30), ("Producto", 38), ("Cant.", 6), ("Último ingreso", 13),
                ("Exist. bodega", 8), ("Otras bodegas", 26), ("Exist. total", 8), ("Razón", 34)]
    ultima = get_column_letter(len(columnas))
    ws["A1"] = f"Anulaciones por NHE · anuladas del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("► = producto mencionado en la razón (de esas facturas solo se imprime ese). Existencia actual al último inventario cargado. "
                f"Impreso {datetime.now():%d/%m/%Y %I:%M %p}")
    ws["A2"].font = Font(size=8, italic=True)
    for j, (titulo, ancho) in enumerate(columnas, 1):
        c = ws.cell(row=4, column=j, value=titulo)
        c.font, c.fill, c.border = Font(bold=True, size=9), gris, borde
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = ancho
    n = 5
    lineas = para_imprimir(filas)
    grupos = defaultdict(lambda: defaultdict(list))
    for r in lineas:
        grupos[a_fecha(r["anulada"])][(r["bodega"] or "Sin bodega").strip()].append(r)

    def banda(texto, relleno, tamaño, arriba=None):
        nonlocal n
        ws.merge_cells(f"A{n}:{ultima}{n}")
        c = ws.cell(row=n, column=1, value=texto)
        c.font, c.fill = Font(bold=True, size=tamaño), relleno
        for j in range(1, len(columnas) + 1):
            ws.cell(row=n, column=j).border = Border(left=fino, right=fino, top=arriba or fino, bottom=fino)
        n += 1

    for dia in sorted(grupos, reverse=True):
        facturas_dia = {(r["factura"], r["fecha"]) for b in grupos[dia].values() for r in b}
        banda(f"Anuladas el {dia:%d/%m/%Y}" + f"   ·   {len(facturas_dia)} factura(s)", gris, 11, arriba=grueso)
        for bodega in sorted(grupos[dia]):
            lista = grupos[dia][bodega]
            banda(f"{bodega}   ·   {len({(r['factura'], r['fecha']) for r in lista})} factura(s), "
                  f"{sum(r.get('cantidad') or 0 for r in lista):g} unidad(es)", gris_claro, 10)
            anterior = None
            for r in lista:
                primera = (r["factura"], r["fecha"]) != anterior
                anterior = (r["factura"], r["fecha"])
                ingreso = (f"{r['ultimo_ingreso']:%d/%m/%Y}" + (f" ({r['texto_ingreso'].lower()})" if r.get("texto_ingreso") else "")
                           if r.get("ultimo_ingreso") else ("sin ingresos" if r["codigo"] else ""))
                producto = ("(sin detalle de productos)" if not r["codigo"] else
                            ("► " if r["señalado"] else "") + f"{r['codigo']} {r.get('parte') or ''} · {r.get('producto') or ''} {r.get('marca') or ''}".strip())
                otras = ", ".join(f"{b}: {c:g}" for b, c in r.get("otras_bodegas") or [])
                valores = [r["factura"] if primera else None, r["fecha"] if primera else None,
                           " · ".join(x for x in (r["cliente"], r["vendedor"]) if x) if primera else None, producto, r.get("cantidad"), ingreso,
                           r["existencia_bodega"], otras or ("—" if r["codigo"] else ""), r["existencia_total"], r["razon"] if primera else None]
                for j, v in enumerate(valores, 1):
                    c = ws.cell(row=n, column=j, value=v)
                    c.border = borde
                    c.font = Font(size=9, bold=bool(r["señalado"] and j == 4))
                    c.alignment = Alignment(vertical="top", wrap_text=j in (3, 4, 6, 8, 10), horizontal="right" if j in (5, 7, 9) else None)
                    if j == 2 and v:
                        c.number_format = "dd/mm/yyyy"
                n += 1
        n += 1
    if not lineas:
        ws.cell(row=n, column=1, value="No hay facturas anuladas por NHE en el período.")
    # Página: horizontal, carta, ajustada al ancho, encabezado repetido, número de página
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:4"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.oddFooter.center.text = "Página &P de &N"
    ws.oddFooter.right.text = "Anulaciones por NHE"
    ws.page_setup.blackAndWhite = True
    return ws


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
    hoja_imprimir(wb, filas, desde, hasta)
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"anuladas_nhe_{desde:%Y%m%d}_{hasta:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
