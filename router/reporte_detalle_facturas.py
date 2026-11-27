"""Detalle de facturas: una fila por factura de la facturación cargada (ventas_historicas, «Ventas del día»), con bodega, facturador,
vendedor, número, fecha, cliente, tipo de factura, crédito/contado, subtotal, descuento, IVA y total.

Filtros: fechas (por defecto el último día cargado), crédito / contado, búsqueda (cliente, factura, vendedor, facturador o bodega) y
si se incluyen las anuladas (por defecto no; si se incluyen, se marcan y no suman en los totales). Exporta a Excel.
Abre y cierra su propia conexión."""
import io
from datetime import date, datetime

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from router.accesos import requiere_acceso
from utils.db import get_connection

detalle_facturas_bp = Blueprint("detalle_facturas", __name__)

VERDE = "027266"
MAX_PANTALLA = 3000        # filas en pantalla; el Excel lleva todas
COLUMNAS = [  # clave, encabezado, es dinero
    ("Bodega", "Bodega", False), ("facturador", "Facturador", False), ("vendedor", "Vendedor", False), ("factura", "# Factura", False),
    ("Fecha", "Fecha", False), ("cliente", "Nombre de Cliente", False), ("tipo factura", "Tipo Factura", False),
    ("Credito/Contado", "Crédito/Contado", False), ("SubTotal", "SubTotal", True), ("Descuento", "Descuento", True), ("IVA", "IVA", True),
    ("Total", "Total", True),
]
MONTOS = [c for c, _, dinero in COLUMNAS if dinero]


def ultimo_dia():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas")
        return cursor.fetchone()[0]
    finally:
        cursor.close()
        conn.close()


def leer_filtros(args):
    def fecha(nombre):
        try:
            return datetime.strptime(args.get(nombre) or "", "%Y-%m-%d").date()
        except ValueError:
            return None
    desde, hasta = fecha("desde"), fecha("hasta")
    if not desde and not hasta:
        desde = hasta = ultimo_dia() or date.today()
    desde, hasta = desde or hasta, hasta or desde
    if hasta < desde:
        desde, hasta = hasta, desde
    tipo = (args.get("tipo") or "").upper()
    return {"desde": desde, "hasta": hasta, "tipo": tipo if tipo in ("CONTADO", "CREDITO") else "",
            "buscar": " ".join((args.get("buscar") or "").split())[:80], "anuladas": args.get("anuladas") == "1"}


def leer_facturas(f):
    sql = ["SELECT Bodega, facturador, vendedor, factura, Fecha, cliente, `tipo factura`, `Credito/Contado`, SubTotal, Descuento, IVA, Total,",
           "estado_factura FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s"]
    params = [f["desde"], f["hasta"]]
    if f["tipo"]:
        sql.append("AND UPPER(`Credito/Contado`) = %s")
        params.append(f["tipo"])
    if not f["anuladas"]:
        sql.append("AND COALESCE(estado_factura, '') <> 'Anulado'")
    if f["buscar"]:
        sql.append("AND CONCAT_WS(' ', cliente, factura, vendedor, facturador, Bodega) LIKE %s")
        params.append(f"%{f['buscar']}%")
    sql.append("ORDER BY Fecha, Bodega, CAST(factura AS UNSIGNED), factura")
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(" ".join(sql), params)
        filas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    totales = {c: 0.0 for c in MONTOS}
    n_anuladas = 0
    for r in filas:
        r["anulada"] = r["estado_factura"] == "Anulado"
        if r["anulada"]:
            n_anuladas += 1
            continue
        for c in MONTOS:
            totales[c] += float(r[c] or 0)
    return filas, {"facturas": len(filas) - n_anuladas, "anuladas": n_anuladas, **totales}


@detalle_facturas_bp.route("/detalle_facturas")
@requiere_acceso("detalle_facturas")
def detalle_facturas():
    f = leer_filtros(request.args)
    filas, t = leer_facturas(f)
    return render_template("reporte_detalle_facturas.html", f=f, filas=filas[:MAX_PANTALLA], recortado=len(filas) > MAX_PANTALLA,
                           total_filas=len(filas), t=t, columnas=COLUMNAS)


@detalle_facturas_bp.route("/detalle_facturas/excel")
@requiere_acceso("detalle_facturas")
def excel():
    f = leer_filtros(request.args)
    filas, t = leer_facturas(f)
    wb = Workbook()
    ws = wb.active
    ws.title = "Detalle de facturas"
    periodo = f"{f['desde']:%d/%m/%Y}" + (f" al {f['hasta']:%d/%m/%Y}" if f["hasta"] != f["desde"] else "")
    ws["A1"] = f"Detalle de facturas · {periodo}"
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    notas = [("contado" if f["tipo"] == "CONTADO" else "crédito") if f["tipo"] else None, f"búsqueda «{f['buscar']}»" if f["buscar"] else None,
             "incluye anuladas (no suman)" if f["anuladas"] else "sin anuladas"]
    ws["A2"] = " · ".join(n for n in notas if n) + f" · generado {datetime.now():%d/%m/%Y %I:%M %p}"
    ws["A2"].font = Font(size=9, color="52514E")
    encabezados = [e for _, e, _ in COLUMNAS] + (["Estado"] if f["anuladas"] else [])
    for j, e in enumerate(encabezados, 1):
        c = ws.cell(row=4, column=j, value=e)
        c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=VERDE), Alignment(horizontal="center")
    gris = Font(color="9A9892", strike=True)
    n = 5
    for r in filas:
        for j, (clave, _, dinero) in enumerate(COLUMNAS, 1):
            v = r[clave]
            c = ws.cell(row=n, column=j, value=float(v) if dinero and v is not None else v)
            if dinero:
                c.number_format = "#,##0.00"
            elif clave == "Fecha":
                c.number_format = "dd/mm/yyyy"
            if r["anulada"]:
                c.font = gris
        if f["anuladas"]:
            ws.cell(row=n, column=len(COLUMNAS) + 1, value=r["estado_factura"])
        n += 1
    ws.cell(row=n, column=1, value=f"Total ({t['facturas']} facturas)").font = Font(bold=True)
    for j, (clave, _, dinero) in enumerate(COLUMNAS, 1):
        if dinero:
            c = ws.cell(row=n, column=j, value=round(t[clave], 2))
            c.number_format, c.font = "#,##0.00", Font(bold=True)
    for j, ancho in enumerate([22, 22, 22, 11, 11, 38, 12, 14, 13, 12, 12, 13, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.freeze_panes = "A5"
    if filas:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(encabezados))}{n - 1}"
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"detalle_facturas_{f['desde']:%Y%m%d}_{f['hasta']:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
