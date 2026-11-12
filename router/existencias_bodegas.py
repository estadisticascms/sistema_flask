"""Existencias y costo por bodega: una o varias bodegas a la vez, con total al costo y orden por montos.
El costo es el Monto Final de cada bodega (el mismo valor de inventario del dashboard), en córdobas."""
import io
from datetime import datetime

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso

existencias_bodegas_bp = Blueprint("existencias_bodegas", __name__)

MAX_FILAS_PANTALLA = 2000        # el Excel lleva todo
VERDE, NARANJA = "027266", "FE6412"
SIN_COSTO = float("inf")
ORDENES = {
    "costo_total_desc": ("Mayor costo total", lambda x: (-x["costo_total"], x["codigo"])),
    "costo_total_asc": ("Menor costo total", lambda x: (x["costo_total"], x["codigo"])),
    "costo_unit_desc": ("Mayor costo unitario", lambda x: (-(x["costo_unit"] or 0), x["codigo"])),
    "costo_unit_asc": ("Menor costo unitario", lambda x: (SIN_COSTO if x["costo_unit"] is None else x["costo_unit"], x["codigo"])),
    "existencia_desc": ("Mayor existencia", lambda x: (-x["existencia"], x["codigo"])),
    "marca": ("Marca y descripción", lambda x: (x["marca"], x["nombre"], x["codigo"])),
    "nombre": ("Descripción", lambda x: (x["nombre"], x["codigo"])),
}


def leer_filtros(fuente):
    return {
        "bodegas": sorted({b.strip() for b in fuente.getlist("bodega") if b.strip()}),
        "marcas": sorted({m.strip() for m in fuente.getlist("marca") if m.strip()}),
        "q": (fuente.get("q") or "").strip(),
        "orden": fuente.get("orden") if fuente.get("orden") in ORDENES else "costo_total_desc",
        # por defecto solo unidades reales; la casilla desmarcada manda "0"
        "solo_con_existencia": fuente.get("solo_con_existencia", "1") == "1",
    }


def params_url(f):
    p = [("bodega", b) for b in f["bodegas"]]
    p += [("marca", m) for m in f["marcas"]]
    p += [(k, f[k]) for k in ("q", "orden") if f[k]]
    return p + [("solo_con_existencia", "1" if f["solo_con_existencia"] else "0")]


def descripcion_filtros(f):
    partes = [("Bodega: " if len(f["bodegas"]) == 1 else "Bodegas: ") + ", ".join(f["bodegas"]) if f["bodegas"] else "Todas las bodegas"]
    if f["marcas"]:
        partes.append(("Marca: " if len(f["marcas"]) == 1 else "Marcas: ") + ", ".join(f["marcas"]))
    if f["q"]:
        partes.append(f"Búsqueda: {f['q']}")
    return " · ".join(partes)


def lista_bodegas():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DISTINCT Bodega FROM existencias_bodega WHERE Bodega IS NOT NULL AND Bodega <> '' ORDER BY Bodega")
        return [r[0] for r in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()


def lista_marcas():
    """Marcas presentes en existencias, con cuántos productos tienen unidades."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT TRIM(Marca), COUNT(DISTINCT CASE WHEN Cant_Final > 0 THEN Codigo_producto END) FROM existencias_bodega "
                       "WHERE Marca IS NOT NULL AND TRIM(Marca) <> '' GROUP BY TRIM(Marca) ORDER BY TRIM(Marca)")
        return cursor.fetchall()
    finally:
        cursor.close()
        conn.close()


def clasificacion():
    """({código: (segmento, familia, subfamilia)} de clase_prod, {código: Equivalencias} del catálogo de productos)."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT codigo, segmento, familia, subfamilia FROM clase_prod")
        clase = {str(c).strip(): (s, fa, sf) for c, s, fa, sf in cursor.fetchall()}
        cursor.execute("SELECT `Codigo Producto`, Equivalencias FROM catalogo_productos "
                       "WHERE Equivalencias IS NOT NULL AND TRIM(Equivalencias) <> ''")
        equivalencias = {str(c).strip(): " ".join(e.split()) for c, e in cursor.fetchall() if c}
        return clase, equivalencias
    finally:
        cursor.close()
        conn.close()


def calcular(f):
    where, params = ["(Cant_Final <> 0 OR `Monto Final` <> 0)"], []
    if f["bodegas"]:
        where.append(f"Bodega IN ({', '.join(['%s'] * len(f['bodegas']))})")
        params += f["bodegas"]
    if f["marcas"]:
        where.append(f"TRIM(Marca) IN ({', '.join(['%s'] * len(f['marcas']))})")
        params += f["marcas"]
    for palabra in f["q"].split():
        where.append("(Codigo_producto LIKE %s OR Producto LIKE %s OR `Número Parte` LIKE %s)")
        params += [f"%{palabra}%"] * 3

    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"""
            SELECT Bodega, Codigo_producto, Producto, `Número Parte`, Marca, UOM, Cant_Final, `Monto Final`
            FROM existencias_bodega WHERE {' AND '.join(where)}
        """, params)
        registros = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    productos, bodegas, fuera = {}, {}, 0.0
    for bodega, codigo, nombre, parte, marca, uom, cantidad, monto in registros:
        cantidad, monto = float(cantidad or 0), float(monto or 0)
        # Negativos y saldos de costo sin unidades quedan fuera, pero se informa el monto
        if f["solo_con_existencia"] and cantidad <= 0:
            fuera += monto
            continue
        x = productos.setdefault(codigo, {"codigo": codigo, "nombre": (nombre or "").strip(), "numero_parte": (parte or "").strip(),
                                          "marca": (marca or "").strip() or "Sin marca", "uom": uom or "",
                                          "existencia": 0.0, "costo_total": 0.0, "detalle": []})
        x["existencia"] += cantidad
        x["costo_total"] += monto
        x["detalle"].append((bodega, cantidad))
        b = bodegas.setdefault(bodega, {"bodega": bodega, "productos": 0, "unidades": 0.0, "costo": 0.0})
        b["productos"] += 1
        b["unidades"] += cantidad
        b["costo"] += monto

    filas = list(productos.values())
    for x in filas:
        x["costo_unit"] = x["costo_total"] / x["existencia"] if x["existencia"] > 0 else None
        x["detalle"].sort(key=lambda d: -d[1])
    filas.sort(key=ORDENES[f["orden"]][1])

    total = sum(x["costo_total"] for x in filas)
    por_bodega = sorted(bodegas.values(), key=lambda b: -b["costo"])
    for b in por_bodega:
        b["pct"] = b["costo"] / total * 100 if total else 0
    return filas, por_bodega, {"productos": len(filas), "unidades": sum(x["existencia"] for x in filas), "costo": total,
                               "bodegas": len(por_bodega), "fuera": fuera}


def texto_detalle(x):
    return " · ".join(f"{b}: {c:,.2f}".replace(".00", "") for b, c in x["detalle"])


# -----------------------------
# Vistas
# -----------------------------
@existencias_bodegas_bp.route("/existencias_bodegas")
@requiere_acceso("existencias_bodegas")
def reporte():
    f = leer_filtros(request.args)
    filas, por_bodega, t = calcular(f)
    return render_template("existencias_bodegas.html", f=f, filas=filas[:MAX_FILAS_PANTALLA], t=t, por_bodega=por_bodega,
                           bodegas=lista_bodegas(), marcas=lista_marcas(), ordenes=ORDENES, params=params_url(f),
                           filtros_txt=descripcion_filtros(f), texto_detalle=texto_detalle)


def encabezados(ws, fila, columnas):
    for i, (nombre, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=fila, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre.startswith("Costo total") else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[fila].height = 22


def celda_texto(ws, fila, columna, valor):
    """Texto del inventario sin que un '=' inicial se vuelva fórmula."""
    c = ws.cell(row=fila, column=columna, value=valor or None)
    if c.data_type == "f":
        c.data_type = "s"


@existencias_bodegas_bp.route("/existencias_bodegas/excel")
@requiere_acceso("existencias_bodegas")
def excel():
    f = leer_filtros(request.args)
    filas, por_bodega, t = calcular(f)
    wb = Workbook()
    ws = wb.active
    ws.title = "Existencias"
    ws["A1"] = "Grupo Comersa · Existencias y costo por bodega"
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"] = f"{descripcion_filtros(f)} · Orden: {ORDENES[f['orden']][0]} · Generado {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(size=9, color="52514E")
    ft = 4
    encabezados(ws, ft, [("Código", 12), ("Nº parte", 20), ("Descripción", 42), ("Equivalencias", 34), ("Marca", 16), ("Segmento", 16),
                         ("Familia", 22), ("Subfamilia", 24), ("Unidad", 8), ("Bodegas", 44), ("Existencia", 11), ("Costo unitario C$", 14),
                         ("Costo total C$", 16)])
    clase, equivalencias = clasificacion()
    for n, x in enumerate(filas, start=ft + 1):
        codigo = str(x["codigo"]).strip()
        seg, fam, sub = clase.get(codigo, (None, None, None))
        for col, valor in ((1, x["codigo"]), (2, x["numero_parte"]), (3, x["nombre"]), (4, equivalencias.get(codigo)), (5, x["marca"]),
                           (6, seg or "Sin clasificar"), (7, fam or "Sin clasificar"), (8, sub or "Sin clasificar"), (9, x["uom"]),
                           (10, texto_detalle(x))):
            celda_texto(ws, n, col, valor)
        ws.cell(row=n, column=11, value=x["existencia"]).number_format = "#,##0.##"
        ws.cell(row=n, column=12, value=x["costo_unit"]).number_format = "#,##0.00"
        ws.cell(row=n, column=13, value=x["costo_total"]).number_format = "#,##0.00"
    if filas:
        tot = ft + len(filas) + 1
        ws.cell(row=tot, column=1, value=f"Total · {len(filas):,} productos").font = Font(bold=True)
        for col, formato in ((11, "#,##0.##"), (13, "#,##0.00")):
            letra = get_column_letter(col)
            c = ws.cell(row=tot, column=col, value=f"=SUM({letra}{ft + 1}:{letra}{tot - 1})")
            c.font = Font(bold=True)
            c.number_format = formato
    ws.freeze_panes = ws.cell(row=ft + 1, column=4)
    ws.auto_filter.ref = f"A{ft}:M{ft + max(len(filas), 1)}"

    hoja = wb.create_sheet("Por bodega")
    encabezados(hoja, 1, [("Bodega", 34), ("Productos", 11), ("Unidades", 13), ("Costo total C$", 17), ("% del costo", 11)])
    for n, b in enumerate(por_bodega, start=2):
        celda_texto(hoja, n, 1, b["bodega"])
        hoja.cell(row=n, column=2, value=b["productos"]).number_format = "#,##0"
        hoja.cell(row=n, column=3, value=b["unidades"]).number_format = "#,##0.##"
        hoja.cell(row=n, column=4, value=b["costo"]).number_format = "#,##0.00"
        hoja.cell(row=n, column=5, value=b["pct"] / 100).number_format = "0.0%"
    if por_bodega:
        tot = len(por_bodega) + 2
        hoja.cell(row=tot, column=1, value="Total").font = Font(bold=True)
        for col, formato in ((3, "#,##0.##"), (4, "#,##0.00")):
            letra = get_column_letter(col)
            c = hoja.cell(row=tot, column=col, value=f"=SUM({letra}2:{letra}{tot - 1})")
            c.font = Font(bold=True)
            c.number_format = formato

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"existencias_costo_bodegas_{datetime.now():%Y-%m-%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
