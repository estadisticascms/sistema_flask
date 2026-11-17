"""Existencias y costo por bodega: una o varias bodegas a la vez, con total al costo y orden por montos.
El costo es el Monto Final de cada bodega (el mismo valor de inventario del dashboard), en córdobas."""
import io
from datetime import datetime

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break

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
        "proveedores": sorted({p.strip() for p in fuente.getlist("proveedor") if p.strip()}),
        "q": (fuente.get("q") or "").strip(),
        "orden": fuente.get("orden") if fuente.get("orden") in ORDENES else "costo_total_desc",
        # por defecto solo unidades reales; la casilla desmarcada manda "0"
        "solo_con_existencia": fuente.get("solo_con_existencia", "1") == "1",
    }


def params_url(f):
    p = [("bodega", b) for b in f["bodegas"]]
    p += [("marca", m) for m in f["marcas"]]
    p += [("proveedor", m) for m in f["proveedores"]]
    p += [(k, f[k]) for k in ("q", "orden") if f[k]]
    return p + [("solo_con_existencia", "1" if f["solo_con_existencia"] else "0")]


def descripcion_filtros(f):
    partes = [("Bodega: " if len(f["bodegas"]) == 1 else "Bodegas: ") + ", ".join(f["bodegas"]) if f["bodegas"] else "Todas las bodegas"]
    if f["proveedores"]:
        partes.append(("Proveedor: " if len(f["proveedores"]) == 1 else "Proveedores: ") + ", ".join(f["proveedores"]))
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


def lista_proveedores():
    """Proveedores de última compra (catálogo) con cuántos productos tienen existencia."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT TRIM(p.`Proveedor Ultima Compra`), COUNT(DISTINCT e.Codigo_producto)
            FROM existencias_bodega e JOIN catalogo_productos p ON p.`Codigo Producto` = e.Codigo_producto
            WHERE e.Cant_Final > 0 AND p.`Proveedor Ultima Compra` IS NOT NULL AND TRIM(p.`Proveedor Ultima Compra`) <> ''
            GROUP BY 1 ORDER BY 1
        """)
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
    if f["proveedores"]:
        where.append(f"Codigo_producto IN (SELECT `Codigo Producto` FROM catalogo_productos "
                     f"WHERE TRIM(`Proveedor Ultima Compra`) IN ({', '.join(['%s'] * len(f['proveedores']))}))")
        params += f["proveedores"]
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
                           bodegas=lista_bodegas(), marcas=lista_marcas(), proveedores=lista_proveedores(), ordenes=ORDENES, params=params_url(f),
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


# -----------------------------
# Excel para conteo físico: índice + por familia y subfamilia, una columna por bodega, total del sistema, conteo físico y diferencia
# -----------------------------
SIN_FAMILIA = "Sin clasificar"
SIN_SUBFAMILIA = "Sin subfamilia"


@existencias_bodegas_bp.route("/existencias_bodegas/conteo")
@requiere_acceso("existencias_bodegas")
def excel_conteo():
    f = leer_filtros(request.args)
    filas, _, _ = calcular(f)
    clase, equivalencias = clasificacion()
    bodegas = sorted({b for x in filas for b, _ in x["detalle"]})
    nb = len(bodegas)

    # Columnas: Código, Nº parte, Equivalencia, Descripción, Marca, Unidad, <bodegas>, Total sistema, Conteo físico, Diferencia
    c_bod = 7
    c_tot, c_conteo, c_dif = c_bod + nb, c_bod + nb + 1, c_bod + nb + 2
    L = get_column_letter
    fino, grueso = Side(style="thin", color="000000"), Side(style="medium", color="000000")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq = Alignment(horizontal="left", vertical="center", wrap_text=True)
    NUM = '#,##0;-#,##0;"-"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Conteo físico"
    ws["A1"] = "Grupo Comersa · Inventario por familia y bodega"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = (f"{descripcion_filtros(f)} · Existencias del sistema al {datetime.now():%d/%m/%Y %H:%M}. "
                f"Columna {L(c_conteo)} para el conteo físico; {L(c_dif)} calcula la diferencia (conteo − sistema).")
    ws["A2"].font = Font(italic=True, size=9)

    enc = 4
    nombres = ["Código", "Nº parte", "Equivalencia", "Descripción", "Marca", "Unidad"] + bodegas + ["Total sistema", "Conteo físico", "Diferencia"]
    anchos = [11, 18, 34, 40, 11, 7] + [8] * nb + [8, 13, 11]
    for i, (nombre, ancho) in enumerate(zip(nombres, anchos), start=1):
        c = ws.cell(row=enc, column=i, value=nombre)
        c.font = Font(bold=True, size=9)
        c.border = Border(left=fino, right=fino, top=grueso, bottom=grueso)
        vertical = c_bod <= i <= c_tot
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90 if vertical else 0)
        ws.column_dimensions[L(i)].width = ancho
    ws.row_dimensions[enc].height = 70 if nb else 30

    # Agrupar por familia y subfamilia (sin clasificar al final) y ordenar por número de parte
    grupos = {}
    for x in filas:
        _, fam, sub = clase.get(str(x["codigo"]).strip(), (None, None, None))
        fam = (fam or "").strip() or SIN_FAMILIA
        sub = (sub or "").strip() or SIN_SUBFAMILIA
        grupos.setdefault(fam, {}).setdefault(sub, []).append(x)
    orden_fam = sorted(grupos, key=lambda k: (k == SIN_FAMILIA, k.upper()))

    def titulo(n, texto, tamano, relleno):
        c = ws.cell(row=n, column=1, value=texto)
        c.font = Font(bold=True, size=tamano)
        ws.merge_cells(start_row=n, start_column=1, end_row=n, end_column=c_dif)
        for col in range(1, c_dif + 1):
            celda = ws.cell(row=n, column=col)
            celda.border = Border(top=grueso, bottom=grueso, left=grueso if col == 1 else None,
                                  right=grueso if col == c_dif else None)
            if relleno:
                celda.fill = PatternFill("solid", fgColor=relleno)
        ws.row_dimensions[n].height = 18

    def subtotal(n, texto, articulos, ini, fin, tamano, relleno=None):
        # SUBTOTAL ignora los subtotales anidados, así familia y total general no cuentan dos veces
        ws.cell(row=n, column=1, value=texto)
        ws.cell(row=n, column=6, value=f"{articulos:,} artículo{'s' if articulos != 1 else ''}")
        for col in sumables:
            ws.cell(row=n, column=col, value=f"=SUBTOTAL(9,{L(col)}{ini}:{L(col)}{fin})").number_format = NUM
        for col in range(1, c_dif + 1):
            c = ws.cell(row=n, column=col)
            c.font = Font(bold=True, size=tamano)
            c.border = Border(left=fino, right=fino, top=grueso, bottom=Side(style="double", color="000000"))
            if relleno:
                c.fill = PatternFill("solid", fgColor=relleno)
            if col >= 6:
                c.alignment = centro

    def filas_productos(ws, n, prods):
        for x in prods:
            codigo = str(x["codigo"]).strip()
            por_bod = {}
            for b, cant in x["detalle"]:
                por_bod[b] = por_bod.get(b, 0) + cant
            for col, valor in ((1, x["codigo"]), (2, x["numero_parte"]), (3, equivalencias.get(codigo)), (4, x["nombre"]),
                               (5, x["marca"]), (6, x["uom"])):
                celda_texto(ws, n, col, valor)
                ws.cell(row=n, column=col).alignment = centro if col in (5, 6) else izq
            for j, b in enumerate(bodegas):
                ws.cell(row=n, column=c_bod + j, value=por_bod.get(b, 0))
            ws.cell(row=n, column=c_tot, value=f"=SUM({L(c_bod)}{n}:{L(c_tot - 1)}{n})" if nb else x["existencia"])
            ws.cell(row=n, column=c_dif, value=f'=IF({L(c_conteo)}{n}="","",{L(c_conteo)}{n}-{L(c_tot)}{n})')
            for col in range(1, c_dif + 1):
                c = ws.cell(row=n, column=col)
                c.border = borde
                c.font = Font(size=8, bold=(col == c_tot))
                if col >= c_bod:
                    c.number_format = NUM
                    c.alignment = centro
            n += 1
        return n

    n = enc + 1
    sumables = list(range(c_bod, c_conteo + 1)) + [c_dif]
    primera_prod = n
    indice = []                                   # (familia, subfamilia | None, artículos, existencia, fila)
    for fam in orden_fam:
        subs = grupos[fam]
        fila_fam = ini_fam = n                    # sin fila de familia: cada subfamilia lleva "FAMILIA | Subfamilia"
        pos_fam = len(indice)
        for sub in sorted(subs, key=lambda k: (k == SIN_SUBFAMILIA, k.upper())):
            prods = sorted(subs[sub], key=lambda x: (x["numero_parte"] or "", x["codigo"]))
            indice.append((fam, sub, len(prods), sum(x["existencia"] for x in prods), n))
            titulo(n, f"{fam.upper()} | {sub}", 11, "EDEDED")
            n += 1
            n = filas_productos(ws, n, prods)
            subtotal(n, f"Subtotal {sub}", len(prods), n - len(prods), n - 1, 8)
            n += 1
        articulos = sum(len(p) for p in subs.values())
        indice.insert(pos_fam, (fam, None, articulos, sum(x["existencia"] for p in subs.values() for x in p), fila_fam))
        subtotal(n, f"Subtotal familia {fam}", articulos, ini_fam, n - 1, 9, "D9D9D9")
        ws.row_breaks.append(Break(id=n))          # cada familia empieza en hoja nueva al imprimir
        n += 1
    if filas:
        n += 1
        ws.cell(row=n, column=1, value="TOTAL GENERAL")
        ws.cell(row=n, column=6, value=f"{len(filas):,} artículos")
        for col in sumables:
            ws.cell(row=n, column=col, value=f"=SUBTOTAL(9,{L(col)}{primera_prod}:{L(col)}{n - 2})").number_format = NUM
        for col in range(1, c_dif + 1):
            c = ws.cell(row=n, column=col)
            c.font = Font(bold=True, size=9)
            c.border = Border(top=grueso, bottom=grueso, left=grueso if col == 1 else fino, right=grueso if col == c_dif else fino)
            if col >= 6:
                c.alignment = centro
    else:
        ws.cell(row=n, column=1, value="No hay productos con esos filtros.")

    hoja_indice(wb, ws, f, indice, len(filas), sum(x["existencia"] for x in filas))

    # Impresión: horizontal, a lo ancho de una hoja, encabezado repetido y número de página
    ws.freeze_panes = ws.cell(row=enc + 1, column=1)
    ws.print_title_rows = f"{enc}:{enc}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top, ws.page_margins.bottom = 0.5, 0.6
    ws.oddFooter.center.text = "Página &P de &N · &F"
    ws.oddFooter.center.size = 8

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = "_".join(x for x in ["inventario_conteo", "_".join(f["proveedores"] or f["marcas"])[:40], f"{datetime.now():%Y-%m-%d}"] if x)
    nombre = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in nombre)
    return send_file(salida, as_attachment=True, download_name=f"{nombre}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def hoja_indice(wb, ws, f, indice, articulos, existencia):
    """Primera hoja: familia, subfamilia, artículos y existencias, con vínculo a la sección del conteo."""
    hi = wb.create_sheet("Índice", 0)
    wb.active = 0
    hi["A1"] = "Grupo Comersa · Índice del inventario por familia y subfamilia"
    hi["A1"].font = Font(bold=True, size=13, color=VERDE)
    hi["A2"] = f"{descripcion_filtros(f)} · Existencias del sistema al {datetime.now():%d/%m/%Y %H:%M}. Clic en el nombre para ir a su sección del conteo."
    hi["A2"].font = Font(italic=True, size=9, color="52514E")
    fino = Side(style="thin", color="BFBFBF")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    encabezados(hi, 4, [("Familia | Subfamilia", 58), ("Nº de ítems", 12), ("Total existencias", 16)])
    hoja = ws.title.replace("'", "''")
    # Una fila por subfamilia ("Familia | Subfamilia") y al cierre de cada familia su total
    filas_indice, total_fam = [], None
    for fam, sub, items, exist, fila in indice + [(None, None, 0, 0, 0)]:
        if sub is None:
            if total_fam:
                filas_indice.append(total_fam)
            total_fam = (f"Total {fam}", items, exist, fila, True) if fam is not None else None
        else:
            filas_indice.append((f"{fam} | {sub}", items, exist, fila, False))
    n = 5
    for texto, items, exist, fila, es_fam in filas_indice:
        for col, valor in enumerate((texto, items, exist), start=1):
            c = hi.cell(row=n, column=col, value=valor)
            if c.data_type == "f":
                c.data_type = "s"
            c.border = borde
            c.font = Font(bold=es_fam, size=10 if es_fam else 9)
            if es_fam:
                c.fill = PatternFill("solid", fgColor="E6F1EF")
            if col >= 2:
                c.number_format = "#,##0" if col == 2 else "#,##0.##"
        vinculo = hi.cell(row=n, column=1)
        vinculo.hyperlink = Hyperlink(ref=vinculo.coordinate, location=f"'{hoja}'!A{fila}")
        vinculo.font = Font(bold=es_fam, size=10 if es_fam else 9, color="0563C1", underline="single")
        n += 1
    hi.cell(row=n, column=1, value="TOTAL GENERAL")
    hi.cell(row=n, column=2, value=articulos).number_format = "#,##0"
    hi.cell(row=n, column=3, value=existencia).number_format = "#,##0.##"
    for col in range(1, 4):
        c = hi.cell(row=n, column=col)
        c.font = Font(bold=True)
        c.border = Border(top=Side(style="medium"), bottom=Side(style="double"))
    hi.freeze_panes = "A5"
    hi.print_title_rows = "4:4"
    hi.page_setup.fitToWidth, hi.page_setup.fitToHeight = 1, 0
    hi.sheet_properties.pageSetUpPr.fitToPage = True
    hi.oddFooter.center.text = "Página &P de &N"
    hi.oddFooter.center.size = 8
