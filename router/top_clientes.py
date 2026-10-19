"""Top de clientes por vendedor: venta neta ajustada (sin IVA, menos devoluciones) en un período."""
import io
from datetime import datetime

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.dashboard import ultima_fecha_con_ventas
from router.ventas_vendedor import EJERCITO, rangos_rapidos

top_clientes_bp = Blueprint("top_clientes", __name__)

TAMANOS = [5, 10, 20]


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def ventas_por_cliente(fecha_inicio, fecha_fin):
    """{(vendedor, cliente): datos} con crédito/contado netos de devoluciones.
    Igual que el reporte de vendedores: el cliente Ejército cuenta como su propio vendedor."""
    vendedor_sql = "CASE WHEN {c} = %s THEN %s ELSE {v} END"
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"""
            SELECT {vendedor_sql.format(c='cliente', v='vendedor')} AS vend, cliente,
                   SUM(CASE WHEN tipo_venta = 'Credito' THEN neto ELSE 0 END),
                   SUM(CASE WHEN tipo_venta = 'Contado' THEN neto ELSE 0 END),
                   COUNT(*), MAX(Fecha)
            FROM ventas_historicas
            WHERE Fecha BETWEEN %s AND %s
            GROUP BY vend, cliente
        """, (EJERCITO, EJERCITO, fecha_inicio, fecha_fin))
        ventas = cursor.fetchall()
        cursor.execute(f"""
            SELECT {vendedor_sql.format(c='cliente', v='vendedor')} AS vend, cliente,
                   SUM(CASE WHEN tipofact = 'Credito' THEN totaldv / 1.15 ELSE 0 END),
                   SUM(CASE WHEN tipofact = 'Contado' THEN totaldv / 1.15 ELSE 0 END)
            FROM devoluciones
            WHERE fechadv BETWEEN %s AND %s AND estado = 'Cerrado'
            GROUP BY vend, cliente
        """, (EJERCITO, EJERCITO, fecha_inicio, fecha_fin))
        devoluciones = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    datos = {}
    for vend, cliente, cred, cont, facturas, ultima in ventas:
        datos[(vend, cliente)] = {"vendedor": vend, "cliente": cliente, "credito": float(cred or 0),
                                  "contado": float(cont or 0), "devoluciones": 0.0,
                                  "facturas": int(facturas), "ultima": ultima}
    for vend, cliente, dcred, dcont in devoluciones:
        d = datos.setdefault((vend, cliente), {"vendedor": vend, "cliente": cliente, "credito": 0.0, "contado": 0.0,
                                               "devoluciones": 0.0, "facturas": 0, "ultima": None})
        d["credito"] -= float(dcred or 0)
        d["contado"] -= float(dcont or 0)
        d["devoluciones"] += float(dcred or 0) + float(dcont or 0)
    for d in datos.values():
        d["total"] = d["credito"] + d["contado"]
    return list(datos.values())


def ranking(clientes, n):
    """Ordena clientes, toma los primeros n y resume el resto."""
    clientes = sorted(clientes, key=lambda c: -c["total"])
    total = sum(c["total"] for c in clientes)
    top, resto = clientes[:n], clientes[n:]
    for i, c in enumerate(top, start=1):
        c["posicion"] = i
        c["pct"] = c["total"] / total * 100 if total else 0
    return {
        "top": top,
        "total": total,
        "credito": sum(c["credito"] for c in clientes),
        "contado": sum(c["contado"] for c in clientes),
        "clientes": len(clientes),
        "facturas": sum(c["facturas"] for c in clientes),
        "resto": {"clientes": len(resto), "total": sum(c["total"] for c in resto),
                  "credito": sum(c["credito"] for c in resto), "contado": sum(c["contado"] for c in resto),
                  "facturas": sum(c["facturas"] for c in resto)} if resto else None,
        "concentracion": sum(c["total"] for c in top) / total * 100 if total else 0,
    }


def datos_reporte(args):
    """Calcula el reporte a partir de los parámetros de la URL (lo usan la pantalla y el Excel)."""
    ultima = ultima_fecha_con_ventas() or datetime.today().date()
    if isinstance(ultima, datetime):
        ultima = ultima.date()
    fecha_fin = fecha_param(args.get("fecha_fin"), ultima)
    fecha_inicio = min(fecha_param(args.get("fecha_inicio"), fecha_fin.replace(day=1)), fecha_fin)
    try:
        n = int(args.get("n", 10))
    except ValueError:
        n = 10
    n = n if n in TAMANOS else 10
    vendedor = (args.get("vendedor") or "").strip() or None

    filas = ventas_por_cliente(fecha_inicio, fecha_fin)
    por_vendedor = {}
    for f in filas:
        por_vendedor.setdefault(f["vendedor"], []).append(f)
    vendedores = sorted(por_vendedor, key=lambda v: -sum(c["total"] for c in por_vendedor[v]))
    if vendedor not in por_vendedor:
        vendedor = None

    if vendedor:
        bloques = [(vendedor, ranking(por_vendedor[vendedor], n))]
        empresa = None
    else:
        bloques = [(v, ranking(por_vendedor[v], n)) for v in vendedores]
        # Top de la empresa: un cliente atendido por varios vendedores se suma
        combinados = {}
        for f in filas:
            c = combinados.setdefault(f["cliente"], {"cliente": f["cliente"], "credito": 0.0, "contado": 0.0, "total": 0.0,
                                                     "facturas": 0, "ultima": None, "vendedores": set()})
            for k in ("credito", "contado", "total", "facturas"):
                c[k] += f[k]
            c["vendedores"].add(f["vendedor"])
            if f["ultima"] and (c["ultima"] is None or f["ultima"] > c["ultima"]):
                c["ultima"] = f["ultima"]
        empresa = ranking(list(combinados.values()), n)

    return {"fecha_inicio": fecha_inicio, "fecha_fin": fecha_fin, "ultima": ultima, "n": n,
            "vendedor": vendedor, "vendedores": vendedores, "bloques": bloques, "empresa": empresa}


@top_clientes_bp.route("/top_clientes_vendedor", methods=["GET"])
@requiere_acceso("top_clientes_vendedor")
def top_clientes_vendedor():
    d = datos_reporte(request.args)
    return render_template(
        "top_clientes_vendedor.html",
        fecha_inicio=d["fecha_inicio"], fecha_fin=d["fecha_fin"], rangos=rangos_rapidos(d["ultima"]),
        vendedor=d["vendedor"], vendedores=d["vendedores"], n=d["n"], tamanos=TAMANOS,
        bloques=d["bloques"], empresa=d["empresa"],
    )


# -----------------------------
# Descarga en Excel
# -----------------------------
VERDE, NARANJA = "027266", "FE6412"
MONEDA, PORCENTAJE, FECHA = "#,##0.00", "0.0%", "dd/mm/yyyy"


def _hoja(wb, titulo, encabezados, filas, subtitulo, formatos, primera=False):
    """Hoja con título, encabezado con colores corporativos, autofiltro y formato por columna."""
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"] = subtitulo
    ws["A2"].font = Font(size=9, color="52514E")
    fila_tit = 4
    for i, (nombre, ancho) in enumerate(encabezados, start=1):
        c = ws.cell(row=fila_tit, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre == "Venta neta" else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[fila_tit].height = 30
    linea = Border(bottom=Side(style="thin", color="E6E4DF"))
    for r, valores in enumerate(filas, start=fila_tit + 1):
        for col, v in enumerate(valores, start=1):
            if formatos.get(col) == MONEDA and isinstance(v, float):
                v = round(v, 2)  # montos a centavos, para que las sumas en Excel cuadren con la pantalla
            c = ws.cell(row=r, column=col, value=v)
            c.border = linea
            if col in formatos:
                c.number_format = formatos[col]
    ws.freeze_panes = ws.cell(row=fila_tit + 1, column=1)
    ws.auto_filter.ref = f"A{fila_tit}:{get_column_letter(len(encabezados))}{fila_tit + max(len(filas), 1)}"
    return ws, fila_tit


def _fila_total(ws, fila_tit, n_filas, columnas):
    """Fila TOTAL con fórmulas SUM. columnas = {número_columna: formato}."""
    if not n_filas:
        return
    tot = fila_tit + n_filas + 1
    ws.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    for col, formato in columnas.items():
        letra = get_column_letter(col)
        c = ws.cell(row=tot, column=col, value=f"=SUM({letra}{fila_tit + 1}:{letra}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = formato


def _cliente(c):
    return c["cliente"].lstrip("*"), ("Ocasional" if c["cliente"].startswith("*") else "")


@top_clientes_bp.route("/top_clientes_vendedor/excel", methods=["GET"])
@requiere_acceso("top_clientes_vendedor")
def descargar_excel():
    d = datos_reporte(request.args)
    n = d["n"]
    subtitulo = (f"Período {d['fecha_inicio']:%d/%m/%Y} al {d['fecha_fin']:%d/%m/%Y} · Top {n} · "
                 f"Venta neta ajustada (sin IVA, menos devoluciones) · Generado {datetime.now():%d/%m/%Y %H:%M}")
    wb = Workbook()
    primera = True

    # 1) Top de la empresa (solo con todos los vendedores)
    if d["empresa"]:
        e = d["empresa"]
        filas = [[c["posicion"], *_cliente(c), len(c["vendedores"]), c["facturas"], c["credito"], c["contado"],
                  c["total"], c["pct"] / 100, c["ultima"]] for c in e["top"]]
        _hoja(wb, "Empresa",
              [("#", 5), ("Cliente", 46), ("Tipo", 10), ("Vendedores", 10), ("Facturas", 9), ("Crédito", 15),
               ("Contado", 15), ("Venta neta", 16), ("% del total", 10), ("Última compra", 13)],
              filas, f"{subtitulo} · {e['clientes']:,} clientes con compra · los {n} mejores concentran "
                     f"{e['concentracion']:.1f}% (total empresa C$ {e['total']:,.2f})",
              {6: MONEDA, 7: MONEDA, 8: MONEDA, 9: PORCENTAJE, 10: FECHA}, primera=True)
        primera = False

    # 2) Top por vendedor: una fila por cliente (lista para filtrar o hacer tablas dinámicas)
    filas = [[vend, c["posicion"], *_cliente(c), c["facturas"], c["credito"], c["contado"], c["total"],
              c["pct"] / 100, c["ultima"]]
             for vend, r in d["bloques"] for c in r["top"]]
    titulo = (f"Top {d['vendedor']}" if d["vendedor"] else "Por vendedor")[:31]
    for prohibido in "[]:*?/\\":
        titulo = titulo.replace(prohibido, "")
    _hoja(wb, titulo,
          [("Vendedor", 28), ("#", 5), ("Cliente", 46), ("Tipo", 10), ("Facturas", 9), ("Crédito", 15),
           ("Contado", 15), ("Venta neta", 16), ("% del vendedor", 11), ("Última compra", 13)],
          filas, subtitulo, {6: MONEDA, 7: MONEDA, 8: MONEDA, 9: PORCENTAJE, 10: FECHA}, primera=primera)

    # 3) Resumen por vendedor
    filas = [[vend, r["clientes"], r["facturas"], r["credito"], r["contado"], r["total"], r["concentracion"] / 100,
              (r["resto"]["clientes"] if r["resto"] else 0), (r["resto"]["total"] if r["resto"] else 0)]
             for vend, r in d["bloques"]]
    ws, ft = _hoja(wb, "Resumen vendedores",
                   [("Vendedor", 28), ("Clientes", 9), ("Facturas", 9), ("Crédito", 15), ("Contado", 15),
                    ("Venta neta", 16), (f"% en top {n}", 10), ("Otros clientes", 10), ("Venta otros clientes", 16)],
                   filas, subtitulo, {4: MONEDA, 5: MONEDA, 6: MONEDA, 7: PORCENTAJE, 9: MONEDA})
    _fila_total(ws, ft, len(filas), {2: "#,##0", 3: "#,##0", 4: MONEDA, 5: MONEDA, 6: MONEDA, 8: "#,##0", 9: MONEDA})

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = "top_clientes_" + (d["vendedor"].replace(" ", "_").replace(".", "") if d["vendedor"] else "vendedores")
    nombre += f"_{d['fecha_inicio']}_al_{d['fecha_fin']}.xlsx"
    return send_file(salida, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
