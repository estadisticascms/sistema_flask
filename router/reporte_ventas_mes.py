"""Ventas netas por mes: comparativo de varios meses (de 1 a 12) con vendedores ordenados de mayor a menor venta
(contado, crédito y devoluciones), y ranking de los 10 productos y las 10 marcas con más venta."""
import io
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.reporte_armados import nombre_limpio

reporte_ventas_mes_bp = Blueprint("reporte_ventas_mes", __name__)

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MAX_MESES = 12
TOP = 10
EJERCITO = "EJERCITO DE NICARAGUA"
VERDE, NARANJA = "027266", "FE6412"


def sumar_meses(anio, mes, n):
    total = anio * 12 + (mes - 1) + n
    return total // 12, total % 12 + 1


def leer_mes(texto):
    try:
        d = datetime.strptime((texto or "")[:7], "%Y-%m")
        return d.year, d.month
    except ValueError:
        return None


def ultima_venta():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas")
    ultima = cursor.fetchone()[0]
    cursor.close()
    conn.close()
    return ultima or date.today()


def leer_periodo(args):
    """Meses 'YYYY-MM' (desde / hasta). Acepta los enlaces viejos con fecha_inicio / fecha_fin. Por defecto: los 3 meses
    que terminan en el mes de la última venta cargada."""
    ultima = ultima_venta()
    hasta = leer_mes(args.get("hasta") or args.get("fecha_fin")) or (ultima.year, ultima.month)
    desde = leer_mes(args.get("desde") or args.get("fecha_inicio")) or sumar_meses(*hasta, -2)
    if desde > hasta:
        desde, hasta = hasta, desde
    ajustado = False
    if (hasta[0] * 12 + hasta[1]) - (desde[0] * 12 + desde[1]) >= MAX_MESES:
        desde, ajustado = sumar_meses(*hasta, -(MAX_MESES - 1)), True
    meses = []
    a, m = desde
    while (a, m) <= hasta:
        meses.append(f"{a}-{m:02d}")
        a, m = sumar_meses(a, m, 1)
    inicio = date(*desde, 1)
    fin_mes = date(*sumar_meses(*hasta, 1), 1) - timedelta(days=1)
    fin = min(fin_mes, ultima) if ultima >= inicio else fin_mes
    return {"desde": f"{desde[0]}-{desde[1]:02d}", "hasta": f"{hasta[0]}-{hasta[1]:02d}", "meses": meses,
            "inicio": inicio, "fin": fin, "parcial": fin < fin_mes, "ajustado": ajustado, "ultima": ultima}


def etiqueta(mes):
    return f"{MESES[int(mes[5:]) - 1][:3]} {mes[:4]}"


def variacion(actual, anterior):
    return (actual / anterior - 1) * 100 if anterior else None


# -----------------------------
# Vendedores: contado, crédito y devoluciones (sin IVA) por mes, sin facturas anuladas
# -----------------------------
def ventas_vendedores(p):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT d.Mes, d.Vendedor, MAX(e.tipo), SUM(d.contado), SUM(d.credito), SUM(d.dev_contado), SUM(d.dev_credito)
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes,
                   CASE WHEN v.cliente = '{EJERCITO}' THEN '{EJERCITO}' ELSE v.Vendedor END AS Vendedor,
                   SUM(CASE WHEN v.tipo_venta = 'Contado' THEN v.neto ELSE 0 END) AS contado,
                   SUM(CASE WHEN v.tipo_venta = 'Credito' THEN v.neto ELSE 0 END) AS credito,
                   0 AS dev_contado, 0 AS dev_credito
            FROM ventas_historicas v
            WHERE v.fecha BETWEEN %s AND %s AND v.estado_factura <> 'Anulado'
            GROUP BY 1, 2
            UNION ALL
            SELECT DATE_FORMAT(r.fechadv, '%Y-%m'),
                   CASE WHEN r.cliente = '{EJERCITO}' THEN '{EJERCITO}' ELSE r.Vendedor END,
                   0, 0,
                   SUM(CASE WHEN r.tipofact = 'Contado' THEN r.totaldv ELSE 0 END) / 1.15,
                   SUM(CASE WHEN r.tipofact = 'Credito' THEN r.totaldv ELSE 0 END) / 1.15
            FROM devoluciones r
            WHERE r.fechadv BETWEEN %s AND %s AND r.estado = 'Cerrado'
            GROUP BY 1, 2
        ) d
        LEFT JOIN (SELECT tipo, nombre FROM cat_vendedor GROUP BY tipo, nombre) e ON e.nombre = d.Vendedor
        GROUP BY d.Mes, d.Vendedor
    """, (p["inicio"], p["fin"], p["inicio"], p["fin"]))
    registros = cursor.fetchall()
    cursor.close()
    conn.close()

    vacio = lambda: {"contado": 0.0, "credito": 0.0, "devoluciones": 0.0, "neto": 0.0}
    vendedores, totales = {}, {m: vacio() for m in p["meses"] + ["total"]}
    for mes, vendedor, tipo, contado, credito, dev_co, dev_cr in registros:
        if mes not in totales:
            continue
        x = vendedores.setdefault(vendedor or "Sin vendedor", {
            "vendedor": EJERCITO.title() if vendedor == EJERCITO else nombre_limpio(vendedor),
            "tipo": "Ejército" if vendedor == EJERCITO else (tipo or "Sin tipo").title(),
            "meses": {m: vacio() for m in p["meses"]}, "total": vacio()})
        valores = {"contado": float(contado or 0), "credito": float(credito or 0), "devoluciones": float((dev_co or 0) + (dev_cr or 0))}
        valores["neto"] = valores["contado"] + valores["credito"] - valores["devoluciones"]
        for destino in (x["meses"][mes], x["total"], totales[mes], totales["total"]):
            for k, v in valores.items():
                destino[k] += v
    filas = sorted(vendedores.values(), key=lambda x: -x["total"]["neto"])
    for x in filas:
        x["pct"] = x["total"]["neto"] / totales["total"]["neto"] * 100 if totales["total"]["neto"] else 0
        x["var"] = variacion(x["meses"][p["meses"][-1]]["neto"], x["meses"][p["meses"][-2]]["neto"]) if len(p["meses"]) > 1 else None
    return filas, totales


# -----------------------------
# Productos y marcas: venta facturada (sin devoluciones) por mes, desde ventas_por_producto
# -----------------------------
def ventas_productos(p):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT LEFT(Fecha, 7) AS Mes, Codigo, MAX(`Descrip Prod.`), MAX(Numero_parte), MAX(Marca),
               SUM(`Sub-Total 2`), SUM(`Cantidad Vendida`), SUM(`Cantidad Vendida` * `Costo Unit.`)
        FROM ventas_por_producto
        WHERE Fecha BETWEEN %s AND %s AND Estatus <> 'Anulado'
        GROUP BY Mes, Codigo
    """, (str(p["inicio"]), str(p["fin"])))  # Fecha es texto 'YYYY-MM-DD' en esta tabla
    registros = cursor.fetchall()
    cursor.close()
    conn.close()

    productos, marcas = {}, {}
    total = 0.0
    for mes, codigo, desc, parte, marca, venta, unidades, costo in registros:
        if mes not in p["meses"]:
            continue
        venta, unidades, costo = float(venta or 0), float(unidades or 0), float(costo or 0)
        marca = (marca or "").strip() or "Sin marca"
        total += venta
        x = productos.setdefault(codigo, {"codigo": codigo, "descripcion": (desc or "").strip(), "numero_parte": (parte or "").strip(),
                                          "marca": marca, "meses": {}, "venta": 0.0, "unidades": 0.0, "costo": 0.0})
        y = marcas.setdefault(marca, {"marca": marca, "meses": {}, "venta": 0.0, "unidades": 0.0, "costo": 0.0, "codigos": set()})
        for z in (x, y):
            z["meses"][mes] = z["meses"].get(mes, 0.0) + venta
            z["venta"] += venta
            z["unidades"] += unidades
            z["costo"] += costo
        y["codigos"].add(codigo)

    def ranking(grupo):
        top = sorted(grupo.values(), key=lambda z: -z["venta"])[:TOP]
        for z in top:
            z["margen"] = (z["venta"] - z["costo"]) / z["venta"] * 100 if z["venta"] else None
            z["pct"] = z["venta"] / total * 100 if total else 0
            z["var"] = variacion(z["meses"].get(p["meses"][-1], 0), z["meses"].get(p["meses"][-2], 0)) if len(p["meses"]) > 1 else None
            if "codigos" in z:
                z["productos"] = len(z.pop("codigos"))
        return top

    return ranking(productos), ranking(marcas), total


def calcular(p):
    vendedores, totales = ventas_vendedores(p)
    productos, marcas, total_productos = ventas_productos(p)
    resumen = []
    for i, m in enumerate(p["meses"]):
        t = totales[m]
        resumen.append({"mes": m, "etiqueta": etiqueta(m), **t,
                        "var": variacion(t["neto"], totales[p["meses"][i - 1]]["neto"]) if i else None,
                        "pct_credito": t["credito"] / (t["contado"] + t["credito"]) * 100 if t["contado"] + t["credito"] else None})
    return {"vendedores": vendedores, "totales": totales, "resumen": resumen,
            "productos": productos, "marcas": marcas, "total_productos": total_productos}


# -----------------------------
# Vistas
# -----------------------------
@reporte_ventas_mes_bp.route("/reporte_ventas_mes", methods=["GET"])
@requiere_acceso("reporte_ventas_mes")
def reporte_ventas_mes():
    p = leer_periodo(request.args)
    a, m = map(int, p["hasta"].split("-"))
    rapidos = [(n, f"{sumar_meses(a, m, -(n - 1))[0]}-{sumar_meses(a, m, -(n - 1))[1]:02d}") for n in (3, 6, 12)]
    return render_template("reporte_ventas_mes.html", p=p, r=calcular(p), etiqueta=etiqueta, rapidos=rapidos, max_meses=MAX_MESES)


def encabezados(ws, fila, columnas, naranja=()):
    for i, (nombre, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=fila, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre in naranja else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[fila].height = 30


def titulo(ws, texto, p):
    ws["A1"] = f"Grupo Comersa · {texto}"
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    nota = f" (último mes al {p['fin']:%d/%m/%Y})" if p["parcial"] else ""
    ws["A2"] = f"{etiqueta(p['meses'][0])} a {etiqueta(p['meses'][-1])}{nota} · Generado {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(size=9, color="52514E")


@reporte_ventas_mes_bp.route("/reporte_ventas_mes/excel")
@requiere_acceso("reporte_ventas_mes")
def descargar_excel():
    p = leer_periodo(request.args)
    r = calcular(p)
    wb = Workbook()
    dinero = "#,##0.00"

    # Vendedores: por mes contado, crédito, devoluciones y neto
    ws = wb.active
    ws.title = "Vendedores"
    titulo(ws, "Ventas netas por vendedor", p)
    columnas = [("#", 5), ("Vendedor", 28), ("Tipo", 18)]
    for m in p["meses"] + ["total"]:
        nombre = "Total" if m == "total" else etiqueta(m)
        columnas += [(f"{nombre} Contado", 14), (f"{nombre} Crédito", 14), (f"{nombre} Devol.", 13), (f"{nombre} Neto", 15)]
    columnas += [("% del total", 9)]
    encabezados(ws, 4, columnas, naranja={c for c, _ in columnas if c.endswith("Neto")})
    campos = ("contado", "credito", "devoluciones", "neto")
    for n, x in enumerate(r["vendedores"], start=5):
        fila = [n - 4, x["vendedor"], x["tipo"]]
        for m in p["meses"]:
            fila += [x["meses"][m][k] for k in campos]
        fila += [x["total"][k] for k in campos] + [x["pct"] / 100]
        for i, v in enumerate(fila, start=1):
            c = ws.cell(row=n, column=i, value=v)
            if i > 3:
                c.number_format = "0.0%" if i == len(fila) else dinero
    tot = 5 + len(r["vendedores"])
    ws.cell(row=tot, column=2, value="Total").font = Font(bold=True)
    for i in range(4, len(columnas)):
        letra = get_column_letter(i)
        c = ws.cell(row=tot, column=i, value=f"=SUM({letra}5:{letra}{tot - 1})")
        c.font, c.number_format = Font(bold=True), dinero
    ws.freeze_panes = "D5"

    # Productos y marcas: venta por mes
    for nombre_hoja, filas, clave in (("Top 10 productos", r["productos"], "producto"), ("Top 10 marcas", r["marcas"], "marca")):
        hoja = wb.create_sheet(nombre_hoja)
        titulo(hoja, f"{nombre_hoja} por venta (facturado sin IVA, sin restar devoluciones)", p)
        if clave == "producto":
            columnas = [("#", 5), ("Código", 11), ("Nº parte", 18), ("Descripción", 38), ("Marca", 16)]
        else:
            columnas = [("#", 5), ("Marca", 22), ("Productos vendidos", 11)]
        columnas += [(etiqueta(m), 14) for m in p["meses"]] + [("Total", 15), ("Unidades", 11), ("Margen", 9), ("% de la venta", 10)]
        encabezados(hoja, 4, columnas, naranja={"Total"})
        for n, z in enumerate(filas, start=5):
            fila = [n - 4] + ([z["codigo"], z["numero_parte"], z["descripcion"], z["marca"]] if clave == "producto" else [z["marca"], z["productos"]])
            base = len(fila)
            fila += [z["meses"].get(m, 0.0) for m in p["meses"]] + [z["venta"], z["unidades"],
                                                                     (z["margen"] or 0) / 100, z["pct"] / 100]
            for i, v in enumerate(fila, start=1):
                c = hoja.cell(row=n, column=i, value=v)
                if c.data_type == "f":
                    c.data_type = "s"
                if i > base:
                    c.number_format = "0.0%" if i > len(fila) - 2 else ("#,##0.##" if i == len(fila) - 2 else dinero)
        hoja.freeze_panes = "C5"

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"ventas_netas_{p['desde']}_a_{p['hasta']}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
