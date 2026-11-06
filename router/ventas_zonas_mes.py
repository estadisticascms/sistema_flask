"""Ventas por vendedor y departamento, comparativo por mes (URL y permiso: ventas_zonas_mes).

Nivel 1: departamentos × meses, con mapa, vendedor principal y los segmentos de producto de cada departamento (se despliegan).
Nivel 2: clientes del departamento × meses con su vendedor, y los segmentos del departamento por mes.
Ventana: clientes de un departamento -> productos de un cliente, con edición de segmento / familia / subfamilia
(la misma ventana de Ventas por departamento: static/lista_ventas.js).

Venta neta sin IVA (Sub-Total 2 de ventas_por_producto) sin facturas anuladas. El departamento sale de la ficha del cliente,
normalizado igual que en Ventas por departamento (se reconocen municipios y nombres mal escritos)."""
import io
import time
from datetime import date, datetime

from flask import Blueprint, render_template, request, send_file, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.clase_prod import opciones as opciones_clase
from router.dashboard import ultima_fecha_con_ventas
from router.ventas_bodegas import en_cache
from router.ventas_departamento import (DEPARTAMENTOS, OCASIONAL, SIN_DEPTO, SIN_SEGMENTO, cartera_por_departamento, catalogo_basico,
                                        clasificacion, departamentos_clientes, mapa, normal, tiene_acceso)
from router.ventas_departamento_mes import (clave, fecha_param, fichas_clientes, lista_vendedores, meses_entre, nombre_depto,
                                            restar_meses)
from router.ventas_vendedor import rangos_rapidos

ventas_zonas_mes_bp = Blueprint("ventas_zonas_mes", __name__)

FUERA = (SIN_DEPTO, OCASIONAL)   # no se pintan en el mapa
GRUPOS = DEPARTAMENTOS + [SIN_DEPTO, OCASIONAL]
VERDE, NARANJA = "027266", "FE6412"


def compacto(texto):
    return normal(texto).replace(" ", "")


def depto_param(valor):
    valor = (valor or "").strip().upper()
    return next((g for g in GRUPOS if g.upper() == valor), "")


def nombre(g):
    return g if g in FUERA else nombre_depto(g)


_VENTAS = {}


def ventas_periodo(desde, hasta):
    """(cliente, vendedor, mes, código, monto, unidades, última fecha, descripción); 5 minutos en memoria."""
    llave, ahora = (desde, hasta), time.time()
    if llave in _VENTAS and ahora - _VENTAS[llave][0] < 300:
        return _VENTAS[llave][1]
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT TRIM(Cliente), TRIM(Vendedor), LEFT(Fecha, 7), Codigo, SUM(`Sub-Total 2`), SUM(`Cantidad Vendida`),
                   MAX(Fecha), MAX(`Descrip Prod.`)
            FROM ventas_por_producto WHERE Fecha BETWEEN %s AND %s AND Estatus <> 'Anulado'
            GROUP BY TRIM(Cliente), TRIM(Vendedor), LEFT(Fecha, 7), Codigo
        """, (desde.isoformat(), hasta.isoformat()))
        filas = [(c, v or "", m, cod, float(monto or 0), float(u or 0), str(ult)[:10] if ult else None, desc)
                 for c, v, m, cod, monto, u, ult, desc in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()
    for k in [k for k, v in _VENTAS.items() if ahora - v[0] >= 300]:
        del _VENTAS[k]
    _VENTAS[llave] = (ahora, filas)
    return filas


def leer_filtros(args):
    ultima = ultima_fecha_con_ventas() or date.today()
    if isinstance(ultima, datetime):
        ultima = ultima.date()
    hasta = fecha_param(args.get("hasta"), ultima)
    desde = min(fecha_param(args.get("desde"), restar_meses(hasta, 5)), hasta)   # por defecto: 6 meses
    return {"desde": desde, "hasta": hasta, "ultima": ultima, "vendedor": (args.get("vendedor") or "").strip(),
            "depto": depto_param(args.get("depto")),
            "segmento": (args.get("segmento") or "").strip(), "familia": "", "q": (args.get("q") or "").strip()}


def filas_filtradas(f):
    """Ventas del período con vendedor y segmento aplicados: (departamento, cliente, vendedor, mes, código, seg, fam, sub, monto, unidades, última, desc)."""
    deptos, clase = departamentos_clientes(), clasificacion()
    for cliente, vendedor, mes, codigo, monto, unidades, ultima, desc in ventas_periodo(f["desde"], f["hasta"]):
        if f["vendedor"] and vendedor != f["vendedor"]:
            continue
        seg, fam, sub = clase.get(codigo, (SIN_SEGMENTO, SIN_SEGMENTO, SIN_SEGMENTO))
        if f["segmento"] and seg != f["segmento"]:
            continue
        depto = deptos.get(clave(cliente), OCASIONAL)
        yield depto, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc


def principal(montos):
    """Nombre con más venta y los demás (de mayor a menor)."""
    orden = [n for n, _ in sorted(montos.items(), key=lambda x: -x[1]) if n]
    return (orden[0] if orden else ""), orden[1:]


def calcular(f):
    meses = meses_entre(f["desde"], f["hasta"])
    claves = [m for m, _ in meses]
    fichas = fichas_clientes()
    por_grupo, por_cliente, segs_total = {}, {}, {}
    for depto, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc in filas_filtradas(f):
        z = por_grupo.setdefault(depto, {"meses": {}, "total": 0.0, "clientes": set(), "segmentos": {}, "vendedores": {}})
        z["meses"][mes] = z["meses"].get(mes, 0.0) + monto
        z["total"] += monto
        z["clientes"].add(clave(cliente))
        s = z["segmentos"].setdefault(seg, {"segmento": seg, "meses": {}, "total": 0.0})
        s["meses"][mes] = s["meses"].get(mes, 0.0) + monto
        s["total"] += monto
        z["vendedores"][vendedor] = z["vendedores"].get(vendedor, 0.0) + monto
        segs_total[seg] = segs_total.get(seg, 0.0) + monto
        if f["depto"] and depto == f["depto"]:
            c = por_cliente.setdefault(clave(cliente), {"cliente": cliente, "meses": {}, "total": 0.0, "vendedores": {},
                                                        "ultima": None, "segmentos": {}})
            c["meses"][mes] = c["meses"].get(mes, 0.0) + monto
            c["total"] += monto
            c["vendedores"][vendedor] = c["vendedores"].get(vendedor, 0.0) + monto
            c["segmentos"][seg] = c["segmentos"].get(seg, 0.0) + monto
            if ultima and (c["ultima"] is None or ultima > c["ultima"]):
                c["ultima"] = ultima

    total = sum(z["total"] for z in por_grupo.values())
    cartera = cartera_por_departamento()
    filas = []
    for g in GRUPOS:
        z = por_grupo.get(g, {"meses": {}, "total": 0.0, "clientes": set(), "segmentos": {}, "vendedores": {}})
        c = cartera.get(g, {"total": 0, "activos": 0})
        vend, otros = principal(z["vendedores"])
        segmentos = sorted(z["segmentos"].values(), key=lambda x: -x["total"])
        for s in segmentos:
            s["pct"] = s["total"] / z["total"] * 100 if z["total"] else 0
        filas.append({"depto": g, "nombre": nombre(g), "meses": z["meses"], "total": z["total"], "compraron": len(z["clientes"]),
                      "cartera": c["total"], "activos": c["activos"],
                      "pct": z["total"] / total * 100 if total else 0, "segmentos": segmentos,
                      "vendedor": vend, "vendedor_pct": z["vendedores"].get(vend, 0) / z["total"] * 100 if z["total"] else 0,
                      "otros_vendedores": otros})
    filas.sort(key=lambda x: (x["depto"] in FUERA, x["depto"] == OCASIONAL, -x["total"]))
    for x in filas:   # comparativo: último mes contra el anterior
        ult, ant = (x["meses"].get(claves[-1], 0), x["meses"].get(claves[-2], 0)) if len(claves) > 1 else (0, 0)
        x["variacion"] = (ult - ant) / ant * 100 if ant else None

    fila_sel = next((x for x in filas if x["depto"] == f["depto"]), None)
    clientes = []
    if f["depto"]:
        for k, c in por_cliente.items():
            vend, otros = principal(c["vendedores"])
            ficha = fichas.get(k, {})
            seg1 = max(c["segmentos"].items(), key=lambda x: x[1])[0] if c["segmentos"] else ""
            clientes.append({**c, "vendedor": vend or ficha.get("vendedor", ""), "otros_vendedores": otros,
                             "asignado": ficha.get("vendedor", ""), "segmento": seg1})
        for palabra in f["q"].upper().split():
            clientes = [c for c in clientes if palabra in (c["cliente"] or "").upper() or palabra in (c["vendedor"] or "").upper()]
        clientes.sort(key=lambda c: (-c["total"], c["cliente"] or ""))
        total_grupo = fila_sel["total"] if fila_sel else 0
        for i, c in enumerate(clientes, start=1):
            c["posicion"] = i
            c["pct"] = c["total"] / total_grupo * 100 if total_grupo else 0

    siguiente = date(f["hasta"].year + f["hasta"].month // 12, f["hasta"].month % 12 + 1, 1)
    return {"meses": meses, "filas": filas, "total": total, "clientes": clientes, "fila_sel": fila_sel,
            "por_mes": {m: sum(x["meses"].get(m, 0) for x in filas) for m in claves},
            "clientes_mes": {m: sum(c["meses"].get(m, 0) for c in clientes) for m in claves},
            "segmentos_todos": [s for s, _ in sorted(segs_total.items(), key=lambda x: -x[1])] if not f["segmento"] else [f["segmento"]],
            "parcial": (siguiente - f["hasta"]).days > 1,
            "maximo": max((v for x in filas if x["depto"] not in FUERA for v in x["meses"].values()), default=0),
            "mapa_datos": {x["depto"]: {"nombre": x["nombre"], "total": round(x["total"], 2), "clientes": x["compraron"], "cartera": x["cartera"],
                                        "vendedor": x["vendedor"], "vendedor_pct": round(x["vendedor_pct"]),
                                        "segmentos": [[s["segmento"], round(s["total"], 2)] for s in x["segmentos"][:6]]}
                           for x in filas if x["depto"] not in FUERA and x["total"] > 0},
            "maximo_cliente": max((v for c in clientes for v in c["meses"].values()), default=0)}


def todos_los_segmentos():
    def leer():
        return sorted({s for s, _, _ in clasificacion().values()}, key=lambda s: (s == SIN_SEGMENTO, s))
    return en_cache("vzm_segmentos", leer)


@ventas_zonas_mes_bp.route("/ventas_zonas_mes")
@requiere_acceso("ventas_zonas_mes")
def reporte_ventas_zonas_mes():
    f = leer_filtros(request.args)
    d = calcular(f)
    return render_template("ventas_zonas_mes.html", f=f, d=d, vendedores=lista_vendedores(), grupos=GRUPOS, nombre=nombre,
                           fuera=FUERA, ocasional=OCASIONAL, segmentos=todos_los_segmentos(), mapa=mapa(),
                           rangos=rangos_rapidos(f["ultima"]))


def lista_grupo(f, depto, tipo, solo_cliente=None):
    """Clientes (con su vendedor) o productos de un departamento; con solo_cliente, los productos de ese cliente."""
    catalogo = catalogo_basico() if tipo == "productos" else {}
    filas, total = {}, 0.0
    for z, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc in filas_filtradas(f):
        if z != depto or (solo_cliente and cliente != solo_cliente):
            continue
        total += monto
        if tipo == "clientes":
            x = filas.setdefault(cliente, {"cliente": cliente, "monto": 0.0, "unidades": 0.0, "productos": set(), "vendedores": {}})
            x["productos"].add(codigo)
            x["vendedores"][vendedor] = x["vendedores"].get(vendedor, 0.0) + monto
        else:
            numero_parte, marca = catalogo.get(codigo, ("", ""))
            x = filas.setdefault(codigo, {"codigo": codigo, "descripcion": desc, "numero_parte": numero_parte, "marca": marca,
                                          "segmento": seg, "familia": fam, "subfamilia": sub,
                                          "monto": 0.0, "unidades": 0.0, "clientes": set()})
            x["clientes"].add(cliente)
        x["monto"] += monto
        x["unidades"] += unidades
    lista = sorted(filas.values(), key=lambda x: -x["monto"])
    for x in lista:
        for campo in ("productos", "clientes"):
            if campo in x:
                x[campo] = len(x[campo])
        if "vendedores" in x:
            x["vendedor"], x["otros_vendedores"] = principal(x.pop("vendedores"))
        x["pct"] = x["monto"] / total * 100 if total else 0
    return lista, total


def texto_filtro(f):
    return " · ".join(x for x in (f["vendedor"], f["segmento"]) if x)


def hoja_excel(ws, titulo, nota, columnas):
    ws["A1"], ws["A2"] = titulo, nota
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"].font = Font(size=9, color="52514E")
    for i, (nombre, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=4, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre == "Total C$" else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "B5"


def totales_excel(ws, fila, desde_col, hasta_col):
    ws.cell(row=fila, column=1, value="TOTAL").font = Font(bold=True)
    for i in range(desde_col, hasta_col + 1):
        l = get_column_letter(i)
        c = ws.cell(row=fila, column=i, value=f"=SUM({l}5:{l}{fila - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00"


def enviar(wb, nombre):
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def nota_excel(f):
    filtro = texto_filtro(f)
    return (f"Del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y} · venta neta sin IVA, sin anuladas"
            f"{' · ' + filtro if filtro else ''} · Generado {datetime.now():%d/%m/%Y %H:%M}")


@ventas_zonas_mes_bp.route("/ventas_zonas_mes/lista")
@requiere_acceso("ventas_zonas_mes")
def lista():
    """Clientes o productos de un departamento para la ventana (HTML) o en Excel."""
    f = leer_filtros(request.args)
    tipo = "productos" if request.args.get("tipo") == "productos" else "clientes"
    depto = depto_param(request.args.get("ver"))
    if not depto:
        return "Departamento no válido", 400
    cliente = (request.args.get("cliente") or "").strip() if tipo == "productos" else ""
    filas, total = lista_grupo(f, depto, tipo, cliente or None)
    if request.args.get("formato") != "excel":
        editar, segmentos, arbol = tipo == "productos" and tiene_acceso("clase_prod"), [], {}
        if editar:
            conn = get_connection()
            cursor = conn.cursor(dictionary=True)
            try:
                segmentos, arbol = opciones_clase(cursor)
            finally:
                cursor.close()
                conn.close()
        return render_template("ventas_departamento_lista.html", f=f, tipo=tipo, depto=depto, filas=filas, total=total,
                               filtro=texto_filtro(f), editar=editar, segmentos=segmentos, arbol=arbol, sin_segmento=SIN_SEGMENTO,
                               cliente=cliente, unidad="depto.",
                               excel_url=url_for("ventas_zonas_mes.lista", formato="excel", **request.args.to_dict()))

    wb = Workbook()
    ws = wb.active
    ws.title = tipo.capitalize()
    titulo = f"{'Clientes' if tipo == 'clientes' else 'Productos'} de {cliente + ' · ' if cliente else ''}{nombre(depto)}"
    if tipo == "clientes":
        columnas = [("#", 5), ("Cliente", 42), ("Vendedor", 28), ("Otros vendedores", 26), ("Total C$", 15), ("% del depto.", 10),
                    ("Unidades", 11), ("Productos distintos", 11)]
        valores = [[n, x["cliente"], x["vendedor"], ", ".join(x["otros_vendedores"]) or None, round(x["monto"], 2), x["pct"] / 100,
                    x["unidades"], x["productos"]] for n, x in enumerate(filas, 1)]
    else:
        columnas = [("#", 5), ("Código", 12), ("Nº parte", 20), ("Descripción", 45), ("Marca", 16), ("Segmento", 20), ("Familia", 26),
                    ("Subfamilia", 30), ("Total C$", 15), ("% del cliente" if cliente else "% del depto.", 10), ("Unidades", 11), ("Clientes", 10)]
        valores = [[n, x["codigo"], x["numero_parte"], x["descripcion"], x["marca"], x["segmento"], x["familia"], x["subfamilia"],
                    round(x["monto"], 2), x["pct"] / 100, x["unidades"], x["clientes"]] for n, x in enumerate(filas, 1)]
        if cliente:
            columnas, valores = columnas[:-1], [v[:-1] for v in valores]
    hoja_excel(ws, titulo, nota_excel(f), columnas)
    for r, fila in enumerate(valores, start=5):
        for i, v in enumerate(fila, start=1):
            c = ws.cell(row=r, column=i, value=v)
            encab = columnas[i - 1][0]
            c.number_format = ("#,##0.00" if encab == "Total C$" else "0.0%" if encab.startswith("%")
                               else "#,##0.##" if encab == "Unidades" else c.number_format)
    col = [c[0] for c in columnas].index("Total C$") + 1
    fila_total = 5 + len(valores)
    ws.cell(row=fila_total, column=col - 1, value="TOTAL").font = Font(bold=True)
    c = ws.cell(row=fila_total, column=col, value=round(total, 2))
    c.font = Font(bold=True)
    c.number_format = "#,##0.00"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(columnas))}{max(fila_total - 1, 5)}"
    return enviar(wb, f"{tipo}_{compacto(cliente or depto).lower()}_{f['desde']}_al_{f['hasta']}.xlsx")


@ventas_zonas_mes_bp.route("/ventas_zonas_mes/excel")
@requiere_acceso("ventas_zonas_mes")
def descargar_excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    meses = d["meses"]
    nota = nota_excel(f)
    wb = Workbook()

    ws = wb.active
    ws.title = "Departamentos"
    hoja_excel(ws, "Ventas por vendedor y departamento, por mes", nota, [("Departamento", 20)] + [(e, 13) for _, e in meses]
               + [("Total C$", 15), ("% del total", 9), ("Vendedor principal", 28), ("% del vendedor", 9), ("Compraron", 10),
                  ("Clientes en cartera", 10), ("Var. último mes", 10)])
    for n, x in enumerate(d["filas"], start=5):
        valores = ([x["nombre"]] + [round(x["meses"].get(m, 0), 2) or None for m, _ in meses]
                   + [round(x["total"], 2), x["pct"] / 100, x["vendedor"] or None, x["vendedor_pct"] / 100 if x["vendedor"] else None,
                      x["compraron"], x["cartera"], x["variacion"] / 100 if x["variacion"] is not None else None])
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=n, column=i, value=v)
            if 1 < i <= len(meses) + 2:
                c.number_format = "#,##0.00"
        for col in (len(meses) + 3, len(meses) + 5):
            ws.cell(row=n, column=col).number_format = "0.0%"
        ws.cell(row=n, column=len(meses) + 8).number_format = "+0.0%;-0.0%"
    totales_excel(ws, 5 + len(d["filas"]), 2, len(meses) + 2)

    ws2 = wb.create_sheet("Segmentos por departamento")
    hoja_excel(ws2, "Segmentos de producto por departamento y mes", nota, [("Departamento", 20), ("Segmento", 24)] + [(e, 13) for _, e in meses]
               + [("Total C$", 15), ("% del depto.", 9)])
    n = 5
    for x in d["filas"]:
        if f["depto"] and x["depto"] != f["depto"]:
            continue
        for s in x["segmentos"]:
            valores = [x["nombre"], s["segmento"]] + [round(s["meses"].get(m, 0), 2) or None for m, _ in meses] + [round(s["total"], 2), s["pct"] / 100]
            for i, v in enumerate(valores, start=1):
                c = ws2.cell(row=n, column=i, value=v)
                if 2 < i <= len(meses) + 3:
                    c.number_format = "#,##0.00"
            ws2.cell(row=n, column=len(meses) + 4).number_format = "0.0%"
            n += 1
    ws2.freeze_panes = "C5"
    ws2.auto_filter.ref = f"A4:{get_column_letter(len(meses) + 4)}{max(n - 1, 5)}"

    if f["depto"]:
        ws3 = wb.create_sheet("Clientes")
        hoja_excel(ws3, f"Clientes de {nombre(f['depto'])} por mes", nota, [("Cliente", 40), ("Vendedor", 28), ("Otros vendedores", 26),
                                                                          ("Segmento principal", 20)]
                   + [(e, 13) for _, e in meses] + [("Total C$", 15), ("% del depto.", 9), ("Última compra", 12)])
        for n, c in enumerate(d["clientes"], start=5):
            valores = ([c["cliente"], c["vendedor"], ", ".join(c["otros_vendedores"]) or None, c["segmento"]]
                       + [round(c["meses"].get(m, 0), 2) or None for m, _ in meses]
                       + [round(c["total"], 2), c["pct"] / 100, datetime.strptime(c["ultima"], "%Y-%m-%d").date() if c["ultima"] else None])
            for i, v in enumerate(valores, start=1):
                cel = ws3.cell(row=n, column=i, value=v)
                if 4 < i <= len(meses) + 5:
                    cel.number_format = "#,##0.00"
            ws3.cell(row=n, column=len(meses) + 6).number_format = "0.0%"
            ws3.cell(row=n, column=len(meses) + 7).number_format = "dd/mm/yyyy"
        totales_excel(ws3, 5 + len(d["clientes"]), 5, len(meses) + 5)

    sufijo = f"_{compacto(f['depto']).lower()}" if f["depto"] else ""
    return enviar(wb, f"ventas_vendedor_departamento{sufijo}_{f['desde']}_al_{f['hasta']}.xlsx")
