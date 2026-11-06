"""Ventas por zona, comparativo por mes (zonas de la tabla zona_cliente: municipio -> zona).

Nivel 1: zonas × meses, con vendedor principal y los segmentos de producto de cada zona (se despliegan).
Nivel 2: clientes de la zona × meses con su vendedor, y los segmentos de la zona por mes.
Ventana: clientes de una zona -> productos de un cliente, con edición de segmento / familia / subfamilia
(la misma ventana de Ventas por departamento: static/lista_ventas.js).

Venta neta sin IVA (Sub-Total 2 de ventas_por_producto) sin facturas anuladas. La zona sale del municipio de la ficha
del cliente; si el municipio no está en zona_cliente, se usa la zona de la cabecera de su departamento."""
import difflib
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
from router.ventas_departamento import (OCASIONAL, SIN_SEGMENTO, catalogo_basico, clasificacion, departamento_de, normal,
                                        tiene_acceso)
from router.ventas_departamento_mes import clave, fecha_param, lista_vendedores, meses_entre, restar_meses
from router.ventas_vendedor import rangos_rapidos

ventas_zonas_mes_bp = Blueprint("ventas_zonas_mes", __name__)

SIN_ZONA = "Sin zona"            # está en la tabla clientes pero sin municipio ni departamento reconocible
FUERA = (SIN_ZONA, OCASIONAL)
VERDE, NARANJA = "027266", "FE6412"
# Cabecera de cada departamento: zona para los clientes cuyo municipio no está en zona_cliente
CABECERA = {"BOACO": "BOACO", "CARAZO": "JINOTEPE", "CHINANDEGA": "CHINANDEGA", "CHONTALES": "JUIGALPA", "ESTELI": "ESTELI",
            "GRANADA": "GRANADA", "JINOTEGA": "JINOTEGA", "LEON": "LEON", "MADRIZ": "SOMOTO", "MANAGUA": "MANAGUA", "MASAYA": "MASAYA",
            "MATAGALPA": "MATAGALPA", "NUEVA SEGOVIA": "OCOTAL", "RACCN": "PUERTO CABEZAS", "RACCS": "BLUEFIELDS",
            "RIO SAN JUAN": "SAN CARLOS", "RIVAS": "RIVAS"}


def compacto(texto):
    return normal(texto).replace(" ", "")


def zonas():
    """Zonas en su orden (grupozona) y municipio (sin espacios ni tildes) -> zona."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT nombre_zona, municipio FROM zona_cliente WHERE nombre_zona <> '' ORDER BY grupozona, idzona")
            orden, municipios = [], {}
            for zona, municipio in cursor.fetchall():
                zona = zona.strip()
                if zona not in orden:
                    orden.append(zona)
                municipios.setdefault(compacto(municipio), zona)   # un municipio repetido se queda con la primera zona
            return {"orden": orden, "municipios": municipios}
        finally:
            cursor.close()
            conn.close()
    return en_cache("vzm_zonas", leer)


def zona_de(departamento, municipio, z):
    """(zona, cómo se asignó): por municipio, o por la cabecera del departamento."""
    for texto in (municipio, departamento):
        k = compacto(texto)
        if len(k) < 3:
            continue
        if k in z["municipios"]:
            return z["municipios"][k], "municipio"
        parecido = difflib.get_close_matches(k, z["municipios"].keys(), n=1, cutoff=0.86)
        if parecido:
            return z["municipios"][parecido[0]], "municipio"
    depto = departamento_de(departamento, municipio)
    if depto and compacto(CABECERA.get(depto)) in z["municipios"]:
        return z["municipios"][compacto(CABECERA[depto])], "departamento"
    return SIN_ZONA, None


def fichas_clientes():
    """Clientes por nombre: zona, vendedor asignado y estado (10 minutos en memoria)."""
    def leer():
        z = zonas()
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT nombre, departamento, municipio, vendedor, estado FROM clientes")
            fichas = {}
            for nombre, depto, municipio, vendedor, estado in cursor.fetchall():
                k = clave(nombre)
                if not k:
                    continue
                actual = fichas.get(k)
                zona, como = zona_de(depto, municipio, z)
                if actual is None or (estado == "Activo" and not actual["activo"]) or (actual["zona"] == SIN_ZONA and zona != SIN_ZONA):
                    fichas[k] = {"nombre": " ".join(nombre.split()), "zona": zona, "como": como,
                                 "vendedor": (vendedor or "").strip(), "activo": estado == "Activo" or bool(actual and actual["activo"])}
            return fichas
        finally:
            cursor.close()
            conn.close()
    return en_cache("vzm_fichas", leer)


def cartera_por_zona():
    cartera = {}
    for ficha in fichas_clientes().values():
        c = cartera.setdefault(ficha["zona"], {"total": 0, "activos": 0, "por_depto": 0})
        c["total"] += 1
        c["activos"] += ficha["activo"]
        c["por_depto"] += ficha["como"] == "departamento"
    return cartera


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
    zona = (args.get("zona") or "").strip()
    return {"desde": desde, "hasta": hasta, "ultima": ultima, "vendedor": (args.get("vendedor") or "").strip(),
            "zona": zona if zona in zonas()["orden"] or zona in FUERA else "",
            "segmento": (args.get("segmento") or "").strip(), "familia": "", "q": (args.get("q") or "").strip()}


def filas_filtradas(f):
    """Ventas del período con vendedor y segmento aplicados: (zona, cliente, vendedor, mes, código, seg, fam, sub, monto, unidades, última, desc)."""
    fichas, clase = fichas_clientes(), clasificacion()
    for cliente, vendedor, mes, codigo, monto, unidades, ultima, desc in ventas_periodo(f["desde"], f["hasta"]):
        if f["vendedor"] and vendedor != f["vendedor"]:
            continue
        seg, fam, sub = clase.get(codigo, (SIN_SEGMENTO, SIN_SEGMENTO, SIN_SEGMENTO))
        if f["segmento"] and seg != f["segmento"]:
            continue
        zona = fichas.get(clave(cliente), {}).get("zona", OCASIONAL)
        yield zona, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc


def principal(montos):
    """Nombre con más venta y los demás (de mayor a menor)."""
    orden = [n for n, _ in sorted(montos.items(), key=lambda x: -x[1]) if n]
    return (orden[0] if orden else ""), orden[1:]


def calcular(f):
    meses = meses_entre(f["desde"], f["hasta"])
    claves = [m for m, _ in meses]
    fichas = fichas_clientes()
    por_zona, por_cliente, segs_total = {}, {}, {}
    for zona, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc in filas_filtradas(f):
        z = por_zona.setdefault(zona, {"meses": {}, "total": 0.0, "clientes": set(), "segmentos": {}, "vendedores": {}})
        z["meses"][mes] = z["meses"].get(mes, 0.0) + monto
        z["total"] += monto
        z["clientes"].add(clave(cliente))
        s = z["segmentos"].setdefault(seg, {"segmento": seg, "meses": {}, "total": 0.0})
        s["meses"][mes] = s["meses"].get(mes, 0.0) + monto
        s["total"] += monto
        z["vendedores"][vendedor] = z["vendedores"].get(vendedor, 0.0) + monto
        segs_total[seg] = segs_total.get(seg, 0.0) + monto
        if f["zona"] and zona == f["zona"]:
            c = por_cliente.setdefault(clave(cliente), {"cliente": cliente, "meses": {}, "total": 0.0, "vendedores": {},
                                                        "ultima": None, "segmentos": {}})
            c["meses"][mes] = c["meses"].get(mes, 0.0) + monto
            c["total"] += monto
            c["vendedores"][vendedor] = c["vendedores"].get(vendedor, 0.0) + monto
            c["segmentos"][seg] = c["segmentos"].get(seg, 0.0) + monto
            if ultima and (c["ultima"] is None or ultima > c["ultima"]):
                c["ultima"] = ultima

    total = sum(z["total"] for z in por_zona.values())
    cartera = cartera_por_zona()
    filas = []
    for nombre in zonas()["orden"] + [SIN_ZONA, OCASIONAL]:
        z = por_zona.get(nombre, {"meses": {}, "total": 0.0, "clientes": set(), "segmentos": {}, "vendedores": {}})
        c = cartera.get(nombre, {"total": 0, "activos": 0, "por_depto": 0})
        vend, otros = principal(z["vendedores"])
        segmentos = sorted(z["segmentos"].values(), key=lambda x: -x["total"])
        for s in segmentos:
            s["pct"] = s["total"] / z["total"] * 100 if z["total"] else 0
        filas.append({"zona": nombre, "meses": z["meses"], "total": z["total"], "compraron": len(z["clientes"]),
                      "cartera": c["total"], "activos": c["activos"], "por_depto": c["por_depto"],
                      "pct": z["total"] / total * 100 if total else 0, "segmentos": segmentos,
                      "vendedor": vend, "vendedor_pct": z["vendedores"].get(vend, 0) / z["total"] * 100 if z["total"] else 0,
                      "otros_vendedores": otros})
    filas.sort(key=lambda x: (x["zona"] in FUERA, x["zona"] == OCASIONAL, -x["total"]))
    for x in filas:   # comparativo: último mes contra el anterior
        ult, ant = (x["meses"].get(claves[-1], 0), x["meses"].get(claves[-2], 0)) if len(claves) > 1 else (0, 0)
        x["variacion"] = (ult - ant) / ant * 100 if ant else None

    zona_fila = next((x for x in filas if x["zona"] == f["zona"]), None)
    clientes = []
    if f["zona"]:
        for k, c in por_cliente.items():
            vend, otros = principal(c["vendedores"])
            ficha = fichas.get(k, {})
            seg1 = max(c["segmentos"].items(), key=lambda x: x[1])[0] if c["segmentos"] else ""
            clientes.append({**c, "vendedor": vend or ficha.get("vendedor", ""), "otros_vendedores": otros,
                             "asignado": ficha.get("vendedor", ""), "segmento": seg1})
        for palabra in f["q"].upper().split():
            clientes = [c for c in clientes if palabra in (c["cliente"] or "").upper() or palabra in (c["vendedor"] or "").upper()]
        clientes.sort(key=lambda c: (-c["total"], c["cliente"] or ""))
        total_zona = zona_fila["total"] if zona_fila else 0
        for i, c in enumerate(clientes, start=1):
            c["posicion"] = i
            c["pct"] = c["total"] / total_zona * 100 if total_zona else 0

    siguiente = date(f["hasta"].year + f["hasta"].month // 12, f["hasta"].month % 12 + 1, 1)
    return {"meses": meses, "filas": filas, "total": total, "clientes": clientes, "zona_fila": zona_fila,
            "por_mes": {m: sum(x["meses"].get(m, 0) for x in filas) for m in claves},
            "clientes_mes": {m: sum(c["meses"].get(m, 0) for c in clientes) for m in claves},
            "segmentos_todos": [s for s, _ in sorted(segs_total.items(), key=lambda x: -x[1])] if not f["segmento"] else [f["segmento"]],
            "parcial": (siguiente - f["hasta"]).days > 1,
            "maximo": max((v for x in filas if x["zona"] not in FUERA for v in x["meses"].values()), default=0),
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
    return render_template("ventas_zonas_mes.html", f=f, d=d, vendedores=lista_vendedores(), zonas=zonas()["orden"],
                           fuera=FUERA, sin_zona=SIN_ZONA, ocasional=OCASIONAL, segmentos=todos_los_segmentos(),
                           rangos=rangos_rapidos(f["ultima"]))


def lista_zona(f, zona, tipo, solo_cliente=None):
    """Clientes (con su vendedor) o productos de una zona; con solo_cliente, los productos de ese cliente."""
    catalogo = catalogo_basico() if tipo == "productos" else {}
    filas, total = {}, 0.0
    for z, cliente, vendedor, mes, codigo, seg, fam, sub, monto, unidades, ultima, desc in filas_filtradas(f):
        if z != zona or (solo_cliente and cliente != solo_cliente):
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
    """Clientes o productos de una zona para la ventana (HTML) o en Excel."""
    f = leer_filtros(request.args)
    tipo = "productos" if request.args.get("tipo") == "productos" else "clientes"
    zona = (request.args.get("ver") or "").strip()
    if zona not in zonas()["orden"] and zona not in FUERA:
        return "Zona no válida", 400
    cliente = (request.args.get("cliente") or "").strip() if tipo == "productos" else ""
    filas, total = lista_zona(f, zona, tipo, cliente or None)
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
        return render_template("ventas_departamento_lista.html", f=f, tipo=tipo, depto=zona, filas=filas, total=total,
                               filtro=texto_filtro(f), editar=editar, segmentos=segmentos, arbol=arbol, sin_segmento=SIN_SEGMENTO,
                               cliente=cliente, unidad="zona",
                               excel_url=url_for("ventas_zonas_mes.lista", formato="excel", **request.args.to_dict()))

    wb = Workbook()
    ws = wb.active
    ws.title = tipo.capitalize()
    titulo = f"{'Clientes' if tipo == 'clientes' else 'Productos'} de {cliente + ' · ' if cliente else ''}zona {zona}"
    if tipo == "clientes":
        columnas = [("#", 5), ("Cliente", 42), ("Vendedor", 28), ("Otros vendedores", 26), ("Total C$", 15), ("% de la zona", 10),
                    ("Unidades", 11), ("Productos distintos", 11)]
        valores = [[n, x["cliente"], x["vendedor"], ", ".join(x["otros_vendedores"]) or None, round(x["monto"], 2), x["pct"] / 100,
                    x["unidades"], x["productos"]] for n, x in enumerate(filas, 1)]
    else:
        columnas = [("#", 5), ("Código", 12), ("Nº parte", 20), ("Descripción", 45), ("Marca", 16), ("Segmento", 20), ("Familia", 26),
                    ("Subfamilia", 30), ("Total C$", 15), ("% del cliente" if cliente else "% de la zona", 10), ("Unidades", 11), ("Clientes", 10)]
        valores = [[n, x["codigo"], x["numero_parte"], x["descripcion"], x["marca"], x["segmento"], x["familia"], x["subfamilia"],
                    round(x["monto"], 2), x["pct"] / 100, x["unidades"], x["clientes"]] for n, x in enumerate(filas, 1)]
        if cliente:
            columnas, valores = columnas[:-1], [v[:-1] for v in valores]
    hoja_excel(ws, titulo, nota_excel(f), columnas)
    for r, fila in enumerate(valores, start=5):
        for i, v in enumerate(fila, start=1):
            c = ws.cell(row=r, column=i, value=v)
            nombre = columnas[i - 1][0]
            c.number_format = ("#,##0.00" if nombre == "Total C$" else "0.0%" if nombre.startswith("%")
                               else "#,##0.##" if nombre == "Unidades" else c.number_format)
    col = [c[0] for c in columnas].index("Total C$") + 1
    fila_total = 5 + len(valores)
    ws.cell(row=fila_total, column=col - 1, value="TOTAL").font = Font(bold=True)
    c = ws.cell(row=fila_total, column=col, value=round(total, 2))
    c.font = Font(bold=True)
    c.number_format = "#,##0.00"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(columnas))}{max(fila_total - 1, 5)}"
    return enviar(wb, f"{tipo}_zona_{compacto(cliente or zona).lower()}_{f['desde']}_al_{f['hasta']}.xlsx")


@ventas_zonas_mes_bp.route("/ventas_zonas_mes/excel")
@requiere_acceso("ventas_zonas_mes")
def descargar_excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    meses = d["meses"]
    nota = nota_excel(f)
    wb = Workbook()

    ws = wb.active
    ws.title = "Zonas"
    hoja_excel(ws, "Ventas por zona y mes", nota, [("Zona", 18)] + [(e, 13) for _, e in meses]
               + [("Total C$", 15), ("% del total", 9), ("Vendedor principal", 28), ("% del vendedor", 9), ("Compraron", 10),
                  ("Clientes en cartera", 10), ("Var. último mes", 10)])
    for n, x in enumerate(d["filas"], start=5):
        valores = ([x["zona"]] + [round(x["meses"].get(m, 0), 2) or None for m, _ in meses]
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

    ws2 = wb.create_sheet("Segmentos por zona")
    hoja_excel(ws2, "Segmentos de producto por zona y mes", nota, [("Zona", 18), ("Segmento", 24)] + [(e, 13) for _, e in meses]
               + [("Total C$", 15), ("% de la zona", 9)])
    n = 5
    for x in d["filas"]:
        if f["zona"] and x["zona"] != f["zona"]:
            continue
        for s in x["segmentos"]:
            valores = [x["zona"], s["segmento"]] + [round(s["meses"].get(m, 0), 2) or None for m, _ in meses] + [round(s["total"], 2), s["pct"] / 100]
            for i, v in enumerate(valores, start=1):
                c = ws2.cell(row=n, column=i, value=v)
                if 2 < i <= len(meses) + 3:
                    c.number_format = "#,##0.00"
            ws2.cell(row=n, column=len(meses) + 4).number_format = "0.0%"
            n += 1
    ws2.freeze_panes = "C5"
    ws2.auto_filter.ref = f"A4:{get_column_letter(len(meses) + 4)}{max(n - 1, 5)}"

    if f["zona"]:
        ws3 = wb.create_sheet("Clientes")
        hoja_excel(ws3, f"Clientes de la zona {f['zona']} por mes", nota, [("Cliente", 40), ("Vendedor", 28), ("Otros vendedores", 26),
                                                                          ("Segmento principal", 20)]
                   + [(e, 13) for _, e in meses] + [("Total C$", 15), ("% de la zona", 9), ("Última compra", 12)])
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

    sufijo = f"_{compacto(f['zona']).lower()}" if f["zona"] else ""
    return enviar(wb, f"ventas_zona_mes{sufijo}_{f['desde']}_al_{f['hasta']}.xlsx")
