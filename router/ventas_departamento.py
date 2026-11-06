"""Ventas por departamento: venta neta (sin IVA) de productos por departamento del cliente, con segmento,
familia y subfamilia (Clasificación de productos) y mapa de Nicaragua.

El departamento sale de la tabla clientes (cruce por nombre, como los reportes de zonas). Ese campo trae
municipios, regiones con nombres viejos y errores de escritura, así que se normaliza a los 17 departamentos."""
import difflib
import io
import json
import os
import re
import time
from datetime import date, datetime

from flask import Blueprint, current_app, jsonify, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import sin_tildes
from router.dashboard import ultima_fecha_con_ventas
from router import sugerencias_clase
from router.clase_prod import LARGO_CAMPO, opciones as opciones_clase
from router.ventas_bodegas import _CACHE, en_cache
from router.ventas_vendedor import rangos_rapidos

ventas_departamento_bp = Blueprint("ventas_departamento", __name__)

SIN_DEPTO = "Sin departamento"       # está en la tabla clientes pero sin departamento ni municipio reconocible
OCASIONAL = "Cliente ocasional"      # el comprador no está en la tabla clientes (mostrador, nombre escrito distinto)
FUERA_DEL_MAPA = (SIN_DEPTO, OCASIONAL)
SIN_SEGMENTO = "Sin clasificar"
VERDE, NARANJA = "027266", "FE6412"

DEPARTAMENTOS = ["BOACO", "CARAZO", "CHINANDEGA", "CHONTALES", "ESTELI", "GRANADA", "JINOTEGA", "LEON", "MADRIZ", "MANAGUA",
                 "MASAYA", "MATAGALPA", "NUEVA SEGOVIA", "RACCN", "RACCS", "RIO SAN JUAN", "RIVAS"]
ALIAS = {"ATLANTICO NORTE": "RACCN", "RAAN": "RACCN", "COSTA CARIBE NORTE": "RACCN", "REGION AUTONOMA ATLANTICO NORTE": "RACCN",
         "ATLANTICO SUR": "RACCS", "RAAS": "RACCS", "COSTA CARIBE SUR": "RACCS", "REGION AUTONOMA ATLANTICO SUR": "RACCS",
         "SAN JUAN": "RIO SAN JUAN", "RIO SANJUAN": "RIO SAN JUAN", "CHON": "CHONTALES", "MN": "MANAGUA"}
# Municipios (y lugares conocidos o mal escritos que aparecen en la ficha) -> departamento
MUNICIPIOS = {
    "BOACO": "BOACO, CAMOAPA, SAN JOSE DE LOS REMATES, SAN LORENZO, SANTA LUCIA, TEUSTEPE",
    "CARAZO": "JINOTEPE, DIRIAMBA, SAN MARCOS, DOLORES, EL ROSARIO, LA CONQUISTA, LA PAZ DE CARAZO, SANTA TERESA",
    "CHINANDEGA": "CHINANDEGA, CHICHIGALPA, CORINTO, EL REALEJO, EL VIEJO, VIEJO, POSOLTEGA, PUERTO MORAZAN, "
                  "SANTO TOMAS DEL NORTE, SOMOTILLO, VILLANUEVA, CINCO PINOS, SAN FRANCISCO DEL NORTE, SAN PEDRO DEL NORTE, GUANACASTILLO",
    "CHONTALES": "JUIGALPA, ACOYAPA, COMALAPA, CUAPA, EL CORAL, LA LIBERTAD, SAN PEDRO DE LOVAGO, SANTO DOMINGO, SANTO TOMAS, VILLA SANDINO",
    "ESTELI": "ESTELI, CONDEGA, LA TRINIDAD, PUEBLO NUEVO, SAN JUAN DE LIMAY, SAN NICOLAS",
    "GRANADA": "GRANADA, DIRIA, DIRIOMO, NANDAIME",
    "JINOTEGA": "JINOTEGA, EL CUA, LA CONCORDIA, SAN JOSE DE BOCAY, SAN RAFAEL DEL NORTE, SAN SEBASTIAN DE YALI, "
                "SANTA MARIA DE PANTASMA, WIWILI DE JINOTEGA",
    "LEON": "LEON, ACHUAPA, EL JICARAL, EL SAUCE, LA PAZ CENTRO, LARREYNAGA, MALPAISILLO, NAGAROTE, QUEZALGUAQUE, QUETZALGUAQUE, "
            "SANTA ROSA DEL PENON, TELICA, PUERTO SANDINO",
    "MADRIZ": "MADRIZ, SOMOTO, LAS SABANAS, PALACAGUINA, SAN JOSE DE CUSMAPA, SAN JUAN DE RIO COCO, SAN JUAN RIO DE COCO, "
              "SAN J RIO COCOC, SAN LUCAS, TELPANECA, TOTOGALPA, YALAGUINA",
    "MANAGUA": "MANAGUA, CIUDAD SANDINO, EL CRUCERO, MATEARE, SAN FRANCISCO LIBRE, SAN RAFAEL DEL SUR, SAN RAFAEL SUR, "
               "TICUANTEPE, TIPITAPA, VILLA EL CARMEN",
    "MASAYA": "MASAYA, CATARINA, LA CONCEPCION, LA CONCHA, MASATEPE, NANDASMO, NINDIRI, NIQUINOHOMO, SAN JUAN DE ORIENTE, TISMA",
    "MATAGALPA": "MATAGALPA, CIUDAD DARIO, DARIO, EL TUMA LA DALIA, LA DALIA, ESQUIPULAS, MATIGUAS, MUY MUY, RANCHO GRANDE, "
                 "RIO BLANCO, SAN DIONISIO, SAN ISIDRO, SAN RAMON, SEBACO, TERRABONA",
    "NUEVA SEGOVIA": "NUEVA SEGOVIA, OCOTAL, CIUDAD ANTIGUA, DIPILTO, EL JICARO, JALAPA, MACUELIZO, MOZONTE, MURRA, QUILALI, "
                     "SAN FERNANDO, WIWILI DE NUEVA SEGOVIA",
    "RACCN": "PUERTO CABEZAS, PUERTO CABEZA, BILWI, BONANZA, MULUKUKU, PRINZAPOLKA, ROSITA, SIUNA, WASLALA, WASPAM, WASPAN",
    "RACCS": "BLUEFIELDS, CORN ISLAND, DESEMBOCADURA DE RIO GRANDE, EL AYOTE, EL RAMA, RAMA, EL TORTUGUERO, KUKRA HILL, "
             "LA CRUZ DE RIO GRANDE, LAGUNA DE PERLAS, MUELLE DE LOS BUEYES, NUEVA GUINEA, PAIWAS",
    "RIO SAN JUAN": "SAN CARLOS, EL ALMENDRO, EL CASTILLO, MORRITO, SAN JUAN DEL NORTE, SAN MIGUELITO",
    "RIVAS": "RIVAS, ALTAGRACIA, BELEN, BUENOS AIRES, CARDENAS, MOYOGALPA, POTOSI, SAN JORGE, SAN JUAN DEL SUR, TOLA, OMETEPE",
}
LUGARES = {d: d for d in DEPARTAMENTOS}
LUGARES.update(ALIAS)
for _depto, _lista in MUNICIPIOS.items():
    for _m in _lista.split(","):
        LUGARES.setdefault(_m.strip(), _depto)


def normal(texto):
    t = sin_tildes(texto or "").upper()
    t = re.sub(r"[^A-Z ]", " ", t)
    return " ".join(t.split())


def departamento_de(*textos):
    """Primer texto que se pueda reconocer como departamento (o municipio de uno)."""
    for texto in textos:
        t = normal(texto)
        if not t or t in ("XXX", "NA", "N A"):
            continue
        if t in LUGARES:
            return LUGARES[t]
        # empieza con un lugar conocido ("MANAGUA89213002", "MATAGALPA CENTRO")
        for lugar in sorted(LUGARES, key=len, reverse=True):
            if len(lugar) >= 4 and t.startswith(lugar):
                return LUGARES[lugar]
        # errores de escritura (MANAGAU, JUGALPA, DIRAMBA) o nombres cortados (CHINANDE)
        parecido = difflib.get_close_matches(t, LUGARES.keys(), n=1, cutoff=0.82)
        if parecido:
            return LUGARES[parecido[0]]
        cortado = [l for l in LUGARES if len(t) >= 5 and l.startswith(t)]
        if len(cortado) == 1:
            return LUGARES[cortado[0]]
    return None


def departamentos_clientes():
    """Nombre de cliente normalizado -> departamento, o SIN_DEPTO si está en la tabla sin ubicación (10 minutos en memoria).
    Quien no aparece aquí no está en la tabla clientes: es un cliente ocasional."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT nombre, departamento, municipio FROM clientes")
            resultado = {}
            for nombre, depto, municipio in cursor.fetchall():
                clave = " ".join((nombre or "").upper().split())
                d = departamento_de(depto, municipio) or SIN_DEPTO
                if clave and (clave not in resultado or resultado[clave] == SIN_DEPTO):
                    resultado[clave] = d
            return resultado
        finally:
            cursor.close()
            conn.close()
    return en_cache("vd_deptos_clientes", leer)


def cartera_por_departamento():
    """Clientes de la tabla clientes por departamento (nombres distintos), total y activos."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT nombre, departamento, municipio, estado FROM clientes")
            por_nombre = {}
            for nombre, depto, municipio, estado in cursor.fetchall():
                clave = " ".join((nombre or "").upper().split())
                if not clave:
                    continue
                d = departamento_de(depto, municipio) or SIN_DEPTO
                actual = por_nombre.get(clave)
                # un mismo nombre puede estar dos veces (activo e inactivo): cuenta una vez, activo si alguno lo está
                if actual is None or (actual[0] == SIN_DEPTO and d != SIN_DEPTO):
                    por_nombre[clave] = (d, estado == "Activo" or (actual is not None and actual[1]))
                elif estado == "Activo":
                    por_nombre[clave] = (actual[0], True)
            cartera = {}
            for d, activo in por_nombre.values():
                c = cartera.setdefault(d, {"total": 0, "activos": 0})
                c["total"] += 1
                c["activos"] += activo
            return cartera
        finally:
            cursor.close()
            conn.close()
    return en_cache("vd_cartera", leer)


def clasificacion():
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT codigo, segmento, familia, subfamilia FROM clase_prod")
            return {c: ((s or "").strip() or SIN_SEGMENTO, (f or "").strip() or SIN_SEGMENTO, (sf or "").strip() or SIN_SEGMENTO)
                    for c, s, f, sf in cursor.fetchall()}
        finally:
            cursor.close()
            conn.close()
    return en_cache("vd_clasificacion", leer)


def catalogo_basico():
    """Código -> (número de parte, marca) del catálogo de productos (10 minutos en memoria)."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT `Codigo Producto`, `Numero Parte`, Marca FROM catalogo_productos WHERE `Codigo Producto` IS NOT NULL")
            return {c: ((np or "").strip(), (m or "").strip()) for c, np, m in cursor.fetchall()}
        finally:
            cursor.close()
            conn.close()
    return en_cache("vd_catalogo", leer)


def mapa():
    def leer():
        with open(os.path.join(current_app.static_folder, "mapa_nicaragua.json"), encoding="utf-8") as f:
            return json.load(f)
    return en_cache("vd_mapa", leer, segundos=86400)


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def leer_filtros(args):
    ultima = ultima_fecha_con_ventas() or date.today()
    if isinstance(ultima, datetime):
        ultima = ultima.date()
    hasta = fecha_param(args.get("hasta"), ultima)
    desde = min(fecha_param(args.get("desde"), ultima.replace(day=1)), hasta)
    return {"desde": desde, "hasta": hasta, "ultima": ultima,
            "depto": (args.get("depto") or "").strip().upper(), "segmento": (args.get("segmento") or "").strip(),
            "familia": (args.get("familia") or "").strip()}


_VENTAS = {}   # (desde, hasta) -> (hora, filas): la página y las listas de clientes/productos leen lo mismo


def ventas_periodo(desde, hasta):
    """(cliente, código, monto, unidades, descripción) del período; 5 minutos en memoria (ventas_por_producto no tiene índices)."""
    clave, ahora = (desde, hasta), time.time()
    if clave in _VENTAS and ahora - _VENTAS[clave][0] < 300:
        return _VENTAS[clave][1]
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT TRIM(Cliente), Codigo, SUM(`Sub-Total 2`), SUM(`Cantidad Vendida`), MAX(`Descrip Prod.`)
            FROM ventas_por_producto WHERE Fecha BETWEEN %s AND %s AND Estatus <> 'Anulado'
            GROUP BY TRIM(Cliente), Codigo
        """, (desde.isoformat(), hasta.isoformat()))
        filas = [(c, cod, float(m or 0), float(u or 0), desc) for c, cod, m, u, desc in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()
    for k in [k for k, v in _VENTAS.items() if ahora - v[0] >= 300]:
        del _VENTAS[k]
    _VENTAS[clave] = (ahora, filas)
    return filas


def nivel_de(f):
    """Qué desglosan las columnas: segmentos, las familias del segmento elegido o las subfamilias de la familia elegida."""
    return "subfamilia" if f["familia"] else "familia" if f["segmento"] else "segmento"


def pasa_filtro(f, seg, fam):
    return (not f["segmento"] or seg == f["segmento"]) and (not f["familia"] or fam == f["familia"])


def clave_cliente(cliente):
    return " ".join((cliente or "").upper().split())


def calcular(f):
    deptos, clase = departamentos_clientes(), clasificacion()
    nivel = nivel_de(f)
    por_depto = {}      # depto -> {total (con filtro), general (sin filtro), grupos:{}, clientes:set}
    detalle = {}        # (depto, segmento, familia, subfamilia) -> [monto, unidades]   (con filtro)
    seg_todos, fam_opciones = {}, {}
    total = total_general = 0.0
    for cliente, codigo, monto, unidades, _desc in ventas_periodo(f["desde"], f["hasta"]):
        depto = deptos.get(clave_cliente(cliente), OCASIONAL)
        seg, fam, sub = clase.get(codigo, (SIN_SEGMENTO, SIN_SEGMENTO, SIN_SEGMENTO))
        d = por_depto.setdefault(depto, {"total": 0.0, "general": 0.0, "grupos": {}, "clientes": set()})
        d["general"] += monto
        total_general += monto
        seg_todos[seg] = seg_todos.get(seg, 0.0) + monto
        if not f["segmento"] or seg == f["segmento"]:
            fam_opciones[fam] = fam_opciones.get(fam, 0.0) + monto
        if not pasa_filtro(f, seg, fam):
            continue
        grupo = sub if nivel == "subfamilia" else fam if nivel == "familia" else seg
        d["total"] += monto
        d["grupos"][grupo] = d["grupos"].get(grupo, 0.0) + monto
        d["clientes"].add(cliente)
        x = detalle.setdefault((depto, seg, fam, sub), [0.0, 0.0])
        x[0] += monto
        x[1] += unidades
        total += monto

    totales_grupo = {}
    for d in por_depto.values():
        for g, v in d["grupos"].items():
            totales_grupo[g] = totales_grupo.get(g, 0.0) + v
    grupos = [g for g, _ in sorted(totales_grupo.items(), key=lambda x: -x[1])]

    cartera = cartera_por_departamento()
    tabla = []
    for nombre, d in por_depto.items():
        c = cartera.get(nombre, {"total": 0, "activos": 0})
        tabla.append({"depto": nombre, "total": d["total"], "general": d["general"], "grupos": d["grupos"], "clientes": len(d["clientes"]),
                      "cartera": c["total"], "activos": c["activos"],
                      "pct": d["total"] / total * 100 if total else 0,
                      "peso": d["total"] / d["general"] * 100 if d["general"] else 0})
    for nombre, c in cartera.items():
        if nombre not in por_depto:
            tabla.append({"depto": nombre, "total": 0.0, "general": 0.0, "grupos": {}, "clientes": 0, "cartera": c["total"],
                          "activos": c["activos"], "pct": 0, "peso": 0})
    tabla.sort(key=lambda x: (x["depto"] in FUERA_DEL_MAPA, x["depto"] == OCASIONAL, -x["total"], -x["general"]))
    principales = grupos[:5]   # columnas de la tabla; el resto va en "Otros"
    for x in tabla:
        x["otros"] = x["total"] - sum(x["grupos"].get(g, 0) for g in principales)
    otros_total = total - sum(totales_grupo.get(g, 0) for g in principales)

    # Familia / subfamilia (del departamento elegido, o de todo; ya con segmento/familia aplicados)
    familias = {}
    for (depto, seg, fam, sub), (monto, unidades) in detalle.items():
        if f["depto"] and depto.upper() != f["depto"]:
            continue
        fa = familias.setdefault(fam, {"familia": fam, "monto": 0.0, "unidades": 0.0, "subfamilias": {}})
        fa["monto"] += monto
        fa["unidades"] += unidades
        sf = fa["subfamilias"].setdefault(sub, {"subfamilia": sub, "monto": 0.0, "unidades": 0.0})
        sf["monto"] += monto
        sf["unidades"] += unidades
    lista_fam = sorted(familias.values(), key=lambda x: -x["monto"])
    for fa in lista_fam:
        fa["subfamilias"] = sorted(fa["subfamilias"].values(), key=lambda x: -x["monto"])
    total_fam = sum(fa["monto"] for fa in lista_fam)

    sin = por_depto.get(SIN_DEPTO, {"total": 0.0})["total"]
    ocasional = por_depto.get(OCASIONAL, {"total": 0.0})["total"]
    return {"tabla": tabla, "nivel": nivel, "filtrado": bool(f["segmento"] or f["familia"]), "grupos": grupos,
            "principales": principales, "otros_total": otros_total, "totales_grupo": totales_grupo,
            "total": total, "total_general": total_general,
            "segmentos_todos": [s for s, _ in sorted(seg_todos.items(), key=lambda x: -x[1])],
            "familias_opciones": [fa for fa, _ in sorted(fam_opciones.items(), key=lambda x: -x[1])],
            "sin_depto": sin, "ocasional": ocasional,
            "familias": lista_fam, "total_familias": total_fam, "detalle": detalle,
            "mapa_datos": {d["depto"]: {"total": round(d["total"], 2), "segmentos": {g: round(v, 2) for g, v in d["grupos"].items()},
                                        "clientes": d["clientes"], "cartera": d["cartera"], "activos": d["activos"]}
                           for d in tabla if d["depto"] not in FUERA_DEL_MAPA and d["total"] > 0},
            "cartera_total": sum(c["total"] for c in cartera.values()),
            "compraron_total": sum(x["clientes"] for x in tabla)}


def texto_filtro(f):
    return " · ".join(x for x in (f["segmento"], f["familia"]) if x)


def nombre_depto(depto):
    for d in DEPARTAMENTOS + [SIN_DEPTO, OCASIONAL]:
        if d.upper() == (depto or "").upper():
            return d
    return None


def lista_depto(f, depto, tipo, solo_cliente=None):
    """Clientes o productos que compraron en un departamento, con los filtros de segmento/familia.
    Con solo_cliente: los productos que compró ese cliente."""
    deptos, clase = departamentos_clientes(), clasificacion()
    catalogo = catalogo_basico() if tipo == "productos" else {}
    filas, total = {}, 0.0
    for cliente, codigo, monto, unidades, desc in ventas_periodo(f["desde"], f["hasta"]):
        if deptos.get(clave_cliente(cliente), OCASIONAL) != depto or (solo_cliente and cliente != solo_cliente):
            continue
        seg, fam, sub = clase.get(codigo, (SIN_SEGMENTO, SIN_SEGMENTO, SIN_SEGMENTO))
        if not pasa_filtro(f, seg, fam):
            continue
        total += monto
        if tipo == "clientes":
            x = filas.setdefault(cliente, {"cliente": cliente, "monto": 0.0, "unidades": 0.0, "productos": set()})
            x["productos"].add(codigo)
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
        x["pct"] = x["monto"] / total * 100 if total else 0
    return lista, total


NIVELES = {"segmento": "Segmento", "familia": "Familia", "subfamilia": "Subfamilia"}


@ventas_departamento_bp.route("/ventas_departamento")
@requiere_acceso("ventas_departamento")
def reporte():
    f = leer_filtros(request.args)
    d = calcular(f)
    return render_template("ventas_departamento.html", f=f, d=d, mapa=mapa(), rangos=rangos_rapidos(f["ultima"]),
                           departamentos=DEPARTAMENTOS, sin_depto=SIN_DEPTO, ocasional=OCASIONAL, fuera_del_mapa=FUERA_DEL_MAPA,
                           filtro=texto_filtro(f), nombre_nivel=NIVELES[d["nivel"]])


def encabezado_excel(ws, columnas, fila=4):
    for i, (nombre, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=fila, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre.startswith("Total C$") else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho


def titulo_excel(ws, titulo, nota):
    ws["A1"], ws["A2"] = titulo, nota
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"].font = Font(size=9, color="52514E")


def enviar_libro(wb, nombre):
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


NOTA_EXCEL = "Venta neta sin IVA (después de descuentos, sin facturas anuladas). Departamento según la ficha del cliente."


@ventas_departamento_bp.route("/ventas_departamento/excel")
@requiere_acceso("ventas_departamento")
def descargar_excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    filtro = texto_filtro(f)
    wb = Workbook()
    titulo = f"Ventas por departamento del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y}" + (f" · {filtro}" if filtro else "")

    ws = wb.active
    ws.title = "Departamento"
    titulo_excel(ws, titulo, NOTA_EXCEL)
    grupos = d["grupos"]
    columnas = ([("Departamento", 20)] + [(g, 14) for g in grupos]
                + [("Total C$" + (" (filtro)" if d["filtrado"] else ""), 15), ("% del total", 10)]
                + ([("Venta total del depto. C$", 15), ("% de la venta del depto.", 11)] if d["filtrado"] else [])
                + [("Clientes que compraron", 11), ("Clientes en cartera", 10), ("Activos en cartera", 10), ("% que compró", 10)])
    encabezado_excel(ws, columnas)
    col_total = len(grupos) + 2
    for n, x in enumerate(d["tabla"], start=5):
        valores = ([x["depto"]] + [round(x["grupos"].get(g, 0), 2) or None for g in grupos]
                   + [round(x["total"], 2), x["pct"] / 100]
                   + ([round(x["general"], 2), x["peso"] / 100] if d["filtrado"] else [])
                   + [x["clientes"], x["cartera"], x["activos"], x["clientes"] / x["cartera"] if x["cartera"] else None])
        for i, v in enumerate(valores, start=1):
            c = ws.cell(row=n, column=i, value=v)
            nombre = columnas[i - 1][0]
            c.number_format = "0.0%" if nombre.startswith("%") else ("#,##0" if nombre.startswith(("Clientes", "Activos")) else "#,##0.00")
        ws.cell(row=n, column=col_total).font = Font(bold=True)
    tot = 5 + len(d["tabla"])
    ws.cell(row=tot, column=1, value="TOTAL").font = Font(bold=True)
    sumar = list(range(2, col_total + 1)) + ([col_total + 2] if d["filtrado"] else [])
    for i in sumar:
        l = get_column_letter(i)
        c = ws.cell(row=tot, column=i, value=f"=SUM({l}5:{l}{tot - 1})")
        c.font = Font(bold=True)
        c.number_format = "#,##0.00"
    ws.freeze_panes = "B5"

    ws2 = wb.create_sheet("Familia y subfamilia")
    titulo_excel(ws2, titulo + " · detalle", NOTA_EXCEL)
    encabezado_excel(ws2, [("Departamento", 20), ("Segmento", 22), ("Familia", 30), ("Subfamilia", 36), ("Total C$", 15), ("Unidades", 11)])
    filas = sorted(d["detalle"].items(), key=lambda x: (x[0][0] in FUERA_DEL_MAPA, x[0][0], -x[1][0]))
    for n, ((depto, seg, fam, sub), (monto, unidades)) in enumerate(filas, start=5):
        for i, v in enumerate([depto, seg, fam, sub, round(monto, 2), unidades], start=1):
            ws2.cell(row=n, column=i, value=v)
        ws2.cell(row=n, column=5).number_format = "#,##0.00"
        ws2.cell(row=n, column=6).number_format = "#,##0.##"
    ws2.freeze_panes = "A5"
    ws2.auto_filter.ref = f"A4:F{4 + max(len(filas), 1)}"
    return enviar_libro(wb, f"ventas_departamento_{f['desde']}_al_{f['hasta']}.xlsx")


@ventas_departamento_bp.route("/ventas_departamento/lista")
@requiere_acceso("ventas_departamento")
def lista():
    """Clientes (tipo=clientes) o productos (tipo=productos) de un departamento; HTML para la ventana o Excel."""
    f = leer_filtros(request.args)
    tipo = "productos" if request.args.get("tipo") == "productos" else "clientes"
    depto = nombre_depto(request.args.get("ver"))
    if not depto:
        return "Departamento no válido", 400
    cliente = (request.args.get("cliente") or "").strip() if tipo == "productos" else ""
    filas, total = lista_depto(f, depto, tipo, cliente or None)
    filtro = texto_filtro(f)
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
        return render_template("ventas_departamento_lista.html", f=f, tipo=tipo, depto=depto, filas=filas, total=total, filtro=filtro,
                               editar=editar, segmentos=segmentos, arbol=arbol, sin_segmento=SIN_SEGMENTO,
                               cliente=cliente, unidad="depto.",
                               excel_url=url_for("ventas_departamento.lista", formato="excel", **request.args.to_dict()))

    wb = Workbook()
    ws = wb.active
    ws.title = tipo.capitalize()
    titulo_excel(ws, f"{'Clientes' if tipo == 'clientes' else 'Productos'} de {cliente + ' · ' if cliente else ''}{depto} del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y}"
                 + (f" · {filtro}" if filtro else ""), NOTA_EXCEL)
    if tipo == "clientes":
        columnas = [("#", 5), ("Cliente", 45), ("Total C$", 15), ("% del depto.", 10), ("Unidades", 11), ("Productos distintos", 11)]
        valores = [[n, x["cliente"], round(x["monto"], 2), x["pct"] / 100, x["unidades"], x["productos"]] for n, x in enumerate(filas, 1)]
    else:
        columnas = [("#", 5), ("Código", 12), ("Nº parte", 20), ("Descripción", 45), ("Marca", 16), ("Segmento", 20), ("Familia", 26),
                    ("Subfamilia", 30), ("Total C$", 15), ("% del cliente" if cliente else "% del depto.", 10), ("Unidades", 11), ("Clientes", 10)]
        valores = [[n, x["codigo"], x["numero_parte"], x["descripcion"], x["marca"], x["segmento"], x["familia"], x["subfamilia"],
                    round(x["monto"], 2), x["pct"] / 100, x["unidades"], x["clientes"]] for n, x in enumerate(filas, 1)]
        if cliente:   # un solo cliente: la columna Clientes no dice nada
            columnas, valores = columnas[:-1], [v[:-1] for v in valores]
    encabezado_excel(ws, columnas)
    for r, fila in enumerate(valores, start=5):
        for i, v in enumerate(fila, start=1):
            c = ws.cell(row=r, column=i, value=v)
            nombre = columnas[i - 1][0]
            if nombre == "Total C$":
                c.number_format = "#,##0.00"
            elif nombre.startswith("%"):
                c.number_format = "0.0%"
            elif nombre == "Unidades":
                c.number_format = "#,##0.##"
    tot = 5 + len(valores)
    col = [c[0] for c in columnas].index("Total C$") + 1
    ws.cell(row=tot, column=col - 1, value="TOTAL").font = Font(bold=True)
    c = ws.cell(row=tot, column=col, value=round(total, 2))
    c.font = Font(bold=True)
    c.number_format = "#,##0.00"
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(columnas))}{max(tot - 1, 5)}"
    return enviar_libro(wb, f"{tipo}_{normal(cliente or depto).replace(' ', '_').lower()}_{f['desde']}_al_{f['hasta']}.xlsx")


def tiene_acceso(vista):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT permitido FROM accesos_vistas WHERE usuario = %s AND vista = %s", (session.get("usuario"), vista))
        fila = cursor.fetchone()
        return bool(fila and fila[0])
    finally:
        cursor.close()
        conn.close()


@ventas_departamento_bp.route("/ventas_departamento/clasificar", methods=["POST"])
def clasificar():
    """Cambia segmento, familia y subfamilia de un producto desde la lista de productos (tabla clase_prod).
    La usan Ventas por departamento y Ventas por zona mensual; pide el permiso de «Clasificación de productos».
    El grupo de equivalencia no se toca."""
    if not tiene_acceso("clase_prod"):
        return jsonify(ok=False, error="No tiene permiso de Clasificación de productos."), 403
    codigo = (request.form.get("codigo") or "").strip()
    datos = {c: (request.form.get(c) or "").strip() or None for c in ("segmento", "familia", "subfamilia")}
    largos = [c for c, v in datos.items() if v and len(v) > LARGO_CAMPO]
    if largos:
        return jsonify(ok=False, error=f"{', '.join(largos)}: máximo {LARGO_CAMPO} caracteres."), 400
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT 1 FROM catalogo_productos WHERE `Codigo Producto` = %s LIMIT 1", (codigo,))
        if not codigo or cursor.fetchone() is None:
            return jsonify(ok=False, error=f"El código {codigo} no existe en el catálogo."), 400
        cursor.execute("""
            INSERT INTO clase_prod (codigo, segmento, familia, subfamilia) VALUES (%s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE segmento = VALUES(segmento), familia = VALUES(familia), subfamilia = VALUES(subfamilia)
        """, (codigo, datos["segmento"], datos["familia"], datos["subfamilia"]))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return jsonify(ok=False, error=f"Error de base de datos: {e}"), 500
    finally:
        cursor.close()
        conn.close()
    _CACHE.pop("vd_clasificacion", None)   # el reporte toma la clasificación nueva al recargar
    sugerencias_clase.olvidar()
    return jsonify(ok=True, codigo=codigo, **{c: v or SIN_SEGMENTO for c, v in datos.items()})
