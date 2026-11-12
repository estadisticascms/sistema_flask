"""Validación de facturas: cruza las tres fuentes de una factura y muestra lo que no cuadra.

- Transacciones (transacciones_inventario): salidas por FACTURA BIENES / MIXTA, costo = total_mn.
- Facturas por producto (ventas_por_producto): cantidad × «Costo Unit.», venta neta = Sub-Total 2.
- Facturas históricas (ventas_historicas): «Costo Total» por factura y bodega.

El número de factura se repite entre bodegas y años: una factura es número + fecha. Si una fuente tiene la misma factura con otra
fecha (hasta 15 días de diferencia) se junta y se marca «Fecha distinta».

Problemas que se detectan por factura:
  duplicada        filas idénticas repetidas en transacciones (misma transacción, producto, bodega, cantidad y costo) de más
  dif_costo        el costo no coincide entre las fuentes (más de la tolerancia)
  dif_cantidad     las unidades de transacciones y de facturas por producto no coinciden
  sin_historica    no está en facturas históricas
  sin_detalle      no está en facturas por producto
  sin_transaccion  no tiene salida de inventario (y tiene costo)
  anulada_con_mov  anulada en facturas por producto pero con transacción o en históricas
  fecha_distinta   la misma factura con fechas diferentes entre las fuentes
  numero_invalido  número de factura 0 o vacío
Las facturas anuladas sin movimiento son normales: se cuentan aparte y no son problema."""
import io
from collections import defaultdict
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from router.accesos import requiere_acceso
from router.ventas_bodegas import en_cache
from utils.db import get_connection

validacion_facturas_bp = Blueprint("validacion_facturas", __name__)

MARGEN_DIAS = 15            # se cargan días de más a cada lado para encontrar la misma factura con otra fecha
TOLERANCIA = 1.0            # C$ de diferencia de costo que se aceptan (redondeos)
MAX_FILAS = 1000
VERDE, NARANJA = "027266", "FE6412"
FUENTES = {"T": "Transacciones", "P": "Facturas por producto", "H": "Facturas históricas"}
PROBLEMAS = {
    "duplicada": ("Transacción duplicada", "Filas repetidas en transacciones: el costo y las unidades salen de más", "#b42318"),
    "dif_costo": ("Costo distinto", "El costo no coincide entre transacciones, facturas por producto e históricas", "#FE6412"),
    "dif_cantidad": ("Cantidad distinta", "Las unidades de transacciones y de facturas por producto no coinciden", "#D9A441"),
    "sin_historica": ("No está en históricas", "La factura no aparece en facturas históricas", "#6D597A"),
    "sin_detalle": ("No está en facturas por producto", "La factura no aparece en el detalle por producto", "#4F6D8A"),
    "sin_transaccion": ("Sin transacción", "La factura tiene costo pero no tiene salida de inventario", "#8a4b08"),
    "anulada_con_mov": ("Anulada con movimiento", "Anulada en facturas por producto pero con transacción o en históricas", "#9a3412"),
    "fecha_distinta": ("Fecha distinta", "La misma factura tiene fechas diferentes entre las fuentes", "#2F6DB5"),
    "numero_invalido": ("Número inválido", "Número de factura 0 o vacío", "#52514e"),
}
ESTADOS = {"problemas": "Con problemas", "todas": "Todas", "ok": "Cuadran", "anuladas": "Anuladas sin movimiento",
           **{k: v[0] for k, v in PROBLEMAS.items()}}


def fecha_param(valor):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return None


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return fecha_param(str(v or "")[:10])


def num_factura(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    return str(v).strip()


def ultimo_dia():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas")
        return a_fecha(cursor.fetchone()[0])
    finally:
        cursor.close()
        conn.close()


def leer_filtros(args):
    desde, hasta = fecha_param(args.get("desde")), fecha_param(args.get("hasta"))
    if not desde and not hasta:
        hasta = ultimo_dia() or date.today()
        desde = hasta.replace(day=1)
    desde, hasta = desde or hasta, hasta or desde
    if hasta < desde:
        desde, hasta = hasta, desde
    try:
        tol = max(float(args.get("tol") or TOLERANCIA), 0)
    except ValueError:
        tol = TOLERANCIA
    return {"desde": desde, "hasta": hasta, "tol": tol, "estado": args.get("estado") if args.get("estado") in ESTADOS else "problemas",
            "q": (args.get("q") or "").strip(), "dia": fecha_param(args.get("dia"))}


# ---------------------------------------------------------------- carga
def cargar(a, b):
    """Líneas de las tres fuentes entre `a` y `b` (con los días de margen ya incluidos). Se guarda 10 minutos."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT id_documento_origen, fecha, bodega, tipo_documento_origen, id_transaccion, codigo_producto, producto,
                       cantidad, costo_unitario_mn, total_mn
                FROM transacciones_inventario
                WHERE fecha BETWEEN %s AND %s AND tipo_transaccion = 'SALIDA' AND tipo_documento_origen LIKE 'FACTURA%%'
            """, (a, b))
            T = [{"num": num_factura(r[0]), "fecha": a_fecha(r[1]), "bodega": r[2] or "", "tipo": r[3] or "", "id_trans": r[4],
                  "codigo": r[5] or "", "producto": r[6] or "", "cantidad": float(r[7] or 0), "costo_u": float(r[8] or 0),
                  "costo": float(r[9] or 0)} for r in cursor.fetchall()]
            # ventas_por_producto no tiene índices y la fecha es texto 'AAAA-MM-DD'
            cursor.execute("""
                SELECT `No. Factura`, Fecha, `Tipo Factura`, Estatus, Cliente, Vendedor, Bodega, Codigo, `Descrip Prod.`,
                       `Cantidad Vendida`, `Costo Unit.`, `Sub-Total 2`
                FROM ventas_por_producto WHERE Fecha BETWEEN %s AND %s
            """, (a.isoformat(), b.isoformat()))
            P = [{"num": num_factura(r[0]), "fecha": a_fecha(r[1]), "tipo": r[2] or "", "estatus": r[3] or "", "cliente": r[4] or "",
                  "vendedor": r[5] or "", "bodega": r[6] or "", "codigo": r[7] or "", "producto": r[8] or "",
                  "cantidad": float(r[9] or 0), "costo_u": float(r[10] or 0), "costo": float(r[9] or 0) * float(r[10] or 0),
                  "venta": float(r[11] or 0)} for r in cursor.fetchall()]
            cursor.execute("""
                SELECT factura, Fecha, `tipo factura`, tipo_venta, Bodega, cliente, vendedor, estado_factura, neto, `Costo Total`, Total
                FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s
            """, (a, b))
            H = [{"num": num_factura(r[0]), "fecha": a_fecha(r[1]), "tipo": r[2] or "", "tipo_venta": r[3] or "", "bodega": r[4] or "",
                  "cliente": r[5] or "", "vendedor": r[6] or "", "estado": r[7] or "", "neto": float(r[8] or 0),
                  "costo": float(r[9] or 0), "total": float(r[10] or 0)} for r in cursor.fetchall()]
            return {"T": T, "P": P, "H": H}
        finally:
            cursor.close()
            conn.close()
    return en_cache(f"validacion_facturas_{a}_{b}", leer)


def agrupar(lineas):
    """Junta las líneas de las tres fuentes en facturas (número + fecha; fechas a 15 días o menos son la misma factura)."""
    por_num = defaultdict(list)          # número -> [(fecha, fuente, línea)]
    for fuente in ("T", "P", "H"):
        for x in lineas[fuente]:
            if x["fecha"]:
                por_num[x["num"]].append((x["fecha"], fuente, x))
    facturas = []
    for num, items in por_num.items():
        items.sort(key=lambda i: i[0])
        grupo, ultimo = [], None
        for it in items:
            if grupo and (it[0] - ultimo).days > MARGEN_DIAS:
                facturas.append((num, grupo))
                grupo = []
            grupo.append(it)
            ultimo = it[0]
        facturas.append((num, grupo))
    return facturas


def armar(num, grupo, tol):
    lin = {"T": [], "P": [], "H": []}
    fechas = {"T": set(), "P": set(), "H": set()}
    for fecha, fuente, x in grupo:
        lin[fuente].append(x)
        fechas[fuente].add(fecha)
    T, P, H = lin["T"], lin["P"], lin["H"]
    fecha = min(fechas["H"] or fechas["P"] or fechas["T"])
    anulada = bool(P) and all(x["estatus"] == "Anulado" for x in P)
    Pv = [x for x in P if x["estatus"] != "Anulado"]          # líneas válidas (sin anuladas)
    f = {"num": num, "fecha": fecha, "id": f"{num}|{fecha.isoformat()}", "anulada": anulada,
         "fechas": sorted(set().union(*fechas.values())),
         "fecha_fuente": {k: sorted(v) for k, v in fechas.items()},
         "tipo": (H[0]["tipo"] if H else (P[0]["tipo"] if P else (T[0]["tipo"] if T else ""))),
         "cliente": (H[0]["cliente"] if H else (P[0]["cliente"] if P else "")),
         "vendedor": (H[0]["vendedor"] if H else (P[0]["vendedor"] if P else "")),
         "bodegas": sorted({x["bodega"] for x in T + P + H if x["bodega"]}),
         "estado": ", ".join(sorted({x["estado"] for x in H if x["estado"]} or {x["estatus"] for x in P if x["estatus"]})),
         "en": {"T": bool(T), "P": bool(P), "H": bool(H)},
         "T_lineas": len(T), "T_cant": sum(x["cantidad"] for x in T), "T_costo": sum(x["costo"] for x in T),
         "P_lineas": len(Pv), "P_cant": sum(x["cantidad"] for x in Pv), "P_costo": sum(x["costo"] for x in Pv),
         "P_venta": sum(x["venta"] for x in Pv),
         "H_filas": len(H), "H_costo": sum(x["costo"] for x in H), "H_neto": sum(x["neto"] for x in H)}
    # Filas repetidas en transacciones: idénticas y más veces que en la factura por producto
    rep_t = defaultdict(list)
    for x in T:
        rep_t[(x["id_trans"], x["codigo"], x["bodega"], x["cantidad"], round(x["costo"], 2))].append(x)
    en_p = defaultdict(int)
    for x in Pv:
        en_p[(x["codigo"], x["cantidad"])] += 1
    extra_filas = extra_costo = extra_cant = 0.0
    for (_, codigo, _, cant, costo), filas in rep_t.items():
        sobra = len(filas) - max(en_p.get((codigo, cant), 0), 1)
        if sobra > 0:
            extra_filas += sobra
            extra_costo += sobra * costo
            extra_cant += sobra * cant
    f["T_repetidas"], f["T_repetido_costo"] = int(extra_filas), extra_costo
    f["dif_TP"] = f["T_costo"] - f["P_costo"] if T and Pv else None
    f["dif_TH"] = f["T_costo"] - f["H_costo"] if T and H else None
    f["dif_PH"] = f["P_costo"] - f["H_costo"] if Pv and H else None
    f["dif_cant"] = f["T_cant"] - f["P_cant"] if T and Pv else None
    difs = [abs(v) for v in (f["dif_TP"], f["dif_TH"], f["dif_PH"]) if v is not None]
    f["dif_max"] = max(difs) if difs else 0.0

    p = []
    if num in ("", "0"):
        p.append("numero_invalido")
    if f["T_repetidas"]:
        p.append("duplicada")
    if anulada:
        if T or H:
            p.append("anulada_con_mov")
    else:
        # Sin contar las filas repetidas (ya se marcan como «duplicada»): ¿queda alguna diferencia?
        t_costo, t_cant = f["T_costo"] - extra_costo, f["T_cant"] - extra_cant
        difs_sin_rep = [abs(v) for v in ((t_costo - f["P_costo"]) if T and Pv else None, (t_costo - f["H_costo"]) if T and H else None,
                                         f["dif_PH"]) if v is not None]
        if difs_sin_rep and max(difs_sin_rep) > tol:
            p.append("dif_costo")
        if f["dif_cant"] is not None and abs(t_cant - f["P_cant"]) > 0.001:
            p.append("dif_cantidad")
        if (T or Pv) and not H:
            p.append("sin_historica")
        if (T or H) and not P:
            p.append("sin_detalle")
        if (Pv or H) and not T and (f["P_costo"] > tol or f["H_costo"] > tol):
            p.append("sin_transaccion")
    if sum(1 for v in fechas.values() if v) > 1 and len(set().union(*fechas.values())) > 1:
        p.append("fecha_distinta")
    f["problemas"] = p
    f["estado_val"] = "problema" if p else ("anulada" if anulada else "ok")
    return f, lin


def calcular(f):
    lineas = cargar(f["desde"] - timedelta(days=MARGEN_DIAS), f["hasta"] + timedelta(days=MARGEN_DIAS))
    facturas = []
    for num, grupo in agrupar(lineas):
        if not any(f["desde"] <= fe <= f["hasta"] for fe, _, _ in grupo):
            continue
        fac, _ = armar(num, grupo, f["tol"])
        if not (f["desde"] <= fac["fecha"] <= f["hasta"]):
            continue
        facturas.append(fac)

    # ---- Resumen general y por día ----
    def nuevo():
        return {"facturas": 0, "ok": 0, "problemas": 0, "anuladas": 0, "T_costo": 0.0, "P_costo": 0.0, "H_costo": 0.0,
                "T_repetido_costo": 0.0, **{k: 0 for k in PROBLEMAS}}
    total, por_dia = nuevo(), defaultdict(nuevo)
    for x in facturas:
        for dest in (total, por_dia[x["fecha"]]):
            dest["facturas"] += 1
            dest[{"problema": "problemas", "anulada": "anuladas", "ok": "ok"}[x["estado_val"]]] += 1
            for k in ("T_costo", "P_costo", "H_costo", "T_repetido_costo"):
                dest[k] += x[k]
            for k in x["problemas"]:
                dest[k] += 1
    for dest in [total] + list(por_dia.values()):
        dest["dif_TP"], dest["dif_TH"], dest["dif_PH"] = (dest["T_costo"] - dest["P_costo"], dest["T_costo"] - dest["H_costo"],
                                                          dest["P_costo"] - dest["H_costo"])
    dias = [{"fecha": k, **v} for k, v in sorted(por_dia.items())]
    monto_problema = {k: sum(x["dif_max"] for x in facturas if k in x["problemas"]) for k in PROBLEMAS}
    monto_problema["duplicada"] = sum(x["T_repetido_costo"] for x in facturas)

    # ---- Tabla (filtros) ----
    filas = facturas
    if f["dia"]:
        filas = [x for x in filas if x["fecha"] == f["dia"]]
    if f["q"]:
        filas = [x for x in filas if f["q"] in x["num"] or f["q"].upper() in x["cliente"].upper()]
    e = f["estado"]
    if e == "problemas":
        filas = [x for x in filas if x["problemas"]]
    elif e == "ok":
        filas = [x for x in filas if x["estado_val"] == "ok"]
    elif e == "anuladas":
        filas = [x for x in filas if x["estado_val"] == "anulada"]
    elif e in PROBLEMAS:
        filas = [x for x in filas if e in x["problemas"]]
    filas.sort(key=lambda x: (-(x["T_repetido_costo"] + x["dif_max"]), x["fecha"], x["num"]))
    return {"facturas": facturas, "filas": filas, "total": total, "dias": dias, "monto_problema": monto_problema}


def indexar(lineas):
    """{número: {"T": [...], "P": [...], "H": [...]}} para buscar rápido las líneas de una factura."""
    def armar_indice():
        ind = defaultdict(lambda: {"T": [], "P": [], "H": []})
        for k, v in lineas.items():
            for x in v:
                ind[x["num"]][k].append(x)
        return dict(ind)
    return armar_indice()


def detalle_factura(num, fecha, tol, desde=None, hasta=None, indice=None):
    """Factura con sus líneas comparadas por producto. Usa la carga del período si está en la caché."""
    if indice is None:
        if desde and hasta and desde <= fecha <= hasta:
            a, b = desde - timedelta(days=MARGEN_DIAS), hasta + timedelta(days=MARGEN_DIAS)
        else:
            a, b = fecha - timedelta(days=MARGEN_DIAS), fecha + timedelta(days=MARGEN_DIAS)
        lineas = cargar(a, b)
        propias = {k: [x for x in v if x["num"] == num] for k, v in lineas.items()}
    else:
        propias = indice.get(num, {"T": [], "P": [], "H": []})
    grupos = sorted(agrupar(propias), key=lambda g: min(abs((fe - fecha).days) for fe, _, _ in g[1]))
    for n, grupo in grupos[:1]:
        if min(abs((fe - fecha).days) for fe, _, _ in grupo) <= MARGEN_DIAS:
            fac, lin = armar(n, grupo, tol)
            break
    else:
        return None, [], []
    productos = defaultdict(lambda: {"codigo": "", "producto": "", "T_filas": 0, "T_cant": 0.0, "T_costo": 0.0,
                                     "P_filas": 0, "P_cant": 0.0, "P_costo": 0.0, "P_venta": 0.0, "anulada": False})
    for x in lin["T"]:
        p = productos[x["codigo"]]
        p["codigo"], p["producto"] = x["codigo"], p["producto"] or x["producto"]
        p["T_filas"] += 1; p["T_cant"] += x["cantidad"]; p["T_costo"] += x["costo"]
    for x in lin["P"]:
        p = productos[x["codigo"]]
        p["codigo"], p["producto"] = x["codigo"], x["producto"] or p["producto"]
        if x["estatus"] == "Anulado":
            p["anulada"] = True
            continue
        p["P_filas"] += 1; p["P_cant"] += x["cantidad"]; p["P_costo"] += x["costo"]; p["P_venta"] += x["venta"]
    filas = []
    for p in productos.values():
        p["T_unit"] = p["T_costo"] / p["T_cant"] if p["T_cant"] else None
        p["P_unit"] = p["P_costo"] / p["P_cant"] if p["P_cant"] else None
        p["dif_cant"], p["dif_costo"] = p["T_cant"] - p["P_cant"], p["T_costo"] - p["P_costo"]
        if not p["T_filas"] and not p["P_filas"]:
            p["estado"] = "Anulado"
        elif not p["T_filas"]:
            p["estado"] = "Falta en transacciones"
        elif not p["P_filas"]:
            p["estado"] = "Falta en facturas por producto"
        elif p["T_filas"] > p["P_filas"] and abs(p["T_cant"] - p["P_cant"] * p["T_filas"] / p["P_filas"]) < 0.001:
            p["estado"] = "Repetida en transacciones"
        elif abs(p["dif_cant"]) > 0.001:
            p["estado"] = "Cantidad distinta"
        elif abs(p["dif_costo"]) > tol:
            p["estado"] = "Costo distinto"
        else:
            p["estado"] = "Cuadra"
        filas.append(p)
    filas.sort(key=lambda p: (p["estado"] == "Cuadra", -abs(p["dif_costo"]), p["codigo"]))
    return fac, filas, lin["H"]


# ---------------------------------------------------------------- rutas
@validacion_facturas_bp.route("/validacion_facturas")
@requiere_acceso("validacion_facturas")
def reporte():
    f = leer_filtros(request.args)
    d = calcular(f)
    return render_template("validacion_facturas.html", f=f, d=d, problemas=PROBLEMAS, estados=ESTADOS, fuentes=FUENTES,
                           max_filas=MAX_FILAS)


@validacion_facturas_bp.route("/validacion_facturas/detalle")
@requiere_acceso("validacion_facturas")
def detalle():
    num = (request.args.get("factura") or "").strip()
    fecha = fecha_param(request.args.get("fecha"))
    f = leer_filtros(request.args)
    fac, productos, historicas = (detalle_factura(num, fecha, f["tol"], f["desde"], f["hasta"]) if fecha else (None, [], []))
    return render_template("validacion_facturas_detalle.html", fac=fac, productos=productos, historicas=historicas,
                           problemas=PROBLEMAS, fuentes=FUENTES, tol=f["tol"], num=num)


@validacion_facturas_bp.route("/validacion_facturas/excel")
@requiere_acceso("validacion_facturas")
def excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    wb = Workbook()
    nota = (f"Del {f['desde']:%d/%m/%Y} al {f['hasta']:%d/%m/%Y} · tolerancia C$ {f['tol']:,.2f} · costo: transacciones = total_mn, "
            f"facturas por producto = cantidad × costo unitario (sin anuladas), históricas = Costo Total · Generado {datetime.now():%d/%m/%Y %H:%M}")
    C = "#,##0.00"

    def hoja(ws, titulo, cols):
        ws["A1"], ws["A2"] = titulo, nota
        ws["A1"].font = Font(bold=True, size=14, color=VERDE)
        ws["A2"].font = Font(size=9, color="52514E")
        for i, (nombre, ancho) in enumerate(cols, start=1):
            c = ws.cell(row=4, column=i, value=nombre)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NARANJA if nombre.startswith(("Dif", "Problema")) else VERDE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.row_dimensions[4].height = 32
        ws.freeze_panes = "B5"

    def escribir(ws, filas, formatos, sumas=()):
        for n, valores in enumerate(filas, start=5):
            for i, (v, fmt) in enumerate(zip(valores, formatos), start=1):
                c = ws.cell(row=n, column=i, value=v)
                if fmt:
                    c.number_format = fmt
        ult = 4 + len(filas)
        ws.auto_filter.ref = f"A4:{get_column_letter(len(formatos))}{max(ult, 5)}"
        if sumas and filas:
            ws.cell(row=ult + 1, column=1, value="TOTAL").font = Font(bold=True)
            for i in sumas:
                col = get_column_letter(i)
                c = ws.cell(row=ult + 1, column=i, value=f"=SUBTOTAL(9,{col}5:{col}{ult})")
                c.font = Font(bold=True)
                c.number_format = formatos[i - 1] or "#,##0"

    ws = wb.active
    ws.title = "Por día"
    hoja(ws, "Validación de facturas: resumen por día", [("Fecha", 12), ("Facturas", 10), ("Cuadran", 10), ("Con problemas", 10),
         ("Anuladas", 10), ("Costo transacciones C$", 16), ("Costo por producto C$", 16), ("Costo históricas C$", 16),
         ("Dif. transacciones − producto", 15), ("Dif. transacciones − históricas", 15), ("Dif. producto − históricas", 15),
         ("Costo repetido en transacciones", 15)] + [(v[0], 12) for v in PROBLEMAS.values()])
    escribir(ws, [[x["fecha"], x["facturas"], x["ok"], x["problemas"], x["anuladas"], round(x["T_costo"], 2), round(x["P_costo"], 2),
                   round(x["H_costo"], 2), round(x["dif_TP"], 2), round(x["dif_TH"], 2), round(x["dif_PH"], 2),
                   round(x["T_repetido_costo"], 2)] + [x[k] for k in PROBLEMAS] for x in d["dias"]],
             ["dd/mm/yyyy", "#,##0", "#,##0", "#,##0", "#,##0", C, C, C, C, C, C, C] + ["#,##0"] * len(PROBLEMAS),
             sumas=list(range(2, 13 + len(PROBLEMAS))))

    ws = wb.create_sheet("Facturas con problemas")
    hoja(ws, "Facturas con problemas", [("Factura", 11), ("Fecha", 12), ("Fechas en las fuentes", 26), ("Tipo", 10), ("Cliente", 32),
         ("Vendedor", 22), ("Bodegas", 20), ("Estado", 14), ("Líneas transacciones", 10), ("Cant. transacciones", 11),
         ("Costo transacciones C$", 15), ("Líneas por producto", 10), ("Cant. por producto", 11), ("Costo por producto C$", 15),
         ("Venta neta C$", 14), ("Filas históricas", 9), ("Costo históricas C$", 15), ("Dif. transacciones − producto", 14),
         ("Dif. transacciones − históricas", 14), ("Dif. producto − históricas", 14), ("Dif. cantidad", 10),
         ("Filas repetidas", 9), ("Costo repetido C$", 13), ("Problemas", 60)])
    malas = sorted((x for x in d["facturas"] if x["problemas"]), key=lambda x: (x["fecha"], x["num"]))
    escribir(ws, [[x["num"], x["fecha"], " · ".join(f"{FUENTES[k][:5]}. {', '.join(fe.strftime('%d/%m/%y') for fe in v)}"
                                                     for k, v in x["fecha_fuente"].items() if v),
                   x["tipo"], x["cliente"], x["vendedor"], ", ".join(x["bodegas"]), x["estado"], x["T_lineas"], x["T_cant"],
                   round(x["T_costo"], 2), x["P_lineas"], x["P_cant"], round(x["P_costo"], 2), round(x["P_venta"], 2), x["H_filas"],
                   round(x["H_costo"], 2), *(round(v, 2) if v is not None else None for v in (x["dif_TP"], x["dif_TH"], x["dif_PH"], x["dif_cant"])),
                   x["T_repetidas"], round(x["T_repetido_costo"], 2), " · ".join(PROBLEMAS[k][0] for k in x["problemas"])] for x in malas],
             [None, "dd/mm/yyyy", None, None, None, None, None, None, "#,##0", "#,##0.##", C, "#,##0", "#,##0.##", C, C, "#,##0", C,
              C, C, C, "#,##0.##", "#,##0", C, None], sumas=[11, 14, 15, 17, 18, 19, 20, 23])

    # Detalle por producto de las facturas con problemas (solo las líneas que no cuadran)
    ws = wb.create_sheet("Productos que no cuadran")
    hoja(ws, "Detalle por producto de las facturas con problemas", [("Factura", 11), ("Fecha", 12), ("Código", 12), ("Producto", 36),
         ("Filas transacciones", 10), ("Cant. transacciones", 11), ("Costo transacciones C$", 15), ("Filas por producto", 10),
         ("Cant. por producto", 11), ("Costo por producto C$", 15), ("Dif. cantidad", 10), ("Dif. costo C$", 13), ("Problema", 28)])
    filas = []
    indice = indexar(cargar(f["desde"] - timedelta(days=MARGEN_DIAS), f["hasta"] + timedelta(days=MARGEN_DIAS)))
    for x in malas:
        _, prods, _ = detalle_factura(x["num"], x["fecha"], f["tol"], indice=indice)
        filas += [[x["num"], x["fecha"], p["codigo"], p["producto"], p["T_filas"], p["T_cant"], round(p["T_costo"], 2), p["P_filas"],
                   p["P_cant"], round(p["P_costo"], 2), p["dif_cant"], round(p["dif_costo"], 2), p["estado"]]
                  for p in prods if p["estado"] not in ("Cuadra", "Anulado")]
    escribir(ws, filas, [None, "dd/mm/yyyy", None, None, "#,##0", "#,##0.##", C, "#,##0", "#,##0.##", C, "#,##0.##", C, None],
             sumas=[7, 10, 12])

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"validacion_facturas_{f['desde']:%Y%m%d}_{f['hasta']:%Y%m%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
