"""Pedidos de mercadería importada: registro, tiempos de cada etapa, reportes y pedido sugerido por tiempo de entrega.

Tablas propias (docs/sql/imp_pedidos.sql): imp_pedidos (encabezado), imp_pedido_detalle (líneas) e imp_stock_seguridad
(días de stock de seguridad por producto). El catálogo, las existencias, las ventas y la cantidad mínima (MOQ) salen de las
tablas que ya tiene el sistema: catalogo_productos, existencias_bodega, ventas_por_producto y cantidad_minima_compra.

Tiempos (días): resta de dos fechas del pedido; si falta alguna, queda vacío (no 0).

Sugerido por tiempo de entrega (por proveedor):
- Tiempo de entrega = promedio de «Días total» (fecha pedido → recepción en bodega) de los pedidos recibidos del proveedor
  (del transporte elegido; si no hay, de todos sus pedidos). Sin historial se usa el valor que escriba el usuario.
- Consumo diario = unidades vendidas en los últimos N meses de ventas cargadas ÷ días del período.
- Stock = existencia sin bodegas NHE (igual que el sugerido de compras) + lo que viene en pedidos abiertos (pedido − recibido).
- Fecha para pedir = hoy + (días de cobertura − tiempo de entrega − días de seguridad). Si ya pasó: «Pedir ya».
- Cantidad sugerida = consumo diario × (tiempo de entrega + seguridad + días a cubrir) − stock − en camino,
  subida al múltiplo de la cantidad mínima de compra."""
import io
import json
import math
import os
import uuid
from datetime import date, datetime, timedelta

import pandas as pd
from flask import Blueprint, flash, jsonify, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import sin_tildes
from router.cargar_ventas import ErrorArchivo, a_fecha as fecha_texto, a_texto, leer_filas, limpiar_temporales, ruta_temporal
from router.cantidades_minimas import ajustar_a_minimo, minimos_por_codigo

pedidos_importacion_bp = Blueprint("pedidos_importacion", __name__)
VISTA = "pedidos_importacion"
VERDE, NARANJA = "027266", "FE6412"
PREFIJO_TEMP = "carga_imp_pedidos_"
DIAS_MES = 30.44

ESTADOS = ["Cotizando", "Pedido", "En producción", "En tránsito", "En aduana", "Recibido", "Cancelado"]
CERRADOS = {"Recibido", "Cancelado"}
TRANSPORTES = ["Marítimo", "Aéreo", "Terrestre"]
COLOR_ESTADO = {"Cotizando": "secondary", "Pedido": "info", "En producción": "primary", "En tránsito": "warning",
                "En aduana": "warning", "Recibido": "success", "Cancelado": "dark"}

# Fechas del pedido en el orden en que ocurren
FECHAS = [
    ("f_solicitud_cot", "Solicitud cotización"),
    ("f_cot_recibida", "Cotización recibida"),
    ("f_pedido", "Fecha pedido"),
    ("f_embarque", "Fecha embarque"),
    ("f_llegada_estimada", "Llegada estimada"),
    ("f_llegada_real", "Llegada real"),
    ("f_liberacion", "Liberación aduana"),
    ("f_recepcion", "Recepción bodega"),
]
# Tiempos: clave, nombre, fecha desde, fecha hasta, explicación
TIEMPOS = [
    ("resp_cot", "Días respuesta cotización", "f_solicitud_cot", "f_cot_recibida", "Cotización recibida − Solicitud cotización"),
    ("decidir", "Días para decidir", "f_cot_recibida", "f_pedido", "Fecha pedido − Cotización recibida"),
    ("proveedor", "Días proveedor", "f_pedido", "f_embarque", "Fecha embarque − Fecha pedido"),
    ("transito", "Días tránsito", "f_embarque", "f_llegada_real", "Llegada real − Fecha embarque"),
    ("aduana", "Días aduana", "f_llegada_real", "f_liberacion", "Liberación aduana − Llegada real"),
    ("total", "Días total", "f_pedido", "f_recepcion", "Recepción bodega − Fecha pedido"),
    ("ciclo", "Días ciclo completo", "f_solicitud_cot", "f_recepcion", "Recepción bodega − Solicitud cotización"),
    ("atraso", "Días atraso", "f_llegada_estimada", "f_llegada_real", "Llegada real − Llegada estimada"),
]
# Fecha en que empieza la etapa de cada estado abierto (para «días en su etapa actual»)
INICIO_ETAPA = {"Cotizando": ["f_solicitud_cot"], "Pedido": ["f_pedido", "f_cot_recibida"], "En producción": ["f_pedido"],
                "En tránsito": ["f_embarque"], "En aduana": ["f_llegada_real"]}
CAMPOS_TEXTO = [("orden_compra", "N° orden de compra", 40), ("proveedor", "Proveedor", 150), ("pais", "País del proveedor", 60),
                ("cotizacion", "N° cotización", 60), ("factura", "N° factura", 60), ("descripcion", "Descripción", 500)]

MESES_CONSUMO = [3, 6, 12]
DIAS_CUBRIR = [30, 45, 60, 90, 120, 180]
SEGURIDAD_DEFECTO = 15


# -----------------------------
# Utilidades
# -----------------------------
def clave(texto):
    return " ".join(sin_tildes(a_texto(texto)).upper().split())


def ffecha(d):
    return d.strftime("%d/%m/%Y") if d else ""


def a_fecha(valor):
    """Fecha de Excel o texto dd/mm/aaaa (también aaaa-mm-dd y número de serie de Excel) -> date, None o 'error'."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return date(1899, 12, 30) + timedelta(days=int(valor)) if 20000 < valor < 80000 else "error"
    texto = a_texto(valor)
    if not texto:
        return None
    try:
        return fecha_texto(texto) or "error"
    except Exception:
        return "error"


def a_numero(valor):
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    try:
        return float(a_texto(valor).replace(",", "").replace("$", "").replace("C", "").replace("U", "").strip())
    except ValueError:
        return "error"


def normal_estado(texto):
    t = clave(texto)
    for e in ESTADOS:
        if clave(e) == t:
            return e
    alias = {"TRANSITO": "En tránsito", "ADUANA": "En aduana", "PRODUCCION": "En producción", "PEDIDO": "Pedido",
             "RECIBIDO": "Recibido", "CANCELADO": "Cancelado", "COTIZANDO": "Cotizando", "COTIZACION": "Cotizando"}
    return next((v for k, v in alias.items() if k in t), None)


def normal_transporte(texto):
    t = clave(texto)
    if not t:
        return ""
    if t.startswith("MAR") or "BARCO" in t or "BUQUE" in t:
        return "Marítimo"
    if t.startswith("AER") or "AVION" in t:
        return "Aéreo"
    if t.startswith("TER") or "CAMION" in t:
        return "Terrestre"
    return None


def dias(desde, hasta):
    return (hasta - desde).days if desde and hasta else None


def tiempos(p):
    return {k: dias(p.get(d), p.get(h)) for k, _, d, h, _ in TIEMPOS}


def estado_por_fechas(p):
    """Estado que corresponde a la última fecha llenada (para sugerirlo; el usuario decide)."""
    if p.get("f_recepcion"):
        return "Recibido"
    if p.get("f_llegada_real") or p.get("f_liberacion"):
        return "En aduana"
    if p.get("f_embarque"):
        return "En tránsito"
    if p.get("f_pedido"):
        return "Pedido"
    return "Cotizando"


def enriquecer(p, hoy):
    p["t"] = tiempos(p)
    p["abierto"] = p["estado"] not in CERRADOS
    p["etapa_desde"], p["dias_etapa"], p["atraso_hoy"] = None, None, None
    if p["abierto"]:
        inicio = next((p[c] for c in INICIO_ETAPA.get(p["estado"], []) if p.get(c)), None)
        if not inicio:
            llenas = [p[c] for c, _ in FECHAS if c != "f_llegada_estimada" and p.get(c)]
            inicio = max(llenas) if llenas else None
        if inicio:
            p["etapa_desde"], p["dias_etapa"] = inicio, (hoy - inicio).days
        if p.get("f_llegada_estimada") and not p.get("f_llegada_real") and p["f_llegada_estimada"] < hoy:
            p["atraso_hoy"] = (hoy - p["f_llegada_estimada"]).days
    p["sugerencia_estado"] = estado_por_fechas(p) if p["abierto"] else None
    return p


def promedio(valores):
    v = [x for x in valores if x is not None]
    return (sum(v) / len(v), len(v)) if v else (None, 0)


def promedios_por(pedidos, grupo):
    """[{grupo, pedidos, tiempos: {clave: (promedio, n)}}] ordenado por cantidad de pedidos."""
    grupos = {}
    for p in pedidos:
        grupos.setdefault(grupo(p) or "Sin dato", []).append(p)
    filas = []
    for nombre, lista in grupos.items():
        filas.append({"grupo": nombre, "pedidos": len(lista), "recibidos": sum(1 for p in lista if p["estado"] == "Recibido"),
                      "tiempos": {k: promedio(p["t"][k] for p in lista) for k, *_ in TIEMPOS}})
    filas.sort(key=lambda f: (-f["pedidos"], f["grupo"]))
    return filas


# -----------------------------
# Datos
# -----------------------------
def leer_pedidos(f=None, ids=None):
    f = f or {}
    where, params = [], []
    if ids is not None:
        where.append(f"p.id IN ({', '.join(['%s'] * len(ids))})" if ids else "1 = 0")
        params += list(ids)
    if f.get("estado") == "abiertos":
        where.append("p.estado NOT IN ('Recibido', 'Cancelado')")
    elif f.get("estado") in ESTADOS:
        where.append("p.estado = %s"); params.append(f["estado"])
    if f.get("proveedor"):
        where.append("p.proveedor = %s"); params.append(f["proveedor"])
    if f.get("transporte") in TRANSPORTES:
        where.append("p.transporte = %s"); params.append(f["transporte"])
    if f.get("desde"):
        where.append("p.f_pedido >= %s"); params.append(f["desde"])
    if f.get("hasta"):
        where.append("p.f_pedido <= %s"); params.append(f["hasta"])
    for palabra in (f.get("q") or "").split():
        where.append("(p.orden_compra LIKE %s OR p.proveedor LIKE %s OR p.factura LIKE %s OR p.cotizacion LIKE %s OR p.descripcion LIKE %s "
                     "OR EXISTS (SELECT 1 FROM imp_pedido_detalle x WHERE x.pedido_id = p.id AND (x.codigo LIKE %s OR x.descripcion LIKE %s)))")
        params += [f"%{palabra}%"] * 7
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(f"""
            SELECT p.*, COUNT(d.id) AS lineas, SUM(d.cantidad_pedida) AS unidades, SUM(d.cantidad_recibida) AS recibidas,
                   SUM(d.cantidad_pedida * COALESCE(d.costo_unitario, 0)) AS monto
            FROM imp_pedidos p LEFT JOIN imp_pedido_detalle d ON d.pedido_id = p.id
            {('WHERE ' + ' AND '.join(where)) if where else ''}
            GROUP BY p.id
            ORDER BY (p.estado IN ('Recibido', 'Cancelado')), COALESCE(p.f_pedido, p.f_solicitud_cot, DATE(p.creado)) DESC, p.id DESC
        """, params)
        pedidos = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    hoy = date.today()
    for p in pedidos:
        for k in ("unidades", "recibidas", "monto"):
            p[k] = float(p[k]) if p[k] is not None else None
        enriquecer(p, hoy)
    return pedidos


def leer_pedido(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM imp_pedidos WHERE id = %s", (id_,))
        p = cursor.fetchone()
        if not p:
            return None
        cursor.execute("SELECT codigo, descripcion, cantidad_pedida, cantidad_recibida, costo_unitario FROM imp_pedido_detalle "
                       "WHERE pedido_id = %s ORDER BY id", (id_,))
        p["lineas"] = [{k: (float(v) if k.startswith(("cantidad", "costo")) and v is not None else v) for k, v in l.items()}
                       for l in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()
    return enriquecer(p, date.today())


def lista_proveedores():
    """Proveedores para escoger: los de los pedidos y los de última compra del catálogo."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT DISTINCT TRIM(proveedor) FROM imp_pedidos WHERE proveedor IS NOT NULL AND TRIM(proveedor) <> ''")
        propios = {r[0] for r in cursor.fetchall()}
        cursor.execute("SELECT TRIM(`Proveedor Ultima Compra`), COUNT(*) FROM catalogo_productos "
                       "WHERE `Proveedor Ultima Compra` IS NOT NULL AND TRIM(`Proveedor Ultima Compra`) <> '' GROUP BY 1")
        catalogo = dict(cursor.fetchall())
    finally:
        cursor.close()
        conn.close()
    return sorted(propios | set(catalogo), key=str.upper), catalogo


def productos_catalogo(codigos):
    if not codigos:
        return {}
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"SELECT `Codigo Producto`, Nombre FROM catalogo_productos WHERE `Codigo Producto` IN ({', '.join(['%s'] * len(codigos))})",
                       list(codigos))
        return {c: n for c, n in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()


def filtros_lista(args):
    return {"estado": args.get("estado", "abiertos") if args.get("estado", "abiertos") in ESTADOS + ["abiertos", ""] else "abiertos",
            "proveedor": (args.get("proveedor") or "").strip(), "transporte": args.get("transporte") or "",
            "q": (args.get("q") or "").strip()}


def fecha_param(valor):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return None


# -----------------------------
# Lista
# -----------------------------
@pedidos_importacion_bp.route("/pedidos_importacion")
@requiere_acceso(VISTA)
def lista():
    f = filtros_lista(request.args)
    pedidos = leer_pedidos(f)
    abiertos = leer_pedidos({"estado": "abiertos"}) if f["estado"] != "abiertos" or f["proveedor"] or f["transporte"] or f["q"] else pedidos
    resumen = {"abiertos": len(abiertos), "atrasados": sum(1 for p in abiertos if p["atraso_hoy"]),
               "camino": sum(1 for p in abiertos if p["estado"] in ("En tránsito", "En aduana")),
               "produccion": sum(1 for p in abiertos if p["estado"] in ("Pedido", "En producción"))}
    proveedores, _ = lista_proveedores()
    return render_template("pedidos_importacion.html", pedidos=pedidos, f=f, resumen=resumen, estados=ESTADOS, transportes=TRANSPORTES,
                           proveedores=proveedores, color=COLOR_ESTADO, ffecha=ffecha, tiempos_def=TIEMPOS)


# -----------------------------
# Nuevo / editar
# -----------------------------
def leer_formulario(form):
    """-> (datos del encabezado, líneas, errores, avisos)."""
    datos, errores, avisos = {}, [], []
    for campo, nombre, largo in CAMPOS_TEXTO:
        datos[campo] = " ".join((form.get(campo) or "").split())[:largo] or None
    datos["observaciones"] = (form.get("observaciones") or "").strip() or None
    datos["estado"] = form.get("estado") if form.get("estado") in ESTADOS else "Cotizando"
    datos["transporte"] = form.get("transporte") if form.get("transporte") in TRANSPORTES else None
    for campo in ["fecha_factura"] + [c for c, _ in FECHAS]:
        valor = (form.get(campo) or "").strip()
        datos[campo] = fecha_param(valor) if valor else None
        if valor and not datos[campo]:
            errores.append(f"La fecha «{valor}» no es válida.")
    if not datos["orden_compra"]:
        errores.append("Escriba el N° de orden de compra.")

    lineas = []
    columnas = [form.getlist(c) for c in ("l_codigo", "l_descripcion", "l_pedida", "l_recibida", "l_costo")]
    for n, (cod, desc, ped, rec, costo) in enumerate(zip(*columnas), start=1):
        cod, desc = " ".join(cod.split()).upper()[:20], " ".join(desc.split())[:300]
        ped, rec, costo = a_numero(ped), a_numero(rec), a_numero(costo)
        if not (cod or desc or ped or rec or costo):
            continue
        for valor, nombre in ((ped, "cantidad pedida"), (rec, "cantidad recibida"), (costo, "costo unitario")):
            if valor == "error" or (valor is not None and valor < 0):
                errores.append(f"Línea {n} ({cod or desc}): la {nombre} no es un número válido.")
        lineas.append({"codigo": cod or None, "descripcion": desc or None, "cantidad_pedida": ped if isinstance(ped, float) else 0,
                       "cantidad_recibida": rec if isinstance(rec, float) else None, "costo_unitario": costo if isinstance(costo, float) else None})
    avisos += revisar(datos, lineas)
    return datos, lineas, errores, avisos


def revisar(datos, lineas):
    """Avisos que no impiden guardar: fechas al revés, estado y códigos que no están en el catálogo."""
    avisos = []
    t = tiempos(datos)
    for k, nombre, d, h, _ in TIEMPOS:
        if k != "atraso" and t[k] is not None and t[k] < 0:
            avisos.append(f"{nombre} da {t[k]}: «{dict(FECHAS)[h]}» es anterior a «{dict(FECHAS)[d]}». Revise las fechas.")
    if datos["estado"] == "Recibido" and not datos.get("f_recepcion"):
        avisos.append("El pedido está «Recibido» pero no tiene fecha de recepción en bodega: no entra en el tiempo total.")
    sugerido = estado_por_fechas(datos)
    if datos["estado"] not in CERRADOS and ESTADOS.index(sugerido) > ESTADOS.index(datos["estado"]):
        avisos.append(f"Según las fechas, el pedido ya estaría «{sugerido}».")
    codigos = {l["codigo"] for l in lineas if l["codigo"]}
    faltan = sorted(codigos - set(productos_catalogo(codigos)))
    if faltan:
        avisos.append(f"Códigos que no están en el catálogo: {', '.join(faltan[:10])}{'…' if len(faltan) > 10 else ''}.")
    return avisos


def completar_descripciones(lineas):
    nombres = productos_catalogo({l["codigo"] for l in lineas if l["codigo"] and not l["descripcion"]})
    for l in lineas:
        if l["codigo"] and not l["descripcion"]:
            l["descripcion"] = (nombres.get(l["codigo"]) or "").strip()[:300] or None


CAMPOS_GUARDAR = [c for c, _, _ in CAMPOS_TEXTO] + ["observaciones", "estado", "transporte", "fecha_factura"] + [c for c, _ in FECHAS]


def guardar(cursor, id_, datos, lineas, usuario, reemplazar_lineas=True):
    valores = [datos.get(c) for c in CAMPOS_GUARDAR]
    if id_:
        cursor.execute(f"UPDATE imp_pedidos SET {', '.join(c + ' = %s' for c in CAMPOS_GUARDAR)}, actualizado_por = %s WHERE id = %s",
                       valores + [usuario, id_])
    else:
        cursor.execute(f"INSERT INTO imp_pedidos ({', '.join(CAMPOS_GUARDAR)}, creado_por, actualizado_por) "
                       f"VALUES ({', '.join(['%s'] * (len(CAMPOS_GUARDAR) + 2))})", valores + [usuario, usuario])
        id_ = cursor.lastrowid
    if reemplazar_lineas:
        cursor.execute("DELETE FROM imp_pedido_detalle WHERE pedido_id = %s", (id_,))
        if lineas:
            cursor.executemany("INSERT INTO imp_pedido_detalle (pedido_id, codigo, descripcion, cantidad_pedida, cantidad_recibida, costo_unitario) "
                               "VALUES (%s, %s, %s, %s, %s, %s)",
                               [(id_, l["codigo"], l["descripcion"], l["cantidad_pedida"], l["cantidad_recibida"], l["costo_unitario"])
                                for l in lineas])
    return id_


@pedidos_importacion_bp.route("/pedidos_importacion/nuevo", methods=["GET", "POST"])
@pedidos_importacion_bp.route("/pedidos_importacion/<int:id_>", methods=["GET", "POST"])
@requiere_acceso(VISTA)
def formulario(id_=None):
    proveedores, _ = lista_proveedores()
    contexto = {"estados": ESTADOS, "transportes": TRANSPORTES, "fechas": FECHAS, "tiempos_def": TIEMPOS, "proveedores": proveedores,
                "ffecha": ffecha, "color": COLOR_ESTADO}
    if request.method == "GET":
        if id_:
            p = leer_pedido(id_)
            if not p:
                flash("El pedido no existe.", "warning")
                return redirect(url_for("pedidos_importacion.lista"))
            return render_template("pedido_importacion_form.html", p=p, lineas=p["lineas"], avisos=revisar(p, p["lineas"]), **contexto)
        return render_template("pedido_importacion_form.html", p={"estado": "Cotizando"}, lineas=[], avisos=[], **contexto)

    datos, lineas, errores, avisos = leer_formulario(request.form)
    if datos["orden_compra"]:
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT id FROM imp_pedidos WHERE orden_compra = %s AND id <> %s", (datos["orden_compra"], id_ or 0))
            if cursor.fetchone():
                errores.append(f"Ya existe otro pedido con la orden de compra {datos['orden_compra']}.")
        finally:
            cursor.close()
            conn.close()
    if errores:
        datos["id"] = id_
        enriquecer(datos, date.today())
        return render_template("pedido_importacion_form.html", p=datos, lineas=lineas, errores=errores, avisos=avisos, **contexto)
    completar_descripciones(lineas)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        nuevo_id = guardar(cursor, id_, datos, lineas, session.get("usuario"))
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó. Error de base de datos: {e}", "danger")
        return redirect(request.url)
    finally:
        cursor.close()
        conn.close()
    flash(f"✅ Pedido {datos['orden_compra']} guardado." + (" Revise: " + " ".join(avisos) if avisos else ""), "success" if not avisos else "warning")
    return redirect(url_for("pedidos_importacion.formulario", id_=nuevo_id))


@pedidos_importacion_bp.route("/pedidos_importacion/<int:id_>/eliminar", methods=["POST"])
@requiere_acceso(VISTA)
def eliminar(id_):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT orden_compra FROM imp_pedidos WHERE id = %s", (id_,))
        fila = cursor.fetchone()
        cursor.execute("DELETE FROM imp_pedidos WHERE id = %s", (id_,))
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash(f"Se eliminó el pedido {fila[0] if fila else ''}.", "info")
    return redirect(url_for("pedidos_importacion.lista"))


@pedidos_importacion_bp.route("/pedidos_importacion/producto")
@requiere_acceso(VISTA)
def producto():
    """Descripción del producto para llenar la línea al escribir el código."""
    codigo = (request.args.get("codigo") or "").strip().upper()
    nombre = productos_catalogo({codigo}).get(codigo) if codigo else None
    return jsonify({"codigo": codigo, "descripcion": (nombre or "").strip(), "existe": nombre is not None})


# -----------------------------
# Reportes y Excel
# -----------------------------
def filtros_reporte(args):
    return {"desde": fecha_param(args.get("desde")), "hasta": fecha_param(args.get("hasta")),
            "proveedor": (args.get("proveedor") or "").strip(), "transporte": args.get("transporte") or ""}


def datos_reporte(f):
    pedidos = [p for p in leer_pedidos(f) if p["estado"] != "Cancelado"]
    abiertos = sorted([p for p in leer_pedidos({**f, "estado": "abiertos"})], key=lambda p: -(p["dias_etapa"] or -1))
    return {
        "pedidos": pedidos,
        "por_proveedor": promedios_por(pedidos, lambda p: p["proveedor"]),
        "por_transporte": promedios_por(pedidos, lambda p: p["transporte"]),
        "general": promedios_por(pedidos, lambda p: "Todos")[0] if pedidos else None,
        "abiertos": abiertos,
        "atrasados": sorted([p for p in abiertos if p["atraso_hoy"]], key=lambda p: -p["atraso_hoy"]),
        "llegaron_tarde": sorted([p for p in pedidos if (p["t"]["atraso"] or 0) > 0], key=lambda p: -p["t"]["atraso"]),
    }


@pedidos_importacion_bp.route("/pedidos_importacion/reportes")
@requiere_acceso(VISTA)
def reportes():
    f = filtros_reporte(request.args)
    proveedores, _ = lista_proveedores()
    return render_template("pedidos_importacion_reportes.html", f=f, d=datos_reporte(f), tiempos_def=TIEMPOS, transportes=TRANSPORTES,
                           proveedores=proveedores, ffecha=ffecha, color=COLOR_ESTADO,
                           params={k: (v.isoformat() if isinstance(v, date) else v) for k, v in f.items() if v})


def hoja(wb, titulo, encabezados, filas, formatos=None, primera=False):
    ws = wb.active if primera else wb.create_sheet(titulo)
    ws.title = titulo
    for i, (nombre, ancho) in enumerate(encabezados, start=1):
        c = ws.cell(row=1, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre.startswith(("Días total", "Cantidad sugerida", "Fecha para pedir")) else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[1].height = 32
    for n, fila in enumerate(filas, start=2):
        for i, v in enumerate(fila, start=1):
            c = ws.cell(row=n, column=i, value=v)
            if isinstance(v, (date, datetime)):
                c.number_format = "dd/mm/yyyy"
            elif isinstance(v, float):
                c.number_format = (formatos or {}).get(i, "#,##0.0")
            elif isinstance(v, str) and v.startswith("="):
                c.data_type = "s"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(encabezados))}{max(len(filas) + 1, 1)}"
    return ws


def filas_promedio(grupos):
    return [[g["grupo"], g["pedidos"], g["recibidos"]] + [round(g["tiempos"][k][0], 1) if g["tiempos"][k][0] is not None else None
                                                         for k, *_ in TIEMPOS] for g in grupos]


@pedidos_importacion_bp.route("/pedidos_importacion/excel")
@requiere_acceso(VISTA)
def excel():
    f = filtros_reporte(request.args)
    d = datos_reporte(f)
    wb = Workbook()
    nombres_t = [(nombre, 11) for _, nombre, *_ in TIEMPOS]
    hoja(wb, "Resumen de pedidos",
         [("N° orden de compra", 16), ("Proveedor", 28), ("País", 12), ("N° cotización", 13), ("N° factura", 13), ("Fecha factura", 11),
          ("Descripción", 34), ("Transporte", 11), ("Estado", 12)] + [(n, 11) for _, n in FECHAS] + nombres_t
         + [("Días en etapa actual", 10), ("Días atrasado hoy", 10), ("Líneas", 8), ("Unidades pedidas", 11), ("Unidades recibidas", 11),
            ("Monto (cant. × costo)", 14), ("Observaciones", 40)],
         [[p["orden_compra"], p["proveedor"], p["pais"], p["cotizacion"], p["factura"], p["fecha_factura"], p["descripcion"], p["transporte"],
           p["estado"]] + [p[c] for c, _ in FECHAS] + [p["t"][k] for k, *_ in TIEMPOS]
          + [p["dias_etapa"], p["atraso_hoy"], p["lineas"], p["unidades"], p["recibidas"], p["monto"], p["observaciones"]] for p in d["pedidos"]],
         formatos={9 + len(FECHAS) + len(TIEMPOS) + 4: "#,##0.##", 9 + len(FECHAS) + len(TIEMPOS) + 5: "#,##0.##",
                   9 + len(FECHAS) + len(TIEMPOS) + 6: "#,##0.00"}, primera=True)
    enc_prom = [("Proveedor", 30), ("Pedidos", 9), ("Recibidos", 9)] + nombres_t
    hoja(wb, "Promedio por proveedor", enc_prom, filas_promedio(d["por_proveedor"] + ([{**d["general"], "grupo": "TODOS"}] if d["general"] else [])))
    hoja(wb, "Promedio por transporte", [("Transporte", 16)] + enc_prom[1:],
         filas_promedio(d["por_transporte"] + ([{**d["general"], "grupo": "TODOS"}] if d["general"] else [])))
    hoja(wb, "Abiertos", [("N° orden de compra", 16), ("Proveedor", 28), ("Transporte", 11), ("Estado", 12), ("En la etapa desde", 12),
                          ("Días en etapa actual", 11), ("Llegada estimada", 12), ("Días atrasado hoy", 11), ("Descripción", 40)],
         [[p["orden_compra"], p["proveedor"], p["transporte"], p["estado"], p["etapa_desde"], p["dias_etapa"], p["f_llegada_estimada"],
           p["atraso_hoy"], p["descripcion"]] for p in d["abiertos"]])
    hoja(wb, "Atrasados", [("N° orden de compra", 16), ("Proveedor", 28), ("Transporte", 11), ("Estado", 12), ("Llegada estimada", 12),
                           ("Llegada real", 12), ("Días de atraso", 10), ("Situación", 26)],
         [[p["orden_compra"], p["proveedor"], p["transporte"], p["estado"], p["f_llegada_estimada"], None, p["atraso_hoy"], "Abierto: aún no llega"]
          for p in d["atrasados"]]
         + [[p["orden_compra"], p["proveedor"], p["transporte"], p["estado"], p["f_llegada_estimada"], p["f_llegada_real"], p["t"]["atraso"],
             "Llegó tarde"] for p in d["llegaron_tarde"]])
    ayuda = wb.create_sheet("Cómo se calcula")
    ayuda.column_dimensions["A"].width = 30
    ayuda.column_dimensions["B"].width = 60
    for n, (nombre, formula) in enumerate([("Tiempo", "Cálculo (días)")] + [(nombre, expl) for _, nombre, _, _, expl in TIEMPOS]
                                          + [("", ""), ("Vacío", "Falta alguna de las dos fechas: no se calcula ni entra en el promedio."),
                                             ("Días en etapa actual", "Hoy − fecha en que empezó la etapa del estado actual."),
                                             ("Días atrasado hoy", "Hoy − llegada estimada, si ya pasó y todavía no llega.")], start=1):
        ayuda.cell(row=n, column=1, value=nombre).font = Font(bold=n == 1)
        ayuda.cell(row=n, column=2, value=formula).font = Font(bold=n == 1)
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"pedidos_importacion_{date.today():%Y-%m-%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------
# Importar pedidos desde Excel
# -----------------------------
COLUMNAS_IMPORTAR = {
    "orden_compra": {"NORDENDECOMPRA", "NOORDENDECOMPRA", "ORDENDECOMPRA", "ORDENCOMPRA", "NUMEROORDENDECOMPRA", "OC", "NOC", "NOOC"},
    "proveedor": {"PROVEEDOR"},
    "pais": {"PAISDELPROVEEDOR", "PAIS", "PAISPROVEEDOR"},
    "cotizacion": {"NCOTIZACION", "NOCOTIZACION", "COTIZACION", "NUMEROCOTIZACION", "NUMERODECOTIZACION"},
    "factura": {"NFACTURA", "NOFACTURA", "FACTURA", "NUMEROFACTURA", "NUMERODEFACTURA"},
    "fecha_factura": {"FECHAFACTURA", "FECHADEFACTURA"},
    "descripcion": {"DESCRIPCION"},
    "transporte": {"TRANSPORTE"},
    "estado": {"ESTADO"},
    "f_solicitud_cot": {"SOLICITUDCOTIZACION", "FECHASOLICITUDCOTIZACION", "SOLICITUDDECOTIZACION"},
    "f_cot_recibida": {"COTIZACIONRECIBIDA", "FECHACOTIZACIONRECIBIDA"},
    "f_pedido": {"FECHAPEDIDO", "FECHADEPEDIDO"},
    "f_embarque": {"FECHAEMBARQUE", "FECHADEEMBARQUE", "EMBARQUE"},
    "f_llegada_estimada": {"LLEGADAESTIMADA", "FECHALLEGADAESTIMADA", "ETA"},
    "f_llegada_real": {"LLEGADAREAL", "FECHALLEGADAREAL"},
    "f_liberacion": {"LIBERACIONADUANA", "FECHALIBERACIONADUANA", "LIBERACIONDEADUANA"},
    "f_recepcion": {"RECEPCIONBODEGA", "FECHARECEPCIONBODEGA", "RECEPCIONENBODEGA"},
    "observaciones": {"OBSERVACIONES", "OBSERVACION", "NOTAS"},
}
COLUMNAS_DETALLE = {
    "orden_compra": COLUMNAS_IMPORTAR["orden_compra"],
    "codigo": {"CODIGO", "CODIGODEPRODUCTO", "CODIGOPRODUCTO"},
    "descripcion": {"DESCRIPCION"},
    "cantidad_pedida": {"CANTIDADPEDIDA", "CANTIDAD", "CANTPEDIDA"},
    "cantidad_recibida": {"CANTIDADRECIBIDA", "CANTRECIBIDA", "RECIBIDA"},
    "costo_unitario": {"COSTOUNITARIO", "COSTO", "COSTOUNIT"},
}
ENC_PLANTILLA = ["N° orden de compra", "Proveedor", "País del proveedor", "N° cotización", "N° factura", "Fecha factura", "Descripción",
                 "Transporte", "Estado"] + [n for _, n in FECHAS] + ["Observaciones"]
ENC_DETALLE = ["N° orden de compra", "Código de producto", "Descripción", "Cantidad pedida", "Cantidad recibida", "Costo unitario"]


def nombre_columna(valor):
    return "".join(ch for ch in sin_tildes(a_texto(valor)).upper() if ch.isalnum())


def ubicar(filas, columnas, requerida):
    for i, fila in enumerate(filas[:20]):
        nombres = [nombre_columna(v) for v in fila]
        mapa = {campo: next((j for j, n in enumerate(nombres) if n in opciones), None) for campo, opciones in columnas.items()}
        if mapa[requerida] is not None and sum(v is not None for v in mapa.values()) >= 2:
            return mapa, i + 1
    return None, None


def leer_importacion(archivo):
    """-> (pedidos {oc: {campo: valor}}, detalle {oc: [líneas]}, errores, avisos). Celda vacía = no se cambia ese dato."""
    contenido = archivo.read()
    try:
        hojas = pd.ExcelFile(io.BytesIO(contenido)).sheet_names
    except Exception as e:
        raise ErrorArchivo(f"no se pudo abrir como Excel ({e})")
    hoja_det = next((h for h in hojas if clave(h) == "DETALLE"), None)
    hoja_ped = next((h for h in hojas if h != hoja_det), hojas[0])
    filas = leer_filas(io.BytesIO(contenido), hoja_ped)
    mapa, inicio = ubicar(filas, COLUMNAS_IMPORTAR, "orden_compra")
    if mapa is None:
        raise ErrorArchivo("no se encontró la columna «N° orden de compra» (use la plantilla)")
    pedidos, errores, avisos = {}, [], []
    for n, fila in enumerate(filas[inicio:], start=inicio + 1):
        valor = lambda campo: fila[mapa[campo]] if mapa[campo] is not None and mapa[campo] < len(fila) else None
        oc = " ".join(a_texto(valor("orden_compra")).split())[:40]
        if not oc:
            if any(a_texto(v) for v in fila):
                errores.append(f"Fila {n}: no tiene N° de orden de compra.")
            continue
        p = {}
        for campo, _, largo in CAMPOS_TEXTO[1:]:
            if mapa[campo] is not None and a_texto(valor(campo)):
                p[campo] = " ".join(a_texto(valor(campo)).split())[:largo]
        if mapa["observaciones"] is not None and a_texto(valor("observaciones")):
            p["observaciones"] = a_texto(valor("observaciones"))
        if mapa["estado"] is not None and a_texto(valor("estado")):
            p["estado"] = normal_estado(valor("estado"))
            if not p["estado"]:
                errores.append(f"Fila {n} ({oc}): el estado «{a_texto(valor('estado'))}» no es válido ({', '.join(ESTADOS)}).")
                continue
        if mapa["transporte"] is not None and a_texto(valor("transporte")):
            p["transporte"] = normal_transporte(valor("transporte"))
            if not p["transporte"]:
                errores.append(f"Fila {n} ({oc}): el transporte «{a_texto(valor('transporte'))}» no es Marítimo, Aéreo ni Terrestre.")
                continue
        mal = False
        for campo in ["fecha_factura"] + [c for c, _ in FECHAS]:
            if mapa[campo] is None:
                continue
            fecha = a_fecha(valor(campo))
            if fecha == "error":
                errores.append(f"Fila {n} ({oc}): la fecha «{a_texto(valor(campo))}» no es válida (use dd/mm/aaaa).")
                mal = True
            elif fecha:
                p[campo] = fecha
        if mal:
            continue
        if oc in pedidos:
            avisos.append(f"La orden {oc} viene más de una vez: se juntan los datos y queda el último valor de cada columna.")
        pedidos.setdefault(oc, {}).update(p)

    detalle = {}
    if hoja_det:
        filas = leer_filas(io.BytesIO(contenido), hoja_det)
        mapa, inicio = ubicar(filas, COLUMNAS_DETALLE, "orden_compra")
        if mapa is None:
            errores.append("La hoja «Detalle» no tiene la columna «N° orden de compra»: no se cargaron líneas.")
        else:
            for n, fila in enumerate(filas[inicio:], start=inicio + 1):
                valor = lambda campo: fila[mapa[campo]] if mapa[campo] is not None and mapa[campo] < len(fila) else None
                oc = " ".join(a_texto(valor("orden_compra")).split())[:40]
                cod = " ".join(a_texto(valor("codigo")).split()).upper()[:20]
                desc = " ".join(a_texto(valor("descripcion")).split())[:300]
                if not oc or not (cod or desc):
                    continue
                nums = {c: a_numero(valor(c)) for c in ("cantidad_pedida", "cantidad_recibida", "costo_unitario")}
                if any(v == "error" or (v is not None and v < 0) for v in nums.values()):
                    errores.append(f"Detalle fila {n} ({oc} {cod}): cantidad o costo no válido.")
                    continue
                detalle.setdefault(oc, []).append({"codigo": cod or None, "descripcion": desc or None,
                                                   "cantidad_pedida": nums["cantidad_pedida"] or 0,
                                                   "cantidad_recibida": nums["cantidad_recibida"], "costo_unitario": nums["costo_unitario"]})
            sin_pedido = sorted(set(detalle) - set(pedidos))
            if sin_pedido:
                avisos.append(f"El detalle trae órdenes que no están en la hoja de pedidos ({', '.join(sin_pedido[:8])}): "
                              f"se cargan si el pedido ya existe en el sistema.")
    return pedidos, detalle, errores, avisos


def a_json(valor):
    return valor.isoformat() if isinstance(valor, date) else valor


@pedidos_importacion_bp.route("/pedidos_importacion/importar", methods=["GET", "POST"])
@requiere_acceso(VISTA)
def importar():
    if request.method == "GET":
        return render_template("pedidos_importacion_importar.html", ffecha=ffecha)
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename.lower().endswith((".xlsx", ".xls")):
        return render_template("pedidos_importacion_importar.html", errores_archivo=["Seleccione un archivo de Excel (.xlsx o .xls)."], ffecha=ffecha)
    try:
        pedidos, detalle, errores, avisos = leer_importacion(archivo)
    except ErrorArchivo as e:
        return render_template("pedidos_importacion_importar.html", errores_archivo=[f"{archivo.filename}: {e}."], ffecha=ffecha)
    except Exception as e:
        return render_template("pedidos_importacion_importar.html", errores_archivo=[f"{archivo.filename}: error al leer – {e}"], ffecha=ffecha)

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM imp_pedidos")
        existentes = {r["orden_compra"]: r for r in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()
    detalle = {oc: l for oc, l in detalle.items() if oc in pedidos or oc in existentes}
    filas = []
    for oc, p in pedidos.items():
        actual = existentes.get(oc)
        final = {**(actual or {"estado": None}), **p, "orden_compra": oc}
        if not final.get("estado"):
            final["estado"] = estado_por_fechas(final)
        cambios = [c for c in p if actual and actual.get(c) != p[c]]
        filas.append({"oc": oc, "nuevo": actual is None, "cambios": len(cambios), "p": final, "lineas": len(detalle.get(oc, [])),
                      "avisos": [a for a in revisar({**{c: None for c in CAMPOS_GUARDAR}, **final}, [])
                                 if not a.startswith("Códigos")]})
    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as fh:
        json.dump({"pedidos": {oc: {k: a_json(v) for k, v in p.items()} for oc, p in pedidos.items()}, "detalle": detalle,
                   "archivo": archivo.filename}, fh, ensure_ascii=False)
    session["imp_pedidos_token"] = token
    return render_template("pedidos_importacion_importar.html", validado=True, token=token, archivo=archivo.filename, filas=filas,
                           errores=errores, avisos=avisos, nuevos=sum(f["nuevo"] for f in filas),
                           actualizados=sum(not f["nuevo"] for f in filas), lineas=sum(len(l) for l in detalle.values()),
                           detalle_sin_pedido=sorted(set(detalle) - set(pedidos)), ffecha=ffecha, color=COLOR_ESTADO)


@pedidos_importacion_bp.route("/pedidos_importacion/importar/confirmar", methods=["POST"])
@requiere_acceso(VISTA)
def confirmar_importacion():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("imp_pedidos_token") or not os.path.exists(ruta):
        flash("⚠️ La revisión expiró o ya se cargó. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("pedidos_importacion.importar"))
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    usuario = session.get("usuario")
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    nuevos = actualizados = 0
    try:
        conn.start_transaction()
        cursor.execute("SELECT * FROM imp_pedidos")
        existentes = {r["orden_compra"]: r for r in cursor.fetchall()}
        ids = {oc: r["id"] for oc, r in existentes.items()}
        for oc, p in datos["pedidos"].items():
            p = {k: (date.fromisoformat(v) if k.startswith("f") and k != "factura" and isinstance(v, str) else v) for k, v in p.items()}
            actual = existentes.get(oc)
            final = {**{c: None for c in CAMPOS_GUARDAR}, **(actual or {}), **p, "orden_compra": oc}
            if not final.get("estado"):
                final["estado"] = estado_por_fechas(final)
            ids[oc] = guardar(cursor, actual["id"] if actual else None, final, [], usuario, reemplazar_lineas=False)
            nuevos += actual is None
            actualizados += actual is not None
        lineas = 0
        for oc, lista in datos["detalle"].items():
            if oc not in ids:
                continue
            completar_descripciones(lista)
            cursor.execute("DELETE FROM imp_pedido_detalle WHERE pedido_id = %s", (ids[oc],))
            cursor.executemany("INSERT INTO imp_pedido_detalle (pedido_id, codigo, descripcion, cantidad_pedida, cantidad_recibida, costo_unitario) "
                               "VALUES (%s, %s, %s, %s, %s, %s)",
                               [(ids[oc], l["codigo"], l["descripcion"], l["cantidad_pedida"], l["cantidad_recibida"], l["costo_unitario"])
                                for l in lista])
            lineas += len(lista)
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó. Error de base de datos: {e}", "danger")
        return redirect(url_for("pedidos_importacion.importar"))
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop("imp_pedidos_token", None)
    flash(f"✅ Importación lista: {nuevos} pedido(s) nuevo(s), {actualizados} actualizado(s)" + (f", {lineas} línea(s) de detalle." if lineas else "."),
          "success")
    return redirect(url_for("pedidos_importacion.lista", estado=""))


@pedidos_importacion_bp.route("/pedidos_importacion/plantilla")
@requiere_acceso(VISTA)
def plantilla():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos"
    for i, nombre in enumerate(ENC_PLANTILLA, start=1):
        c = ws.cell(row=1, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre == "N° orden de compra" else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = 30 if nombre in ("Proveedor", "Descripción", "Observaciones") else 15
    ws.row_dimensions[1].height = 32
    for fila in range(2, 502):
        for i, nombre in enumerate(ENC_PLANTILLA, start=1):
            ws.cell(row=fila, column=i).number_format = "dd/mm/yyyy" if nombre.startswith(("Fecha", "Solicitud", "Cotización recibida", "Llegada",
                                                                                            "Liberación", "Recepción")) else "@"
    ws.freeze_panes = "B2"
    det = wb.create_sheet("Detalle")
    for i, nombre in enumerate(ENC_DETALLE, start=1):
        c = det.cell(row=1, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=VERDE)
        det.column_dimensions[get_column_letter(i)].width = 40 if nombre == "Descripción" else 16
    for fila in range(2, 1002):
        for col in (1, 2):
            det.cell(row=fila, column=col).number_format = "@"
    ayuda = wb.create_sheet("Instrucciones")
    ayuda.column_dimensions["A"].width = 110
    for n, (texto, negrita) in enumerate([
        ("Importar pedidos de importación", True), ("", False),
        ("Hoja «Pedidos»: una fila por pedido. Solo el N° de orden de compra es obligatorio; si ya existe en el sistema, se actualiza.", False),
        ("Celda vacía = ese dato no se cambia (no se borra lo que ya tenía el pedido).", False),
        ("Fechas en formato dd/mm/aaaa (o fecha de Excel). Todas son opcionales.", False),
        (f"Transporte: {', '.join(TRANSPORTES)}.   Estado: {', '.join(ESTADOS)}.", False),
        ("   Si el estado viene vacío en un pedido nuevo, se pone según la última fecha llenada.", False),
        ("Hoja «Detalle» (opcional): una fila por producto, con el N° de orden de compra del pedido al que pertenece.", False),
        ("   Si una orden trae detalle, sus líneas reemplazan las que tenía el pedido. La descripción se completa del catálogo si viene vacía.", False),
    ], start=1):
        ayuda.cell(row=n, column=1, value=texto).font = Font(bold=True, size=13, color=VERDE) if negrita else Font(size=10)
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name="plantilla_pedidos_importacion.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------
# Pedido sugerido por tiempo de entrega
# -----------------------------
def entero(valor, defecto, minimo=0, maximo=3650):
    try:
        v = int(float(valor))
    except (TypeError, ValueError):
        return defecto
    return min(max(v, minimo), maximo)


def parametros_sugerido(args):
    return {"proveedor": (args.get("proveedor") or "").strip(), "transporte": args.get("transporte") if args.get("transporte") in TRANSPORTES else "",
            "meses": entero(args.get("meses"), 6) if entero(args.get("meses"), 6) in MESES_CONSUMO else 6,
            "dias_cubrir": entero(args.get("dias_cubrir"), 90) if entero(args.get("dias_cubrir"), 90) in DIAS_CUBRIR else 90,
            "seguridad": entero(args.get("seguridad"), SEGURIDAD_DEFECTO, 0, 365),
            "entrega_manual": entero(args.get("entrega_manual"), None, 1, 730) if (args.get("entrega_manual") or "").strip() else None,
            "sin_venta": args.get("sin_venta") == "1"}


def mismo_proveedor(a, b):
    a, b = clave(a), clave(b)
    return bool(a and b) and (a == b or a in b or b in a)


def tiempo_entrega(proveedor, transporte):
    """Promedios del historial del proveedor: (resumen, tiempo de entrega en días o None)."""
    historial = [p for p in leer_pedidos() if mismo_proveedor(p["proveedor"], proveedor) and p["estado"] != "Cancelado"]
    con_transporte = [p for p in historial if p["transporte"] == transporte] if transporte else historial
    base = con_transporte if any(p["t"]["total"] is not None for p in con_transporte) else historial
    r = {k: promedio(p["t"][k] for p in base) for k, *_ in TIEMPOS}
    r["previo"] = promedio(dias(p["f_solicitud_cot"], p["f_pedido"]) for p in base)   # de pedir cotización a hacer el pedido
    r["pedidos"], r["usa_todos"] = len(base), bool(transporte) and base is historial and bool(historial)
    r["abiertos"] = [p for p in historial if p["abierto"]]
    return r, (round(r["total"][0]) if r["total"][0] is not None else None)


def calcular_sugerido(prm):
    hoy = date.today()
    hist, entrega_hist = tiempo_entrega(prm["proveedor"], prm["transporte"])
    entrega = prm["entrega_manual"] or entrega_hist
    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Productos del proveedor (última compra en el catálogo) y los que ya se le pidieron en pedidos registrados
        cursor.execute("SELECT `Codigo Producto`, Nombre, `Numero Parte`, Marca, `Ultimo Costo`, `Ultimo Precio FOB`, `Proveedor Ultima Compra` "
                       "FROM catalogo_productos WHERE `Proveedor Ultima Compra` IS NOT NULL AND TRIM(`Proveedor Ultima Compra`) <> ''")
        productos = {r[0]: {"codigo": r[0], "nombre": (r[1] or "").strip(), "parte": r[2], "marca": r[3], "costo": float(r[4] or 0),
                            "fob": float(r[5]) if r[5] else None} for r in cursor.fetchall() if clave(r[6]) == clave(prm["proveedor"])}
        cursor.execute("SELECT d.codigo, d.cantidad_pedida, d.cantidad_recibida, p.estado, p.proveedor, p.orden_compra, p.f_llegada_estimada "
                       "FROM imp_pedido_detalle d JOIN imp_pedidos p ON p.id = d.pedido_id WHERE d.codigo IS NOT NULL")
        en_camino, pedidos_cod, extra = {}, {}, set()
        for cod, ped, rec, estado, prov, oc, eta in cursor.fetchall():
            if not mismo_proveedor(prov, prm["proveedor"]):
                continue
            if cod not in productos:
                extra.add(cod)
            if estado not in CERRADOS and estado != "Cotizando":
                pend = max(float(ped or 0) - float(rec or 0), 0)
                if pend:
                    en_camino[cod] = en_camino.get(cod, 0) + pend
                    pedidos_cod.setdefault(cod, []).append(f"{oc}: {pend:,.0f}" + (f" (llega {ffecha(eta)})" if eta else ""))
        if extra:
            cursor.execute(f"SELECT `Codigo Producto`, Nombre, `Numero Parte`, Marca, `Ultimo Costo`, `Ultimo Precio FOB` FROM catalogo_productos "
                           f"WHERE `Codigo Producto` IN ({', '.join(['%s'] * len(extra))})", sorted(extra))
            for r in cursor.fetchall():
                productos[r[0]] = {"codigo": r[0], "nombre": (r[1] or "").strip(), "parte": r[2], "marca": r[3], "costo": float(r[4] or 0),
                                   "fob": float(r[5]) if r[5] else None}
        codigos = sorted(productos)
        if not codigos:
            return {"filas": [], "hist": hist, "entrega": entrega, "entrega_hist": entrega_hist, "hoy": hoy, "periodo": None}
        lista = ", ".join(["%s"] * len(codigos))
        cursor.execute("SELECT MAX(Fecha) FROM ventas_por_producto")
        fin = cursor.fetchone()[0]
        fin = date.fromisoformat(str(fin)[:10]) if fin else hoy
        inicio = fin - timedelta(days=round(prm["meses"] * DIAS_MES)) + timedelta(days=1)
        cursor.execute(f"SELECT Codigo, SUM(`Cantidad Vendida`) FROM ventas_por_producto WHERE Fecha BETWEEN %s AND %s AND Estatus <> 'Anulado' "
                       f"AND Codigo IN ({lista}) GROUP BY Codigo", [inicio.isoformat(), fin.isoformat()] + codigos)
        ventas = {c: float(n or 0) for c, n in cursor.fetchall()}
        cursor.execute(f"SELECT Codigo_producto, SUM(Cant_Final), SUM(CASE WHEN UPPER(Bodega) LIKE '%NHE%' THEN Cant_Final ELSE 0 END) "
                       f"FROM existencias_bodega WHERE Codigo_producto IN ({lista}) GROUP BY Codigo_producto", codigos)
        stock = {c: float(t or 0) - float(n or 0) for c, t, n in cursor.fetchall()}
        cursor.execute(f"SELECT codigo, dias FROM imp_stock_seguridad WHERE codigo IN ({lista})", codigos)
        seguridad = dict(cursor.fetchall())
        minimos = minimos_por_codigo(cursor, codigos)
    finally:
        cursor.close()
        conn.close()

    dias_periodo = (fin - inicio).days + 1
    filas = []
    for cod in codigos:
        x = productos[cod]
        vendido = ventas.get(cod, 0.0)
        if vendido <= 0 and not prm["sin_venta"] and not en_camino.get(cod):
            continue
        diario = vendido / dias_periodo if vendido > 0 else 0.0
        disp, camino = stock.get(cod, 0.0), en_camino.get(cod, 0.0)
        seg = seguridad.get(cod, prm["seguridad"])
        fila = {**x, "vendido": vendido, "mensual": diario * DIAS_MES, "diario": diario, "stock": disp, "camino": camino,
                "detalle_camino": "; ".join(pedidos_cod.get(cod, [])), "seguridad": seg, "seguridad_propia": cod in seguridad,
                "minimo": minimos.get(cod), "cobertura": None, "reorden": None, "fecha_pedir": None, "dias_para_pedir": None,
                "fecha_cotizar": None, "sugerido": 0, "urgencia": "sin_venta"}
        if diario > 0:
            fila["cobertura"] = max(disp, 0) / diario + camino / diario
            if entrega:
                fila["reorden"] = diario * (entrega + seg)
                fila["dias_para_pedir"] = math.floor(fila["cobertura"] - entrega - seg)
                fila["fecha_pedir"] = hoy + timedelta(days=fila["dias_para_pedir"])
                if hist["previo"][0] is not None:
                    fila["fecha_cotizar"] = fila["fecha_pedir"] - timedelta(days=round(hist["previo"][0]))
                bruto = math.ceil(diario * (entrega + seg + prm["dias_cubrir"]) - max(disp, 0) - camino - 1e-9)
                fila["sugerido"] = ajustar_a_minimo(max(bruto, 0), fila["minimo"])
                d = fila["dias_para_pedir"]
                fila["urgencia"] = "ya" if d <= 0 else "pronto" if d <= 30 else "despues" if d <= prm["dias_cubrir"] else "cubierto"
            else:
                fila["urgencia"] = "sin_tiempo"
        fila["valor"] = fila["sugerido"] * fila["costo"] if fila["sugerido"] else 0
        fila["valor_fob"] = fila["sugerido"] * fila["fob"] if fila["sugerido"] and fila["fob"] else None
        filas.append(fila)
    orden = {"ya": 0, "pronto": 1, "despues": 2, "cubierto": 3, "sin_tiempo": 4, "sin_venta": 5}
    filas.sort(key=lambda f: (orden[f["urgencia"]], f["dias_para_pedir"] if f["dias_para_pedir"] is not None else 99999, f["codigo"]))
    con_fecha = [f for f in filas if f["fecha_pedir"]]
    proxima = min((f["fecha_pedir"] for f in con_fecha), default=None)
    return {"filas": filas, "hist": hist, "entrega": entrega, "entrega_hist": entrega_hist, "hoy": hoy, "periodo": (inicio, fin),
            "proxima": proxima, "proxima_cotizar": min((f["fecha_cotizar"] for f in con_fecha if f["fecha_cotizar"]), default=None),
            "pedir_ya": sum(1 for f in filas if f["urgencia"] == "ya"), "pronto": sum(1 for f in filas if f["urgencia"] == "pronto"),
            "total_sugerido": sum(f["valor"] for f in filas), "total_fob": sum(f["valor_fob"] or 0 for f in filas),
            "con_sugerido": sum(1 for f in filas if f["sugerido"])}


URGENCIAS = {"ya": ("Pedir ya", "danger"), "pronto": ("En 30 días o menos", "warning"), "despues": ("Dentro del período a cubrir", "info"),
             "cubierto": ("Cubierto", "success"), "sin_tiempo": ("Falta tiempo de entrega", "secondary"), "sin_venta": ("Sin venta", "light")}


@pedidos_importacion_bp.route("/pedidos_importacion/sugerido")
@requiere_acceso(VISTA)
def sugerido():
    prm = parametros_sugerido(request.args)
    proveedores, catalogo = lista_proveedores()
    r = calcular_sugerido(prm) if prm["proveedor"] else None
    return render_template("pedidos_importacion_sugerido.html", prm=prm, r=r, proveedores=proveedores, conteo=catalogo, transportes=TRANSPORTES,
                           meses_opc=MESES_CONSUMO, dias_opc=DIAS_CUBRIR, urgencias=URGENCIAS, ffecha=ffecha, tiempos_def=TIEMPOS,
                           params={k: v for k, v in request.args.items()})


@pedidos_importacion_bp.route("/pedidos_importacion/sugerido/excel")
@requiere_acceso(VISTA)
def sugerido_excel():
    prm = parametros_sugerido(request.args)
    if not prm["proveedor"]:
        return redirect(url_for("pedidos_importacion.sugerido"))
    r = calcular_sugerido(prm)
    wb = Workbook()
    ws = hoja(wb, "Sugerido", [
        ("Código", 12), ("Nº parte", 16), ("Descripción", 38), ("Marca", 14), ("Consumo mensual", 10), ("Stock disponible", 10),
        ("En camino", 10), ("Detalle en camino", 30), ("Días de cobertura", 10), ("Stock seguridad (días)", 10), ("Punto de reorden (unid.)", 11),
        ("Fecha para pedir", 12), ("Días para pedir", 9), ("Fecha para cotizar", 12), ("Situación", 18), ("Cant. mínima", 9),
        ("Cantidad sugerida", 11), ("Último costo C$", 12), ("Valor sugerido C$", 14), ("Precio FOB US$", 11), ("Valor FOB US$", 13)],
        [[f["codigo"], f["parte"], f["nombre"], f["marca"], round(f["mensual"], 2), f["stock"], f["camino"] or None, f["detalle_camino"] or None,
          round(f["cobertura"], 1) if f["cobertura"] is not None else None, f["seguridad"],
          round(f["reorden"], 1) if f["reorden"] is not None else None, f["fecha_pedir"], f["dias_para_pedir"], f["fecha_cotizar"],
          URGENCIAS[f["urgencia"]][0], f["minimo"], f["sugerido"] or None, f["costo"] or None, f["valor"] or None, f["fob"], f["valor_fob"]]
         for f in r["filas"]], formatos={5: "#,##0.0", 6: "#,##0.##", 7: "#,##0.##", 18: "#,##0.00", 19: "#,##0.00", 20: "#,##0.00", 21: "#,##0.00"},
        primera=True)
    ws.insert_rows(1, 3)
    h = r["hist"]
    ws["A1"] = f"Pedido sugerido por tiempo de entrega · {prm['proveedor']}" + (f" · {prm['transporte']}" if prm["transporte"] else "")
    ws["A1"].font = Font(bold=True, size=13, color=VERDE)
    ws["A2"] = (f"Tiempo de entrega usado: {r['entrega'] or 'sin dato'} días"
                + (" (escrito a mano)" if prm["entrega_manual"] else f" (promedio de {h['total'][1]} pedido(s) recibido(s))" if r["entrega"] else "")
                + f" · Seguridad general {prm['seguridad']} días · Cubrir {prm['dias_cubrir']} días · Consumo de "
                + (f"{ffecha(r['periodo'][0])} a {ffecha(r['periodo'][1])}" if r["periodo"] else "—")
                + f" · Generado {datetime.now():%d/%m/%Y %H:%M}")
    ws["A2"].font = Font(size=9, color="52514E")
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:U{max(len(r['filas']) + 4, 4)}"
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = "".join(ch if ch.isalnum() else "_" for ch in prm["proveedor"])[:40]
    return send_file(salida, as_attachment=True, download_name=f"sugerido_tiempo_entrega_{nombre}_{date.today():%Y-%m-%d}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@pedidos_importacion_bp.route("/pedidos_importacion/seguridad", methods=["POST"])
@requiere_acceso(VISTA)
def guardar_seguridad():
    """Stock de seguridad por producto: número = días propios; vacío = usa el valor general."""
    guardar_, quitar = [], []
    for k, v in request.form.items():
        if not k.startswith("seg_"):
            continue
        cod, v = k[4:], v.strip()
        if v == "":
            quitar.append(cod)
        else:
            dias_ = entero(v, None, 0, 365)
            if dias_ is not None:
                guardar_.append((cod, dias_, session.get("usuario")))
    conn = get_connection()
    cursor = conn.cursor()
    try:
        if guardar_:
            cursor.executemany("INSERT INTO imp_stock_seguridad (codigo, dias, actualizado_por) VALUES (%s, %s, %s) "
                               "ON DUPLICATE KEY UPDATE dias = VALUES(dias), actualizado_por = VALUES(actualizado_por)", guardar_)
        if quitar:
            cursor.execute(f"DELETE FROM imp_stock_seguridad WHERE codigo IN ({', '.join(['%s'] * len(quitar))})", quitar)
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash(f"✅ Stock de seguridad guardado: {len(guardar_)} producto(s) con días propios" + (f", {len(quitar)} vuelven al valor general." if quitar else "."),
          "success")
    volver = request.form.get("volver") or ""
    return redirect(volver if volver.startswith("/") and not volver.startswith("//") else url_for("pedidos_importacion.sugerido"))
