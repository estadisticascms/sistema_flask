"""Entregas · parte 12: tablero y reportes (sección 8 de la especificación). Pantalla del administrador: /entregas/reportes.

Período (desde / hasta, por defecto el mes en curso). Cada reporte dice qué fecha usa:
- Tablero: la situación de ahora (envíos abiertos), más las órdenes y entregas de hoy.
- Tiempo por etapa (envíos creados en el período), en HORAS LABORALES (horario y feriados de Catálogos › Configuración,
  router/entregas_horario.py), con el historial ent_envio_eventos:
    orden → lista para facturar (POR_FACTURAR, la última vez que llegó), separando las órdenes con traslado de las que no;
    lista → facturada; facturada → empacada; empacada → entregada; y la duración de los traslados (aplicado → recibido).
- Entregas a tiempo (paradas entregadas en el período): a tiempo si la hora de entrega no pasa de la hora requerida de salida.
- Faltantes (envíos creados en el período): líneas marcadas como faltante por bodega (la bodega donde faltó, aunque después se
  cambiara) y por producto; venta perdida = líneas quitadas y la parte no disponible de las que fueron parciales (lo mismo que se
  escribe en ventas_perdidas), valorada con el subtotal de la orden.
- Fletes por transportista (paradas entregadas en el período): monto de los recibos; por envío, separando lo que paga el cliente de
  lo que paga la empresa (flete_paga). Lo que no se detalló por envío queda «sin detallar».
- Conciliación (facturas registradas en el período): cada factura de ent_envio_facturas se busca por número en la facturación que
  se carga a la aplicación (ventas_historicas). Si no aparece y su fecha es posterior a la última facturación cargada, se indica
  que falta cargar esa facturación.
"""
import io
import re
import statistics
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file

from router.entregas import contexto_entregas, requiere_rol
from router.entregas_alerta import ahora_nicaragua, duracion
from router.entregas_catalogos import TIPOS_TRANSPORTISTA
from router.entregas_catalogos import leer_configuracion
from router.entregas_flujo import ABIERTOS
from router.entregas_horario import Horario
from router.entregas_orden import NOMBRE_ESTADO
from utils.db import get_connection

entregas_reportes_bp = Blueprint("entregas_reportes", __name__)
entregas_reportes_bp.context_processor(contexto_entregas)

ETAPAS = [   # clave, nombre, estado inicial (None = creación de la orden), estado final
    ("orden_lista", "Orden → lista para facturar", None, "POR_FACTURAR"),
    ("lista_facturada", "Lista → facturada", "POR_FACTURAR", "FACTURADA"),
    ("facturada_empacada", "Facturada → empacada", "FACTURADA", "EMPACADA"),
    ("empacada_entregada", "Empacada → entregada", "EMPACADA", "ENTREGADA"),
]


def lista_sql(valores):
    return ", ".join(["%s"] * len(valores))


def minutos(a, b):
    return (b - a).total_seconds() / 60


def resumen_tiempos(valores):
    """Lista de minutos -> {n, promedio, mediana, maximo} (con textos)."""
    if not valores:
        return {"n": 0}
    prom, med, mx = statistics.mean(valores), statistics.median(valores), max(valores)
    return {"n": len(valores), "promedio": prom, "mediana": med, "maximo": mx,
            "promedio_texto": duracion(prom), "mediana_texto": duracion(med), "maximo_texto": duracion(mx)}


def sin_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", (texto or "").upper()) if unicodedata.category(c) != "Mn").strip()


# -----------------------------
# Tablero
# -----------------------------
def tablero(cursor):
    momento = ahora_nicaragua()
    cursor.execute(f"""SELECT estado, COUNT(*) AS n, SUM(riesgo_confirmado = 1 OR riesgo = 'rojo') AS rojo,
                              SUM(hora_requerida < %s) AS atrasados
                       FROM ent_envios WHERE estado IN ({lista_sql(ABIERTOS)}) GROUP BY estado""", [momento] + list(ABIERTOS))
    por_estado = {r["estado"]: r for r in cursor.fetchall()}
    estados = [{"estado": e, "nombre": NOMBRE_ESTADO[e], "n": por_estado.get(e, {}).get("n", 0),
                "rojo": int(por_estado.get(e, {}).get("rojo") or 0), "atrasados": int(por_estado.get(e, {}).get("atrasados") or 0)}
               for e in ABIERTOS]
    hoy = momento.date()
    cursor.execute("SELECT COUNT(*) AS n FROM ent_envios WHERE DATE(creado_en) = %s", (hoy,))
    ordenes_hoy = cursor.fetchone()["n"]
    cursor.execute("""SELECT COUNT(*) AS n FROM ent_envios e JOIN ent_ruta_paradas p ON p.id = e.parada_id
                      WHERE e.estado = 'ENTREGADA' AND DATE(p.entregado_en) = %s""", (hoy,))
    entregados_hoy = cursor.fetchone()["n"]
    return {"estados": estados, "abiertos": sum(e["n"] for e in estados), "rojo": sum(e["rojo"] for e in estados),
            "atrasados": sum(e["atrasados"] for e in estados), "ordenes_hoy": ordenes_hoy, "entregados_hoy": entregados_hoy}


# -----------------------------
# Tiempo por etapa
# -----------------------------
def tiempos(cursor, desde, hasta):
    horario = Horario(leer_configuracion(cursor))
    minutos = horario.minutos      # minutos laborales entre dos momentos
    cursor.execute("SELECT id, creado_en FROM ent_envios WHERE creado_en >= %s AND creado_en < %s AND estado <> 'ANULADA'",
                   (desde, hasta + timedelta(days=1)))
    envios = {e["id"]: e for e in cursor.fetchall()}
    vacio = {"etapas": [], "traslado": resumen_tiempos([]), "envios": 0, "horario": horario.texto()}
    if not envios:
        vacio["etapas"] = [{"clave": c, "nombre": n, **resumen_tiempos([])} for c, n, _, _ in ETAPAS]
        vacio["etapas"][0].update(con=resumen_tiempos([]), sin=resumen_tiempos([]))
        return vacio
    ids = list(envios)
    cursor.execute(f"""SELECT envio_id, a_estado, creado_en FROM ent_envio_eventos
                       WHERE campo = 'estado' AND envio_id IN ({lista_sql(ids)}) ORDER BY creado_en, id""", ids)
    llegadas = defaultdict(dict)        # envío -> estado -> última vez que llegó
    for ev in cursor.fetchall():
        llegadas[ev["envio_id"]][ev["a_estado"]] = ev["creado_en"]
    cursor.execute(f"SELECT DISTINCT envio_id FROM ent_envio_lineas WHERE traslado_estado <> 'NO_REQUIERE' AND envio_id IN ({lista_sql(ids)})", ids)
    con_traslado = {r["envio_id"] for r in cursor.fetchall()}
    valores = defaultdict(list)
    for eid, e in envios.items():
        ll = llegadas[eid]
        for clave, _, inicio, fin in ETAPAS:
            a = e["creado_en"] if inicio is None else ll.get(inicio)
            b = ll.get(fin)
            if a and b and b >= a:
                valores[clave].append(minutos(a, b))
                if clave == "orden_lista":
                    valores["orden_lista_con" if eid in con_traslado else "orden_lista_sin"].append(minutos(a, b))
    # Traslados: de aplicado a recibido, por línea
    cursor.execute(f"""SELECT linea_id, a_estado, creado_en FROM ent_envio_eventos
                       WHERE campo = 'traslado' AND a_estado IN ('APLICADO', 'RECIBIDO') AND envio_id IN ({lista_sql(ids)})
                       ORDER BY creado_en, id""", ids)
    marcas = defaultdict(dict)
    for ev in cursor.fetchall():
        marcas[ev["linea_id"]].setdefault(ev["a_estado"], ev["creado_en"])
    traslados = [minutos(m["APLICADO"], m["RECIBIDO"]) for m in marcas.values() if "APLICADO" in m and "RECIBIDO" in m]
    etapas = []
    for clave, nombre, _, _ in ETAPAS:
        fila = {"clave": clave, "nombre": nombre, **resumen_tiempos(valores[clave])}
        if clave == "orden_lista":
            fila["con"] = resumen_tiempos(valores["orden_lista_con"])
            fila["sin"] = resumen_tiempos(valores["orden_lista_sin"])
        etapas.append(fila)
    return {"etapas": etapas, "traslado": resumen_tiempos(traslados), "envios": len(envios), "horario": horario.texto()}


# -----------------------------
# Entregas a tiempo
# -----------------------------
def a_tiempo(cursor, desde, hasta):
    cursor.execute("""SELECT e.id, e.pedido_num, e.cliente, e.destino, e.hora_requerida, p.entregado_en, t.nombre AS transportista
                      FROM ent_envios e JOIN ent_ruta_paradas p ON p.id = e.parada_id JOIN ent_transportistas t ON t.id = e.transportista_id
                      WHERE e.estado = 'ENTREGADA' AND p.entregado_en >= %s AND p.entregado_en < %s ORDER BY p.entregado_en""",
                   (desde, hasta + timedelta(days=1)))
    filas = cursor.fetchall()
    por_t = defaultdict(lambda: {"total": 0, "a_tiempo": 0})
    atrasados = []
    for f in filas:
        f["atraso"] = minutos(f["hora_requerida"], f["entregado_en"])
        ok = f["atraso"] <= 0
        por_t[f["transportista"]]["total"] += 1
        por_t[f["transportista"]]["a_tiempo"] += ok
        if not ok:
            f["atraso_texto"] = duracion(f["atraso"])
            atrasados.append(f)
    total, bien = len(filas), sum(1 for f in filas if f["atraso"] <= 0)
    transportistas = [{"transportista": t, **v, "pct": v["a_tiempo"] * 100 / v["total"]} for t, v in sorted(por_t.items())]
    return {"total": total, "a_tiempo": bien, "pct": bien * 100 / total if total else None, "transportistas": transportistas,
            "atrasados": sorted(atrasados, key=lambda f: -f["atraso"])}


# -----------------------------
# Faltantes y venta perdida
# -----------------------------
CAMBIO = re.compile(r"^Cambio de bodega: (.+) → (.+)$")


def faltantes(cursor, desde, hasta):
    cursor.execute("""SELECT l.id, l.envio_id, l.codigo, l.producto, l.marca, l.cantidad, l.cant_disponible, l.subtotal, l.existencia,
                             b.nombre AS bodega
                      FROM ent_envio_lineas l JOIN ent_envios e ON e.id = l.envio_id JOIN ent_bodegas b ON b.id = l.bodega_id
                      WHERE e.creado_en >= %s AND e.creado_en < %s""", (desde, hasta + timedelta(days=1)))
    lineas = {l["id"]: l for l in cursor.fetchall()}
    vacio = {"bodegas": [], "productos": [], "lineas": 0, "unidades": 0, "monto": 0}
    if not lineas:
        return vacio
    ids = list(lineas)
    cursor.execute(f"""SELECT linea_id, de_estado, a_estado, comentario FROM ent_envio_eventos
                       WHERE campo = 'existencia' AND linea_id IN ({lista_sql(ids)}) ORDER BY creado_en, id""", ids)
    eventos = defaultdict(list)
    for ev in cursor.fetchall():
        eventos[ev["linea_id"]].append(ev)
    bodegas = defaultdict(lambda: {"faltantes": 0, "unidades": 0.0, "monto": 0.0, "quitadas": 0, "parciales": 0, "cambiadas": 0})
    productos = {}
    total_lineas = 0
    for lid, l in lineas.items():
        evs = eventos.get(lid, [])
        cambios = [CAMBIO.match(ev["comentario"] or "") for ev in evs]
        primera = next((m for m in cambios if m), None)
        bodega_actual = primera.group(1) if primera else l["bodega"]
        donde_falto = []
        for ev, m in zip(evs, cambios):
            if m:
                bodega_actual = m.group(2)
            elif ev["a_estado"] == "FALTANTE" and ev["de_estado"] != "FALTANTE":
                donde_falto.append(bodega_actual)
        if not donde_falto:
            continue
        total_lineas += 1
        for b in donde_falto:
            bodegas[b]["faltantes"] += 1
        cantidad = float(l["cantidad"])
        perdida = 0.0
        if l["existencia"] == "QUITADA":
            perdida = cantidad
            bodegas[l["bodega"]]["quitadas"] += 1
        elif l["existencia"] == "CONFIRMADA" and l["cant_disponible"] is not None and float(l["cant_disponible"]) < cantidad:
            perdida = cantidad - float(l["cant_disponible"])
            bodegas[l["bodega"]]["parciales"] += 1
        if any(cambios):
            bodegas[donde_falto[0]]["cambiadas"] += 1
        monto = float(l["subtotal"] or 0) * perdida / cantidad if cantidad else 0
        bodegas[l["bodega"]]["unidades"] += perdida
        bodegas[l["bodega"]]["monto"] += monto
        p = productos.setdefault((l["codigo"], l["marca"]), {"codigo": l["codigo"], "producto": l["producto"], "marca": l["marca"],
                                                              "veces": 0, "unidades": 0.0, "monto": 0.0, "bodegas": set()})
        p["veces"] += 1
        p["unidades"] += perdida
        p["monto"] += monto
        p["bodegas"].update(donde_falto)
    lista_b = sorted(({"bodega": b, **v} for b, v in bodegas.items()), key=lambda x: (-x["faltantes"], x["bodega"]))
    lista_p = sorted(productos.values(), key=lambda x: (-x["veces"], -x["monto"], x["codigo"]))
    for p in lista_p:
        p["bodegas"] = ", ".join(sorted(p["bodegas"]))
    return {"bodegas": lista_b, "productos": lista_p, "lineas": total_lineas,
            "unidades": sum(b["unidades"] for b in lista_b), "monto": sum(b["monto"] for b in lista_b)}


# -----------------------------
# Fletes por transportista
# -----------------------------
def fletes(cursor, desde, hasta):
    cursor.execute("""SELECT p.id, p.monto_recibo, t.nombre AS transportista, t.tipo FROM ent_ruta_paradas p
                      JOIN ent_transportistas t ON t.id = p.transportista_id
                      WHERE p.entregado_en >= %s AND p.entregado_en < %s""", (desde, hasta + timedelta(days=1)))
    paradas = cursor.fetchall()
    filas = {}
    for p in paradas:
        f = filas.setdefault(p["transportista"], {"transportista": p["transportista"], "tipo": TIPOS_TRANSPORTISTA.get(p["tipo"], p["tipo"]),
                                                  "paradas": 0, "envios": 0, "recibos": 0.0, "cliente": 0.0, "empresa": 0.0,
                                                  "sin_detallar": 0.0})
        f["paradas"] += 1
        f["recibos"] += float(p["monto_recibo"] or 0)
    if paradas:
        pids = [p["id"] for p in paradas]
        cursor.execute(f"""SELECT e.parada_id, e.flete_paga, e.flete_monto, t.nombre AS transportista FROM ent_envios e
                           JOIN ent_ruta_paradas p ON p.id = e.parada_id JOIN ent_transportistas t ON t.id = p.transportista_id
                           WHERE e.estado = 'ENTREGADA' AND e.parada_id IN ({lista_sql(pids)})""", pids)
        for e in cursor.fetchall():
            f = filas[e["transportista"]]
            f["envios"] += 1
            if e["flete_monto"] is not None:
                f["empresa" if e["flete_paga"] == "empresa" else "cliente"] += float(e["flete_monto"])
    for f in filas.values():
        f["sin_detallar"] = round(f["recibos"] - f["cliente"] - f["empresa"], 2)
    lista = sorted(filas.values(), key=lambda f: -f["recibos"])
    totales = {k: sum(f[k] for f in lista) for k in ("paradas", "envios", "recibos", "cliente", "empresa", "sin_detallar")}
    return {"filas": lista, "totales": totales}


# -----------------------------
# Conciliación de facturas
# -----------------------------
def numero_limpio(numero):
    digitos = re.sub(r"\D", "", numero or "")
    return digitos.lstrip("0") or digitos


def conciliacion(cursor, desde, hasta):
    cursor.execute("""SELECT f.factura_num, f.registrada_en, e.id AS envio_id, e.pedido_num, e.cliente, e.estado, b.nombre AS bodega,
                             b.nombre_en_orden
                      FROM ent_envio_facturas f JOIN ent_envios e ON e.id = f.envio_id JOIN ent_bodegas b ON b.id = f.bodega_id
                      WHERE f.registrada_en >= %s AND f.registrada_en < %s ORDER BY f.registrada_en""", (desde, hasta + timedelta(days=1)))
    facturas = cursor.fetchall()
    cursor.execute("SELECT MAX(Fecha) AS f FROM ventas_historicas")
    ultima = cursor.fetchone()["f"]
    resumen = {"total": len(facturas), "encontradas": 0, "observadas": 0, "no_encontradas": 0, "sin_cargar": 0, "ultima_carga": ultima}
    if not facturas:
        return {"filas": [], "resumen": resumen}
    numeros = sorted({f["factura_num"].strip() for f in facturas} | {numero_limpio(f["factura_num"]) for f in facturas} - {""})
    cursor.execute(f"""SELECT factura, Fecha, cliente, Bodega, Total, estado_factura FROM ventas_historicas
                       WHERE factura IN ({lista_sql(numeros)})""", numeros)
    en_sistema = defaultdict(list)
    for v in cursor.fetchall():
        en_sistema[v["factura"].strip()].append(v)
    veces = defaultdict(set)
    for f in facturas:
        veces[numero_limpio(f["factura_num"]) or f["factura_num"]].add(f["envio_id"])
    for f in facturas:
        candidatos = en_sistema.get(f["factura_num"].strip()) or en_sistema.get(numero_limpio(f["factura_num"])) or []
        f["sistema"] = candidatos[0] if len(candidatos) == 1 else None
        obs = []
        if len(veces[numero_limpio(f["factura_num"]) or f["factura_num"]]) > 1:
            obs.append("el mismo número está registrado en otro envío")
        if len(candidatos) > 1:
            obs.append(f"el número aparece {len(candidatos)} veces en la facturación cargada")
        v = f["sistema"]
        if v:
            if v["estado_factura"] == "Anulado":
                obs.append("anulada en el sistema de facturación")
            if sin_tildes(v["Bodega"]) not in (sin_tildes(f["bodega"]), sin_tildes(f["nombre_en_orden"])):
                obs.append(f"bodega en la facturación: {v['Bodega']}")
            if sin_tildes(v["cliente"]).lstrip("*") != sin_tildes(f["cliente"]).lstrip("*"):
                obs.append(f"cliente en la facturación: {v['cliente']}")
            f["resultado"] = "observada" if obs else "encontrada"
        elif candidatos:
            f["resultado"] = "observada"
        elif ultima is None or f["registrada_en"].date() > ultima:
            f["resultado"] = "sin_cargar"
            obs.append(f"falta cargar la facturación del {f['registrada_en'].strftime('%d/%m/%Y')}")
        else:
            f["resultado"] = "no_encontrada"
            obs.append("no está en la facturación cargada")
        f["observaciones"] = obs
        resumen[{"encontrada": "encontradas", "observada": "observadas", "no_encontrada": "no_encontradas", "sin_cargar": "sin_cargar"}[f["resultado"]]] += 1
    orden = {"no_encontrada": 0, "observada": 1, "sin_cargar": 2, "encontrada": 3}
    return {"filas": sorted(facturas, key=lambda f: (orden[f["resultado"]], f["registrada_en"])), "resumen": resumen}


# -----------------------------
# Pantalla y Excel
# -----------------------------
def leer_periodo(args):
    hoy = ahora_nicaragua().date()

    def fecha(texto, defecto):
        try:
            return datetime.strptime(texto or "", "%Y-%m-%d").date()
        except ValueError:
            return defecto
    desde = fecha(args.get("desde"), hoy.replace(day=1))
    hasta = fecha(args.get("hasta"), hoy)
    if hasta < desde:
        desde, hasta = hasta, desde
    return desde, hasta


def calcular(cursor, desde, hasta):
    d = datetime.combine(desde, datetime.min.time())
    h = datetime.combine(hasta, datetime.min.time())
    return {"tablero": tablero(cursor), "tiempos": tiempos(cursor, d, h), "a_tiempo": a_tiempo(cursor, d, h),
            "faltantes": faltantes(cursor, d, h), "fletes": fletes(cursor, d, h), "conciliacion": conciliacion(cursor, d, h)}


def leer_todo(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        return calcular(cursor, desde, hasta)
    finally:
        cursor.close()
        conn.close()


@entregas_reportes_bp.route("/entregas/reportes")
@requiere_rol("entregas_reportes", "admin")
def reportes():
    desde, hasta = leer_periodo(request.args)
    return render_template("entregas_reportes.html", r=leer_todo(desde, hasta), desde=desde, hasta=hasta, ahora=ahora_nicaragua(),
                           activo="entregas_reportes")


@entregas_reportes_bp.route("/entregas/reportes/excel")
@requiere_rol("entregas_reportes", "admin")
def excel():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    desde, hasta = leer_periodo(request.args)
    r = leer_todo(desde, hasta)
    wb = Workbook()
    verde, blanco = PatternFill("solid", fgColor="027266"), Font(color="FFFFFF", bold=True)
    periodo = f"Período: {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}"

    def hoja(titulo, encabezados, filas, anchos=None, nota=None):
        ws = wb.active if wb.active.title == "Sheet" else wb.create_sheet()
        ws.title = titulo
        ws["A1"] = titulo
        ws["A1"].font = Font(bold=True, size=13, color="027266")
        ws["A2"] = periodo + (f" · {nota}" if nota else "")
        ws["A2"].font = Font(italic=True, color="666666")
        for j, e in enumerate(encabezados, 1):
            c = ws.cell(row=4, column=j, value=e)
            c.fill, c.font, c.alignment = verde, blanco, Alignment(wrap_text=True, vertical="center")
        for i, fila in enumerate(filas, 5):
            for j, v in enumerate(fila, 1):
                c = ws.cell(row=i, column=j, value=v)
                if isinstance(v, float):
                    c.number_format = "#,##0.00"
                elif isinstance(v, datetime):
                    c.number_format = "dd/mm/yyyy hh:mm AM/PM"
                elif isinstance(v, date):
                    c.number_format = "dd/mm/yyyy"
        for j, ancho in enumerate(anchos or [], 1):
            ws.column_dimensions[get_column_letter(j)].width = ancho
        ws.freeze_panes = "A5"
        if filas:
            ws.auto_filter.ref = f"A4:{get_column_letter(len(encabezados))}{4 + len(filas)}"
        return ws

    h = lambda m: round(m / 60, 2) if m is not None else None  # noqa: E731  (minutos -> horas)
    filas_t = []
    for e in r["tiempos"]["etapas"]:
        filas_t.append([e["nombre"], e["n"], h(e.get("promedio")), h(e.get("mediana")), h(e.get("maximo"))])
        if e["clave"] == "orden_lista":
            for k, texto in (("sin", "   sin traslado"), ("con", "   con traslado")):
                filas_t.append([texto, e[k]["n"], h(e[k].get("promedio")), h(e[k].get("mediana")), h(e[k].get("maximo"))])
    t = r["tiempos"]["traslado"]
    filas_t.append(["Traslado entre bodegas (aplicado → recibido, por línea)", t["n"], h(t.get("promedio")), h(t.get("mediana")), h(t.get("maximo"))])
    hoja("Tiempo por etapa", ["Etapa", "Envíos", "Promedio (h)", "Mediana (h)", "Máximo (h)"], filas_t, [52, 10, 14, 14, 14],
         f"envíos creados en el período; horas laborales ({r['tiempos']['horario']})")
    a = r["a_tiempo"]
    hoja("Entregas a tiempo", ["Transportista", "Entregados", "A tiempo", "% a tiempo"],
         [[x["transportista"], x["total"], x["a_tiempo"], round(x["pct"], 1)] for x in a["transportistas"]]
         + [["TOTAL", a["total"], a["a_tiempo"], round(a["pct"], 1) if a["pct"] is not None else None]], [34, 12, 12, 12],
         "paradas entregadas en el período")
    hoja("Atrasados", ["Pedido", "Cliente", "Transportista", "Destino", "Hora requerida", "Entregado", "Atraso (min)"],
         [[x["pedido_num"], x["cliente"], x["transportista"], x["destino"], x["hora_requerida"], x["entregado_en"], round(x["atraso"])]
          for x in a["atrasados"]], [10, 34, 24, 16, 20, 20, 12])
    f = r["faltantes"]
    hoja("Faltantes por bodega", ["Bodega", "Líneas faltantes", "Quitadas", "Parciales", "Cambiadas de bodega", "Unidades perdidas", "Venta perdida C$"],
         [[x["bodega"], x["faltantes"], x["quitadas"], x["parciales"], x["cambiadas"], x["unidades"], round(x["monto"], 2)] for x in f["bodegas"]],
         [30, 14, 10, 10, 14, 14, 16], "envíos creados en el período")
    hoja("Faltantes por producto", ["Nº parte", "Producto", "Marca", "Veces faltante", "Bodegas", "Unidades perdidas", "Venta perdida C$"],
         [[x["codigo"], x["producto"], x["marca"], x["veces"], x["bodegas"], x["unidades"], round(x["monto"], 2)] for x in f["productos"]],
         [18, 40, 14, 12, 30, 14, 16])
    fl = r["fletes"]
    hoja("Fletes", ["Transportista", "Tipo", "Paradas", "Envíos", "Recibos C$", "Paga el cliente C$", "Paga la empresa C$", "Sin detallar C$"],
         [[x["transportista"], x["tipo"], x["paradas"], x["envios"], round(x["recibos"], 2), round(x["cliente"], 2), round(x["empresa"], 2),
           round(x["sin_detallar"], 2)] for x in fl["filas"]], [30, 18, 10, 10, 14, 16, 16, 14], "paradas entregadas en el período")
    c = r["conciliacion"]
    hoja("Conciliación", ["Factura", "Bodega", "Pedido", "Cliente", "Registrada", "Resultado", "Fecha facturación", "Total facturado C$",
                          "Observaciones"],
         [[x["factura_num"], x["bodega"], x["pedido_num"], x["cliente"], x["registrada_en"],
           {"encontrada": "Encontrada", "observada": "Revisar", "no_encontrada": "No encontrada", "sin_cargar": "Falta cargar"}[x["resultado"]],
           x["sistema"]["Fecha"] if x["sistema"] else None, float(x["sistema"]["Total"]) if x["sistema"] else None, "; ".join(x["observaciones"])]
          for x in c["filas"]], [12, 22, 10, 32, 20, 14, 14, 16, 50], "facturas registradas en el período")
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name=f"entregas_reportes_{desde.isoformat()}_{hasta.isoformat()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
