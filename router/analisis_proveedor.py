"""Análisis de proveedor (una sola pantalla + Excel). Responde, en orden:
1) cuánto hemos traído  2) cómo lo hemos vendido  3) desde cuándo y en qué marca se puso lento
4) si se mueve mejor en los últimos meses  5) qué no volver a comprar y qué reponer.

- Proveedor: «Proveedor Ultima Compra» del catálogo (desde 2023 el sistema no trae el proveedor en las transacciones).
- Todo en córdobas, costo contra venta:
  · Ingresado: entradas por ORDEN COMPRA desde la fecha, a su costo REAL en C$ (total_mn de la transacción).
  · Ventas: unidades y venta neta en C$ (Sub-Total 2 de ventas_por_producto: sin IVA, con descuento, sin anuladas) y su costo.
  · Inventario y NHE: cantidad × ÚLTIMO costo del catálogo (C$); si el producto tiene costo 0, se estima con FOB × TC_FOB.
  Ingresado (costo real) e inventario (último costo) tienen bases distintas: no se restan entre sí en ningún indicador.
- FOB en US$ («Ultimo Precio FOB» del catálogo) solo como referencia en los totales.
- «12 meses» = los últimos 12 meses completos (el mes en curso se muestra aparte y no entra en promedios ni tendencias).
- Las métricas (clasificación, rotación, GMROI, tendencia, lento desde, recomendaciones) están en router/metricas_proveedor.py:
  la pantalla y el Excel usan los mismos resultados de calcular()."""
import io
from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.ventas_bodegas import en_cache
from router import metricas_proveedor as M

analisis_proveedor_bp = Blueprint("analisis_proveedor", __name__)

HINT = "/*+ INDEX(transacciones_inventario idx_transacciones_producto_fecha) */"
MESES = M.MESES
MAX_FILAS_PANTALLA = 1500
MAX_FILAS_RECOMENDACION = 60         # en pantalla; en el Excel van todas
MIN_INVENTARIO_MARCA = 50000         # en la tabla de rentabilidad, marcas con menos inventario se agrupan en «Otras»
TOP_MARCAS_GRAFICO = 6
TC_FOB = 36.6243                     # tipo de cambio para estimar en C$ el costo de un producto con último costo 0 (FOB × TC)
VERDE, NARANJA = "027266", "FE6412"
ESTADOS = {"": "Todos", "rota": "Se vende", "lento": "Lento", "sin_venta": "Sin venta en 12 meses", "agotado": "Agotado",
           "nuevo": "Nuevo – sin evaluar", "no_recomprar": "No volver a comprar", "reponer": "Reponer",
           "lento_reciente": "Compras recientes ya lentas", "nhe": "Con NHE", "ingreso": "Ingresaron desde la fecha",
           "negativo": "Inventario negativo", "utilidad_neg": "Utilidad negativa", "costo_estimado": "Costo estimado con FOB"}
FILTROS_RAPIDOS = ["rota", "lento", "sin_venta", "agotado", "nuevo", "no_recomprar", "reponer"]
ORDENES = {"inventario": "Mayor inventario a costo", "ingreso": "Mayor ingreso a costo", "vendido": "Mayor venta neta",
           "utilidad": "Mayor utilidad", "sin_venta": "Más tiempo sin venta", "marca": "Marca y descripción"}
restar_meses = M.restar_meses


def fecha_param(valor, defecto):
    try:
        return datetime.strptime(valor or "", "%Y-%m-%d").date()
    except ValueError:
        return defecto


def a_fecha(v):
    return v.date() if isinstance(v, datetime) else v


def lista_proveedores():
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT TRIM(`Proveedor Ultima Compra`), COUNT(*) FROM catalogo_productos "
                           "WHERE `Proveedor Ultima Compra` <> '' GROUP BY 1 ORDER BY 1")
            return [(p, n) for p, n in cursor.fetchall() if p]
        finally:
            cursor.close()
            conn.close()
    return en_cache("analisis_proveedor_lista", leer)


def marcas_del_proveedor(proveedor):
    """Marcas del proveedor con cuántos productos tiene y la existencia total."""
    if not proveedor:
        return []
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COALESCE(NULLIF(TRIM(Marca), ''), 'SIN MARCA'), COUNT(*), SUM(COALESCE(`Existencia Actual`, 0)) "
                       "FROM catalogo_productos WHERE TRIM(`Proveedor Ultima Compra`) = %s GROUP BY 1 ORDER BY 2 DESC", (proveedor,))
        return [{"marca": m, "productos": n, "existencia": float(e or 0)} for m, n, e in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()


def leer_filtros(args):
    hoy = date.today()
    proveedor = (args.get("proveedor") or "").strip()
    nombres = {p for p, _ in lista_proveedores()}
    if proveedor and proveedor not in nombres:   # escrito a mano: se busca sin importar mayúsculas
        proveedor = next((p for p in nombres if p.upper() == proveedor.upper()), proveedor)
    return {"proveedor": proveedor, "marcas": sorted({m.strip() for m in args.getlist("marca") if m.strip()}),
            "desde": min(fecha_param(args.get("desde"), restar_meses(hoy, 6)), hoy), "hoy": hoy,
            "q": (args.get("q") or "").strip(), "estado": args.get("estado") if args.get("estado") in ESTADOS else "",
            "orden": args.get("orden") if args.get("orden") in ORDENES else "inventario",
            "existencia": "con" if args.get("existencia") == "con" else ""}


def ventas_netas(inicio, desde):
    """Venta por producto y mes desde `inicio`: ({código: [(mes, desde_la_fecha, unidades, venta_neta, costo)]}, última fecha con venta).
    ventas_por_producto no tiene índices: se recorre una vez y se guarda 10 minutos (sirve para cualquier proveedor)."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT Codigo, LEFT(Fecha, 7), Fecha >= %s, SUM(`Cantidad Vendida`), SUM(`Sub-Total 2`),
                       SUM(`Cantidad Vendida` * `Costo Unit.`), MAX(Fecha)
                FROM ventas_por_producto WHERE Fecha >= %s AND Estatus <> 'Anulado'
                GROUP BY Codigo, LEFT(Fecha, 7), Fecha >= %s
            """, (desde.isoformat(), inicio.isoformat(), desde.isoformat()))
            por_codigo, corte = defaultdict(list), ""
            for cod, mes, desde_fecha, cant, monto, costo, ultima in cursor.fetchall():
                por_codigo[cod].append((mes, bool(desde_fecha), float(cant or 0), float(monto or 0), float(costo or 0)))
                corte = max(corte, str(ultima or "")[:10])
            return dict(por_codigo), fecha_param(corte, None)
        finally:
            cursor.close()
            conn.close()
    return en_cache(f"analisis_proveedor_ventas3_{inicio}_{desde}", leer)


def en_lista(cursor, sql, codigos, extra=()):
    filas, codigos = [], list(codigos)
    for i in range(0, len(codigos), 800):
        bloque = codigos[i:i + 800]
        cursor.execute(sql.replace("{IN}", ", ".join(["%s"] * len(bloque))), tuple(bloque) + tuple(extra))
        filas += cursor.fetchall()
    return filas


CAMPOS_SUMA = ("ingreso", "vendido", "disponible", "nhe", "ingreso_costo", "venta_neta", "costo_vendido", "inventario_costo",
               "nhe_costo", "ingreso_fob", "inventario_fob", "nhe_fob", "venta_12", "monto_12", "costo_12")


def nuevo_total():
    t = defaultdict(float)
    t["monto_meses"] = None     # se crean al acumular (36 meses completos + el mes en curso)
    return t


def acumular(destino, x):
    for campo in CAMPOS_SUMA:
        destino[campo] += x[campo] or 0
    if destino["monto_meses"] is None:
        destino["monto_meses"], destino["unid_meses"] = [0.0] * len(x["monto_serie"]), [0.0] * len(x["unid_serie"])
    for i, (mto, u) in enumerate(zip(x["monto_serie"], x["unid_serie"])):
        destino["monto_meses"][i] += mto
        destino["unid_meses"][i] += u
    destino["productos"] += 1
    destino["inv_" + x["estado"]] += x["inventario_costo"]
    if x["estado"] == "sin_venta" and x["disponible"] > 0:
        destino["sin_venta"] += 1
        destino["sin_venta_costo"] += x["inventario_costo"]
    destino["costo_estimado"] += x["costo_estimado"]
    destino["sin_costo"] += x["costo"] <= 0


def resumir(m, dias, n36, sin_venta_filas):
    """Indicadores de una marca o del total (m: acumulado de acumular())."""
    m["rent"] = M.rentabilidad(m["venta_neta"], m["costo_vendido"], m["inventario_costo"], dias, m["monto_12"], m["costo_12"])
    serie_m, serie_u = (m["monto_meses"] or [0.0] * (n36 + 1))[:n36], (m["unid_meses"] or [0.0] * (n36 + 1))[:n36]
    m["monto_prev"], m["unid_prev"] = sum(serie_m[n36 - 24:n36 - 12]), sum(serie_u[n36 - 24:n36 - 12])
    m["cambio"] = M.cambio_pct(m["monto_12"], m["monto_prev"])
    m["cambio_unid"] = M.cambio_pct(m["venta_12"], m["unid_prev"])
    m["nivel"] = M.nivel_tendencia(m["cambio"], m["monto_12"])
    m["nivel_unid"] = M.nivel_tendencia(m["cambio_unid"], m["venta_12"])
    i = M.lento_desde(serie_m, primero=n36 - 24)
    m["lento_desde_idx"] = None if i is None else i - (n36 - 24)      # índice dentro de los 24 meses del mapa de calor
    m["heat"] = serie_m[n36 - 24:]
    m["mediana_ultima_venta"] = M.mediana_fecha(x["ultima_venta"] for x in sin_venta_filas)
    lento = m["inv_lento"] + m["inv_sin_venta"]
    m["pct_lento"] = lento / m["inventario_costo"] * 100 if m["inventario_costo"] > 0 else None
    m["dejar"] = M.dejar_de_comprar(m["cambio"], m["pct_lento"], m["monto_12"])
    m["pct_vendido"] = m["vendido"] / m["ingreso"] * 100 if m["ingreso"] > 0 else None
    return m


def calcular(f):
    vacio = {"filas": [], "marcas": [], "total": {}, "meses": [], "temporada": (None, None, None), "sin_actividad": 0}
    if not f["proveedor"]:
        return vacio
    hoy = f["hoy"]
    inicio_hist = min(f["desde"], restar_meses(hoy, 37))
    todas, corte = ventas_netas(inicio_hist, f["desde"])
    # Mes en curso = el de la última venta cargada (si no termina el mes): va aparte, sin entrar en promedios ni tendencias
    corte = min(corte or hoy, hoy)
    if corte.day == monthrange(corte.year, corte.month)[1]:
        ref, mes_actual = corte + timedelta(days=1), None
    else:
        ref, mes_actual = corte, f"{corte.year}-{corte.month:02d}"
    meses36 = M.lista_meses(ref, 36)            # meses completos, del más antiguo al más reciente
    n36 = len(meses36)
    meses24, meses12 = meses36[-24:], set(meses36[-12:])
    serie_meses = meses36 + ([mes_actual] if mes_actual else [])
    pos_mes = {m: i for i, m in enumerate(serie_meses)}
    fin_mes_anterior = restar_meses(ref, 0) - timedelta(days=1)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        where, params = ["TRIM(p.`Proveedor Ultima Compra`) = %s"], [f["proveedor"]]
        if f["marcas"]:
            where.append(f"COALESCE(NULLIF(TRIM(p.Marca), ''), 'SIN MARCA') IN ({', '.join(['%s'] * len(f['marcas']))})")
            params += f["marcas"]
        cursor.execute(f"""
            SELECT /*+ NO_INDEX(c) */ p.`Codigo Producto`, p.Nombre, p.`Numero Parte`, COALESCE(NULLIF(TRIM(p.Marca), ''), 'SIN MARCA'),
                   p.Categoria, p.`Ultimo Precio FOB`, p.`Fecha Ultima Compra`, c.segmento, c.familia, c.subfamilia, p.`Ultimo Costo`
            FROM catalogo_productos p LEFT JOIN clase_prod c ON c.codigo = p.`Codigo Producto`
            WHERE {' AND '.join(where)}
        """, params)
        productos = {}
        for r in cursor.fetchall():
            fob, costo = (float(r[5]) if r[5] else None), float(r[10] or 0)
            estimado = costo <= 0 and bool(fob)
            productos[r[0]] = {"codigo": r[0], "nombre": r[1] or "", "parte": r[2] or "", "marca": r[3], "categoria": r[4] or "",
                               "fob": fob, "ultima_compra": r[6],
                               "segmento": r[7] or "Sin clasificar", "familia": r[8] or "Sin clasificar", "subfamilia": r[9] or "Sin clasificar",
                               # costo unitario en C$: último costo del catálogo; si está en 0, FOB × tipo de cambio (estimado)
                               "costo": fob * TC_FOB if estimado else costo, "costo_estimado": estimado}
        if not productos:
            return vacio
        codigos = list(productos)
        # Ingresos por ORDEN COMPRA con su costo real en C$ (total_mn) y último ingreso
        desde_sql = f"'{f['desde'].isoformat()}'"     # fecha (objeto date): segura para ir en el texto de la consulta
        movs = en_lista(cursor, f"""
            SELECT {HINT} codigo_producto, fecha >= {desde_sql},
                   SUM(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' THEN cantidad ELSE 0 END),
                   SUM(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' THEN total_mn ELSE 0 END),
                   MAX(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' THEN fecha END)
            FROM transacciones_inventario
            WHERE codigo_producto IN ({{IN}}) AND fecha >= %s AND tipo_documento_origen = 'ORDEN COMPRA'
            GROUP BY codigo_producto, fecha >= {desde_sql}
        """, codigos, (inicio_hist,))
        # Primer ingreso: la primera entrada por orden de compra; si no tiene, su primer movimiento
        primer_ingreso = {c: a_fecha(oc or primero) for c, oc, primero in en_lista(cursor,
            f"SELECT {HINT} codigo_producto, MIN(CASE WHEN tipo_transaccion = 'ENTRADA' AND tipo_documento_origen = 'ORDEN COMPRA' "
            "THEN fecha END), MIN(fecha) FROM transacciones_inventario WHERE codigo_producto IN ({IN}) GROUP BY codigo_producto", codigos)}
        ultima_venta = dict(en_lista(cursor, f"SELECT {HINT} codigo_producto, MAX(fecha) FROM transacciones_inventario "
                                             "WHERE codigo_producto IN ({IN}) AND tipo_transaccion = 'SALIDA' "
                                             "AND tipo_documento_origen LIKE 'FACTURA%%' GROUP BY codigo_producto", codigos))
        existencias = {r[0]: (float(r[1] or 0), float(r[2] or 0)) for r in en_lista(cursor,
            "SELECT Codigo_producto, SUM(Cant_Final), SUM(CASE WHEN UPPER(Bodega) LIKE '%NHE%' THEN Cant_Final ELSE 0 END) "
            "FROM existencias_bodega WHERE Codigo_producto IN ({IN}) GROUP BY Codigo_producto", codigos)}
    finally:
        cursor.close()
        conn.close()

    ingreso, ingreso_costo, ultimo_ingreso = defaultdict(float), defaultdict(float), {}
    for cod, desde_fecha, cant, costo, ult_ing in movs:
        if desde_fecha:
            ingreso[cod] += float(cant or 0)
            ingreso_costo[cod] += float(costo or 0)
        ult_ing = a_fecha(ult_ing)
        if ult_ing and (cod not in ultimo_ingreso or ult_ing > ultimo_ingreso[cod]):
            ultimo_ingreso[cod] = ult_ing

    filas, sin_actividad = [], 0
    for cod, p in productos.items():
        total, nhe = existencias.get(cod, (0.0, 0.0))
        disponible = total - nhe
        ing = ingreso.get(cod, 0.0)
        ven = vn = cv = 0.0
        unid_serie, monto_serie = [0.0] * len(serie_meses), [0.0] * len(serie_meses)
        venta_12 = monto_12 = costo_12 = 0.0
        for mes, desde_fecha, cant, monto, costo in todas.get(cod, ()):
            if desde_fecha:
                ven += cant
                vn += monto
                cv += costo
            i = pos_mes.get(mes)
            if i is not None:
                unid_serie[i] += cant
                monto_serie[i] += monto
            if mes in meses12:
                venta_12 += cant
                monto_12 += monto
                costo_12 += costo
        ventas_recientes = sum(unid_serie[n36 - 24:])     # 24 meses completos + el mes en curso
        if not (ing or ven or abs(total) > 1e-9 or ventas_recientes):
            sin_actividad += 1
            continue
        if f["existencia"] == "con" and disponible <= 0:   # «Solo con existencia»: disponible (sin NHE) mayor que cero
            continue
        ult = a_fecha(ultima_venta.get(cod))
        dias_sin = (hoy - ult).days if ult else None
        primero = primer_ingreso.get(cod)
        meses_inv = M.meses_entre(primero, hoy)
        mensual = M.venta_mensual(venta_12, M.meses_entre(primero, fin_mes_anterior))
        cob = M.cobertura(disponible, mensual)
        estado, comentario = M.clasificar(meses_inv, dias_sin, ult, disponible, mensual)
        rec = M.recomendacion(estado, disponible, cob, venta_12)
        if nhe > 0:
            comentario += f" · {nhe:,.0f} en NHE"
        por_mes_anio = defaultdict(float)
        for m, v in zip(meses24, unid_serie[n36 - 24:n36]):
            por_mes_anio[int(m[5:])] += v
        temp, temp_pct, _ = M.temporada(por_mes_anio)
        fob, cu = p["fob"], p["costo"]
        filas.append({**p, "ingreso": ing, "ingreso_costo": ingreso_costo.get(cod, 0.0),
                      "vendido": ven, "venta_neta": vn, "costo_vendido": cv, "utilidad": vn - cv,
                      "margen": (vn - cv) / vn * 100 if vn > 0 else None,
                      "venta_12": venta_12, "monto_12": monto_12, "costo_12": costo_12,
                      "existencia": total, "nhe": nhe, "disponible": disponible,
                      "inventario_costo": max(disponible, 0) * cu, "nhe_costo": nhe * cu if nhe > 0 else 0.0,
                      # FOB en US$ solo como referencia en los totales generales
                      "ingreso_fob": ing * fob if fob else 0.0, "inventario_fob": max(disponible, 0) * fob if fob else 0.0,
                      "nhe_fob": nhe * fob if fob and nhe > 0 else 0.0,
                      "ultima_venta": ult, "dias_sin_venta": dias_sin, "ultimo_ingreso": ultimo_ingreso.get(cod),
                      "primer_ingreso": primero, "meses_inventario": meses_inv,
                      "temporada": temp, "temporada_pct": temp_pct, "mensual": mensual, "cobertura": cob,
                      "estado": estado, "comentario": comentario, "recomendacion": rec,
                      "venta_meses": unid_serie[n36 - 24:n36], "unid_serie": unid_serie, "monto_serie": monto_serie})

    # ---------- Resumen por marca, total y rotación (antes de los filtros de la tabla, para que cuadre con el proveedor) ----------
    total, por_marca = nuevo_total(), defaultdict(nuevo_total)
    sin_venta_inv = defaultdict(list)       # productos sin venta en 12 meses con inventario, por marca (para la mediana)
    por_estado = {k: {"estado": k, "nombre": nombre, "corto": corto, "color": color, "productos": 0, "disponible": 0.0,
                      "inventario_costo": 0.0, "venta_neta": 0.0, "monto_12": 0.0, "marcas": defaultdict(lambda: defaultdict(float))}
                  for k, nombre, corto, color in M.ESTADOS}
    sin_venta_sin_existencia = 0
    for x in filas:
        acumular(total, x)
        acumular(por_marca[x["marca"]], x)
        if x["estado"] == "sin_venta":
            if x["disponible"] <= 0:          # sin venta y sin existencia: no cuenta en la rotación
                sin_venta_sin_existencia += 1
                continue
            sin_venta_inv[x["marca"]].append(x)
        e = por_estado[x["estado"]]
        e["productos"] += 1
        e["disponible"] += max(x["disponible"], 0)
        e["inventario_costo"] += x["inventario_costo"]
        e["venta_neta"] += x["venta_neta"]
        e["monto_12"] += x["monto_12"]
        em = e["marcas"][x["marca"]]
        for campo in ("inventario_costo", "venta_neta", "monto_12"):
            em[campo] += x[campo]
        em["productos"] += 1
        em["disponible"] += max(x["disponible"], 0)
    rotacion = list(por_estado.values())
    tot_prod = sum(e["productos"] for e in rotacion)
    tot_inv = sum(e["inventario_costo"] for e in rotacion)
    tot_12 = sum(e["monto_12"] for e in rotacion)
    for e in rotacion:
        e["pct_productos"] = e["productos"] / tot_prod * 100 if tot_prod else 0
        e["pct_inventario"] = e["inventario_costo"] / tot_inv * 100 if tot_inv else 0
        e["pct_venta_12"] = e["monto_12"] / tot_12 * 100 if tot_12 else 0
        e["marcas"] = sorted(({"marca": m, **v, "pct": v["inventario_costo"] / e["inventario_costo"] * 100 if e["inventario_costo"] else 0}
                              for m, v in e["marcas"].items()), key=lambda y: (-y["inventario_costo"], -y["productos"]))
    rot = {e["estado"]: e for e in rotacion}

    dias = (hoy - f["desde"]).days + 1
    for m, v in por_marca.items():          # se quedan como defaultdict: los campos que no se acumularon valen 0
        v["marca"] = m
        resumir(v, dias, n36, sin_venta_inv[m])
    marcas = sorted(por_marca.values(), key=lambda x: -x["inventario_costo"])
    for m in marcas:
        m["lento_desde"] = M.etiqueta_mes(meses24[m["lento_desde_idx"]]) if m["lento_desde_idx"] is not None else None
    if filas:
        resumir(total, dias, n36, [x for v in sin_venta_inv.values() for x in v])
    rent_total = total.get("rent")

    # Rentabilidad por marca: las de menos de C$ 50,000 de inventario se agrupan en «Otras»
    grandes = [m for m in marcas if m["inventario_costo"] >= MIN_INVENTARIO_MARCA]
    chicas = [m for m in marcas if m["inventario_costo"] < MIN_INVENTARIO_MARCA]
    marcas_rent = list(grandes)
    if len(chicas) > 1:
        otras = defaultdict(float)
        for m in chicas:
            for campo in CAMPOS_SUMA + ("productos",):
                otras[campo] += m[campo]
        otras = {"marca": f"Otras ({len(chicas)} marcas)", "otras": True, "detalle": ", ".join(m["marca"] for m in chicas), **otras}
        otras["rent"] = M.rentabilidad(otras["venta_neta"], otras["costo_vendido"], otras["inventario_costo"], dias,
                                       otras["monto_12"], otras["costo_12"])
        marcas_rent.append(otras)
    else:
        marcas_rent += chicas

    # Gráfico «ingresado, vendido y tenemos»: top 6 marcas + Otras
    def fila_grafico(nombre, grupo):
        return {"marca": nombre, **{c: sum(m[c] for m in grupo) for c in ("ingreso_costo", "costo_vendido", "inventario_costo",
                                                                            "ingreso", "vendido", "disponible")}}
    marcas_grafico = [fila_grafico(m["marca"], [m]) for m in marcas[:TOP_MARCAS_GRAFICO]]
    if len(marcas) > TOP_MARCAS_GRAFICO:
        marcas_grafico.append(fila_grafico(f"Otras ({len(marcas) - TOP_MARCAS_GRAFICO})", marcas[TOP_MARCAS_GRAFICO:]))

    # Ventas por mes: 24 meses completos + el mes en curso aparte (barra clara, con proyección)
    serie_m = total.get("monto_meses") or [0.0] * len(serie_meses)
    serie_u = total.get("unid_meses") or [0.0] * len(serie_meses)
    pm3 = M.promedio_movil(serie_m[:n36])
    actual_m, actual_u = (serie_m[n36], serie_u[n36]) if mes_actual else (None, None)
    dias_mes = monthrange(corte.year, corte.month)[1]
    por_mes_anio = defaultdict(float)
    for m, v in zip(meses24, serie_m[n36 - 24:n36]):
        por_mes_anio[int(m[5:])] += v
    temp_total = M.temporada(por_mes_anio) if sum(serie_u[n36 - 24:n36]) >= M.MIN_UNIDADES_TEMPORADA else (None, None, None)
    temp_meses = M.meses_temporada(temp_total[2])
    graf_meses = {
        "meses": meses24, "parcial": bool(mes_actual),
        "etiquetas": [M.etiqueta_mes(m) for m in meses24] + ([f"{M.etiqueta_mes(mes_actual)} (al {corte:%d/%m})"] if mes_actual else []),
        "monto": serie_m[n36 - 24:n36] + ([actual_m] if mes_actual else []),
        "unid": serie_u[n36 - 24:n36] + ([actual_u] if mes_actual else []),
        "pm3": pm3[n36 - 24:n36] + ([None] if mes_actual else []),
        # mismo mes del año anterior (para el mes en curso, el mes completo del año anterior)
        "anterior_monto": serie_m[n36 - 36:n36 - 12] + ([serie_m[n36 - 12]] if mes_actual else []),
        "anterior_unid": serie_u[n36 - 36:n36 - 12] + ([serie_u[n36 - 12]] if mes_actual else []),
        "proyeccion": actual_m * dias_mes / corte.day if mes_actual else None,
        "temporada": [i for i, m in enumerate(meses24) if int(m[5:]) in temp_meses],
    }

    # Alerta: compras de este año que ya están lentas o sin venta (sin contar los nuevos)
    recientes = [x for x in filas if x["estado"] in ("lento", "sin_venta") and x["disponible"] > 0
                 and x["ultimo_ingreso"] and x["ultimo_ingreso"].year == hoy.year]
    rec_marca = defaultdict(float)
    for x in recientes:
        rec_marca[x["marca"]] += x["inventario_costo"]
    alerta = {"productos": len(recientes), "inventario_costo": sum(x["inventario_costo"] for x in recientes),
              "marcas": sorted(rec_marca.items(), key=lambda kv: -kv[1]), "anio": hoy.year}

    # Recomendaciones
    no_recomprar = sorted((x for x in filas if x["recomendacion"] in (M.NO_RECOMPRAR, M.NO_RECOMPRAR_12, M.REVISAR_NUEVO)),
                          key=lambda x: -x["inventario_costo"])
    reponer = sorted((x for x in filas if x["recomendacion"] == M.REPONER), key=lambda x: -x["costo_12"])
    marcas_dejar = [m for m in marcas if m["dejar"]]

    # Calidad de datos
    negativos = [x for x in filas if x["existencia"] < 0 or x["disponible"] < 0]
    util_neg = [x for x in filas if x["venta_neta"] > 0 and x["utilidad"] < 0]
    calidad = {"negativos": negativos, "util_neg": util_neg, "util_neg_monto": sum(x["utilidad"] for x in util_neg),
               "costo_estimado": int(total["costo_estimado"]), "sin_costo": int(total["sin_costo"]),
               "estimados": [x for x in filas if x["costo_estimado"]]}

    d = {"total": total, "marcas": marcas, "marcas_rent": marcas_rent, "marcas_grafico": marcas_grafico, "rotacion": rotacion,
         "rot": rot, "rent": rent_total, "dias": dias, "temporada": temp_total, "sin_actividad": sin_actividad,
         "sin_venta_sin_existencia": sin_venta_sin_existencia, "graf_meses": graf_meses, "meses": meses24,
         "etiquetas_24": [M.etiqueta_mes(m) for m in meses24], "mes_actual": mes_actual, "corte": corte, "alerta": alerta,
         "no_recomprar": no_recomprar, "reponer": reponer, "marcas_dejar": marcas_dejar, "calidad": calidad,
         "agotados_venden": rot["agotado"]["productos"], "periodo_12": (M.etiqueta_mes(meses36[-12]), M.etiqueta_mes(meses36[-1])),
         "periodo_prev": (M.etiqueta_mes(meses36[-24]), M.etiqueta_mes(meses36[-13]))}
    d["graf_rot"] = [{k: e[k] for k in ("estado", "corto", "nombre", "color", "productos", "inventario_costo", "monto_12",
                                         "pct_productos", "pct_inventario", "pct_venta_12")} for e in rotacion]
    d["respuestas"] = respuestas(f, d) if filas else []

    # ---------- Filtros de la tabla ----------
    q = f["q"].upper()
    if q:
        filas = [x for x in filas if q in x["codigo"].upper() or q in x["parte"].upper() or q in x["nombre"].upper()]
    filtro = {"nhe": lambda x: x["nhe"] > 0, "ingreso": lambda x: x["ingreso"] > 0,
              "sin_venta": lambda x: x["estado"] == "sin_venta" and x["disponible"] > 0,
              "no_recomprar": lambda x: x["recomendacion"] in (M.NO_RECOMPRAR, M.NO_RECOMPRAR_12, M.REVISAR_NUEVO),
              "reponer": lambda x: x["recomendacion"] == M.REPONER,
              "negativo": lambda x: x["existencia"] < 0 or x["disponible"] < 0,
              "utilidad_neg": lambda x: x["venta_neta"] > 0 and x["utilidad"] < 0,
              "costo_estimado": lambda x: x["costo_estimado"]}.get(f["estado"])
    if f["estado"] == "lento_reciente":
        ids = {id(x) for x in recientes}
        filtro = lambda x: id(x) in ids
    if filtro:
        filas = [x for x in filas if filtro(x)]
    elif f["estado"]:
        filas = [x for x in filas if x["estado"] == f["estado"]]
    clave = {"inventario": lambda x: -x["inventario_costo"], "ingreso": lambda x: -x["ingreso_costo"],
             "vendido": lambda x: -(x["venta_neta"] or 0), "utilidad": lambda x: -(x["utilidad"] or 0),
             "sin_venta": lambda x: (x["dias_sin_venta"] is not None, -(x["dias_sin_venta"] or 0), -x["inventario_costo"]),
             "marca": lambda x: (x["marca"], x["nombre"])}[f["orden"]]
    filas.sort(key=clave)
    d["filas"] = filas
    return d


def respuestas(f, d):
    """Las 5 preguntas del reporte respondidas en texto (hoja «Resumen» del Excel y resumen de la pantalla)."""
    t, r, rot = d["total"], d["rent"], d["rot"]
    c, pct = (lambda v: f"C$ {v:,.0f}"), (lambda v: "—" if v is None else f"{v:+.0f}%")
    n_marcas = len(d["marcas"])
    p1 = (f"Desde el {f['desde']:%d/%m/%Y} ingresaron {t['ingreso']:,.0f} unidades por {c(t['ingreso_costo'])} a costo real, "
          f"de {int(t['productos']):,} productos en {n_marcas} marcas (FOB ref. US$ {t['ingreso_fob']:,.0f}).")
    p2 = (f"Se vendieron {t['vendido']:,.0f} unidades por {c(t['venta_neta'])} de venta neta, con utilidad de {c(r['utilidad'])} "
          f"(margen {r['margen'] or 0:.1f}%). Se ha vendido el {t['pct_vendido'] or 0:.0f}% de las unidades ingresadas. "
          f"Rotación de 12 meses {r['rotacion_12'] or 0:.2f} veces al año (histórica {r['rotacion_hist'] or 0:.2f}); "
          f"GMROI 12 meses {r['gmroi_12'] or 0:.2f}. {r['veredicto']}.")
    lentas = [m for m in d["marcas"] if m["lento_desde"]]
    sv = rot["sin_venta"]
    p3 = (f"El {t['pct_lento'] or 0:.0f}% del inventario a costo está lento o sin venta; {sv['productos']:,} productos con inventario "
          f"({c(sv['inventario_costo'])}) no han tenido venta en los últimos 12 meses. "
          + (("Marcas que se pusieron lentas: " + ", ".join(f"{m['marca']} desde {m['lento_desde']}" for m in lentas) + ".")
             if lentas else "Ninguna marca cayó por debajo del 70% de su venta del año anterior de forma sostenida."))
    suben = [m["marca"] for m in d["marcas"] if m["nivel"] == "sube"]
    caen = [m["marca"] for m in d["marcas"] if m["nivel"] == "baja"]
    p4 = (f"Últimos 12 meses ({d['periodo_12'][0]}–{d['periodo_12'][1]}) contra los 12 anteriores: venta en C$ {pct(t['cambio'])}, "
          f"en unidades {pct(t['cambio_unid'])}. " + (f"Suben: {', '.join(suben)}. " if suben else "")
          + (f"Caen: {', '.join(caen)}." if caen else ""))
    nr = d["no_recomprar"]
    p5 = (f"No volver a comprar: {len(nr)} productos con {c(sum(x['inventario_costo'] for x in nr))} de inventario a costo. "
          f"Reponer: {len(d['reponer'])} productos que se venden y están agotados o con menos de 2 meses de inventario. "
          + (f"Marcas que conviene dejar de comprar: {', '.join(m['marca'] for m in d['marcas_dejar'])}." if d["marcas_dejar"] else ""))
    return [("1. ¿Cuánto hemos traído?", p1), ("2. ¿Cómo lo hemos vendido?", p2), ("3. ¿Desde cuándo y en qué marca se puso lento?", p3),
            ("4. ¿Se mueve mejor en los últimos meses?", p4), ("5. ¿Qué no volver a comprar y qué reponer?", p5)]


@analisis_proveedor_bp.route("/analisis_proveedor")
@requiere_acceso("analisis_proveedor")
def reporte():
    f = leer_filtros(request.args)
    marcas_prov = marcas_del_proveedor(f["proveedor"])
    d = calcular(f)
    return render_template("analisis_proveedor.html", f=f, d=d, proveedores=lista_proveedores(), marcas_prov=marcas_prov,
                           estados=ESTADOS, filtros_rapidos=FILTROS_RAPIDOS, ordenes=ORDENES, max_filas=MAX_FILAS_PANTALLA,
                           max_rec=MAX_FILAS_RECOMENDACION, tc_fob=TC_FOB, M=M, abrev=M.abreviar,
                           min_inv_marca=MIN_INVENTARIO_MARCA)


# ---------------------------------------------------------------- Excel
C, P, U, F = "#,##0.00", "0.0%", "#,##0", "dd/mm/yyyy"


def _hoja(wb, nombre, titulo, nota, columnas, primera=False):
    ws = wb.active if primera else wb.create_sheet(nombre)
    ws.title = nombre
    ws["A1"], ws["A2"] = titulo, nota
    ws["A1"].font = Font(bold=True, size=14, color=VERDE)
    ws["A2"].font = Font(size=9, color="52514E")
    for i, (col, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=4, column=i, value=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if any(k in col for k in ("NHE", "Comentario", "Veredicto", "Recomendación")) else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[4].height = 36
    ws.freeze_panes = "B5"
    return ws


def _filas(ws, filas, formatos, sumas=(), etiqueta_total="TOTAL"):
    """Escribe las filas desde la 5, pone autofiltro y una fila de totales con fórmulas en las columnas `sumas` (1 = A)."""
    for n, valores in enumerate(filas, start=5):
        for i, (v, fmt) in enumerate(zip(valores, formatos), start=1):
            c = ws.cell(row=n, column=i, value=v)
            if fmt:
                c.number_format = fmt
    ultima = 4 + len(filas)
    ncol = len(formatos)
    ws.auto_filter.ref = f"A4:{get_column_letter(ncol)}{max(ultima, 5)}"
    if sumas and filas:
        tot = ultima + 1
        ws.cell(row=tot, column=1, value=etiqueta_total).font = Font(bold=True)
        for i in sumas:
            col = get_column_letter(i)
            c = ws.cell(row=tot, column=i, value=f"=SUBTOTAL(9,{col}5:{col}{ultima})")
            c.font = Font(bold=True)
            c.number_format = formatos[i - 1] or U
        return tot
    return None


def _r(v, d=2):
    return round(v, d) if v is not None else None


def _pct(v):
    return v / 100 if v is not None else None


@analisis_proveedor_bp.route("/analisis_proveedor/excel")
@requiere_acceso("analisis_proveedor")
def excel():
    f = leer_filtros(request.args)
    d = calcular(f)
    wb = Workbook()
    nota = (f"Proveedor: {f['proveedor']} · Marcas: {', '.join(f['marcas']) or 'todas'}{' · Solo con existencia' if f['existencia'] else ''} · "
            f"Ingresos y ventas desde el {f['desde']:%d/%m/%Y} · 12 meses = {d.get('periodo_12', ('', ''))[0]}–{d.get('periodo_12', ('', ''))[1]} "
            f"(meses completos) · Todo en C$: ingresado a costo real, inventario y NHE a último costo (bases distintas: no se restan), "
            f"ventas netas (sin IVA, con descuento, sin anuladas) contra su costo · FOB US$ solo de referencia · "
            f"Generado {datetime.now():%d/%m/%Y %H:%M}")
    t = d.get("total") or {}

    # --- Resumen: las 5 respuestas ---
    ws = _hoja(wb, "Resumen", f"Análisis de proveedor: {f['proveedor']}", nota, [("Pregunta", 44), ("Respuesta", 120)], primera=True)
    for n, (preg, resp) in enumerate(d.get("respuestas", []), start=5):
        ws.cell(row=n, column=1, value=preg).font = Font(bold=True, color=VERDE)
        c = ws.cell(row=n, column=2, value=resp)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=n, column=1).alignment = Alignment(vertical="top")
        ws.row_dimensions[n].height = 62
    n = 5 + len(d.get("respuestas", [])) + 1
    if d.get("rent"):
        r = d["rent"]
        for et, v, fmt in (("Ingresado a costo real C$", t["ingreso_costo"], C), ("Venta neta C$", t["venta_neta"], C),
                           ("Utilidad bruta C$", r["utilidad"], C), ("Inventario a último costo C$", t["inventario_costo"], C),
                           ("Rotación histórica (veces/año)", r["rotacion_hist"], "0.00"), ("Rotación 12 meses (veces/año)", r["rotacion_12"], "0.00"),
                           ("GMROI histórico", r["gmroi_hist"], "0.00"), ("GMROI 12 meses", r["gmroi_12"], "0.00")):
            ws.cell(row=n, column=1, value=et)
            ws.cell(row=n, column=2, value=_r(v)).number_format = fmt
            ws.cell(row=n, column=2).alignment = Alignment(horizontal="left")
            n += 1
    ws.freeze_panes = "A5"

    # --- Por marca (todas las marcas) ---
    cols = [("Marca", 22), ("Productos", 10), ("Ingresado (unid.)", 11), ("Ingresado a costo C$", 14), ("Vendido (unid.)", 11),
            ("% vendido de lo ingresado (unid.)", 11), ("Venta neta C$", 14), ("Costo de lo vendido C$", 14), ("Utilidad bruta C$", 14),
            ("Margen %", 9), ("Inventario disponible (unid.)", 12), ("Inventario a costo C$", 14), ("% inventario lento o sin venta", 11),
            ("NHE (unid.)", 10), ("NHE a costo C$", 13), ("Productos sin venta 12m (con inv.)", 11), ("Inventario sin venta 12m C$", 14),
            ("Venta 12 meses C$", 14), ("Rotación histórica", 10), ("Rotación 12 meses", 10), ("GMROI histórico", 10), ("GMROI 12 meses", 10),
            ("Cambio 12m C$", 10), ("Cambio 12m unid.", 10), ("Lento desde", 11), ("Veredicto", 60),
            ("Ingresado FOB US$ (ref.)", 13), ("Inventario FOB US$ (ref.)", 13)]
    fm = [None, U, U, C, U, P, C, C, C, P, U, C, P, U, C, U, C, C, "0.00", "0.00", "0.00", "0.00", P, P, None, None, C, C]
    ws = _hoja(wb, "Por marca", f"Resumen por marca: {f['proveedor']}", nota, cols)
    filas = []
    for m in d.get("marcas", []):
        r = m["rent"]
        filas.append([m["marca"], m["productos"], m["ingreso"], _r(m["ingreso_costo"]), m["vendido"], _pct(m["pct_vendido"]),
                      _r(m["venta_neta"]), _r(m["costo_vendido"]), _r(r["utilidad"]), _pct(r["margen"]), m["disponible"],
                      _r(m["inventario_costo"]), _pct(m["pct_lento"]), m["nhe"], _r(m["nhe_costo"]), m["sin_venta"], _r(m["sin_venta_costo"]),
                      _r(m["monto_12"]), _r(r["rotacion_hist"]), _r(r["rotacion_12"]), _r(r["gmroi_hist"]), _r(r["gmroi_12"]),
                      _pct(m["cambio"]), _pct(m["cambio_unid"]), m["lento_desde"], r["veredicto"], _r(m["ingreso_fob"]), _r(m["inventario_fob"])])
    tot = _filas(ws, filas, fm, sumas=[2, 3, 4, 5, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 27, 28])
    if tot and d.get("rent"):
        r = d["rent"]
        for col, v, fmt in ((6, _pct(t["pct_vendido"]), P), (10, f"=IFERROR(I{tot}/G{tot},\"\")", P), (13, _pct(t["pct_lento"]), P),
                            (19, _r(r["rotacion_hist"]), "0.00"), (20, _r(r["rotacion_12"]), "0.00"), (21, _r(r["gmroi_hist"]), "0.00"),
                            (22, _r(r["gmroi_12"]), "0.00"), (23, _pct(t["cambio"]), P), (24, _pct(t["cambio_unid"]), P), (26, r["veredicto"], None)):
            c = ws.cell(row=tot, column=col, value=v)
            c.font = Font(bold=True)
            if fmt:
                c.number_format = fmt

    # --- Tendencia por marca: venta mensual de 24 meses ---
    etiquetas = d.get("etiquetas_24", [])
    p12, pp = d.get("periodo_12", ("", "")), d.get("periodo_prev", ("", ""))
    cols = ([("Marca", 22), (f"Venta 12m C$ ({p12[0]}–{p12[1]})", 14), (f"12 meses anteriores C$ ({pp[0]}–{pp[1]})", 14), ("Cambio C$", 10),
             ("Unid. 12m", 11), ("Unid. 12 anteriores", 11), ("Cambio unid.", 10), ("Lento desde", 11),
             ("Mediana última venta (sin venta 12m)", 13)] + [(e, 11) for e in etiquetas])
    fm = [None, C, C, P, U, U, P, None, F] + [C] * len(etiquetas)
    ws = _hoja(wb, "Tendencia por marca", "Tendencia por marca: venta neta C$ por mes (meses completos)", nota, cols)
    filas = [[m["marca"], _r(m["monto_12"]), _r(m["monto_prev"]), _pct(m["cambio"]), m["venta_12"], m["unid_prev"], _pct(m["cambio_unid"]),
              m["lento_desde"], m["mediana_ultima_venta"]] + [_r(v) for v in m["heat"]] for m in d.get("marcas", [])]
    tot = _filas(ws, filas, fm, sumas=[2, 3, 5, 6] + list(range(10, 10 + len(etiquetas))))
    if tot:
        ws.cell(row=tot, column=4, value=f"=IFERROR(B{tot}/C{tot}-1,\"\")").number_format = P
        ws.cell(row=tot, column=7, value=f"=IFERROR(E{tot}/F{tot}-1,\"\")").number_format = P

    # --- Detalle ---
    cols = [("Código", 11), ("Nº parte", 18), ("Descripción", 40), ("Marca", 14), ("Segmento", 16), ("Familia", 22), ("Subfamilia", 24),
            ("Categoría", 16), ("Costo unitario C$", 11), ("Primer ingreso", 12), ("Meses en inventario", 10), ("Ingresado (unid.)", 10),
            ("Ingresado a costo C$", 13), ("Último ingreso", 12), ("Vendido (unid.)", 10), ("Venta neta C$", 13), ("Costo de lo vendido C$", 13),
            ("Utilidad C$", 12), ("Margen %", 9), ("Venta 12m (unid.)", 10), ("Venta 12m C$", 13), ("Existencia total", 10), ("NHE (unid.)", 9),
            ("Inventario disponible", 10), ("Inventario a costo C$", 13), ("NHE a costo C$", 12), ("Venta mensual", 10),
            ("Meses de inventario", 10), ("Última venta", 12), ("Temporada de más venta", 12), ("% del año en esa temporada", 10),
            ("Clasificación", 14), ("Recomendación", 22), ("Comentario", 52), ("Precio FOB US$ (ref.)", 11)]
    fm = [None] * 8 + [C, F, "0.0", U, C, F, U, C, C, C, P, U, C, U, U, U, C, C, "#,##0.0", "#,##0.0", F, None, "0%", None, None, None, "#,##0.00##"]
    ws = _hoja(wb, "Detalle", f"Detalle por producto: {f['proveedor']}", nota, cols)
    filas = [[x["codigo"], x["parte"], x["nombre"], x["marca"], x["segmento"], x["familia"], x["subfamilia"], x["categoria"],
              _r(x["costo"]), x["primer_ingreso"], _r(x["meses_inventario"], 1), x["ingreso"], _r(x["ingreso_costo"]), x["ultimo_ingreso"],
              x["vendido"], _r(x["venta_neta"]), _r(x["costo_vendido"]), _r(x["utilidad"]), _pct(x["margen"]), x["venta_12"], _r(x["monto_12"]),
              x["existencia"], x["nhe"], x["disponible"], _r(x["inventario_costo"]), _r(x["nhe_costo"]), _r(x["mensual"]),
              _r(x["cobertura"], 1), x["ultima_venta"], x["temporada"], _pct(x["temporada_pct"]), M.CORTO[x["estado"]],
              x["recomendacion"], x["comentario"] + (" · costo estimado con FOB" if x["costo_estimado"] else ""), x["fob"]]
             for x in d.get("filas", [])]
    _filas(ws, filas, fm, sumas=[12, 13, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 26])

    # --- No volver a comprar / Reponer ---
    cols = [("Código", 11), ("Nº parte", 18), ("Descripción", 40), ("Marca", 14), ("Clasificación", 14), ("Inventario disponible", 11),
            ("Inventario a costo C$", 14), ("Venta 12m (unid.)", 10), ("Meses de inventario", 10), ("Última venta", 12), ("Último ingreso", 12),
            ("Recomendación", 32)]
    fm = [None, None, None, None, None, U, C, U, "#,##0.0", F, F, None]
    ws = _hoja(wb, "No volver a comprar", "No volver a comprar: con inventario y más de 24 meses de cobertura o sin venta en 12 meses", nota, cols)
    _filas(ws, [[x["codigo"], x["parte"], x["nombre"], x["marca"], M.CORTO[x["estado"]], x["disponible"], _r(x["inventario_costo"]),
                 x["venta_12"], _r(x["cobertura"], 1), x["ultima_venta"], x["ultimo_ingreso"], x["recomendacion"]]
                for x in d.get("no_recomprar", [])], fm, sumas=[6, 7, 8])
    if d.get("marcas_dejar"):
        n = ws.max_row + 2
        ws.cell(row=n, column=1, value="Marcas que conviene dejar de comprar (caen más de 30% y más del 85% de su inventario está lento): "
                + ", ".join(m["marca"] for m in d["marcas_dejar"])).font = Font(bold=True, color=NARANJA)
    cols = [("Código", 11), ("Nº parte", 18), ("Descripción", 40), ("Marca", 14), ("Clasificación", 14), ("Inventario disponible", 11),
            ("Venta 12m (unid.)", 10), ("Costo vendido 12m C$", 14), ("Venta mensual", 10), ("Meses de inventario", 10), ("Temporada", 12),
            ("Última venta", 12)]
    fm = [None, None, None, None, None, U, U, C, "#,##0.0", "#,##0.0", None, F]
    ws = _hoja(wb, "Reponer", "Reponer: con venta en 12 meses y agotados o con menos de 2 meses de inventario", nota, cols)
    _filas(ws, [[x["codigo"], x["parte"], x["nombre"], x["marca"], M.CORTO[x["estado"]], x["disponible"], x["venta_12"], _r(x["costo_12"]),
                 _r(x["mensual"]), _r(x["cobertura"], 1), x["temporada"], x["ultima_venta"]] for x in d.get("reponer", [])], fm, sumas=[6, 7, 8])

    # --- Venta por mes ---
    g = d.get("graf_meses") or {"etiquetas": [], "unid": [], "monto": [], "anterior_monto": []}
    ws = _hoja(wb, "Venta por mes", "Ventas por mes (24 meses completos y el mes en curso)", nota,
               [("Mes", 18), ("Unidades vendidas", 12), ("Venta neta C$", 14), ("Mismo mes año anterior C$", 14), ("Variación", 10)])
    filas = [[e, u, _r(v), _r(a), f"=IFERROR(C{i}/D{i}-1,\"\")"] for i, (e, u, v, a) in
             enumerate(zip(g["etiquetas"], g["unid"], g["monto"], g["anterior_monto"]), start=5)]
    _filas(ws, filas, [None, U, C, C, P], sumas=[2, 3, 4], etiqueta_total="TOTAL (incluye el mes en curso)")

    # --- Rotación: por tipo y por marca ---
    ws = _hoja(wb, "Rotación", "Rotación general por tipo y marca", nota,
               [("Rotación", 36), ("Marca", 22), ("Productos", 10), ("% de productos", 10), ("Inventario disponible (unid.)", 12),
                ("Inventario a costo C$", 14), ("% del inventario", 10), ("Venta 12m C$", 14), ("Venta histórica C$", 14)])
    filas = []
    for e in d.get("rotacion", []):
        filas.append([e["nombre"], "TODAS", e["productos"], _pct(e["pct_productos"]), e["disponible"], _r(e["inventario_costo"]),
                      _pct(e["pct_inventario"]), _r(e["monto_12"]), _r(e["venta_neta"])])
        filas += [[e["nombre"], y["marca"], y["productos"], None, y["disponible"], _r(y["inventario_costo"]), _pct(y["pct"]),
                   _r(y["monto_12"]), _r(y["venta_neta"])] for y in e["marcas"]]
    _filas(ws, filas, [None, None, U, P, U, C, P, C, C])
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        if row[1].value == "TODAS":
            for c in row:
                c.font = Font(bold=True)
    ws.cell(row=ws.max_row + 2, column=1, value=f"«Sin venta en 12 meses» cuenta solo productos con inventario; "
            f"{d.get('sin_venta_sin_existencia', 0)} sin venta y sin existencia no se cuentan. En las filas de cada marca, "
            "«% del inventario» es la parte de ese tipo de rotación.").font = Font(size=9, italic=True)

    # --- Calidad de datos ---
    cal = d.get("calidad") or {"negativos": [], "util_neg": [], "costo_estimado": 0}
    ws = _hoja(wb, "Calidad de datos", "Calidad de datos: revisar en el sistema", nota,
               [("Aviso", 34), ("Código", 11), ("Nº parte", 18), ("Descripción", 40), ("Marca", 14), ("Existencia", 10), ("Disponible", 10),
                ("Venta neta C$", 13), ("Utilidad C$", 13), ("Costo unitario C$", 12)])
    filas = ([["Inventario negativo", x["codigo"], x["parte"], x["nombre"], x["marca"], x["existencia"], x["disponible"],
               _r(x["venta_neta"]), _r(x["utilidad"]), _r(x["costo"])] for x in cal["negativos"]]
             + [["Utilidad negativa", x["codigo"], x["parte"], x["nombre"], x["marca"], x["existencia"], x["disponible"],
                 _r(x["venta_neta"]), _r(x["utilidad"]), _r(x["costo"])] for x in cal["util_neg"]]
             + [["Costo estimado con FOB × " + str(TC_FOB), x["codigo"], x["parte"], x["nombre"], x["marca"], x["existencia"],
                 x["disponible"], _r(x["venta_neta"]), _r(x["utilidad"]), _r(x["costo"])] for x in cal.get("estimados", [])])
    _filas(ws, filas, [None, None, None, None, None, U, U, C, C, C])

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = "".join(ch if ch.isalnum() else "_" for ch in f["proveedor"])[:40] or "proveedor"
    return send_file(salida, as_attachment=True, download_name=f"analisis_proveedor_{nombre}_{f['desde']}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
