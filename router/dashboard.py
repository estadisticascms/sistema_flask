
from flask import Blueprint, render_template, request
from datetime import date, datetime, timedelta
from utils.db import get_connection
from router.accesos import requiere_acceso

dashboard_bp = Blueprint("dashboard", __name__)

# -----------------------------
# Venta neta de un período (facturado − devoluciones cerradas sin IVA)
# -----------------------------
def venta_neta(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT (SELECT COALESCE(SUM(neto), 0) FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s)
             - (SELECT COALESCE(SUM(totaldv), 0) / 1.15 FROM devoluciones WHERE fechadv BETWEEN %s AND %s AND estado = 'Cerrado')
    """, (desde, hasta, desde, hasta))
    valor = cursor.fetchone()[0]
    cursor.close()
    conn.close()
    return float(valor or 0)


def mismo_periodo(fecha_dia, meses_atras):
    """Del día 1 al mismo día (o al último día si el mes es más corto) de N meses atrás."""
    anio, mes = fecha_dia.year, fecha_dia.month - meses_atras
    while mes < 1:
        anio, mes = anio - 1, mes + 12
    ultimo = (date(anio + (mes == 12), mes % 12 + 1, 1) - timedelta(days=1)).day
    return date(anio, mes, 1), date(anio, mes, min(fecha_dia.day, ultimo))


# -----------------------------
# Ranking de vendedores de un período
# -----------------------------
def obtener_ranking_vendedores(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT datos.Vendedor,
               SUM(datos.Credito) - SUM(datos.DevCredito) AS Credito,
               SUM(datos.Contado) - SUM(datos.DevContado) AS Contado,
               (SUM(datos.Credito) - SUM(datos.DevCredito)) + (SUM(datos.Contado) - SUM(datos.DevContado)) AS Neto_Ajustado
        FROM (
            SELECT v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE v.fecha BETWEEN %s AND %s GROUP BY v.Vendedor
            UNION ALL
            SELECT r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado' GROUP BY r.Vendedor
        ) datos
        GROUP BY datos.Vendedor
        ORDER BY Neto_Ajustado DESC;
    """
    cursor.execute(query, (desde, hasta, desde, hasta))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

# -----------------------------
# Productos de un período (una sola lectura de ventas_por_producto para top 10 y marcas)
# -----------------------------
def obtener_productos(desde, hasta):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT v.Codigo AS codigo, MAX(v.Numero_parte) AS numero_parte, MAX(v.`Descrip Prod.`) AS descripcion,
               MAX(v.Marca) AS marca, MAX(c.segmento) AS segmento, MAX(c.familia) AS familia,
               SUM(v.`Cantidad Vendida`) AS unidades,
               SUM(v.`Sub-Total 2`) AS venta,
               SUM(v.`Cantidad Vendida` * v.`Costo Unit.`) AS costo,
               COUNT(DISTINCT v.`No. Factura`) AS facturas
        FROM ventas_por_producto v
        LEFT JOIN clase_prod c ON c.codigo = v.Codigo
        WHERE v.Fecha BETWEEN %s AND %s AND v.Estatus <> 'Anulado'
        GROUP BY v.Codigo
        ORDER BY venta DESC
    """
    cursor.execute(query, (str(desde), str(hasta)))  # Fecha es texto 'YYYY-MM-DD' en esta tabla
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        r["venta"] = float(r["venta"] or 0)
        r["costo"] = float(r["costo"] or 0)
        r["unidades"] = float(r["unidades"] or 0)
        r["margen"] = (r["venta"] - r["costo"]) / r["venta"] * 100 if r["venta"] else None
    return rows


def ranking_marcas(productos, limite=10):
    """Agrupa los productos por marca; el resto se pliega en 'Otras marcas'."""
    marcas = {}
    for p in productos:
        m = marcas.setdefault(p["marca"] or "Sin marca", {"Marca": p["marca"] or "Sin marca", "VentasTotales": 0.0,
                                                          "Unidades": 0.0, "Costo": 0.0, "Productos": 0})
        m["VentasTotales"] += p["venta"]
        m["Unidades"] += p["unidades"]
        m["Costo"] += p["costo"]
        m["Productos"] += 1
    filas = sorted(marcas.values(), key=lambda m: -m["VentasTotales"])
    for m in filas:
        m["Margen"] = (m["VentasTotales"] - m["Costo"]) / m["VentasTotales"] * 100 if m["VentasTotales"] else None
    top, resto = filas[:limite], filas[limite:]
    otras = None
    if resto:
        venta = sum(m["VentasTotales"] for m in resto)
        costo = sum(m["Costo"] for m in resto)
        otras = {"Marca": f"Otras {len(resto)} marcas", "VentasTotales": venta, "Costo": costo,
                 "Unidades": sum(m["Unidades"] for m in resto), "Productos": sum(m["Productos"] for m in resto),
                 "Margen": (venta - costo) / venta * 100 if venta else None}
    return filas, top, otras


# -----------------------------
# Indicadores del mes (del día 1 a la fecha elegida) y comparaciones al mismo día
# -----------------------------
def obtener_indicadores_mes(fecha_dia):
    inicio = fecha_dia.replace(day=1)
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT COUNT(*) AS facturas, COUNT(DISTINCT Fecha) AS dias, COUNT(DISTINCT cliente) AS clientes,
               COALESCE(SUM(neto), 0) AS facturado, COALESCE(SUM(utilidad), 0) AS utilidad,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Credito' THEN neto END), 0) AS credito,
               COALESCE(SUM(CASE WHEN tipo_venta = 'Contado' THEN neto END), 0) AS contado
        FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s
    """, (inicio, fecha_dia))
    m = {k: float(v or 0) for k, v in cursor.fetchone().items()}
    # Devoluciones cerradas (sin IVA), separadas por tipo de factura
    cursor.execute("""
        SELECT COALESCE(SUM(totaldv), 0) / 1.15 AS total,
               COALESCE(SUM(CASE WHEN tipofact = 'Credito' THEN totaldv END), 0) / 1.15 AS credito
        FROM devoluciones WHERE fechadv BETWEEN %s AND %s AND estado = 'Cerrado'
    """, (inicio, fecha_dia))
    dev = cursor.fetchone()
    cursor.execute("SELECT COALESCE(SUM(total), 0) AS total, COUNT(*) AS recibos FROM reciboscaja WHERE fecharc BETWEEN %s AND %s",
                   (inicio, fecha_dia))
    rec = cursor.fetchone()
    cursor.close()
    conn.close()

    m["devoluciones"] = float(dev["total"] or 0)
    dev_credito = float(dev["credito"] or 0)
    m["credito"] -= dev_credito
    m["contado"] -= m["devoluciones"] - dev_credito
    m["neto"] = m["facturado"] - m["devoluciones"]
    m["margen"] = m["utilidad"] / m["facturado"] * 100 if m["facturado"] else None
    m["ticket"] = m["facturado"] / m["facturas"] if m["facturas"] else None
    m["promedio"] = m["neto"] / m["dias"] if m["dias"] else None
    m["recuperacion"], m["recibos"] = float(rec["total"] or 0), int(rec["recibos"] or 0)
    m["inicio"] = inicio
    comparaciones = []
    for etiqueta, meses_atras in (("mes anterior", 1), ("mismo mes año anterior", 12)):
        desde, hasta = mismo_periodo(fecha_dia, meses_atras)
        valor = venta_neta(desde, hasta)
        comparaciones.append({"etiqueta": etiqueta, "desde": desde, "hasta": hasta, "valor": valor,
                              "variacion": (m["neto"] / valor - 1) * 100 if valor else None})
    return m, comparaciones


# -----------------------------
# Ingresos de mercadería (órdenes de compra) por proveedor
# -----------------------------
def ingresos_por_proveedor(desde, hasta, limite=10):
    """Desde 2023 el ingreso no trae proveedor: se completa con el último proveedor del catálogo."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT t.id_documento_origen AS documento, t.codigo_producto AS codigo, t.cantidad, t.total_mn,
               COALESCE(NULLIF(TRIM(t.cliente_proveedor), ''), NULLIF(TRIM(c.`Proveedor Ultima Compra`), ''), 'Sin proveedor') AS proveedor
        FROM transacciones_inventario t
        LEFT JOIN catalogo_productos c ON c.`Codigo Producto` = t.codigo_producto
        WHERE t.tipo_transaccion = 'ENTRADA' AND t.tipo_documento_origen = 'ORDEN COMPRA' AND t.fecha BETWEEN %s AND %s
    """, (desde, hasta))
    filas = cursor.fetchall()
    cursor.close()
    conn.close()
    prov = {}
    for r in filas:
        p = prov.setdefault(r["proveedor"], {"proveedor": r["proveedor"], "costo": 0.0, "unidades": 0.0, "documentos": set(), "productos": set()})
        p["costo"] += float(r["total_mn"] or 0)
        p["unidades"] += float(r["cantidad"] or 0)
        p["documentos"].add(r["documento"])
        p["productos"].add(r["codigo"])
    lista = sorted(prov.values(), key=lambda p: -p["costo"])
    for p in lista:
        p["documentos"], p["productos"] = len(p["documentos"]), len(p["productos"])
    resto = lista[limite:]
    return {"todos": lista, "top": lista[:limite], "total": sum(p["costo"] for p in lista),
            "otros": {"proveedor": f"Otros {len(resto)} proveedores", "costo": sum(p["costo"] for p in resto)} if resto else None,
            "documentos": len({r["documento"] for r in filas}), "productos": len({r["codigo"] for r in filas})}


# -----------------------------
# Cartera de crédito al corte (reutiliza el análisis de cartera)
# -----------------------------
def cartera_al(fecha_dia):
    from router.analisis_cartera import calcular, INICIO_DATOS   # aquí: evita importaciones circulares
    from router.ventas_bodegas import en_cache
    f = {"desde": INICIO_DATOS, "hasta": fecha_dia, "corte": fecha_dia, "vendedor": "", "q": "", "reco": ""}
    r = en_cache(f"dashboard_cartera_{fecha_dia}", lambda: calcular(f), segundos=600)
    return {"saldo": r["saldo"], "vencido": r["vencido"], "pct_vencido": r["vencido"] / r["saldo"] * 100 if r["saldo"] else None,
            "dias": r["dias"], "tramos": r["tramos"], "suspender": r["conteo_reco"].get("suspender", 0),
            "revisar": r["conteo_reco"].get("revisar", 0), "desde": INICIO_DATOS}


def obtener_resumen_mensual(fecha_inicio, fecha_fin, credito, contado):
    """Una fila por mes: crédito/contado netos de devoluciones (del reporte mensual) + días, facturas y margen."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT DATE_FORMAT(Fecha, '%Y-%m') AS Mes, COUNT(DISTINCT Fecha) AS dias, COUNT(*) AS facturas,
               COALESCE(SUM(neto), 0) AS neto, COALESCE(SUM(utilidad), 0) AS utilidad, MAX(Fecha) AS ultima
        FROM ventas_historicas WHERE Fecha BETWEEN %s AND %s
        GROUP BY Mes ORDER BY Mes
    """, (fecha_inicio, fecha_fin))
    datos = {r["Mes"]: r for r in cursor.fetchall()}
    cursor.close()
    conn.close()

    filas, anterior = [], None
    for mes in sorted(set(credito) | set(contado) | set(datos)):
        d = datos.get(mes, {})
        cr, co = float(credito.get(mes, 0) or 0), float(contado.get(mes, 0) or 0)
        total, dias = cr + co, int(d.get("dias") or 0)
        promedio = total / dias if dias else None
        ultima = d.get("ultima")
        siguiente_dia = ultima + timedelta(days=1) if ultima else None
        filas.append({
            "mes": mes, "credito": cr, "contado": co, "total": total, "dias": dias,
            "facturas": int(d.get("facturas") or 0),
            "pct_credito": cr / total * 100 if total else None,
            "promedio": promedio,
            # Se compara el promedio diario: así el mes en curso (incompleto) es comparable
            "variacion": (promedio / anterior - 1) * 100 if promedio and anterior else None,
            "margen": float(d["utilidad"]) / float(d["neto"]) * 100 if d.get("neto") else None,
            "parcial_al": ultima.day if ultima and siguiente_dia.month == ultima.month
                          and ultima == datetime.strptime(fecha_fin, "%Y-%m-%d").date() else None,
        })
        anterior = promedio
    return filas


def ultima_fecha_con_ventas():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(Fecha) FROM ventas_historicas")
    fecha = cursor.fetchone()[0]
    cursor.close()
    conn.close()
    return fecha
# -----------------------------
# Ventas mensuales solo Totales
# -----------------------------
def obtener_totales_vendedor_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT Mes, Vendedor,
               (SUM(Credito) - SUM(DevCredito)) + (SUM(Contado) - SUM(DevContado)) AS Total
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes, v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE v.fecha BETWEEN %s AND %s GROUP BY Mes, v.Vendedor
            UNION ALL
            SELECT DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes, r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado' GROUP BY Mes, r.Vendedor
        ) datos
        GROUP BY Mes, Vendedor
        ORDER BY Vendedor, Mes;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    ventas_por_vendedor = {}
    meses = sorted(set(r["Mes"] for r in rows))
    for r in rows:
        vendedor = r["Vendedor"]
        mes = r["Mes"]
        if vendedor not in ventas_por_vendedor:
            ventas_por_vendedor[vendedor] = {}
        ventas_por_vendedor[vendedor][mes] = r["Total"]

    vendedores_ordenados = sorted(
        ventas_por_vendedor.keys(),
        key=lambda v: sum(ventas_por_vendedor[v].values()),
        reverse=True
    )
    return ventas_por_vendedor, meses, vendedores_ordenados

# -----------------------------
# Ventas mensuales con Crédito/Contado
# -----------------------------
def obtener_ventas_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT Mes, Vendedor,
               SUM(Credito) - SUM(DevCredito) AS Credito,
               SUM(Contado) - SUM(DevContado) AS Contado
        FROM (
            SELECT DATE_FORMAT(v.fecha, '%Y-%m') AS Mes, v.Vendedor,
                   SUM(CASE WHEN v.tipo_venta='Credito' THEN v.neto ELSE 0 END) AS Credito,
                   SUM(CASE WHEN v.tipo_venta='Contado' THEN v.neto ELSE 0 END) AS Contado,
                   0 AS DevCredito, 0 AS DevContado
            FROM ventas_historicas v WHERE v.fecha BETWEEN %s AND %s GROUP BY Mes, v.Vendedor
            UNION ALL
            SELECT DATE_FORMAT(r.fechadv, '%Y-%m') AS Mes, r.Vendedor,
                   0 AS Credito, 0 AS Contado,
                   SUM(CASE WHEN r.tipofact='Credito' THEN r.totaldv/1.15 ELSE 0 END) AS DevCredito,
                   SUM(CASE WHEN r.tipofact='Contado' THEN r.totaldv/1.15 ELSE 0 END) AS DevContado
            FROM devoluciones r WHERE r.fechadv BETWEEN %s AND %s AND r.estado='Cerrado' GROUP BY Mes, r.Vendedor
        ) datos
        GROUP BY Mes, Vendedor
        ORDER BY Mes, Vendedor;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin, fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    ventas_por_vendedor = {}
    meses = sorted(set(r["Mes"] for r in rows))
    for r in rows:
        vendedor = r["Vendedor"]
        mes = r["Mes"]
        if vendedor not in ventas_por_vendedor:
            ventas_por_vendedor[vendedor] = {}
        ventas_por_vendedor[vendedor][mes] = {
            "Credito": r["Credito"],
            "Contado": r["Contado"]
        }
    return ventas_por_vendedor, meses


# -----------------------------
# Recuperación mensual
# -----------------------------
def obtener_recuperacion_mes(fecha_inicio, fecha_fin):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    query = """
        SELECT DATE_FORMAT(fecharc, '%Y-%m') AS Mes,
               COUNT(numerorc) AS Cantidad,
               SUM(total) AS TotalCordobas
        FROM reciboscaja
        WHERE fecharc BETWEEN %s AND %s
        GROUP BY Mes
        ORDER BY Mes;
    """
    cursor.execute(query, (fecha_inicio, fecha_fin))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

# -----------------------------
# Ruta principal del dashboard
# -----------------------------
@dashboard_bp.route("/dashboard", methods=["GET"])
@requiere_acceso("dashboard")
def dashboard_view():
    # Por defecto: el último día con ventas cargadas (no "hoy", que suele estar vacío hasta la carga)
    ultima = ultima_fecha_con_ventas()
    defecto = ultima or datetime.today().date()
    try:
        fecha_dia = datetime.strptime(request.args.get("fecha_fin") or str(defecto), "%Y-%m-%d").date()
    except ValueError:
        fecha_dia = defecto
    fecha_fin = fecha_dia.strftime("%Y-%m-%d")
    fecha_inicio = request.args.get("fecha_inicio") or f"{fecha_dia.year}-01-01"

    # Mes de la fecha elegida (del día 1 a esa fecha)
    inicio_mes = fecha_dia.replace(day=1)
    mes, comparaciones = obtener_indicadores_mes(fecha_dia)
    ranking_vendedores_mes = obtener_ranking_vendedores(inicio_mes, fecha_dia)
    productos_mes = obtener_productos(inicio_mes, fecha_dia)
    top_productos = productos_mes[:10]
    ranking_marcas_mes, top_marcas, otras_marcas = ranking_marcas(productos_mes)
    venta_productos_mes = sum(p["venta"] for p in productos_mes)
    ingresos = ingresos_por_proveedor(inicio_mes, fecha_dia)
    cartera = cartera_al(fecha_dia)

    total_credito = sum(r["Credito"] or 0 for r in ranking_vendedores_mes)
    total_contado = sum(r["Contado"] or 0 for r in ranking_vendedores_mes)
    total_neto = sum(r["Neto_Ajustado"] or 0 for r in ranking_vendedores_mes)

    # Reporte mensual con Crédito/Contado
    ventas_detalladas, meses_detallados = obtener_ventas_mes(fecha_inicio, fecha_fin)

    # Reporte mensual solo Totales
    ventas_totales, meses_totales, vendedores_ordenados = obtener_totales_vendedor_mes(fecha_inicio, fecha_fin)

    # Totales generales por mes (para gráficos)
    totales_por_mes_credito = {
        mes_: float(sum(ventas_detalladas[v].get(mes_, {}).get("Credito", 0) or 0 for v in ventas_detalladas))
        for mes_ in meses_detallados
    }
    totales_por_mes_contado = {
        mes_: float(sum(ventas_detalladas[v].get(mes_, {}).get("Contado", 0) or 0 for v in ventas_detalladas))
        for mes_ in meses_detallados
    }


    # Recuperación mensual
    recuperacion_mes = obtener_recuperacion_mes(fecha_inicio, fecha_fin)
    total_general = sum(totales_por_mes_credito.values()) + sum(totales_por_mes_contado.values())
    resumen_mensual = obtener_resumen_mensual(fecha_inicio, fecha_fin, totales_por_mes_credito, totales_por_mes_contado)

    # Navegación por mes: fin del mes anterior / siguiente (sin pasar la última fecha cargada)
    mes_anterior = inicio_mes - timedelta(days=1)
    fin_siguiente = (inicio_mes + timedelta(days=62)).replace(day=1) - timedelta(days=1)
    mes_siguiente = min(fin_siguiente, ultima) if ultima and ultima > fecha_dia else None

    return render_template("dashboard.html",
                           fecha_inicio=fecha_inicio,
                           fecha_fin=fecha_fin,
                           fecha_dia=fecha_dia,
                           mes_anterior=mes_anterior.strftime("%Y-%m-%d"),
                           mes_siguiente=mes_siguiente.strftime("%Y-%m-%d") if mes_siguiente else None,
                           ultima_fecha=ultima,
                           mes=mes,
                           comparaciones=comparaciones,
                           ranking_vendedores_mes=ranking_vendedores_mes,
                           ranking_marcas_mes=ranking_marcas_mes,
                           top_marcas=top_marcas,
                           otras_marcas=otras_marcas,
                           top_productos=top_productos,
                           total_productos_mes=len(productos_mes),
                           venta_productos_mes=venta_productos_mes,
                           ingresos=ingresos,
                           cartera=cartera,
                           total_general=total_general,
                           resumen_mensual=resumen_mensual,
                           total_credito=total_credito,
                           total_contado=total_contado,
                           total_neto=total_neto,
                           ventas_detalladas=ventas_detalladas,
                           meses_detallados=meses_detallados,
                           ventas_totales=ventas_totales,
                           meses_totales=meses_totales,
                           vendedores_ordenados=vendedores_ordenados,
                           totales_por_mes_credito=totales_por_mes_credito,
                           totales_por_mes_contado=totales_por_mes_contado,
                           recuperacion_mes=recuperacion_mes)
