"""Métricas del Análisis de proveedor: funciones puras (sin base de datos ni Flask).
La pantalla y el Excel usan los mismos resultados de estas funciones, así los números siempre coinciden.

Convenciones:
- «12 meses» = los últimos 12 meses COMPLETOS (sin el mes en curso, que solo tiene unos días).
- Montos en C$ a costo o a venta neta; el FOB en US$ solo es referencia.
"""
from datetime import date

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
DIAS_MES = 30.44

MESES_NUEVO = 3                 # con menos meses desde el primer ingreso: «Nuevo – sin evaluar»
COBERTURA_LENTA = 12            # más meses de inventario: lento
COBERTURA_NO_RECOMPRAR = 24     # más meses de inventario: no recomprar hasta bajar de 12
COBERTURA_REPONER = 2           # menos meses de inventario (y con venta): reponer
UMBRAL_TENDENCIA = 10           # % de cambio para decir que sube o cae
UMBRAL_LENTO_DESDE = 0.70       # promedio móvil 3 meses bajo el 70% del promedio del año anterior
CAIDA_DEJAR = -30               # marcas que caen más de 30% …
PCT_LENTO_DEJAR = 85            # … y tienen más del 85% de su inventario lento o sin venta
GMROI_BUENO, GMROI_ACEPTABLE = 1.5, 0.8
MIN_UNIDADES_TEMPORADA = 6

# Clasificación por producto: clave, nombre largo, nombre corto, color fijo en toda la pantalla
ESTADOS = [
    ("rota", "Se vende (12 meses de inventario o menos)", "Se vende", "#027266"),
    ("lento", "Lenta (más de 12 meses de inventario)", "Lento", "#D9A441"),
    ("sin_venta", "Sin venta en 12 meses (con inventario)", "Sin venta 12m", "#FE6412"),
    ("nuevo", "Nuevo – sin evaluar (menos de 3 meses)", "Nuevo", "#2F6DB5"),
    ("agotado", "Agotado (se vende, sin existencia)", "Agotado", "#8A8984"),
]
COLOR = {k: c for k, _, _, c in ESTADOS}
CORTO = {k: c for k, _, c, _ in ESTADOS}

NO_RECOMPRAR = "No recomprar"
NO_RECOMPRAR_12 = "No recomprar hasta bajar de 12 meses"
REVISAR_NUEVO = "Revisar (nuevo)"
REPONER = "Reponer"


# ---------- fechas y meses ----------
def restar_meses(d, n):
    """Día 1 del mes que está `n` meses antes de `d` (n negativo = después)."""
    m = d.month - 1 - n
    return date(d.year + m // 12, m % 12 + 1, 1)


def lista_meses(hoy, n):
    """Los `n` meses completos antes del mes de `hoy`, del más antiguo al más reciente, como 'AAAA-MM'."""
    return [f"{restar_meses(hoy, k).year}-{restar_meses(hoy, k).month:02d}" for k in range(n, 0, -1)]


def etiqueta_mes(mes):
    return f"{MESES[int(mes[5:]) - 1]} {mes[2:4]}"


def meses_entre(desde, hasta):
    """Meses (con decimales) entre dos fechas; None si falta la primera."""
    if not desde:
        return None
    return (hasta - desde).days / DIAS_MES


# ---------- venta mensual y clasificación ----------
def venta_mensual(venta_12m, meses_desde_ingreso):
    """Venta promedio por mes de los últimos 12 meses completos.
    Un producto que ingresó hace menos de 12 meses se divide entre los meses que lleva (mínimo 1), no entre 12."""
    if meses_desde_ingreso is None:
        divisor = 12
    else:
        divisor = max(1.0, min(12.0, meses_desde_ingreso))
    return venta_12m / divisor if venta_12m > 0 else 0.0


def es_nuevo(meses_en_inventario):
    return meses_en_inventario is not None and meses_en_inventario < MESES_NUEVO


def cobertura(disponible, mensual):
    """Meses que alcanza el inventario disponible con la venta mensual; None sin venta."""
    return disponible / mensual if mensual > 0 else None


def plazo_txt(meses):
    if meses is None:
        return ""
    if meses > 120:
        return "más de 10 años"
    if meses > 24:
        return f"{meses / 12:.1f} años"
    return f"{meses:.1f} meses" if meses < 10 else f"{meses:.0f} meses"


def clasificar(meses_en_inventario, dias_sin_venta, ultima_venta, disponible, mensual):
    """(estado, comentario) de un producto.
    Orden: nuevo → sin venta en 12 meses → agotado → lento → se vende."""
    if es_nuevo(meses_en_inventario):
        return "nuevo", f"Nuevo – sin evaluar: {meses_en_inventario:.1f} meses desde su primer ingreso"
    if ultima_venta is None:
        return "sin_venta", "Sin venta: nunca se ha vendido"
    if dias_sin_venta is not None and dias_sin_venta > 365:
        return "sin_venta", f"Sin venta en 12 meses: última el {ultima_venta:%d/%m/%Y} (hace {plazo_txt(dias_sin_venta / DIAS_MES)})"
    if disponible <= 0:
        return "agotado", "Agotado: se vende, revisar reposición"
    if mensual <= 0:
        return "lento", f"Lento: sin venta en los últimos 12 meses completos (última el {ultima_venta:%d/%m/%Y})"
    cob = cobertura(disponible, mensual)
    if cob > COBERTURA_LENTA:
        return "lento", f"Lento: inventario para {plazo_txt(cob)}"
    return "rota", f"Se vende: inventario para {plazo_txt(cob)}"


def recomendacion(estado, disponible, cob, venta_12m):
    """Qué hacer con el producto: no volver a comprar, reponer o nada (None)."""
    sin_venta_12 = venta_12m <= 0
    if disponible > 0 and (sin_venta_12 or (cob is not None and cob > COBERTURA_NO_RECOMPRAR)):
        if estado == "nuevo":
            return REVISAR_NUEVO
        return NO_RECOMPRAR if sin_venta_12 else NO_RECOMPRAR_12
    if not sin_venta_12 and (disponible <= 0 or (cob is not None and cob < COBERTURA_REPONER)):
        return REPONER
    return None


# ---------- rentabilidad ----------
def rotacion(costo_anual, inventario):
    """Veces al año que se vende el inventario a costo."""
    return costo_anual / inventario if inventario > 0 else None


def gmroi(utilidad_anual, inventario):
    """Utilidad bruta al año por cada C$ 1 de inventario a costo."""
    return utilidad_anual / inventario if inventario > 0 else None


def rentabilidad(venta, costo, inventario, dias, venta_12m, costo_12m):
    """Rentabilidad histórica (desde la fecha, anualizada con `dias`) y de los últimos 12 meses.
    El veredicto y los meses de inventario usan la de 12 meses."""
    anual = 365 / dias if dias else 0
    utilidad, utilidad_12 = venta - costo, venta_12m - costo_12m
    r = {"utilidad": utilidad, "margen": utilidad / venta * 100 if venta else None,
         "utilidad_12": utilidad_12, "margen_12": utilidad_12 / venta_12m * 100 if venta_12m else None,
         "rotacion_hist": rotacion(costo * anual, inventario), "gmroi_hist": gmroi(utilidad * anual, inventario),
         "rotacion_12": rotacion(costo_12m, inventario), "gmroi_12": gmroi(utilidad_12, inventario)}
    r["rotacion"], r["gmroi"] = r["rotacion_12"], r["gmroi_12"]
    r["meses_inventario"] = 12 / r["rotacion_12"] if r["rotacion_12"] else None
    g = r["gmroi_12"]
    if venta_12m <= 0:
        if inventario > 0:
            r["veredicto"], r["nivel"] = "Sin ventas en los últimos 12 meses: el inventario no está generando utilidad", "malo"
        else:
            r["veredicto"], r["nivel"] = "Sin ventas ni inventario en los últimos 12 meses", ""
    elif g is None:
        r["veredicto"], r["nivel"] = "Se vende sin inventario disponible: rentable si se repone", "bueno"
    elif g >= GMROI_BUENO:
        r["veredicto"], r["nivel"] = f"Rentable: cada C$ 1 en inventario deja C$ {g:.2f} de utilidad al año", "bueno"
    elif g >= GMROI_ACEPTABLE:
        r["veredicto"], r["nivel"] = f"Aceptable: deja C$ {g:.2f} al año por cada C$ 1 en inventario; mejorar la rotación", "medio"
    else:
        r["veredicto"], r["nivel"] = f"Poco rentable: solo C$ {g:.2f} al año por cada C$ 1 en inventario (inventario lento)", "malo"
    return r


# ---------- tendencia ----------
def cambio_pct(actual, anterior):
    """% de cambio; None si no hay base para comparar."""
    if anterior and anterior > 0:
        return (actual / anterior - 1) * 100
    return None


def nivel_tendencia(pct, actual=0):
    """'sube' (> +10%), 'igual', 'baja' (< −10%); 'nueva' si antes no había venta y ahora sí."""
    if pct is None:
        return "nueva" if actual > 0 else ""
    if pct > UMBRAL_TENDENCIA:
        return "sube"
    if pct < -UMBRAL_TENDENCIA:
        return "baja"
    return "igual"


def promedio_movil(serie, n=3):
    return [None if i < n - 1 else sum(serie[i - n + 1:i + 1]) / n for i in range(len(serie))]


def lento_desde(serie, primero=12, umbral=UMBRAL_LENTO_DESDE, minimo=3):
    """Índice del primer mes (desde `primero`) en que el promedio móvil de 3 meses queda por debajo del `umbral`
    del promedio de los 12 meses anteriores a ese mes y ya no se recupera hasta el final de la serie. None si no pasa.
    Tiene que sostenerse al menos `minimo` meses (una caída de solo los últimos 1–2 meses todavía no cuenta).
    `serie`: venta mensual de meses completos, del más antiguo al más reciente."""
    pm = promedio_movil(serie)
    for i in range(max(primero, 12), len(serie) - minimo + 1):
        ref = sum(serie[i - 12:i]) / 12
        if ref <= 0:
            continue
        if all(pm[j] < umbral * ref for j in range(i, len(serie))):
            return i
    return None


def mediana_fecha(fechas):
    fechas = sorted(x for x in fechas if x)
    return fechas[(len(fechas) - 1) // 2] if fechas else None


def dejar_de_comprar(cambio, pct_lento, venta_12m=None):
    """Marca que conviene dejar de comprar: más del 85% de su inventario está lento o sin venta, y su venta cae más de 30%
    o no vendió nada en los últimos 12 meses."""
    if pct_lento is None or pct_lento <= PCT_LENTO_DEJAR:
        return False
    return (cambio is not None and cambio < CAIDA_DEJAR) or (venta_12m is not None and venta_12m <= 0)


# ---------- temporada ----------
def temporada(por_mes_anio):
    """{1..12: venta} → (texto «Abr–Jun», % de la venta del año, mes inicial) con los 3 meses seguidos que más venden."""
    total = sum(por_mes_anio.values())
    if total < MIN_UNIDADES_TEMPORADA:
        return None, None, None
    mejor, inicio = -1, 1
    for m in range(1, 13):
        suma = sum(por_mes_anio.get((m + k - 1) % 12 + 1, 0) for k in range(3))
        if suma > mejor:
            mejor, inicio = suma, m
    fin = (inicio + 1) % 12 + 1
    return f"{MESES[inicio - 1]}–{MESES[fin - 1]}", mejor / total * 100, inicio


def meses_temporada(inicio):
    return [] if not inicio else [(inicio + k - 1) % 12 + 1 for k in range(3)]


# ---------- formato ----------
def abreviar(v, prefijo="C$ "):
    """91,425,832 → «C$ 91.4 M»; 850,000 → «C$ 850 K»."""
    if v is None:
        return "—"
    a = abs(v)
    if a >= 1e6:
        return f"{prefijo}{v / 1e6:,.1f} M"
    if a >= 1e3:
        return f"{prefijo}{v / 1e3:,.0f} K"
    return f"{prefijo}{v:,.0f}"
