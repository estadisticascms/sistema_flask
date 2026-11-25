"""Entregas · parte 5: alerta de tiempo (sección 4.8 de la especificación).

    tiempo_necesario = minutos_preparacion + minutos_traslado_bodega (si alguna línea requiere traslado)
                     + minutos_facturacion + minutos_empaque + transportista.minutos_traslado
    margen = hora_requerida - ahora - tiempo_necesario

- Rojo (en riesgo): margen < 0, o la orden se envió después de hora_corte y hora_requerida - ahora < minutos_urgencia.
- Amarillo (justo): margen < margen_amarillo.
- Verde: en otro caso.

Avisa pero no bloquea: en rojo el vendedor confirma la hora (riesgo_confirmado = 1) o la cambia.

Los estimados [PENDIENTE] (minutos de preparación, traslado entre bodegas, facturación, empaque y el margen para amarillo) los da la
empresa en Catálogos › Configuración. Mientras falte alguno, el cálculo lo toma como 0 y la alerta lo dice («faltan estimados»);
sin margen para amarillo, la alerta solo distingue rojo y verde. No se inventan valores.

Al recalcular un envío ya en marcha, los pasos que ya terminaron no se vuelven a contar (por ejemplo, después de facturar ya no
se suman los minutos de facturación). La regla de la hora de corte usa la hora en que se envió la orden.

Todas las horas son de Nicaragua (UTC-6), sin depender de la zona del servidor.
"""
import unicodedata
from datetime import datetime, time, timedelta, timezone

from router.entregas_formato import hora12

NICARAGUA = timezone(timedelta(hours=-6))
COMPONENTES = [   # clave de configuración (o del transportista), nombre, cuándo aplica
    ("minutos_preparacion", "Preparación en bodega", "preparacion"),
    ("minutos_traslado_bodega", "Traslado entre bodegas", "traslado"),
    ("minutos_facturacion", "Facturación", "facturacion"),
    ("minutos_empaque", "Empaque", "empaque"),
    ("transportista", "Traslado al transporte", "transporte"),
]
NOMBRE_CONFIG = {"minutos_preparacion": "minutos de preparación", "minutos_traslado_bodega": "minutos de traslado entre bodegas",
                 "minutos_facturacion": "minutos de facturación", "minutos_empaque": "minutos de empaque",
                 "margen_amarillo": "margen para amarillo", "transportista": "minutos de traslado del transportista"}
LETRA_DIA = "LMXJVSD"     # lunes = 0
NOMBRE_DIA = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]
# Estados en que ya terminó cada paso (para recalcular un envío en marcha)
DESPUES_DE_FACTURAR = {"FACTURADA", "LISTA_PARA_RETIRO", "RETIRADA", "EMPACADA"}
DESPUES_DE_PREPARAR = {"LISTA_PARA_RETIRO", "RETIRADA", "EMPACADA"}
SIN_ALERTA = {"EN_RUTA", "ENTREGADA", "ANULADA"}


def ahora_nicaragua():
    return datetime.now(NICARAGUA).replace(tzinfo=None)


def a_entero(valor):
    try:
        return int(str(valor).strip())
    except (TypeError, ValueError):
        return None


def a_hora(valor):
    """'15:00', time o timedelta (TIME de MySQL) -> time, o None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, time):
        return valor
    if isinstance(valor, timedelta):
        segundos = int(valor.total_seconds()) % 86400
        return time(segundos // 3600, segundos % 3600 // 60)
    try:
        h, m = str(valor).split(":")[:2]
        return time(int(h), int(m))
    except (ValueError, TypeError):
        return None


def leer_configuracion(cursor):
    cursor.execute("SELECT clave, valor FROM ent_configuracion")
    return {fila[0] if isinstance(fila, tuple) else fila["clave"]: fila[1] if isinstance(fila, tuple) else fila["valor"]
            for fila in cursor.fetchall()}


def sin_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFKD", (texto or "").strip().upper()) if not unicodedata.combining(c))


def duracion(minutos):
    """Minutos -> «1 h 05 min», «35 min»."""
    m = abs(int(minutos))
    return f"{m // 60} h {m % 60:02d} min" if m >= 60 else f"{m} min"


def proxima_salida(horarios, destino, desde):
    """Primera salida del transportista (para el destino, si tiene horarios para él) que ocurre a partir de «desde».
    -> (datetime, horario) o (None, None). Busca en los próximos 7 días."""
    if not horarios:
        return None, None
    propios = [h for h in horarios if destino and sin_tildes(h["destino"]) == sin_tildes(destino)]
    candidatos = propios or horarios
    mejor = (None, None)
    for d in range(8):
        dia = desde.date() + timedelta(days=d)
        letra = LETRA_DIA[dia.weekday()]
        for h in candidatos:
            dias = [x.strip().upper() for x in (h.get("dias") or "").split(",") if x.strip()]
            if dias and letra not in dias:
                continue
            hora = a_hora(h["hora_salida"])
            if hora is None:
                continue
            salida = datetime.combine(dia, hora)
            if salida >= desde and (mejor[0] is None or salida < mejor[0]):
                mejor = (salida, h)
        if mejor[0]:
            return mejor
    return mejor


def calcular(config, hora_requerida, hay_traslado, transportista=None, horarios=None, destino=None, ahora=None, enviada_en=None,
             hechos=frozenset()):
    """-> dict con nivel ('verde', 'amarillo', 'rojo' o None si no aplica), margen y necesario (minutos), desglose, motivos,
    faltan (estimados que no se han definido), sugerencia de salida y aviso de la hora límite de recepción.
    hechos: pasos ya terminados ('preparacion', 'traslado', 'facturacion', 'empaque') que no se vuelven a contar."""
    ahora = ahora or ahora_nicaragua()
    enviada_en = enviada_en or ahora
    transportista = transportista or {}
    desglose, faltan = [], []
    for clave, nombre, paso in COMPONENTES:
        if paso == "traslado" and not hay_traslado:
            continue
        if paso in hechos:
            desglose.append({"nombre": nombre, "minutos": 0, "hecho": True})
            continue
        valor = a_entero(transportista.get("minutos_traslado")) if clave == "transportista" else a_entero(config.get(clave))
        if valor is None:
            faltan.append(NOMBRE_CONFIG[clave])
        desglose.append({"nombre": nombre, "minutos": valor or 0, "falta": valor is None})
    necesario = sum(d["minutos"] for d in desglose)
    r = {"nivel": None, "necesario": necesario, "necesario_texto": duracion(necesario), "margen_texto": None, "desglose": desglose, "faltan": faltan, "motivos": [], "margen": None,
         "listo": ahora + timedelta(minutes=necesario), "sugerencia": None, "aviso_limite": None, "sin_amarillo": False}

    # Próxima salida y hora límite de recepción del transportista (se informan aunque falte la hora requerida)
    salida, horario = proxima_salida(horarios or [], destino, r["listo"])
    if salida:
        r["sugerencia"] = {"hora": salida, "destino": horario["destino"],
                           "texto": f"Próxima salida que se alcanza: {horario['destino']} {NOMBRE_DIA[salida.weekday()]} "
                                    f"{hora12(salida, '%d/%m')}"}
    limite = a_hora(transportista.get("hora_limite_recepcion"))
    if limite:
        tope = datetime.combine(ahora.date(), limite)
        if r["listo"] > tope:
            r["aviso_limite"] = (f"Ya no se llega hoy a la hora límite de recepción del transporte ({hora12(limite)}): "
                                 f"la mercadería estaría lista a las {hora12(r['listo'])}.")
        elif hora_requerida and hora_requerida.time() > limite and hora_requerida.date() == ahora.date():
            r["aviso_limite"] = (f"La hora requerida ({hora12(hora_requerida)}) es después de la hora límite de recepción del "
                                 f"transporte ({hora12(limite)}).")

    if not hora_requerida:
        return r
    restante = int((hora_requerida - ahora).total_seconds() // 60)
    r["margen"] = restante - necesario
    r["margen_texto"] = duracion(r["margen"])
    corte, urgencia = a_hora(config.get("hora_corte")), a_entero(config.get("minutos_urgencia"))
    amarillo = a_entero(config.get("margen_amarillo"))
    if r["margen"] < 0:
        r["motivos"].append(f"No alcanza el tiempo: se necesitan {duracion(necesario)} y faltan {duracion(max(restante, 0))}"
                            + (" (la hora requerida ya pasó)" if restante < 0 else "") + ".")
    if corte and urgencia is not None and enviada_en.time() > corte and restante < urgencia:
        r["motivos"].append(f"Orden enviada después de las {hora12(corte)} con menos de {duracion(urgencia)} para la hora requerida.")
    if r["motivos"]:
        r["nivel"] = "rojo"
    elif amarillo is None:
        r["nivel"], r["sin_amarillo"] = "verde", True
        faltan.append(NOMBRE_CONFIG["margen_amarillo"])
    elif r["margen"] < amarillo:
        r["nivel"] = "amarillo"
        r["motivos"].append(f"Justo: quedan {duracion(r['margen'])} de margen (menos de {duracion(amarillo)}).")
    else:
        r["nivel"] = "verde"
    return r


def pasos_hechos(envio, lineas):
    """Pasos ya terminados de un envío en marcha, para no volver a sumarlos al recalcular."""
    hechos = set()
    if envio["estado"] in DESPUES_DE_FACTURAR:
        hechos.add("facturacion")
    activas = [l for l in lineas if l["existencia"] != "QUITADA"]
    if envio["estado"] in DESPUES_DE_PREPARAR or (activas and all(l["preparacion"] in ("LISTA", "RETIRADA") for l in activas)):
        hechos.add("preparacion")
    con_traslado = [l for l in activas if l["traslado_estado"] != "NO_REQUIERE"]
    if con_traslado and all(l["traslado_estado"] == "RECIBIDO" for l in con_traslado):
        hechos.add("traslado")
    if envio["estado"] == "EMPACADA":
        hechos.add("empaque")
    return frozenset(hechos)


def alerta_de_envio(cursor, envio, lineas):
    """Alerta recalculada ahora para un envío guardado (None si ya salió, se entregó o se anuló)."""
    if envio["estado"] in SIN_ALERTA:
        return None
    cursor.execute("SELECT minutos_traslado, hora_limite_recepcion FROM ent_transportistas WHERE id = %s", (envio["transportista_id"],))
    transportista = cursor.fetchone()
    cursor.execute("SELECT destino, hora_salida, dias FROM ent_transportista_horarios WHERE transportista_id = %s", (envio["transportista_id"],))
    horarios = cursor.fetchall()
    hay_traslado = any(l["traslado_estado"] != "NO_REQUIERE" and l["existencia"] != "QUITADA" for l in lineas)
    return calcular(leer_configuracion(cursor), envio["hora_requerida"], hay_traslado, transportista, horarios, envio["destino"],
                    enviada_en=envio["creado_en"], hechos=pasos_hechos(envio, lineas))


def alertas_de_envios(cursor, envios, lineas_por_envio):
    """Alertas recalculadas para varios envíos con tres consultas en total (tableros y colas). -> {envio_id: alerta o None}"""
    config = leer_configuracion(cursor)
    ids_t = sorted({e["transportista_id"] for e in envios})
    transportistas, horarios = {}, {}
    if ids_t:
        marcas = ", ".join(["%s"] * len(ids_t))
        cursor.execute(f"SELECT id, minutos_traslado, hora_limite_recepcion FROM ent_transportistas WHERE id IN ({marcas})", ids_t)
        transportistas = {t["id"]: t for t in cursor.fetchall()}
        cursor.execute(f"SELECT transportista_id, destino, hora_salida, dias FROM ent_transportista_horarios WHERE transportista_id IN ({marcas})", ids_t)
        for h in cursor.fetchall():
            horarios.setdefault(h["transportista_id"], []).append(h)
    resultado = {}
    for e in envios:
        lineas = lineas_por_envio.get(e["id"], [])
        if e["estado"] in SIN_ALERTA:
            resultado[e["id"]] = None
            continue
        hay = any(l["traslado_estado"] != "NO_REQUIERE" and l["existencia"] != "QUITADA" for l in lineas)
        resultado[e["id"]] = calcular(config, e["hora_requerida"], hay, transportistas.get(e["transportista_id"]),
                                      horarios.get(e["transportista_id"], []), e["destino"], enviada_en=e["creado_en"],
                                      hechos=pasos_hechos(e, lineas))
    return resultado
