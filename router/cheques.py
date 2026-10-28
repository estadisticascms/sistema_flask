"""Emisión e impresión de cheques.

- Cada cuenta bancaria (cheq_bancos) lleva su propio consecutivo de cheques (proximo_cheque). Al guardar con
  numeración automática se toma ese número y se incrementa dentro de la misma transacción (sin duplicados);
  con numeración manual se valida que el número no exista para ese banco.
- La partida contable (cheq_partidas) es de partida doble: cada línea va al Debe o al Haber y los totales deben cuadrar.
  La cuenta del banco debe quedar al Haber por el monto del cheque (el formulario la propone sola).
- La partida siempre va en córdobas: en cheques en dólares el banco va al Haber por monto × tasa de cambio.
- Catálogos editables por el usuario: bancos (cheq_bancos) y cuentas contables (cheq_cuentas).
- Un cheque impreso se puede editar indicando el motivo. Todo queda en la bitácora (cheq_bitacora): quién, cuándo, desde qué
  equipo e IP y qué cambió (campo, antes y después). Solo los anulados ya no se modifican."""
import csv
import io
import json
import socket
from concurrent.futures import ThreadPoolExecutor, TimeoutError as Tiempo
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from flask import Blueprint, flash, jsonify, redirect, render_template, request, session, url_for

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.bitacora import ip_cliente

cheques_bp = Blueprint("cheques", __name__)

EMPRESA = {"nombre": "COMERCIAL LA ESPERANZA",
           "direccion": "Cementerio del B Monseñor Lezcano 2C Arriba 75 vrs al lago",
           "telefono": "Teléfono: 22266442, Fax: 22660386", "ruc": "No RUC: J0310000115940", "ciudad": "MANAGUA"}
TASA_DEFECTO = Decimal("36.6243")
MONEDAS = {"C$": ("CÓRDOBA", "CÓRDOBAS"), "U$": ("DÓLAR", "DÓLARES")}
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
POR_PAGINA = 50


# -----------------------------
# Monto en letras
# -----------------------------
_UNIDADES = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE",
             "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDÓS",
             "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
_DECENAS = ["", "", "", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS",
             "OCHOCIENTOS", "NOVECIENTOS"]


def _hasta_999(n, apocope):
    """0..999 en letras. apocope: 'UN' en vez de 'UNO' (antes de MIL, MILLONES o la moneda)."""
    if n == 0:
        return ""
    if n == 100:
        return "CIEN"
    c, resto = divmod(n, 100)
    partes = [_CENTENAS[c]] if c else []
    if resto:
        if resto < 30:
            palabra = _UNIDADES[resto]
            if apocope and resto in (1, 21):
                palabra = "UN" if resto == 1 else "VEINTIÚN"
        else:
            d, u = divmod(resto, 10)
            palabra = _DECENAS[d] + ((" Y " + ("UN" if apocope and u == 1 else _UNIDADES[u])) if u else "")
        partes.append(palabra)
    return " ".join(partes)


def entero_a_letras(n):
    if n == 0:
        return "CERO"
    millones, resto = divmod(n, 1_000_000)
    miles, unidades = divmod(resto, 1000)
    partes = []
    if millones:
        partes.append("UN MILLÓN" if millones == 1 else f"{entero_a_letras_simple(millones)} MILLONES")
    if miles:
        partes.append("MIL" if miles == 1 else f"{_hasta_999(miles, True)} MIL")
    if unidades:
        partes.append(_hasta_999(unidades, True))
    return " ".join(partes)


def entero_a_letras_simple(n):
    """Para la parte de millones (hasta 999,999 millones)."""
    miles, unidades = divmod(n, 1000)
    partes = []
    if miles:
        partes.append("MIL" if miles == 1 else f"{_hasta_999(miles, True)} MIL")
    if unidades:
        partes.append(_hasta_999(unidades, True))
    return " ".join(partes)


def monto_en_letras(monto, moneda="C$"):
    """100500 -> 'CIEN MIL QUINIENTOS CÓRDOBAS NETOS'; 1250.50 -> 'MIL DOSCIENTOS CINCUENTA CÓRDOBAS CON 50/100'."""
    monto = Decimal(monto).quantize(Decimal("0.01"), ROUND_HALF_UP)
    entero = int(monto)
    centavos = int((monto - entero) * 100)
    singular, plural = MONEDAS.get(moneda, MONEDAS["C$"])
    texto = f"{entero_a_letras(entero)} {singular if entero == 1 else plural}"
    if entero % 1_000_000 == 0 and entero:
        texto = texto.replace(" MILLÓN ", " MILLÓN DE ").replace(" MILLONES ", " MILLONES DE ")
    return f"{texto} CON {centavos:02d}/100" if centavos else f"{texto} {'NETO' if entero == 1 else 'NETOS'}"


# -----------------------------
# Utilidades
# -----------------------------
def a_decimal(valor):
    try:
        return Decimal(str(valor or "0").replace(",", "").strip() or "0").quantize(Decimal("0.01"), ROUND_HALF_UP)
    except InvalidOperation:
        raise ValueError(f"monto inválido '{valor}'")


def niveles(codigo):
    """Código contable -> (cta, scta, sscta) como en el formato del cheque: 10102014 -> 1010 / 2 / 4."""
    c = (codigo or "").replace("-", "")
    resto = c[7:] if len(c) >= 8 else c[5:]
    return c[:4], (c[4:5] or "0"), ((resto.lstrip("0") or "0") if resto else "0")


def descripcion_cuenta(nombre):
    return (nombre or "").replace(":", ",")


def usuario():
    return session.get("usuario") or ""


# -----------------------------
# Bitácora de cambios
# -----------------------------
_nombres_red = {}
_dns = ThreadPoolExecutor(max_workers=2)


def nombre_en_red(ip):
    """Nombre del equipo según la red (DNS inverso). En la oficina suele dar el nombre de la PC; se espera máximo 1.5 s."""
    if not ip:
        return None
    if ip not in _nombres_red:
        try:
            nombre = _dns.submit(socket.gethostbyaddr, ip).result(timeout=1.5)[0]
        except (Tiempo, OSError):
            nombre = None
        _nombres_red[ip] = nombre if nombre and nombre != ip else None
    return _nombres_red[ip]


def origen():
    """Desde dónde se hace el cambio. Se calcula antes de abrir la transacción (el DNS puede tardar)."""
    ip = ip_cliente()
    return {"ip": ip[:45], "equipo": " ".join((request.form.get("equipo") or "").split())[:100] or None,
            "equipo_red": (nombre_en_red(ip) or "")[:255] or None, "navegador": (request.headers.get("User-Agent") or "")[:255]}


def registrar_cambio(cursor, cheque_id, numero, accion, quien, estado_antes=None, motivo=None, cambios=None):
    cursor.execute("""INSERT INTO cheq_bitacora (cheque_id, numero, fecha_hora, usuario, accion, estado_antes, motivo, cambios,
                                                 ip, equipo, equipo_red, navegador)
                      VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                   (cheque_id, numero, datetime.now().replace(microsecond=0), usuario() or "?", accion, estado_antes,
                    (motivo or "")[:255] or None, json.dumps(cambios, ensure_ascii=False) if cambios else None,
                    quien["ip"], quien["equipo"], quien["equipo_red"], quien["navegador"]))


def leer_bitacora(cursor, cheque_id=None, where="", params=()):
    cursor.execute(f"""SELECT b.*, ch.beneficiario, ch.estado FROM cheq_bitacora b LEFT JOIN cheq_cheques ch ON ch.id = b.cheque_id
                       WHERE {'b.cheque_id = %s' if cheque_id else '1=1'} {where} ORDER BY b.fecha_hora DESC, b.id DESC LIMIT 1000""",
                   ((cheque_id,) if cheque_id else ()) + tuple(params))
    filas = cursor.fetchall()
    for f in filas:
        f["cambios"] = json.loads(f["cambios"]) if f["cambios"] else []
    return filas


def _dinero(moneda, v):
    return f"{moneda} {Decimal(v):,.2f}"


def _lineas_partida(lineas):
    """Cada línea de la partida como texto comparable: 'cuenta · descripción · Debe/Haber monto [ref]'."""
    textos = []
    for p in lineas:
        lado = f"Debe {Decimal(p['debe']):,.2f}" if p["debe"] else f"Haber {Decimal(p['haber']):,.2f}"
        desc = (p["descripcion"] or "").strip()
        desc = desc if len(desc) <= 60 else desc[:57] + "…"
        textos.append(" · ".join(x for x in (p["codigo"], desc, lado, p.get("referencia") or "") if x))
    return textos


def diferencias(anterior, nuevo, partida_antes, partida_nueva):
    """Lista de {campo, antes, despues} con lo que cambió."""
    cambios = []
    campos = [("Banco", anterior["banco_nombre"], nuevo["banco_nombre"]),
              ("Número", anterior["numero"], nuevo["numero"]),
              ("Fecha", anterior["fecha"].strftime("%d/%m/%Y"), nuevo["fecha"].strftime("%d/%m/%Y")),
              ("Beneficiario", anterior["beneficiario"], nuevo["beneficiario"]),
              ("Monto", _dinero(anterior["moneda"], anterior["monto"]), _dinero(nuevo["moneda"], nuevo["monto"])),
              ("Concepto", anterior["concepto"] or "", nuevo["concepto"] or "")]
    if "U$" in (anterior["moneda"], nuevo["moneda"]):
        campos.append(("Tasa de cambio", f"{Decimal(anterior['tasa_cambio'] or 0):.4f}", f"{Decimal(nuevo['tasa_cambio'] or 0):.4f}"))
    for campo, antes, despues in campos:
        if str(antes) != str(despues):
            cambios.append({"campo": campo, "antes": str(antes), "despues": str(despues)})
    viejas, nuevas = _lineas_partida(partida_antes), _lineas_partida(partida_nueva)
    quitadas = [x for x in viejas if x not in nuevas or viejas.count(x) > nuevas.count(x)]
    agregadas = [x for x in nuevas if x not in viejas or nuevas.count(x) > viejas.count(x)]
    if quitadas or agregadas:
        cambios.append({"campo": "Partida contable", "antes": "\n".join(dict.fromkeys(quitadas)) or "—",
                        "despues": "\n".join(dict.fromkeys(agregadas)) or "—"})
    return cambios


def bancos(cursor, solo_activos=False):
    cursor.execute(f"""
        SELECT b.*, c.codigo AS cuenta_codigo, c.nombre AS cuenta_nombre
        FROM cheq_bancos b LEFT JOIN cheq_cuentas c ON c.id = b.cuenta_id
        {'WHERE b.activo = 1' if solo_activos else ''} ORDER BY b.activo DESC, b.moneda, b.nombre
    """)
    return cursor.fetchall()


def cargar_cheque(cursor, id_):
    cursor.execute("""
        SELECT ch.*, b.nombre AS banco_nombre, b.banco, b.numero_cuenta, b.ajuste_x_mm, b.ajuste_y_mm
        FROM cheq_cheques ch JOIN cheq_bancos b ON b.id = ch.banco_id WHERE ch.id = %s
    """, (id_,))
    cheque = cursor.fetchone()
    if cheque:
        cursor.execute("""
            SELECT p.*, c.nombre AS cuenta_nombre FROM cheq_partidas p LEFT JOIN cheq_cuentas c ON c.id = p.cuenta_id
            WHERE p.cheque_id = %s ORDER BY p.orden, p.id
        """, (id_,))
        cheque["partidas"] = cursor.fetchall()
    return cheque


# -----------------------------
# Lista
# -----------------------------
@cheques_bp.route("/cheques")
@requiere_acceso("cheques")
def lista():
    f = {"banco": request.args.get("banco", ""), "estado": request.args.get("estado", ""), "q": (request.args.get("q") or "").strip(),
         "desde": request.args.get("desde") or date.today().replace(day=1).isoformat(),
         "hasta": request.args.get("hasta") or date.today().isoformat()}
    where, params = ["ch.fecha BETWEEN %s AND %s"], [f["desde"], f["hasta"]]
    if f["banco"].isdigit():
        where.append("ch.banco_id = %s"); params.append(int(f["banco"]))
    if f["estado"] in ("Emitido", "Impreso", "Anulado"):
        where.append("ch.estado = %s"); params.append(f["estado"])
    for palabra in f["q"].split():
        where.append("(ch.beneficiario LIKE %s OR ch.concepto LIKE %s OR CAST(ch.numero AS CHAR) = %s)")
        params += [f"%{palabra}%", f"%{palabra}%", palabra]
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(f"""
            SELECT ch.id, ch.numero, ch.numero_manual, ch.fecha, ch.beneficiario, ch.monto, ch.moneda, ch.concepto, ch.estado,
                   ch.veces_impreso, ch.creado_por, b.nombre AS banco_nombre
            FROM cheq_cheques ch JOIN cheq_bancos b ON b.id = ch.banco_id
            WHERE {' AND '.join(where)} ORDER BY ch.fecha DESC, ch.id DESC
        """, params)
        cheques = cursor.fetchall()
        lista_bancos = bancos(cursor)
    finally:
        cursor.close()
        conn.close()
    totales = {}
    for ch in cheques:
        if ch["estado"] != "Anulado":
            totales[ch["moneda"]] = totales.get(ch["moneda"], 0) + ch["monto"]
    return render_template("cheques_lista.html", cheques=cheques, bancos=lista_bancos, f=f, totales=totales)


# -----------------------------
# Crear / editar
# -----------------------------
def monto_cordobas(monto, moneda, tasa):
    """Monto del cheque en córdobas para la partida: los cheques en dólares se convierten con la tasa del cheque."""
    if moneda == "U$":
        return (monto * tasa).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return monto


def leer_formulario(form):
    """Valida el formulario. Devuelve (datos, partidas_usuario, errores)."""
    errores = []
    datos = {"banco_id": form.get("banco_id", ""), "numeracion": "manual" if form.get("numeracion") == "manual" else "auto",
             "numero_manual": (form.get("numero_manual") or "").strip(), "beneficiario": " ".join((form.get("beneficiario") or "").split()).upper(),
             "concepto": (form.get("concepto") or "").strip().upper(), "fecha": form.get("fecha") or ""}
    try:
        datos["fecha"] = datetime.strptime(datos["fecha"], "%Y-%m-%d").date()
    except ValueError:
        errores.append("Fecha inválida.")
    try:
        datos["monto"] = a_decimal(form.get("monto"))
        if datos["monto"] <= 0:
            errores.append("El monto debe ser mayor que cero.")
    except ValueError as e:
        errores.append(str(e)); datos["monto"] = Decimal(0)
    try:
        datos["tasa_cambio"] = Decimal(str(form.get("tasa_cambio") or TASA_DEFECTO).replace(",", ""))
        if datos["tasa_cambio"] <= 0:
            errores.append("La tasa de cambio debe ser mayor que cero."); datos["tasa_cambio"] = TASA_DEFECTO
    except InvalidOperation:
        errores.append("Tasa de cambio inválida."); datos["tasa_cambio"] = TASA_DEFECTO
    if not datos["beneficiario"]:
        errores.append("Indique el beneficiario.")
    if len(datos["beneficiario"]) > 150:
        errores.append("El beneficiario tiene más de 150 caracteres.")
    if not str(datos["banco_id"]).isdigit():
        errores.append("Seleccione el banco.")
    if datos["numeracion"] == "manual" and not datos["numero_manual"].isdigit():
        errores.append("Escriba el número de cheque (solo dígitos) o use la numeración automática.")

    partidas = []
    cuentas, descripciones = form.getlist("cuenta_id"), form.getlist("descripcion")
    debes, haberes, referencias = form.getlist("debe"), form.getlist("haber"), form.getlist("referencia")
    def en(lista, i):
        return (lista[i] if i < len(lista) else "") or ""
    for i, cuenta in enumerate(cuentas):
        debe_txt, haber_txt = en(debes, i).strip(), en(haberes, i).strip()
        if not cuenta and not debe_txt and not haber_txt:
            continue   # fila vacía
        try:
            debe = a_decimal(debe_txt) if debe_txt else Decimal(0)
            haber = a_decimal(haber_txt) if haber_txt else Decimal(0)
        except ValueError as e:
            errores.append(f"Línea {i + 1}: {e}"); continue
        if not cuenta.isdigit():
            errores.append(f"Línea {i + 1}: seleccione la cuenta contable."); continue
        if debe < 0 or haber < 0:
            # Un Debe negativo (p. ej. una retención que se resta) es un Haber, y al revés: se guarda en el lado correcto
            neto = debe - haber
            debe, haber = (neto, Decimal(0)) if neto > 0 else (Decimal(0), -neto)
        if debe and haber:
            errores.append(f"Línea {i + 1}: escriba el valor en Debe o en Haber, no en ambos."); continue
        if not debe and not haber:
            errores.append(f"Línea {i + 1}: escriba el valor en Debe o en Haber."); continue
        partidas.append({"cuenta_id": int(cuenta), "descripcion": en(descripciones, i).strip()[:255], "debe": debe, "haber": haber,
                         "referencia": en(referencias, i).strip()[:100] or None})
    if not partidas and not errores:
        errores.append("Agregue las líneas de la partida (Debe y Haber).")
    total_debe = sum(p["debe"] for p in partidas)
    total_haber = sum(p["haber"] for p in partidas)
    if partidas and total_debe != total_haber:
        errores.append(f"La partida no cuadra: Debe {total_debe:,.2f} y Haber {total_haber:,.2f}. Diferencia {total_debe - total_haber:,.2f}.")
    return datos, partidas, errores


def guardar(datos, partidas, id_=None, quien=None, motivo=None):
    """Inserta o actualiza el cheque y su partida en una transacción, y lo anota en la bitácora.
    Devuelve (id, numero) o lanza ValueError."""
    quien = quien or {"ip": None, "equipo": None, "equipo_red": None, "navegador": None}
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        conn.start_transaction()
        cursor.execute("SELECT * FROM cheq_bancos WHERE id = %s FOR UPDATE", (int(datos["banco_id"]),))
        banco = cursor.fetchone()
        if not banco:
            raise ValueError("El banco no existe.")
        if not banco["cuenta_id"]:
            raise ValueError(f"El banco {banco['nombre']} no tiene cuenta contable asignada (configúrela en Bancos).")
        anterior = cargar_cheque(cursor, id_) if id_ else None
        if anterior and anterior["estado"] == "Anulado":
            raise ValueError("El cheque está anulado; ya no se puede modificar.")
        if anterior and anterior["estado"] == "Impreso" and not (motivo or "").strip():
            raise ValueError("El cheque ya fue impreso: escriba el motivo del cambio.")

        # Número de cheque
        mismo_banco = anterior and anterior["banco_id"] == banco["id"]
        if datos["numeracion"] == "manual":
            numero, manual = int(datos["numero_manual"]), 1
        elif mismo_banco:
            numero, manual = anterior["numero"], anterior["numero_manual"]   # al editar se conserva el número
        else:
            if not banco["proximo_cheque"]:
                raise ValueError(f"Configure el próximo número de cheque de {banco['nombre']} en Bancos, o use numeración manual.")
            numero, manual = banco["proximo_cheque"], 0
        cursor.execute("SELECT id FROM cheq_cheques WHERE banco_id = %s AND numero = %s AND id <> %s", (banco["id"], numero, id_ or 0))
        if cursor.fetchone():
            raise ValueError(f"El cheque No. {numero} de {banco['nombre']} ya existe.")
        # El consecutivo avanza cuando se usa un número igual o mayor al próximo (automático o manual);
        # un manual menor (cheque viejo) no lo mueve
        if numero >= (banco["proximo_cheque"] or 0) and not (mismo_banco and numero == anterior["numero"]):
            cursor.execute("UPDATE cheq_bancos SET proximo_cheque = %s WHERE id = %s", (numero + 1, banco["id"]))

        valores = (banco["id"], numero, manual, datos["fecha"], datos["beneficiario"], datos["monto"], banco["moneda"],
                   datos["tasa_cambio"], monto_en_letras(datos["monto"], banco["moneda"]), datos["concepto"] or None)
        if anterior:
            cursor.execute("""UPDATE cheq_cheques SET banco_id=%s, numero=%s, numero_manual=%s, fecha=%s, beneficiario=%s, monto=%s,
                              moneda=%s, tasa_cambio=%s, monto_letras=%s, concepto=%s, actualizado_por=%s, actualizado_en=NOW()
                              WHERE id=%s""", valores + (usuario(), id_))
            cursor.execute("DELETE FROM cheq_partidas WHERE cheque_id = %s", (id_,))
        else:
            cursor.execute("""INSERT INTO cheq_cheques (banco_id, numero, numero_manual, fecha, beneficiario, monto, moneda, tasa_cambio,
                              monto_letras, concepto, creado_por) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", valores + (usuario(),))
            id_ = cursor.lastrowid

        # Partida doble: tal como la escribió el usuario; el banco tiene que quedar al Haber por el monto
        ids = sorted({p["cuenta_id"] for p in partidas} | {banco["cuenta_id"]})
        cursor.execute(f"SELECT id, codigo, nombre FROM cheq_cuentas WHERE id IN ({', '.join(['%s'] * len(ids))})", ids)
        cuentas = {r["id"]: r for r in cursor.fetchall()}
        faltan = [p for p in partidas if p["cuenta_id"] not in cuentas]
        if faltan:
            raise ValueError("Una de las cuentas seleccionadas ya no existe.")
        credito_banco = sum(p["haber"] - p["debe"] for p in partidas if p["cuenta_id"] == banco["cuenta_id"])
        esperado = monto_cordobas(datos["monto"], banco["moneda"], datos["tasa_cambio"])
        if credito_banco != esperado:
            cb = cuentas[banco["cuenta_id"]]
            detalle = (f"U$ {datos['monto']:,.2f} × {datos['tasa_cambio']} = C$ {esperado:,.2f}" if banco["moneda"] == "U$"
                       else f"C$ {esperado:,.2f}")
            raise ValueError(f"La cuenta del banco ({cb['codigo']} {descripcion_cuenta(cb['nombre'])}) debe ir al Haber por el monto del cheque "
                             f"en córdobas ({detalle}); en la partida tiene C$ {credito_banco:,.2f}.")
        lineas = sorted(partidas, key=lambda p: p["haber"] > 0)   # Debe primero, luego Haber
        for p in lineas:
            p["es_banco"] = 1 if p["cuenta_id"] == banco["cuenta_id"] and p["haber"] else 0
        for orden, p in enumerate(lineas, start=1):
            cuenta = cuentas[p["cuenta_id"]]
            cursor.execute("""INSERT INTO cheq_partidas (cheque_id, orden, cuenta_id, codigo, descripcion, debe, haber, referencia, es_banco)
                              VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                           (id_, orden, cuenta["id"], cuenta["codigo"], p["descripcion"] or descripcion_cuenta(cuenta["nombre"])[:255],
                            p["debe"], p["haber"], p["referencia"], p["es_banco"]))

        if anterior:
            nuevo = {"banco_nombre": banco["nombre"], "numero": numero, "fecha": datos["fecha"], "beneficiario": datos["beneficiario"],
                     "monto": datos["monto"], "moneda": banco["moneda"], "tasa_cambio": datos["tasa_cambio"], "concepto": datos["concepto"] or None}
            partida_nueva = [{"codigo": cuentas[p["cuenta_id"]]["codigo"], "debe": p["debe"], "haber": p["haber"], "referencia": p["referencia"],
                              "descripcion": p["descripcion"] or descripcion_cuenta(cuentas[p["cuenta_id"]]["nombre"])[:255]} for p in lineas]
            cambios = diferencias(anterior, nuevo, anterior["partidas"], partida_nueva)
            if cambios:
                registrar_cambio(cursor, id_, numero, "Editado", quien, anterior["estado"], motivo, cambios)
        else:
            registrar_cambio(cursor, id_, numero, "Creado", quien, None, None,
                             [{"campo": "Monto", "antes": "", "despues": _dinero(banco["moneda"], datos["monto"])},
                              {"campo": "Beneficiario", "antes": "", "despues": datos["beneficiario"]}])
        conn.commit()
        return id_, numero
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()


def valores_iniciales(form, cheque):
    """Lo que se muestra en el formulario: lo enviado (si hubo error), lo guardado (al editar) o valores por defecto."""
    campos = ("fecha", "banco_id", "beneficiario", "monto", "tasa_cambio", "concepto")
    if form:
        return {k: form.get(k, "") for k in campos}
    if cheque:
        return {"fecha": cheque["fecha"].isoformat(), "banco_id": str(cheque["banco_id"]), "beneficiario": cheque["beneficiario"],
                "monto": f"{cheque['monto']:.2f}", "tasa_cambio": f"{cheque['tasa_cambio'] or TASA_DEFECTO}", "concepto": cheque["concepto"] or ""}
    return {"fecha": date.today().isoformat(), "banco_id": "", "beneficiario": "", "monto": "", "tasa_cambio": f"{TASA_DEFECTO}", "concepto": ""}


def datos_formulario(cursor):
    cursor.execute("SELECT DISTINCT beneficiario FROM cheq_cheques ORDER BY beneficiario LIMIT 2000")
    beneficiarios = [r["beneficiario"] for r in cursor.fetchall()]
    lista = bancos(cursor, solo_activos=True)
    for b in lista:
        b["cuenta_texto"] = f"{b['cuenta_codigo']} {descripcion_cuenta(b['cuenta_nombre'])}" if b["cuenta_codigo"] else ""
    return {"bancos": lista, "beneficiarios": beneficiarios, "tasa_defecto": TASA_DEFECTO}


@cheques_bp.route("/cheques/nuevo", methods=["GET", "POST"])
@requiere_acceso("cheques")
def nuevo():
    return formulario(None)


@cheques_bp.route("/cheques/<int:id_>/editar", methods=["GET", "POST"])
@requiere_acceso("cheques")
def editar(id_):
    return formulario(id_)


def formulario(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        extra = datos_formulario(cursor)
        cheque = cargar_cheque(cursor, id_) if id_ else None
    finally:
        cursor.close()
        conn.close()
    if id_ and not cheque:
        flash("El cheque no existe.", "warning")
        return redirect(url_for("cheques.lista"))
    if cheque and cheque["estado"] == "Anulado":
        flash(f"El cheque No. {cheque['numero']} está anulado: solo se puede consultar.", "info")
        return redirect(url_for("cheques.imprimir", id_=id_))

    if request.method == "POST":
        datos, partidas, errores = leer_formulario(request.form)
        if not errores:
            try:
                motivo = " ".join((request.form.get("motivo_cambio") or "").split())
                nuevo_id, numero = guardar(datos, partidas, id_, origen(), motivo)
                if cheque and cheque["estado"] == "Impreso":
                    flash(f"✅ Cheque No. {numero} modificado y anotado en la bitácora. Ya estaba impreso: vuelva a imprimirlo.", "success")
                else:
                    flash(f"✅ Cheque No. {numero} guardado.", "success")
                return redirect(url_for("cheques.imprimir", id_=nuevo_id))
            except ValueError as e:
                errores.append(str(e))
            except Exception as e:
                errores.append(f"Error de base de datos: {e}")
        # Volver a mostrar lo que escribió el usuario
        lineas = [{"cuenta_id": c, "cuenta_texto": x, "descripcion": d, "debe": de, "haber": ha, "referencia": r}
                  for c, x, d, de, ha, r in zip(request.form.getlist("cuenta_id"), request.form.getlist("cuenta_texto"),
                                                request.form.getlist("descripcion"), request.form.getlist("debe"),
                                                request.form.getlist("haber"), request.form.getlist("referencia"))]
        return render_template("cheque_form.html", cheque=cheque, form=request.form, lineas=lineas, errores=errores,
                               val=valores_iniciales(request.form, cheque), **extra)

    lineas = []
    if cheque:
        lineas = [{"cuenta_id": str(p["cuenta_id"]), "cuenta_texto": f"{p['codigo']} {descripcion_cuenta(p['cuenta_nombre'])}",
                   "descripcion": p["descripcion"], "debe": f"{p['debe']:.2f}" if p["debe"] else "",
                   "haber": f"{p['haber']:.2f}" if p["haber"] else "", "referencia": p["referencia"] or ""} for p in cheque["partidas"]]
    return render_template("cheque_form.html", cheque=cheque, form=None, lineas=lineas, errores=[],
                           val=valores_iniciales(None, cheque), **extra)


@cheques_bp.route("/cheques/letras")
@requiere_acceso("cheques")
def letras():
    try:
        return jsonify(texto=monto_en_letras(a_decimal(request.args.get("monto")), request.args.get("moneda", "C$")))
    except (ValueError, InvalidOperation):
        return jsonify(texto="")


@cheques_bp.route("/cheques/cuentas/buscar")
@requiere_acceso("cheques")
def buscar_cuentas():
    q = (request.args.get("q") or "").strip()
    if not q:
        return jsonify([])
    palabras = q.split()[:5]
    where = " AND ".join(["(codigo LIKE %s OR nombre LIKE %s)"] * len(palabras))
    params = [x for p in palabras for x in (f"{p}%", f"%{p}%")]
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute(f"SELECT id, codigo, nombre, cta, scta, sscta FROM cheq_cuentas WHERE activa = 1 AND {where} "
                       f"ORDER BY codigo LIKE %s DESC, LENGTH(codigo), codigo LIMIT 20", params + [f"{palabras[0]}%"])
        filas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    return jsonify([{**r, "descripcion": descripcion_cuenta(r["nombre"])} for r in filas])


# -----------------------------
# Imprimir, anular
# -----------------------------
@cheques_bp.route("/cheques/<int:id_>")
@cheques_bp.route("/cheques/<int:id_>/imprimir")
@requiere_acceso("cheques")
def imprimir(id_):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cheque = cargar_cheque(cursor, id_)
        historial = leer_bitacora(cursor, id_) if cheque else []
    finally:
        cursor.close()
        conn.close()
    if not cheque:
        flash("El cheque no existe.", "warning")
        return redirect(url_for("cheques.lista"))
    # Editado después de la última impresión: lo impreso ya no coincide con lo guardado
    ultimo = next((h["accion"] for h in historial if h["accion"] in ("Editado", "Impreso")), None)
    cheque["reimprimir"] = ultimo == "Editado" and bool(cheque["veces_impreso"])
    for p in cheque["partidas"]:
        p["cta"], p["scta"], p["sscta"] = niveles(p["codigo"])
    return render_template("cheque_imprimir.html", c=cheque, empresa=EMPRESA, mes=MESES[cheque["fecha"].month - 1],
                           ahora=datetime.now(), total_debe=sum(p["debe"] for p in cheque["partidas"]),
                           total_haber=sum(p["haber"] for p in cheque["partidas"]),
                           total_haber_banco=sum(p["haber"] for p in cheque["partidas"] if p["es_banco"]), historial=historial)


@cheques_bp.route("/cheques/<int:id_>/impreso", methods=["POST"])
@requiere_acceso("cheques")
def marcar_impreso(id_):
    quien = origen()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT numero, estado, veces_impreso FROM cheq_cheques WHERE id = %s", (id_,))
        antes = cursor.fetchone()
        cursor.execute("UPDATE cheq_cheques SET veces_impreso = veces_impreso + 1, impreso_en = NOW(), "
                       "estado = IF(estado = 'Anulado', estado, 'Impreso') WHERE id = %s", (id_,))
        if antes:
            modo = "solo cheque" if request.args.get("modo") == "cheque" else "completo"
            registrar_cambio(cursor, id_, antes["numero"], "Impreso", quien, antes["estado"],
                             f"{'Reimpresión' if antes['veces_impreso'] else 'Primera impresión'} ({modo})")
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    return jsonify(ok=True)


@cheques_bp.route("/cheques/<int:id_>/anular", methods=["POST"])
@requiere_acceso("cheques")
def anular(id_):
    motivo = (request.form.get("motivo") or "").strip()
    if not motivo:
        flash("Indique el motivo de la anulación.", "warning")
        return redirect(url_for("cheques.imprimir", id_=id_))
    quien = origen()
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT numero, estado FROM cheq_cheques WHERE id = %s", (id_,))
        antes = cursor.fetchone()
        cursor.execute("UPDATE cheq_cheques SET estado = 'Anulado', motivo_anulacion = %s, actualizado_por = %s, actualizado_en = NOW() "
                       "WHERE id = %s AND estado <> 'Anulado'", (motivo[:255], usuario(), id_))
        ok = cursor.rowcount
        if ok:
            registrar_cambio(cursor, id_, antes["numero"], "Anulado", quien, antes["estado"], motivo,
                             [{"campo": "Estado", "antes": antes["estado"], "despues": "Anulado"}])
        conn.commit()
    finally:
        cursor.close()
        conn.close()
    flash("🚫 Cheque anulado. El número queda registrado como anulado y no se vuelve a usar." if ok else "El cheque ya estaba anulado.",
          "success" if ok else "info")
    return redirect(url_for("cheques.imprimir", id_=id_))


@cheques_bp.route("/cheques/bitacora")
@requiere_acceso("cheques")
def bitacora():
    f = {"desde": request.args.get("desde") or date.today().replace(day=1).isoformat(), "hasta": request.args.get("hasta") or date.today().isoformat(),
         "usuario": request.args.get("usuario", ""), "accion": request.args.get("accion", ""), "q": (request.args.get("q") or "").strip()}
    where, params = " AND b.fecha_hora >= %s AND b.fecha_hora < DATE_ADD(%s, INTERVAL 1 DAY)", [f["desde"], f["hasta"]]
    if f["usuario"]:
        where += " AND b.usuario = %s"; params.append(f["usuario"])
    if f["accion"] in ACCIONES:
        where += " AND b.accion = %s"; params.append(f["accion"])
    if f["q"]:
        where += " AND (b.numero = %s OR ch.beneficiario LIKE %s OR b.motivo LIKE %s OR b.equipo LIKE %s OR b.ip LIKE %s)"
        params += [f["q"] if f["q"].isdigit() else -1, f"%{f['q']}%", f"%{f['q']}%", f"%{f['q']}%", f"{f['q']}%"]
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        filas = leer_bitacora(cursor, None, where, params)
        cursor.execute("SELECT DISTINCT usuario FROM cheq_bitacora ORDER BY usuario")
        usuarios = [r["usuario"] for r in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()
    return render_template("cheques_bitacora.html", filas=filas, f=f, usuarios=usuarios, acciones=ACCIONES)


ACCIONES = ["Creado", "Editado", "Impreso", "Anulado"]


# -----------------------------
# Bancos
# -----------------------------
@cheques_bp.route("/cheques/bancos", methods=["GET", "POST"])
@requiere_acceso("cheques_config")
def bancos_vista():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if request.method == "POST":
            f = request.form
            id_ = f.get("id") or None
            errores = []
            nombre = " ".join((f.get("nombre") or "").split()).upper()
            banco = " ".join((f.get("banco") or "").split()).upper()
            numero_cuenta = (f.get("numero_cuenta") or "").strip()
            moneda = f.get("moneda") if f.get("moneda") in MONEDAS else "C$"
            cuenta_id = int(f["cuenta_id"]) if (f.get("cuenta_id") or "").isdigit() else None
            proximo = (f.get("proximo_cheque") or "").strip()
            if not (nombre and banco and numero_cuenta):
                errores.append("Nombre, banco y número de cuenta son obligatorios.")
            if not cuenta_id:
                errores.append("Seleccione la cuenta contable del banco (va al crédito de cada cheque).")
            if proximo and not proximo.isdigit():
                errores.append("El próximo número de cheque debe tener solo dígitos.")
            try:
                ajuste_x, ajuste_y = Decimal(f.get("ajuste_x_mm") or 0), Decimal(f.get("ajuste_y_mm") or 0)
            except InvalidOperation:
                errores.append("Los ajustes de impresión deben ser números (milímetros)."); ajuste_x = ajuste_y = Decimal(0)
            if proximo and id_:
                cursor.execute("SELECT MAX(numero) AS m FROM cheq_cheques WHERE banco_id = %s", (id_,))
                maximo = cursor.fetchone()["m"]
                if maximo and int(proximo) <= maximo:
                    errores.append(f"El próximo número debe ser mayor que el último cheque registrado ({maximo}).")
            if errores:
                for e in errores:
                    flash(f"⚠️ {e}", "warning")
            else:
                valores = (nombre, banco, numero_cuenta, moneda, cuenta_id, int(proximo) if proximo else None, ajuste_x, ajuste_y,
                           1 if f.get("activo") else 0)
                try:
                    if id_:
                        cursor.execute("""UPDATE cheq_bancos SET nombre=%s, banco=%s, numero_cuenta=%s, moneda=%s, cuenta_id=%s,
                                          proximo_cheque=%s, ajuste_x_mm=%s, ajuste_y_mm=%s, activo=%s WHERE id=%s""", valores + (int(id_),))
                    else:
                        cursor.execute("""INSERT INTO cheq_bancos (nombre, banco, numero_cuenta, moneda, cuenta_id, proximo_cheque,
                                          ajuste_x_mm, ajuste_y_mm, activo) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)""", valores)
                    conn.commit()
                    flash(f"✅ Banco {nombre} guardado.", "success")
                except Exception as e:
                    conn.rollback()
                    flash(f"❌ No se guardó: {e}", "danger")
            return redirect(url_for("cheques.bancos_vista"))
        lista_bancos = bancos(cursor)
        cursor.execute("SELECT banco_id, COUNT(*) AS n, MAX(numero) AS ultimo FROM cheq_cheques GROUP BY banco_id")
        uso = {r["banco_id"]: r for r in cursor.fetchall()}
        cursor.execute("SELECT id, codigo, nombre FROM cheq_cuentas WHERE activa = 1 AND codigo LIKE '1010%%' ORDER BY codigo")
        cuentas_banco = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    return render_template("cheques_bancos.html", bancos=lista_bancos, uso=uso, cuentas_banco=cuentas_banco, monedas=MONEDAS,
                           descripcion_cuenta=descripcion_cuenta)


# -----------------------------
# Cuentas contables
# -----------------------------
@cheques_bp.route("/cheques/cuentas", methods=["GET", "POST"])
@requiere_acceso("cheques_config")
def cuentas_vista():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if request.method == "POST":
            f = request.form
            id_ = f.get("id") or None
            codigo = (f.get("codigo") or "").strip()
            nombre = " ".join((f.get("nombre") or "").split()).upper()
            if not codigo or not nombre:
                flash("⚠️ Código y nombre son obligatorios.", "warning")
            else:
                cta, scta, sscta = niveles(codigo)
                cta, scta, sscta = (f.get("cta") or cta).strip(), (f.get("scta") or scta).strip(), (f.get("sscta") or sscta).strip()
                valores = (codigo[:20], nombre[:255], cta[:8], scta[:4], sscta[:8], (f.get("tipo") or "").strip()[:80] or None,
                           1 if f.get("activa") else 0)
                try:
                    if id_:
                        cursor.execute("UPDATE cheq_cuentas SET codigo=%s, nombre=%s, cta=%s, scta=%s, sscta=%s, tipo=%s, activa=%s WHERE id=%s",
                                       valores + (int(id_),))
                    else:
                        cursor.execute("INSERT INTO cheq_cuentas (codigo, nombre, cta, scta, sscta, tipo, activa, creado_por) "
                                       "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)", valores + (usuario(),))
                    conn.commit()
                    flash(f"✅ Cuenta {codigo} guardada.", "success")
                except Exception as e:
                    conn.rollback()
                    flash(f"❌ No se guardó (¿ya existe esa cuenta con ese nombre?): {e}", "danger")
            return redirect(url_for("cheques.cuentas_vista", q=f.get("volver_q", "")))

        q = (request.args.get("q") or "").strip()
        pagina = max(1, int(request.args.get("pagina", 1) or 1)) if (request.args.get("pagina") or "1").isdigit() else 1
        where, params = ["1=1"], []
        for palabra in q.split():
            where.append("(codigo LIKE %s OR nombre LIKE %s)"); params += [f"{palabra}%", f"%{palabra}%"]
        cursor.execute(f"SELECT COUNT(*) AS n FROM cheq_cuentas WHERE {' AND '.join(where)}", params)
        total = cursor.fetchone()["n"]
        cursor.execute(f"SELECT * FROM cheq_cuentas WHERE {' AND '.join(where)} ORDER BY codigo LIMIT %s OFFSET %s",
                       params + [POR_PAGINA, (pagina - 1) * POR_PAGINA])
        cuentas = cursor.fetchall()
        cursor.execute("SELECT codigo FROM cheq_cuentas GROUP BY codigo HAVING COUNT(*) > 1")
        repetidos = {r["codigo"] for r in cursor.fetchall()}
        cursor.execute("SELECT DISTINCT tipo FROM cheq_cuentas WHERE tipo IS NOT NULL ORDER BY tipo")
        tipos = [r["tipo"] for r in cursor.fetchall()]
    finally:
        cursor.close()
        conn.close()
    return render_template("cheques_cuentas.html", cuentas=cuentas, q=q, pagina=pagina, total=total, por_pagina=POR_PAGINA,
                           repetidos=repetidos, tipos=tipos)


@cheques_bp.route("/cheques/cuentas/importar", methods=["POST"])
@requiere_acceso("cheques_config")
def importar_cuentas():
    """CSV con ; (Número de cuenta;Nombre de la cuenta;Tipo de cuenta;...). Solo agrega las que no existen."""
    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        flash("Seleccione el archivo CSV.", "warning")
        return redirect(url_for("cheques.cuentas_vista"))
    contenido = archivo.read()
    texto = None
    for codificacion in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = contenido.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    filas, omitidas = [], 0
    lector = csv.reader(io.StringIO(texto), delimiter=";")
    next(lector, None)   # encabezado
    for r in lector:
        if len(r) < 2 or not r[0].strip() or not any(ch.isdigit() for ch in r[0]):
            omitidas += 1
            continue
        codigo, nombre = r[0].strip(), " ".join(r[1].split()).upper()
        cta, scta, sscta = niveles(codigo)
        filas.append((codigo[:20], nombre[:255], cta, scta, sscta, (r[2].strip()[:80] if len(r) > 2 else None), usuario()))
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.executemany("INSERT IGNORE INTO cheq_cuentas (codigo, nombre, cta, scta, sscta, tipo, creado_por) VALUES (%s,%s,%s,%s,%s,%s,%s)", filas)
        conn.commit()
        nuevas = cursor.rowcount
    finally:
        cursor.close()
        conn.close()
    flash(f"✅ Importación: {nuevas} cuentas nuevas, {len(filas) - nuevas} ya existían, {omitidas} filas sin número de cuenta omitidas.", "success")
    return redirect(url_for("cheques.cuentas_vista"))
