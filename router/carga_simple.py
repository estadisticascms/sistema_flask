"""Lógica común para cargas por clave primaria (recibos de caja, devoluciones):
comparar con la base de datos, guardar la validación en un archivo temporal y confirmar en una transacción."""
import json
import os
import uuid
from datetime import date
from decimal import Decimal

from flask import session

from utils.db import get_connection
from router.cargar_ventas import ruta_temporal, limpiar_temporales
from router.cargar_catalogo import iguales

VACIOS = {"nan", "none", "null"}   # textos que dejaron cargas anteriores en lugar de vacío


def normal(valor):
    if isinstance(valor, str) and valor.strip().lower() in VACIOS:
        return None
    return valor


def comparar(tabla, clave, campos, numericas, registros):
    """Clasifica los registros del archivo en nuevos, modificados (con sus cambios) y sin cambio."""
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    actuales = {}
    try:
        claves = [r[clave] for r in registros]
        for i in range(0, len(claves), 1000):
            lote = claves[i:i + 1000]
            cursor.execute(f"SELECT * FROM {tabla} WHERE {clave} IN ({', '.join(['%s'] * len(lote))})", lote)
            actuales.update({r[clave]: r for r in cursor.fetchall()})
    finally:
        cursor.close()
        conn.close()

    nuevos, modificados, sin_cambio = [], [], 0
    for r in registros:
        a = actuales.get(r[clave])
        if a is None:
            nuevos.append(r)
            continue
        cambios = [(c, normal(a[c]), r[c]) for c in campos if c != clave and not iguales(c, normal(a[c]), r[c], numericas)]
        if cambios:
            modificados.append({"registro": r, "cambios": cambios})
        else:
            sin_cambio += 1
    return {"nuevos": nuevos, "modificados": modificados, "sin_cambio": sin_cambio}


def _json(valor):
    if isinstance(valor, Decimal):
        return str(valor)
    if isinstance(valor, date):
        return valor.isoformat()
    return valor


def guardar_temporal(prefijo, clave_sesion, campos, dif):
    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, prefijo), "w", encoding="utf-8") as f:
        json.dump({"nuevos": [[_json(r[c]) for c in campos] for r in dif["nuevos"]],
                   "modificados": [[_json(m["registro"][c]) for c in campos] for m in dif["modificados"]]},
                  f, ensure_ascii=False)
    session[clave_sesion] = token
    return token


def confirmar(tabla, clave, campos, prefijo, clave_sesion, token):
    """Inserta los nuevos y actualiza los modificados en una sola transacción.
    Devuelve (nuevos, modificados) o lanza ValueError si la validación expiró."""
    ruta = ruta_temporal(token, prefijo)
    if not ruta or token != session.get(clave_sesion) or not os.path.exists(ruta):
        raise ValueError("La validación expiró o ya fue cargada. Vuelva a subir el archivo.")
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)

    otros = [c for c in campos if c != clave]
    pos = campos.index(clave)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()
        # Por si alguien cargó lo mismo entre la validación y la confirmación, los "nuevos" que ya
        # existan se actualizan en vez de fallar por clave duplicada.
        if datos["nuevos"]:
            cursor.executemany(
                f"INSERT INTO {tabla} ({', '.join(campos)}) VALUES ({', '.join(['%s'] * len(campos))}) "
                f"ON DUPLICATE KEY UPDATE {', '.join(f'{c} = VALUES({c})' for c in otros)}", datos["nuevos"])
        for fila in datos["modificados"]:
            valores = dict(zip(campos, fila))
            cursor.execute(f"UPDATE {tabla} SET {', '.join(f'{c} = %s' for c in otros)} WHERE {clave} = %s",
                           [valores[c] for c in otros] + [fila[pos]])
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop(clave_sesion, None)
    return len(datos["nuevos"]), len(datos["modificados"])
