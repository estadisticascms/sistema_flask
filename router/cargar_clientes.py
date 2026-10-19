"""Actualiza la tabla clientes desde el reporte "Lista de Clientes" del sistema (ListaClientes.xlsx).

Cada fila del archivo se busca por `clave` (RUC_Estado_NombreSinEspacios). Si no aparece, se busca
por RUC: así un cliente que cambió de estado o de nombre se actualiza en vez de duplicarse.
Lo que no se encuentra se agrega. Los clientes que no vienen en el archivo NO se eliminan
(los reportes de zonas cruzan las ventas con el nombre del cliente).
Departamento y municipio solo se cambian cuando el archivo trae un valor (se editan a mano en /clientes)."""
import json
import os
import re
import uuid

from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, ruta_temporal, limpiar_temporales
from router.cargar_catalogo import sin_tildes, iguales
from router.cargar_inventario import arreglar_texto

cargar_clientes_bp = Blueprint("cargar_clientes", __name__)

PREFIJO_TEMP = "carga_clientes_"
MAX_DETALLE = 300

# columna del archivo -> columna de la tabla
COLUMNAS = {
    "Nombre": "nombre", "RUC": "ruc", "Estado": "estado", "Términos Pago": "terminos_pago",
    "Tipo Cliente": "tipo_cliente", "Vendedor": "vendedor", "Precio": "precio", "Descuento": "descuento",
    "% Ret. DGI": "ret_dgi", "% Ret. ALMA": "ret_alma", "Límite Crédito": "limite_credito",
    "Direccion": "direccion", "Departamento": "departamento", "Municipio": "municipio", "Telefono": "telefono",
}
NUMERICAS = {"precio", "descuento", "ret_dgi", "ret_alma", "limite_credito"}
SOLO_SI_TRAE_VALOR = {"departamento", "municipio"}   # no se borran si el archivo viene vacío
LARGOS = {"nombre": 255, "ruc": 50, "terminos_pago": 50, "tipo_cliente": 50, "vendedor": 100,
          "direccion": 255, "departamento": 100, "telefono": 50, "municipio": 45}
CAMPOS = list(COLUMNAS.values())
NOMBRES = {v: k for k, v in COLUMNAS.items()}


def clave_de(c):
    """Misma regla con la que se cargó la tabla: RUC (o NA) _ Estado _ nombre sin espacios."""
    return f"{c['ruc'] or 'NA'}_{c['estado']}_{''.join(c['nombre'].split())}"


def reparar(texto):
    """El sistema exporta algunas tildes/eñes dañadas, a veces dos veces (ZUÃIGA -> ZUÑIGA)."""
    for _ in range(3):
        nuevo = arreglar_texto(texto)
        if nuevo == texto:
            break
        texto = nuevo
    return texto


def a_numero(valor):
    texto = a_texto(valor).replace(",", "")
    return float(texto) if texto else 0.0   # ValueError si no es número


# -----------------------------
# Lectura del archivo
# -----------------------------
def leer_clientes(archivo):
    filas = leer_filas(archivo)
    objetivo = {sin_tildes(c): c for c in COLUMNAS}
    inicio, pos = None, None
    for i, fila in enumerate(filas[:40]):
        encontrados = {objetivo[sin_tildes(a_texto(v))]: j for j, v in enumerate(fila) if sin_tildes(a_texto(v)) in objetivo}
        if len(encontrados) == len(COLUMNAS):
            inicio, pos = i, encontrados
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas de la Lista de Clientes (Nombre, RUC, Estado, Vendedor, Límite Crédito, ...)")

    # "Total de resultados: N" en el encabezado del reporte
    total_reportado = None
    for fila in filas[:inicio]:
        for j, v in enumerate(fila):
            if "total de resultados" in sin_tildes(a_texto(v)):
                for w in fila[j + 1:]:
                    if re.fullmatch(r"\d+(\.0)?", a_texto(w)):
                        total_reportado = int(float(a_texto(w)))
                        break

    clientes, errores, advertencias = [], [], []
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        if not any(a_texto(v) for v in fila):
            continue
        c = {}
        for col_archivo, campo in COLUMNAS.items():
            valor = fila[pos[col_archivo]] if pos[col_archivo] < len(fila) else None
            if campo in NUMERICAS:
                try:
                    c[campo] = a_numero(valor)
                except ValueError:
                    c[campo] = 0.0
                    advertencias.append(f"Fila {num} ({a_texto(fila[pos['Nombre']])}): {col_archivo} no es un número ('{a_texto(valor)}'); se guarda 0")
            else:
                c[campo] = a_texto(valor) or None
        nombre_original = c["nombre"]
        c = {k: reparar(v) if isinstance(v, str) else v for k, v in c.items()}
        if not c["nombre"]:
            errores.append(f"Fila {num}: no tiene nombre; no se carga")
            continue
        estado = (c["estado"] or "").capitalize()
        if estado not in ("Activo", "Inactivo"):
            errores.append(f"Fila {num} ({c['nombre']}): estado '{c['estado'] or ''}' no válido (debe ser Activo o Inactivo); no se carga")
            continue
        c["estado"] = estado
        for campo, largo in LARGOS.items():
            if c[campo] and len(c[campo]) > largo:
                advertencias.append(f"Fila {num} ({c['nombre']}): {NOMBRES[campo]} tiene más de {largo} caracteres; se recorta")
                c[campo] = c[campo][:largo]
        c["clave"] = clave_de(c)
        # Clave con el nombre tal como viene (así se cargaron los clientes antes de reparar el texto)
        c["clave_original"] = clave_de({**c, "nombre": nombre_original[:LARGOS["nombre"]]})
        c["_fila"] = num
        clientes.append(c)

    if not clientes:
        raise ErrorArchivo("no contiene clientes")
    return clientes, errores, advertencias, total_reportado


# -----------------------------
# Comparación con la base de datos
# -----------------------------
def comparar(clientes):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT * FROM clientes")
        actuales = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()

    por_clave = {a["clave"]: a for a in actuales}
    usados = set()      # ids de la tabla ya emparejados con una fila del archivo
    pendientes = []
    emparejados = []    # (cliente del archivo, fila actual)
    for c in clientes:
        a = por_clave.get(c["clave"])
        if a is None or a["id"] in usados:
            a = por_clave.get(c["clave_original"])
        if a is not None and a["id"] not in usados:
            usados.add(a["id"])
            emparejados.append((c, a, "clave"))
        else:
            pendientes.append(c)

    # Segundo intento por RUC: solo si el RUC identifica a un único cliente en ambos lados
    def por_ruc(filas, clave_ruc):
        grupos = {}
        for f in filas:
            ruc = (f[clave_ruc] or "").strip()
            if ruc and ruc.upper() != "NA":
                grupos.setdefault(ruc, []).append(f)
        return grupos

    libres = por_ruc([a for a in actuales if a["id"] not in usados], "ruc")
    archivo_ruc = por_ruc(pendientes, "ruc")
    nuevos = []
    for c in pendientes:
        ruc = (c["ruc"] or "").strip()
        candidatos = libres.get(ruc, [])
        if ruc and len(candidatos) == 1 and len(archivo_ruc.get(ruc, [])) == 1:
            a = candidatos[0]
            usados.add(a["id"])
            emparejados.append((c, a, "ruc"))
        else:
            nuevos.append(c)

    modificados, sin_cambio = [], 0
    for c, a, metodo in emparejados:
        cambios = []
        for campo in CAMPOS:
            if campo in SOLO_SI_TRAE_VALOR and not c[campo]:
                continue
            if not iguales(campo, a[campo], c[campo], NUMERICAS):
                cambios.append((NOMBRES[campo], a[campo], c[campo]))
        if cambios:
            modificados.append({"id": a["id"], "nombre": c["nombre"], "metodo": metodo, "cambios": cambios, "cliente": c,
                                "clave_nueva": c["clave"] if c["clave"] != a["clave"] else None})
        else:
            sin_cambio += 1

    no_vienen = [a for a in actuales if a["id"] not in usados]
    return {"actuales": len(actuales), "nuevos": nuevos, "modificados": modificados, "sin_cambio": sin_cambio,
            "no_vienen": no_vienen, "por_ruc": sum(1 for m in modificados if m["metodo"] == "ruc")}


def resumen_cambios(modificados):
    conteo = {}
    for m in modificados:
        for col, _, _ in m["cambios"]:
            conteo[col] = conteo.get(col, 0) + 1
    return sorted(conteo.items(), key=lambda x: -x[1])


# -----------------------------
# Vistas
# -----------------------------
@cargar_clientes_bp.route("/cargar_clientes", methods=["GET", "POST"])
@requiere_acceso("cargar_clientes")
def cargar_clientes():
    if request.method == "GET":
        return render_template("cargar_clientes.html")

    archivo = request.files.get("archivo_clientes")
    if not archivo or not archivo.filename:
        return render_template("cargar_clientes.html", errores_archivo=["Falta el archivo Lista de Clientes."])
    try:
        clientes, errores, advertencias, total_reportado = leer_clientes(archivo)
    except ErrorArchivo as e:
        return render_template("cargar_clientes.html", errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el archivo correcto?"])
    except Exception as e:
        return render_template("cargar_clientes.html", errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    # Filas repetidas (misma clave): se conserva la primera
    vistos, unicos = {}, []
    for c in clientes:
        if c["clave"] in vistos:
            advertencias.append(f"Fila {c['_fila']}: {c['nombre']} ({c['ruc'] or 'sin RUC'}, {c['estado']}) repetido "
                                f"(igual a la fila {vistos[c['clave']]}); se ignora")
            continue
        vistos[c["clave"]] = c["_fila"]
        unicos.append(c)
    if total_reportado is not None and total_reportado != len(clientes):
        advertencias.insert(0, f"El reporte dice Total de resultados: {total_reportado:,} pero se leyeron {len(clientes):,} filas válidas")
    clientes = unicos

    try:
        dif = comparar(clientes)
    except Exception as e:
        return render_template("cargar_clientes.html", errores_archivo=[f"Error al comparar con la base de datos: {e}"])

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump({
            "nuevos": [{k: c[k] for k in CAMPOS + ["clave"]} for c in dif["nuevos"]],
            "modificados": [{"id": m["id"], "datos": {k: m["cliente"][k] for k in CAMPOS + ["clave"]}} for m in dif["modificados"]],
        }, f, ensure_ascii=False)
    session["carga_clientes_token"] = token

    return render_template("cargar_clientes.html", validado=True, token=token, total=len(clientes), dif=dif,
                           cambios_por_columna=resumen_cambios(dif["modificados"]), max_detalle=MAX_DETALLE,
                           errores=errores, advertencias=advertencias)


@cargar_clientes_bp.route("/cargar_clientes/confirmar", methods=["POST"])
@requiere_acceso("cargar_clientes")
def confirmar_carga_clientes():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_clientes_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cargar_clientes.cargar_clientes"))

    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)

    columnas = CAMPOS + ["clave"]
    sql_insert = f"INSERT INTO clientes ({', '.join(columnas)}) VALUES ({', '.join(['%s'] * len(columnas))})"
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()   # todo o nada
        for m in datos["modificados"]:
            d = m["datos"]
            campos = [k for k in columnas if not (k in SOLO_SI_TRAE_VALOR and not d[k])]
            cursor.execute(f"UPDATE clientes SET {', '.join(f'{k} = %s' for k in campos)} WHERE id = %s",
                           [d[k] for k in campos] + [m["id"]])
        if datos["nuevos"]:
            cursor.executemany(sql_insert, [[c[k] for k in columnas] for c in datos["nuevos"]])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se modificaron los clientes. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_clientes.cargar_clientes"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_clientes_token", None)
    flash(f"✅ Clientes actualizados: {len(datos['modificados']):,} modificados y {len(datos['nuevos']):,} nuevos.", "success")
    return render_template("cargar_clientes.html", cargado=True)
