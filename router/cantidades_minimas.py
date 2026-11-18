"""Cantidad mínima de compra por producto (tabla cantidad_minima_compra, docs/sql/cantidad_minima_compra.sql).

- Se sube en Excel: cada fila trae el código del producto, o el número de parte con su marca y/o proveedor (el número de
  parte se repite entre marcas), la cantidad mínima y, opcional, el número de parte del proveedor (ej. 11AVX260 = 15260).
- Cantidad 0 o vacía = se quita el mínimo; Nº parte proveedor vacío = se conserva el que ya tenía. Si al producto no le queda
  ni mínimo ni Nº parte proveedor, se borra la fila.
- Validar → revisar → confirmar (archivo temporal como las demás cargas).
- El sugerido de compras lo usa para el «Sugerido ajustado» (ajustar_a_minimo)."""
import io
import json
import math
import os
import re
import uuid

from flask import Blueprint, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import sin_tildes
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, limpiar_temporales, ruta_temporal

cantidades_minimas_bp = Blueprint("cantidades_minimas", __name__)

PREFIJO_TEMP = "carga_cant_minima_"
VERDE, NARANJA = "027266", "FE6412"
MAX_FILAS = 20000
ENCABEZADOS = ["Codigo", "Numero de parte", "Marca", "Proveedor", "Numero parte proveedor", "Cantidad minima"]
COLUMNAS = {
    "codigo": {"CODIGO", "CODIGOPRODUCTO", "COD", "CODPRODUCTO"},
    "parte": {"NUMERODEPARTE", "NUMEROPARTE", "NPARTE", "NOPARTE", "NUMPARTE", "PARTE", "CODIGODEPARTE", "CODIGOPARTE"},
    "marca": {"MARCA"},
    "proveedor": {"PROVEEDOR", "PROVEEDORULTIMACOMPRA"},
    "parte_prov": {"NUMEROPARTEPROVEEDOR", "NUMERODEPARTEPROVEEDOR", "NPARTEPROVEEDOR", "PARTEPROVEEDOR", "CODIGOPROVEEDOR",
                   "CODIGODELPROVEEDOR", "CODIGOALTERNO", "CODIGOALTERNATIVO", "CODIGOALTERNATIVOPROVEEDOR"},
    "minimo": {"CANTIDADMINIMA", "CANTMINIMA", "MINIMO", "CANTIDADMINIMADECOMPRA", "MINIMODECOMPRA", "CANTIDAD", "MOQ", "MOLDQTY"},
}


def ajustar_a_minimo(sugerido, minimo):
    """Sube el sugerido al siguiente múltiplo de la cantidad mínima (0 se queda en 0)."""
    if not sugerido or sugerido <= 0 or not minimo or minimo <= 0:
        return sugerido
    return int(math.ceil(sugerido / minimo - 1e-9) * minimo) if float(minimo).is_integer() else round(math.ceil(sugerido / minimo - 1e-9) * minimo, 2)


def minimos_por_codigo(cursor, codigos=None):
    if codigos is not None and not codigos:
        return {}
    cursor.execute("SELECT codigo, minimo FROM cantidad_minima_compra WHERE minimo IS NOT NULL")
    todos = {c: float(m) for c, m in cursor.fetchall()}
    return todos if codigos is None else {c: todos[c] for c in codigos if c in todos}


def partes_proveedor(cursor, codigos=None):
    """{código: número de parte del proveedor}."""
    if codigos is not None and not codigos:
        return {}
    cursor.execute("SELECT codigo, parte_proveedor FROM cantidad_minima_compra WHERE parte_proveedor IS NOT NULL AND parte_proveedor <> ''")
    todos = dict(cursor.fetchall())
    return todos if codigos is None else {c: todos[c] for c in codigos if c in todos}


def registrados(cursor):
    """{código: (mínimo o None, nº parte proveedor o None)} de lo guardado."""
    cursor.execute("SELECT codigo, minimo, parte_proveedor FROM cantidad_minima_compra")
    return {c: (float(m) if m is not None else None, pp or None) for c, m, pp in cursor.fetchall()}


def combinar(actual, nuevo):
    """Lo que queda guardado: mínimo None en el archivo = conservar, 0 = quitar; Nº parte proveedor None = conservar."""
    m_act, pp_act = actual or (None, None)
    m_nuevo, pp_nuevo = nuevo
    return (m_act if m_nuevo is None else (m_nuevo or None), pp_act if pp_nuevo is None else pp_nuevo)


def clave(texto):
    return " ".join(sin_tildes(a_texto(texto)).upper().split())


def nombre_columna(valor):
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(a_texto(valor)).upper())


def a_cantidad(valor):
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = a_texto(valor).replace(",", "").strip()
    if not texto:
        return None
    try:
        return float(texto)
    except ValueError:
        return "error"


def catalogo():
    """Código -> (código, nº parte, marca, descripción); nº parte -> [(código, marca, proveedor)]."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT `Codigo Producto`, `Numero Parte`, Marca, Nombre, `Proveedor Ultima Compra` FROM catalogo_productos "
                       "WHERE `Codigo Producto` IS NOT NULL")
        filas = cursor.fetchall()
    finally:
        cursor.close()
        conn.close()
    por_codigo, por_parte = {}, {}
    for cod, parte, marca, nombre, proveedor in filas:
        por_codigo[clave(cod)] = (cod, parte, marca, nombre)
        if clave(parte):
            por_parte.setdefault(clave(parte), []).append((cod, clave(marca), clave(proveedor)))
    return por_codigo, por_parte


def mismo_proveedor(del_archivo, del_catalogo):
    """«SKF» coincide con «SKF DE CENTROAMERICA, S.A.»: basta que el nombre del archivo esté contenido en el del catálogo."""
    return bool(del_catalogo) and del_archivo in del_catalogo


def leer_archivo(archivo):
    """-> (registros {codigo: (minimo, parte_proveedor)}, detalle de filas, errores, advertencias).
    minimo: None = la columna no viene (se conserva), 0 = quitar. parte_proveedor: None = vacío o sin columna (se conserva)."""
    filas = leer_filas(archivo)
    posiciones, inicio = {}, None
    for i, fila in enumerate(filas[:20]):
        nombres = [nombre_columna(v) for v in fila]
        mapa = {campo: next((j for j, n in enumerate(nombres) if n in opciones), None) for campo, opciones in COLUMNAS.items()}
        if ((mapa["minimo"] is not None or mapa["parte_prov"] is not None)
                and (mapa["codigo"] is not None or mapa["parte"] is not None or mapa["parte_prov"] is not None)):
            posiciones, inicio = mapa, i + 1
            break
    if inicio is None:
        raise ErrorArchivo("no se encontraron las columnas: Codigo (o Numero de parte con Marca o Proveedor) y "
                           "Cantidad minima y/o Numero parte proveedor")

    por_codigo, por_parte = catalogo()
    conn = get_connection()
    cursor = conn.cursor()
    try:
        por_parte_prov = {}
        for cod, (_, pp) in registrados(cursor).items():
            if pp:
                por_parte_prov.setdefault(clave(pp), []).append(cod)
    finally:
        cursor.close()
        conn.close()

    registros, detalle, errores = {}, [], []
    repetidos = 0
    for n, fila in enumerate(filas[inicio:], start=inicio + 1):
        valor = lambda campo: fila[posiciones[campo]] if posiciones[campo] is not None and posiciones[campo] < len(fila) else None
        codigo, parte, marca, proveedor = clave(valor("codigo")), clave(valor("parte")), clave(valor("marca")), clave(valor("proveedor"))
        parte_prov = " ".join(a_texto(valor("parte_prov")).split())[:60] or None
        cantidad = a_cantidad(valor("minimo")) if posiciones["minimo"] is not None else None
        if not (codigo or parte or parte_prov):
            if cantidad not in (None, 0):
                errores.append(f"Fila {n}: tiene cantidad pero no trae código ni número de parte.")
            continue
        quien = codigo or " ".join(x for x in (a_texto(valor("parte")), a_texto(valor("marca")), a_texto(valor("proveedor"))) if x) or parte_prov
        if cantidad == "error" or (cantidad is not None and cantidad < 0):
            errores.append(f"Fila {n} ({quien}): la cantidad mínima «{a_texto(valor('minimo'))}» no es un número válido.")
            continue
        # Producto: por código; si no, por número de parte + marca y/o proveedor (o solo nº parte si es único);
        # si tampoco trae nuestro nº parte, por el Nº parte proveedor ya registrado
        if codigo:
            if codigo not in por_codigo:
                errores.append(f"Fila {n}: el código {a_texto(valor('codigo'))} no existe en el catálogo.")
                continue
            cod = por_codigo[codigo][0]
        elif parte:
            candidatos = [c for c in por_parte.get(parte, [])
                          if (not marca or c[1] == marca) and (not proveedor or mismo_proveedor(proveedor, c[2]))]
            filtro = " y ".join(x for x in ((f"marca {a_texto(valor('marca'))}" if marca else ""),
                                            (f"proveedor {a_texto(valor('proveedor'))}" if proveedor else "")) if x)
            if not candidatos:
                errores.append(f"Fila {n}: no se encontró el número de parte {a_texto(valor('parte'))}"
                               f"{' con ' + filtro if filtro else ''} en el catálogo.")
                continue
            if len(candidatos) > 1:
                errores.append(f"Fila {n}: el número de parte {a_texto(valor('parte'))}{' (' + filtro + ')' if filtro else ''} "
                               f"corresponde a {len(candidatos)} códigos ({', '.join(c[0] for c in candidatos[:5])}); "
                               f"{'póngalo por código' if marca and proveedor else 'agregue la marca o el proveedor, o póngalo por código'}.")
                continue
            cod = candidatos[0][0]
        else:
            candidatos = por_parte_prov.get(clave(parte_prov), [])
            if len(candidatos) != 1:
                errores.append(f"Fila {n}: el Nº parte proveedor {parte_prov} "
                               f"{'no está registrado' if not candidatos else 'está en ' + str(len(candidatos)) + ' productos'}; "
                               f"agregue nuestro número de parte con la marca o el proveedor.")
                continue
            cod = candidatos[0]
        minimo = None if posiciones["minimo"] is None else (round(cantidad, 2) if cantidad else 0.0)
        if cod in registros:
            repetidos += 1
        registros[cod] = (minimo, parte_prov)
        _, p, m, nombre = por_codigo[clave(cod)]
        detalle.append({"fila": n, "codigo": cod, "parte": p, "marca": m, "nombre": nombre, "minimo": minimo, "parte_proveedor": parte_prov})
    advertencias = []
    if repetidos:
        advertencias.append(f"{repetidos} producto(s) vienen más de una vez: queda la última fila del archivo.")
    no_enteros = [d for d in detalle if d["minimo"] and not float(d["minimo"]).is_integer()]
    if no_enteros:
        advertencias.append(f"{len(no_enteros)} producto(s) con cantidad mínima con decimales (ej. {no_enteros[0]['codigo']}: {no_enteros[0]['minimo']}).")
    if posiciones["minimo"] is None:
        advertencias.append("El archivo no trae Cantidad minima: solo se actualiza el Nº parte proveedor y los mínimos quedan como estaban.")
    return registros, detalle, errores, advertencias


# -----------------------------
# Pantalla
# -----------------------------
def lista_actual(q=""):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        where, params = "", []
        for palabra in q.split():
            where += (" AND " if where else " WHERE ") + "(m.codigo LIKE %s OR p.`Numero Parte` LIKE %s OR p.Marca LIKE %s OR p.Nombre LIKE %s OR m.parte_proveedor LIKE %s)"
            params += [f"%{palabra}%"] * 5
        cursor.execute(f"""
            SELECT m.codigo, p.`Numero Parte` AS parte, p.Marca AS marca, p.Nombre AS nombre, TRIM(p.`Proveedor Ultima Compra`) AS proveedor, m.minimo, m.parte_proveedor, m.actualizado, m.actualizado_por, m.archivo
            FROM cantidad_minima_compra m LEFT JOIN catalogo_productos p ON p.`Codigo Producto` = m.codigo {where}
            ORDER BY p.Marca, p.Nombre, m.codigo
        """, params)
        filas = cursor.fetchall()
        cursor.execute("SELECT COUNT(*) AS n FROM cantidad_minima_compra")
        total = cursor.fetchone()["n"]
    finally:
        cursor.close()
        conn.close()
    return filas, total


@cantidades_minimas_bp.route("/cantidades_minimas", methods=["GET", "POST"])
@requiere_acceso("cantidades_minimas")
def pantalla():
    q = (request.args.get("q") or "").strip()
    if request.method == "GET":
        filas, total = lista_actual(q)
        return render_template("cantidades_minimas.html", filas=filas[:1000], total=total, encontrados=len(filas), q=q)

    if request.form.get("accion") == "quitar":
        codigo = (request.form.get("codigo") or "").strip()
        conn = get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM cantidad_minima_compra WHERE codigo = %s", (codigo,))
            conn.commit()
        finally:
            cursor.close()
            conn.close()
        flash(f"Se quitó la cantidad mínima de {codigo}.", "info")
        return redirect(url_for("cantidades_minimas.pantalla", q=q))

    archivo = request.files.get("archivo")
    filas, total = lista_actual(q)
    contexto = {"filas": filas[:1000], "total": total, "encontrados": len(filas), "q": q}
    if not archivo or not archivo.filename:
        return render_template("cantidades_minimas.html", errores_archivo=["Seleccione el Excel con las cantidades mínimas."], **contexto)
    if not archivo.filename.lower().endswith((".xlsx", ".xls")):
        return render_template("cantidades_minimas.html", errores_archivo=[f"{archivo.filename}: debe ser Excel (.xlsx o .xls)."], **contexto)
    try:
        registros, detalle, errores, advertencias = leer_archivo(archivo)
    except ErrorArchivo as e:
        return render_template("cantidades_minimas.html", errores_archivo=[f"{archivo.filename}: {e}. Puede descargar la plantilla."], **contexto)
    except Exception as e:
        return render_template("cantidades_minimas.html", errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"], **contexto)
    if len(registros) > MAX_FILAS:
        return render_template("cantidades_minimas.html", errores_archivo=[f"El archivo trae {len(registros):,} productos; el máximo es {MAX_FILAS:,}."], **contexto)
    if not registros:
        return render_template("cantidades_minimas.html", errores=errores,
                               errores_archivo=[f"{archivo.filename}: ningún producto se pudo identificar."], **contexto)

    # comparar con lo que hay
    conn = get_connection()
    cursor = conn.cursor()
    try:
        actuales = registrados(cursor)
    finally:
        cursor.close()
        conn.close()
    for d in detalle:
        d["anterior"], d["pp_anterior"] = actuales.get(d["codigo"], (None, None))
        d["final_minimo"], d["final_pp"] = combinar(actuales.get(d["codigo"]), registros[d["codigo"]])
    finales = {c: combinar(actuales.get(c), r) for c, r in registros.items()}
    queda = lambda v: v != (None, None)
    nuevos = sum(1 for c, v in finales.items() if queda(v) and c not in actuales)
    cambian = sum(1 for c, v in finales.items() if queda(v) and c in actuales and v != actuales[c])
    iguales = sum(1 for c, v in finales.items() if queda(v) and c in actuales and v == actuales[c])
    quitar = sum(1 for c, v in finales.items() if not queda(v) and c in actuales)
    sin_efecto = sum(1 for c, v in finales.items() if not queda(v) and c not in actuales)

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump({"archivo": archivo.filename[:255], "registros": registros}, f, ensure_ascii=False)
    session["carga_cant_minima_token"] = token
    ultimos = {}
    for d in detalle:    # una fila por producto (la última del archivo)
        ultimos[d["codigo"]] = d
    return render_template("cantidades_minimas.html", validado=True, token=token, archivo=archivo.filename,
                           detalle=list(ultimos.values())[:2000], total_archivo=len(registros), nuevos=nuevos, cambian=cambian,
                           iguales=iguales, quitar=quitar, sin_efecto=sin_efecto, errores=errores, advertencias=advertencias, **contexto)


@cantidades_minimas_bp.route("/cantidades_minimas/confirmar", methods=["POST"])
@requiere_acceso("cantidades_minimas")
def confirmar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_cant_minima_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cantidades_minimas.pantalla"))
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        conn.start_transaction()
        actuales = registrados(cursor)
        finales = {c: combinar(actuales.get(c), tuple(r)) for c, r in datos["registros"].items()}
        guardar = [(c, m, pp, datos["archivo"], session.get("usuario")) for c, (m, pp) in finales.items() if (m, pp) != (None, None)]
        quitar = [c for c, v in finales.items() if v == (None, None) and c in actuales]
        for i in range(0, len(guardar), 1000):
            cursor.executemany("INSERT INTO cantidad_minima_compra (codigo, minimo, parte_proveedor, archivo, actualizado_por) "
                               "VALUES (%s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE minimo = VALUES(minimo), "
                               "parte_proveedor = VALUES(parte_proveedor), archivo = VALUES(archivo), actualizado_por = VALUES(actualizado_por)",
                               guardar[i:i + 1000])
        for i in range(0, len(quitar), 1000):
            lote = quitar[i:i + 1000]
            cursor.execute(f"DELETE FROM cantidad_minima_compra WHERE codigo IN ({', '.join(['%s'] * len(lote))})", lote)
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó. Error de base de datos: {e}", "danger")
        return redirect(url_for("cantidades_minimas.pantalla"))
    finally:
        cursor.close()
        conn.close()
    os.remove(ruta)
    session.pop("carga_cant_minima_token", None)
    flash(f"✅ Guardados {len(guardar):,} producto(s)" + (f"; se quitaron {len(quitar):,}." if quitar else "."), "success")
    return redirect(url_for("cantidades_minimas.pantalla"))


@cantidades_minimas_bp.route("/cantidades_minimas/plantilla")
@requiere_acceso("cantidades_minimas")
def plantilla():
    """Plantilla vacía, o con las cantidades actuales (actuales=1) para corregirlas y volver a subir."""
    con_actuales = request.args.get("actuales") == "1"
    wb = Workbook()
    ws = wb.active
    ws.title = "Cantidades minimas"
    for i, (nombre, ancho) in enumerate(zip(ENCABEZADOS + (["Descripción"] if con_actuales else []), (14, 22, 16, 26, 24, 16, 45)), start=1):
        c = ws.cell(row=1, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if nombre == "Cantidad minima" else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = ancho
    filas = lista_actual()[0] if con_actuales else []
    for n, f in enumerate(filas, start=2):
        for i, v in enumerate([f["codigo"], f["parte"], f["marca"], f["proveedor"], f["parte_proveedor"],
                               float(f["minimo"]) if f["minimo"] is not None else None, f["nombre"]], start=1):
            ws.cell(row=n, column=i, value=v)
    for fila in range(2, max(len(filas) + 2, 502)):
        for col in (1, 2, 3, 4, 5):
            ws.cell(row=fila, column=col).number_format = "@"   # texto: Excel no quita ceros ni cambia números de parte
    ws.freeze_panes = "A2"
    ayuda = wb.create_sheet("Instrucciones")
    ayuda.column_dimensions["A"].width = 105
    for n, (texto, negrita) in enumerate([
        ("Cantidades mínimas de compra", True), ("", False),
        ("Una fila por producto, desde la fila 2 de la hoja «Cantidades minimas».", False),
        ("Codigo: el código del producto (lo más seguro). Si no lo tiene, ponga nuestro Numero de parte con la Marca y/o el Proveedor:", False),
        ("   el número de parte se repite entre marcas; sin marca ni proveedor solo se acepta si corresponde a un único producto.", False),
        ("   Proveedor: el de la última compra en el catálogo; basta una parte del nombre (ej. «SKF»).", False),
        ("Numero parte proveedor (opcional): el código con el que el proveedor conoce el producto (ej. 11AVX260 para nuestro 15260).", False),
        ("   Sale en el Excel del sugerido de compras. Vacío = se conserva el que ya tenía.", False),
        ("Cantidad minima: lo mínimo que se puede comprar (empaque, caja, lote). Con 0 o vacía se quita el mínimo del producto.", False),
        ("En el sugerido de compras, el «Sugerido ajustado» sube el sugerido al siguiente múltiplo del mínimo", False),
        ("   (mínimo 10: sugerido 3 → 10, 13 → 20, 25 → 30). El pedido y el total FOB parten del sugerido ajustado.", False),
        ("Los productos que no vengan en el archivo conservan su cantidad mínima.", False),
    ], start=1):
        c = ayuda.cell(row=n, column=1, value=texto)
        c.font = Font(bold=True, size=13, color=VERDE) if negrita else Font(size=10)
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    return send_file(salida, as_attachment=True, download_name="cantidades_minimas_actuales.xlsx" if con_actuales else "plantilla_cantidades_minimas.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
