import json
import os
import re
import uuid

from flask import Blueprint, render_template, request, redirect, url_for, flash, session, jsonify
from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_ventas import ErrorArchivo, a_texto, leer_filas, buscar_encabezado, ruta_temporal, limpiar_temporales
from router.productos_incompletos import vacio, tiene_valor, enlaces, CAMPOS as CAMPOS_FICHA
from router import sugerencias_clase
from router.cargar_catalogo import sin_tildes
from router.ventas_bodegas import en_cache

clase_prod_bp = Blueprint("clase_prod", __name__)

POR_PAGINA = 50
LARGO_CODIGO = 10
LARGO_CAMPO = 45
CAMPOS = ["segmento", "familia", "subfamilia", "grupo_equival"]
POR_CLASIFICAR = "Por clasificar"
SEGMENTOS_BASE = ["Liviano", "Pesado", "Agrícola", "Industrial", "Mixto (liviano/pesado)", "Universal",
                  "Multi-segmento (genérico)", "Sin definir", "No aplica"]
ORDENES = {
    "marca": "p.Marca, p.Nombre, p.`Codigo Producto`",
    "existencia": "p.`Existencia Actual` DESC, p.Marca, p.Nombre, p.`Codigo Producto`",
    "codigo": "CAST(SUBSTRING(p.`Codigo Producto`, 4) AS UNSIGNED)",
}
# Ficha del catálogo (viene del sistema; aquí solo se consulta): medida, aplicación y equivalencias
FICHAS = {"incompleta": "Falta medida, aplicación o equivalencia", "medida": "Falta medida",
          "aplicacion": "Falta aplicación", "equivalencia": "Falta equivalencia"}
# Clasificación incompleta: falta segmento, familia, subfamilia o grupo de equivalencia
INCOMPLETO = ("(c.codigo IS NULL OR COALESCE(c.familia, '') IN ('', 'Por clasificar') OR COALESCE(c.subfamilia, '') = '' "
              "OR COALESCE(c.segmento, '') IN ('', 'Sin definir') OR COALESCE(c.grupo_equival, '') = '')")
# Productos que no hace falta clasificar ("descartados"):
#  - sin movimiento: ninguna transacción de inventario, o sin entradas y sin existencia
#  - publicidad: catálogos, lapiceros, afiches, calendarios, etc.
PUBLICIDAD = re.compile(r"PUBLICIDAD|PROMOCIONAL|LAPICER|BOLIGRAF|CATALOGO|CALENDAR|AFICHE|POSTER|BROCHU|BOLETIN|"
                        r"BANN?ER|CALCOMANIA|LLAVERO|GORRA|CAMISETA|AGENDA|OBSEQUIO")
DESCARTES = {"1": "Ocultar sin movimiento y publicidad", "0": "Mostrar todos", "solo": "Solo sin movimiento y publicidad"}
PREFIJO_TEMP = "carga_clase_prod_"
LOTE = 2000


# -----------------------------
# Consultas
# -----------------------------
def filtros_desde(args):
    return {
        "q": (args.get("q") or "").strip(),
        "estado": args.get("estado") or "todos",
        "segmento": args.get("segmento") or "",
        "familia": args.get("familia") or "",
        "subfamilia": args.get("subfamilia") or "",
        "grupo": (args.get("grupo") or "").strip(),
        "orden": args.get("orden") if args.get("orden") in ORDENES else "",
        "ficha": args.get("ficha") if args.get("ficha") in FICHAS else "",
        "existencia": "1" if args.get("existencia") == "1" else "",
        # por defecto se ocultan; el formulario manda "0" cuando se desmarca la casilla
        "descartar": args.get("descartar") if args.get("descartar") in DESCARTES else "1",
    }


def descartados():
    """{codigo: motivo} de los productos que no hace falta clasificar (10 minutos en memoria)."""
    def leer():
        conn = get_connection()
        cursor = conn.cursor()
        try:
            # Agrupar por el índice (producto, fecha, tipo) tarda ~0.5 s
            cursor.execute("SELECT /*+ INDEX(transacciones_inventario idx_transacciones_producto_fecha) */ codigo_producto, "
                           "MAX(tipo_transaccion IN ('ENTRADA', 'AJUSTE ENTRADA')) FROM transacciones_inventario GROUP BY codigo_producto")
            movimiento = {r[0]: bool(r[1]) for r in cursor.fetchall()}
            cursor.execute("SELECT `Codigo Producto`, COALESCE(`Existencia Actual`, 0), Nombre, Categoria, Marca "
                           "FROM catalogo_productos WHERE `Codigo Producto` IS NOT NULL")
            resultado = {}
            for cod, existencia, nombre, categoria, marca in cursor.fetchall():
                if (categoria or "").upper() == "PUBLICIDAD" or (marca or "").upper() == "PROMOCIONAL" or PUBLICIDAD.search(sin_tildes(nombre or "").upper()):
                    resultado[cod] = "publicidad"
                elif cod not in movimiento or (existencia <= 0 and not movimiento[cod]):
                    resultado[cod] = "sin movimiento"
            return resultado
        finally:
            cursor.close()
            conn.close()
    return en_cache("clase_descartados", leer)


def condicion_descarte():
    """SQL verdadero para los productos descartados. Los ~9 mil códigos van escritos en la consulta
    (como parámetros, el conector tarda segundos en armarla); solo se aceptan códigos alfanuméricos."""
    def armar():
        codigos = [c for c in descartados() if re.fullmatch(r"[A-Za-z0-9._-]+", c)]
        return f"p.`Codigo Producto` IN ({', '.join(repr(c) for c in codigos) or 'NULL'})"
    return en_cache("clase_sql_descarte", armar)


def condiciones(f):
    where, params = ["p.`Codigo Producto` IS NOT NULL"], []
    if f["descartar"] != "0":
        sql = condicion_descarte()
        where.append(sql if f["descartar"] == "solo" else f"NOT {sql}")
    for palabra in f["q"].split():
        like = f"%{palabra}%"
        where.append("(p.`Codigo Producto` LIKE %s OR p.Nombre LIKE %s OR p.`Numero Parte` LIKE %s "
                     "OR p.Marca LIKE %s OR p.Equivalencias LIKE %s)")
        params += [like] * 5
    if f["estado"] == "clasificados":
        where.append("c.codigo IS NOT NULL AND COALESCE(c.familia, '') NOT IN ('', %s)")
        params.append(POR_CLASIFICAR)
    elif f["estado"] == "por_clasificar":
        where.append("c.codigo IS NOT NULL AND COALESCE(c.familia, '') IN ('', %s)")
        params.append(POR_CLASIFICAR)
    elif f["estado"] == "sin_clasificar":
        where.append("c.codigo IS NULL")
    elif f["estado"] == "incompletos":
        where.append(INCOMPLETO)
    if f["ficha"] == "incompleta":
        where.append("(" + " OR ".join(vacio(col) for col in CAMPOS_FICHA.values()) + ")")
    elif f["ficha"]:
        where.append(vacio(CAMPOS_FICHA[f["ficha"]]))
    if f["existencia"]:
        where.append("p.`Existencia Actual` > 0")
    for campo, col in (("segmento", "c.segmento"), ("familia", "c.familia"), ("subfamilia", "c.subfamilia"), ("grupo", "c.grupo_equival")):
        if f[campo]:
            where.append(f"{col} = %s")
            params.append(f[campo])
    return " AND ".join(where), params


def opciones(cursor):
    cursor.execute("SELECT DISTINCT segmento FROM clase_prod WHERE segmento IS NOT NULL AND segmento <> '' ORDER BY segmento")
    segmentos = list(dict.fromkeys(SEGMENTOS_BASE + [r["segmento"] for r in cursor.fetchall()]))
    cursor.execute("SELECT familia, subfamilia, COUNT(*) AS n FROM clase_prod "
                   "WHERE familia IS NOT NULL AND familia <> '' GROUP BY familia, subfamilia ORDER BY familia, subfamilia")
    arbol = {}
    for r in cursor.fetchall():
        subs = arbol.setdefault(r["familia"], [])
        if r["subfamilia"]:
            subs.append(r["subfamilia"])
    return segmentos, arbol


def contadores(cursor, descartar="1"):
    sql = condicion_descarte()
    filtro = f"AND NOT {sql}" if descartar == "1" else (f"AND {sql}" if descartar == "solo" else "")
    cursor.execute(f"""
        SELECT /*+ NO_INDEX(c) */ COUNT(*) AS total,
               SUM(c.codigo IS NULL) AS sin_clasificar,
               SUM(c.codigo IS NOT NULL AND COALESCE(c.familia, '') IN ('', %s)) AS por_clasificar,
               SUM(c.codigo IS NOT NULL AND COALESCE(c.familia, '') NOT IN ('', %s)) AS clasificados,
               SUM(p.`Existencia Actual` > 0 AND {INCOMPLETO}) AS incompletos,
               SUM(p.`Existencia Actual` > 0 AND ({' OR '.join(vacio(col) for col in CAMPOS_FICHA.values())})) AS ficha
        FROM catalogo_productos p
        LEFT JOIN clase_prod c ON c.codigo = p.`Codigo Producto`
        WHERE p.`Codigo Producto` IS NOT NULL {filtro}
    """, (POR_CLASIFICAR, POR_CLASIFICAR))
    conteo = {k: int(v or 0) for k, v in cursor.fetchone().items()}
    d = descartados()   # siempre sobre todo el catálogo
    conteo["descartados"] = len(d)
    conteo["publicidad"] = sum(1 for m in d.values() if m == "publicidad")
    conteo["sin_movimiento"] = conteo["descartados"] - conteo["publicidad"]
    return conteo


def limpiar(valor):
    valor = (valor or "").strip()
    return valor or None


def validar(codigo, datos):
    errores = []
    if not codigo or len(codigo) > LARGO_CODIGO:
        errores.append(f"código inválido '{codigo}'")
    for campo, valor in datos.items():
        if valor and len(valor) > LARGO_CAMPO:
            errores.append(f"{campo} supera {LARGO_CAMPO} caracteres")
    return errores


def volver_a(destino):
    """Regresa a la lista con los mismos filtros (solo rutas internas de esta vista)."""
    if destino and destino.startswith("/clase_prod") and "//" not in destino:
        return redirect(destino)
    return redirect(url_for("clase_prod.clase_prod"))


# -----------------------------
# Vista principal
# -----------------------------
@clase_prod_bp.route("/clase_prod", methods=["GET"])
@requiere_acceso("clase_prod")
def clase_prod():
    f = filtros_desde(request.args)
    pagina = max(int(request.args.get("page", 1) or 1), 1)
    orden = ORDENES[f["orden"] or ("existencia" if f["estado"] in ("sin_clasificar", "por_clasificar", "incompletos") or f["ficha"] else "marca")]
    where, params = condiciones(f)

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        # NO_INDEX(c): con `Codigo Producto` tipo TEXT, el hash join es ~2x más rápido que usar la PK
        cursor.execute(f"SELECT /*+ NO_INDEX(c) */ COUNT(*) AS n FROM catalogo_productos p LEFT JOIN clase_prod c "
                       f"ON c.codigo = p.`Codigo Producto` WHERE {where}", params)
        total = cursor.fetchone()["n"]
        paginas = max((total + POR_PAGINA - 1) // POR_PAGINA, 1)
        pagina = min(pagina, paginas)

        cursor.execute(f"""
            SELECT p.`Codigo Producto` AS codigo, p.Nombre AS nombre, p.Marca AS marca,
                   p.`Numero Parte` AS numero_parte, p.Equivalencias AS equivalencias,
                   p.`Existencia Actual` AS existencia, p.Medida AS medida, p.`Descripcion Larga` AS aplicacion,
                   c.codigo IS NOT NULL AS tiene_clase, c.segmento, c.familia, c.subfamilia, c.grupo_equival
            FROM catalogo_productos p
            LEFT JOIN clase_prod c ON c.codigo = p.`Codigo Producto`
            WHERE {where}
            ORDER BY {orden}
            LIMIT %s OFFSET %s
        """, params + [POR_PAGINA, (pagina - 1) * POR_PAGINA])
        registros = cursor.fetchall()
        segmentos, arbol = opciones(cursor)
        conteo = contadores(cursor, f["descartar"])
    finally:
        cursor.close()
        conn.close()

    # Sugerencias (equivalencias / aplicación) y ficha del catálogo para las filas de la página
    sugerencias = sugerencias_clase.sugerir([r["codigo"] for r in registros])
    motivos = descartados()
    for r in registros:
        r["descarte"] = motivos.get(r["codigo"])
        r["faltan_ficha"] = [k for k, col in (("medida", "medida"), ("aplicacion", "aplicacion"), ("equivalencia", "equivalencias"))
                             if not tiene_valor(r[col])]
        r["enlaces"] = enlaces(r["numero_parte"], r["marca"]) if r["faltan_ficha"] else {}
        s = sugerencias.get(r["codigo"])
        # Solo se muestra si propone algo distinto a lo guardado
        r["sug"] = s if s and any(s[k] and s[k] != (r[c] or "") for k, c in (("segmento", "segmento"), ("familia", "familia"),
                                  ("subfamilia", "subfamilia"), ("grupo", "grupo_equival"))) else None
        r["incompleto"] = (not r["tiene_clase"] or sugerencias_clase.falta(r["familia"]) or not (r["subfamilia"] or "").strip()
                           or sugerencias_clase.falta(r["segmento"]) or not (r["grupo_equival"] or "").strip())

    return render_template(
        "clase_prod.html",
        registros=registros, f=f, pagina=pagina, paginas=paginas, total=total,
        segmentos=segmentos, arbol=arbol, conteo=conteo, ordenes=list(ORDENES), fichas=FICHAS, descartes=DESCARTES,
        volver=request.full_path.rstrip("?"),
    )


@clase_prod_bp.route("/clase_prod/guardar", methods=["POST"])
@requiere_acceso("clase_prod")
def guardar():
    # La vista guarda cada fila con fetch (sin recargar, para no perder otras filas editadas);
    # sin JavaScript se usa el envío normal del formulario y se redirige.
    es_fetch = request.headers.get("X-Requested-With") == "fetch"
    codigo = (request.form.get("codigo") or "").strip()
    datos = {c: limpiar(request.form.get(c)) for c in CAMPOS}
    errores = validar(codigo, datos)

    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT 1 FROM catalogo_productos WHERE `Codigo Producto` = %s LIMIT 1", (codigo,))
        if cursor.fetchone() is None:
            errores.append(f"el código {codigo} no existe en el catálogo")
        if errores:
            if es_fetch:
                return jsonify(ok=False, error="; ".join(errores)), 400
            flash(f"❌ No se guardó {codigo}: " + "; ".join(errores), "danger")
            return volver_a(request.form.get("volver"))
        cursor.execute("""
            INSERT INTO clase_prod (codigo, segmento, familia, subfamilia, grupo_equival)
            VALUES (%s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE segmento = VALUES(segmento), familia = VALUES(familia),
                                    subfamilia = VALUES(subfamilia), grupo_equival = VALUES(grupo_equival)
        """, (codigo, datos["segmento"], datos["familia"], datos["subfamilia"], datos["grupo_equival"]))
        # Equivalentes sin grupo que la sugerencia propuso juntar con este producto (solo si no tienen grupo)
        agrupar = [c.strip() for c in (request.form.get("agrupar") or "").split(",") if c.strip() and len(c.strip()) <= LARGO_CODIGO]
        agrupados = 0
        if agrupar and datos["grupo_equival"]:
            cursor.execute(f"SELECT `Codigo Producto` FROM catalogo_productos WHERE `Codigo Producto` IN ({', '.join(['%s'] * len(agrupar))})", agrupar)
            for c in [r[0] for r in cursor.fetchall()]:
                cursor.execute("INSERT INTO clase_prod (codigo, grupo_equival) VALUES (%s, %s) "
                               "ON DUPLICATE KEY UPDATE grupo_equival = IF(COALESCE(grupo_equival, '') = '', VALUES(grupo_equival), grupo_equival)",
                               (c, datos["grupo_equival"]))
                agrupados += cursor.rowcount > 0
        conn.commit()
        sugerencias_clase.olvidar()
    except Exception as e:
        conn.rollback()
        if es_fetch:
            return jsonify(ok=False, error=f"error de base de datos: {e}"), 500
        raise
    finally:
        cursor.close()
        conn.close()

    if es_fetch:
        return jsonify(ok=True, codigo=codigo, datos=datos, agrupados=agrupados)
    flash(f"✅ {codigo} guardado: {datos['segmento'] or '—'} / {datos['familia'] or '—'} / {datos['subfamilia'] or '—'}", "success")
    return volver_a(request.form.get("volver"))


@clase_prod_bp.route("/clase_prod/masivo", methods=["POST"])
@requiere_acceso("clase_prod")
def guardar_masivo():
    codigos = sorted({c.strip() for c in request.form.getlist("codigos") if c.strip()})
    datos = {c: limpiar(request.form.get(f"m_{c}")) for c in CAMPOS}
    cambiar = {c: v for c, v in datos.items() if v}  # vacío = no cambiar ese campo
    if not codigos or not cambiar:
        flash("⚠️ Seleccione productos y al menos un campo a asignar.", "warning")
        return volver_a(request.form.get("volver"))
    errores = validar("X", cambiar)
    if errores:
        flash("❌ No se guardó: " + "; ".join(errores), "danger")
        return volver_a(request.form.get("volver"))

    columnas = ", ".join(cambiar)
    actualizar = ", ".join(f"{c} = VALUES({c})" for c in cambiar)
    sql = (f"INSERT INTO clase_prod (codigo, {columnas}) VALUES (%s, {', '.join(['%s'] * len(cambiar))}) "
           f"ON DUPLICATE KEY UPDATE {actualizar}")

    conn = get_connection()
    cursor = conn.cursor()
    try:
        marcas = ", ".join(["%s"] * len(codigos))
        cursor.execute(f"SELECT `Codigo Producto` FROM catalogo_productos WHERE `Codigo Producto` IN ({marcas})", codigos)
        validos = [r[0] for r in cursor.fetchall()]
        cursor.executemany(sql, [(c, *cambiar.values()) for c in validos])
        conn.commit()
        sugerencias_clase.olvidar()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return volver_a(request.form.get("volver"))
    finally:
        cursor.close()
        conn.close()

    detalle = ", ".join(f"{c} = {v}" for c, v in cambiar.items())
    flash(f"✅ {len(validos)} productos actualizados ({detalle}).", "success")
    return volver_a(request.form.get("volver"))


# -----------------------------
# Importar desde el Excel de clasificación
# -----------------------------
def leer_clasificacion(archivo):
    requeridas = ["Codigo", "Segmento", "Familia", "Subfamilia", "Grupo_equiv"]
    aviso_hoja = ""
    try:
        filas = leer_filas(archivo, hoja="Productos")
    except ErrorArchivo as e:
        aviso_hoja = f"{e}; "  # ej. "no tiene la hoja 'Productos' (hojas: Hoja1)"
        archivo.seek(0)
        filas = leer_filas(archivo)
    try:
        inicio, col = buscar_encabezado(filas, requeridas)
    except ErrorArchivo:
        encabezados = {a_texto(v) for fila in filas[:40] for v in fila}
        faltan = [c for c in requeridas if c not in encabezados]
        raise ErrorArchivo(f"{aviso_hoja}faltan las columnas {', '.join(faltan)}. "
                           "Parece un catálogo sin clasificar: use el archivo CATALOGO_ORDENADO…xlsx")

    registros, errores, vistos = {}, [], {}
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        celda = lambda n: fila[col[n]] if col[n] < len(fila) else None
        codigo = a_texto(celda("Codigo"))
        if not codigo:
            continue
        datos = {"segmento": limpiar(a_texto(celda("Segmento"))), "familia": limpiar(a_texto(celda("Familia"))),
                 "subfamilia": limpiar(a_texto(celda("Subfamilia"))), "grupo_equival": limpiar(a_texto(celda("Grupo_equiv")))}
        problemas = validar(codigo, datos)
        if problemas:
            errores.append(f"Fila {num} ({codigo}): " + "; ".join(problemas))
            continue
        if codigo in vistos:
            errores.append(f"Fila {num}: {codigo} repetido (ya está en la fila {vistos[codigo]}); se ignora")
            continue
        vistos[codigo] = num
        registros[codigo] = datos
    if not registros:
        raise ErrorArchivo("no contiene productos clasificados")
    return registros, errores


@clase_prod_bp.route("/clase_prod/importar", methods=["GET", "POST"])
@requiere_acceso("clase_prod")
def importar():
    if request.method == "GET":
        return render_template("clase_prod_importar.html")

    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return render_template("clase_prod_importar.html", errores_archivo=["Seleccione el Excel de clasificación."])
    try:
        registros, errores = leer_clasificacion(archivo)
    except ErrorArchivo as e:
        return render_template("clase_prod_importar.html",
                               errores_archivo=[f"{archivo.filename}: {e}. Debe tener la hoja Productos con Codigo, Segmento, Familia, Subfamilia y Grupo_equiv."])
    except Exception as e:
        return render_template("clase_prod_importar.html", errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"])

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cursor.execute("SELECT codigo, segmento, familia, subfamilia, grupo_equival FROM clase_prod")
        actuales = {r["codigo"]: r for r in cursor.fetchall()}
        cursor.execute("SELECT `Codigo Producto` AS c FROM catalogo_productos WHERE `Codigo Producto` IS NOT NULL")
        catalogo = {r["c"] for r in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    nuevos, modificados, sin_cambio, por_columna = [], [], 0, {}
    for codigo, datos in registros.items():
        actual = actuales.get(codigo)
        if actual is None:
            nuevos.append(codigo)
            continue
        cambios = [(c, actual[c], datos[c]) for c in CAMPOS if (actual[c] or None) != datos[c]]
        if cambios:
            modificados.append({"codigo": codigo, "cambios": cambios})
            for c, _, _ in cambios:
                por_columna[c] = por_columna.get(c, 0) + 1
        else:
            sin_cambio += 1

    advertencias = []
    fuera = sorted(set(registros) - catalogo)
    if fuera:
        advertencias.append(f"{len(fuera)} códigos no existen en catalogo_productos (se guardan igual): " + ", ".join(fuera[:15]) + (" …" if len(fuera) > 15 else ""))
    solo_bd = len(set(actuales) - set(registros))
    if solo_bd:
        advertencias.append(f"{solo_bd} códigos clasificados en la base no vienen en el Excel; se conservan sin cambios.")

    a_guardar = {c: registros[c] for c in nuevos + [m["codigo"] for m in modificados]}
    token = None
    if a_guardar:
        limpiar_temporales()
        token = uuid.uuid4().hex
        with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as fh:
            json.dump(a_guardar, fh, ensure_ascii=False)
        session["carga_clase_prod_token"] = token

    return render_template(
        "clase_prod_importar.html", validado=True, token=token, total=len(registros),
        nuevos=nuevos, modificados=modificados, sin_cambio=sin_cambio,
        por_columna=sorted(por_columna.items(), key=lambda x: -x[1]),
        errores=errores, advertencias=advertencias,
    )


@clase_prod_bp.route("/clase_prod/importar/confirmar", methods=["POST"])
@requiere_acceso("clase_prod")
def confirmar_importar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_clase_prod_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("clase_prod.importar"))

    with open(ruta, encoding="utf-8") as fh:
        registros = json.load(fh)
    filas = [(c, d["segmento"], d["familia"], d["subfamilia"], d["grupo_equival"]) for c, d in registros.items()]

    conn = get_connection()
    cursor = conn.cursor()
    try:
        for i in range(0, len(filas), LOTE):
            cursor.executemany("""
                INSERT INTO clase_prod (codigo, segmento, familia, subfamilia, grupo_equival)
                VALUES (%s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE segmento = VALUES(segmento), familia = VALUES(familia),
                                        subfamilia = VALUES(subfamilia), grupo_equival = VALUES(grupo_equival)
            """, filas[i:i + LOTE])
        conn.commit()
        sugerencias_clase.olvidar()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("clase_prod.importar"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_clase_prod_token", None)
    flash(f"✅ Clasificación importada: {len(filas):,} productos nuevos o actualizados.", "success")
    return redirect(url_for("clase_prod.clase_prod"))
