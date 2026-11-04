"""Cierre de mes de inventario: saldo por bodega al último día de cada mes (unidades, costo C$ y costo U$).

Se carga el "Reporte de Análisis de Inventario" del sistema cortado al fin de mes (/cargar_cierre_inventario); se guarda el saldo
por bodega (inv_cierre_bodega) y el detalle por producto (inv_cierre_detalle). El reporte (/cierre_inventario) muestra los
últimos meses lado a lado, separado en bodegas, bodegas a revisar y bodegas a liquidar, igual que el Excel que se hacía a mano.
La sección y el orden de cada bodega están en inv_bodega_grupo (se editan en /cierre_inventario/bodegas)."""
import io
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta
from decimal import Decimal

from flask import Blueprint, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.cargar_catalogo import sin_tildes, LOTE_INSERT
from router.cargar_ventas import ErrorArchivo, ruta_temporal, limpiar_temporales
from router.cargar_existencias import leer_existencias

cierre_inventario_bp = Blueprint("cierre_inventario", __name__)

PREFIJO_TEMP = "carga_cierre_inv_"
TC_DEFECTO = Decimal("36.70")
VERDE, NARANJA = "027266", "FE6412"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MESES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"]
# Secciones del reporte, en orden; None = bodegas que aún no tienen sección asignada
GRUPOS = [("BODEGAS", "Saldo de inventario por bodega"), ("REVISAR", "Bodegas a revisar"), ("LIQUIDAR", "Bodegas a liquidar"),
          (None, "Otras bodegas (sin clasificar)")]
NOMBRE_GRUPO = {g: n for g, n in GRUPOS}
OPCIONES_MESES = [2, 3, 4, 6, 12]


def clave_bodega(nombre):
    """Mismo nombre aunque cambien espacios, guiones o puntos: 'BODEGA #5 MONTOYA' == 'BODEGA #5 - MONTOYA'."""
    return re.sub(r"[^A-Z0-9]", "", sin_tildes(nombre or "").upper())


def fin_de_mes(d):
    return (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)


def mes_txt(d, corto=False):
    return f"{MESES_CORTO[d.month - 1]}-{d:%y}" if corto else f"{MESES[d.month - 1]} {d.year}"


# -----------------------------
# Base de datos
# -----------------------------
def lista_cierres(cursor):
    cursor.execute("""SELECT c.fecha_corte, c.tc, c.origen, c.archivo, c.cargado_por, c.cargado,
                             COUNT(b.bodega) AS bodegas, COALESCE(SUM(b.unidades), 0) AS unidades, COALESCE(SUM(b.costo_mn), 0) AS costo_mn
                      FROM inv_cierre c LEFT JOIN inv_cierre_bodega b ON b.fecha_corte = c.fecha_corte
                      GROUP BY c.fecha_corte, c.tc, c.origen, c.archivo, c.cargado_por, c.cargado ORDER BY c.fecha_corte""")
    return cursor.fetchall()


def grupos_bodega(cursor):
    cursor.execute("SELECT clave, bodega, grupo, orden FROM inv_bodega_grupo")
    return {f["clave"]: f for f in cursor.fetchall()}


def guardar_cierre(conn, fecha_corte, tc, origen, archivo, usuario, bodegas, detalle=None):
    """Reemplaza el cierre de esa fecha. bodegas: {nombre: (unidades, costo_mn, productos)}; detalle: filas para inv_cierre_detalle.
    No hace commit: quien llama maneja la transacción."""
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM inv_cierre WHERE fecha_corte = %s", (fecha_corte,))   # borra bodegas y detalle (cascade)
        cursor.execute("INSERT INTO inv_cierre (fecha_corte, tc, origen, archivo, cargado_por) VALUES (%s, %s, %s, %s, %s)",
                       (fecha_corte, tc, origen, archivo, usuario))
        cursor.executemany("INSERT INTO inv_cierre_bodega (fecha_corte, bodega, unidades, costo_mn, productos) VALUES (%s, %s, %s, %s, %s)",
                           [(fecha_corte, b, round(u, 4), round(m, 4), p) for b, (u, m, p) in bodegas.items()])
        if detalle:
            sql = ("INSERT INTO inv_cierre_detalle (fecha_corte, bodega, id_producto, codigo, producto, numero_parte, marca, uom, unidades, costo_mn) "
                   "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)")
            filas = [(fecha_corte, *f) for f in detalle]
            for i in range(0, len(filas), LOTE_INSERT):
                cursor.executemany(sql, filas[i:i + LOTE_INSERT])
    finally:
        cursor.close()


# -----------------------------
# Reporte
# -----------------------------
def armar_reporte(cursor, hasta, meses):
    """Últimos `meses` cierres hasta la fecha `hasta`, con las bodegas en sus secciones."""
    todos = lista_cierres(cursor)
    hasta = hasta or (todos[-1]["fecha_corte"] if todos else None)
    elegidos = [c for c in todos if c["fecha_corte"] <= hasta][-meses:] if hasta else []
    if not elegidos:
        return None
    fechas = [c["fecha_corte"] for c in elegidos]
    tcs = {c["fecha_corte"]: float(c["tc"]) for c in elegidos}

    grupos = grupos_bodega(cursor)
    cursor.execute(f"SELECT fecha_corte, bodega, unidades, costo_mn FROM inv_cierre_bodega WHERE fecha_corte IN ({', '.join(['%s'] * len(fechas))})",
                   fechas)
    bodegas = {}
    for f in cursor.fetchall():
        k = clave_bodega(f["bodega"])
        b = bodegas.setdefault(k, {"clave": k, "nombre": f["bodega"], "ultima": f["fecha_corte"], "meses": {}})
        if f["fecha_corte"] >= b["ultima"]:   # se muestra el nombre más reciente
            b["nombre"], b["ultima"] = f["bodega"], f["fecha_corte"]
        m = b["meses"].setdefault(f["fecha_corte"], {"unidades": 0.0, "mn": 0.0})
        m["unidades"] += float(f["unidades"])
        m["mn"] += float(f["costo_mn"])

    def celdas(valores):
        """valores: {fecha: {unidades, mn}} -> lista por mes con el U$ al tipo de cambio de cada cierre."""
        salida = []
        for fe in fechas:
            v = valores.get(fe)
            salida.append(None if v is None else {"unidades": v["unidades"], "mn": v["mn"], "me": v["mn"] / tcs[fe] if tcs[fe] else 0})
        return salida

    def variacion(cel):
        if len(cel) < 2:
            return None
        a, b = cel[-2], cel[-1]
        ant, act = (a or {}).get("mn", 0.0), (b or {}).get("mn", 0.0)
        return {"mn": act - ant, "pct": (act / ant - 1) * 100 if ant > 0 and act >= 0 else None}

    secciones, linea = [], 0
    for g, titulo in GRUPOS:
        filas = []
        for k, b in bodegas.items():
            info = grupos.get(k)
            if (info["grupo"] if info else None) != g:
                continue
            filas.append({"clave": k, "nombre": info["bodega"] if info else b["nombre"], "orden": info["orden"] if info else 999,
                          "valores": b["meses"]})
        if not filas:
            continue
        filas.sort(key=lambda f: (f["orden"], -f["valores"].get(fechas[-1], {}).get("mn", 0), f["nombre"]))
        total = {}
        for f in filas:
            linea += 1
            f["linea"] = linea
            f["celdas"] = celdas(f["valores"])
            f["var"] = variacion(f["celdas"])
            for fe, v in f["valores"].items():
                t = total.setdefault(fe, {"unidades": 0.0, "mn": 0.0})
                t["unidades"] += v["unidades"]
                t["mn"] += v["mn"]
        sub = celdas(total)
        secciones.append({"grupo": g, "titulo": titulo, "filas": filas, "total": sub, "var": variacion(sub)})

    general = {}
    for s in secciones:
        for fe, v in zip(fechas, s["total"]):
            if v:
                t = general.setdefault(fe, {"unidades": 0.0, "mn": 0.0})
                t["unidades"] += v["unidades"]
                t["mn"] += v["mn"]
    total = celdas(general)

    # Comparación del último cierre con el anterior y con el mismo mes del año pasado (aunque no estén en pantalla)
    por_fecha = {c["fecha_corte"]: c for c in todos}
    ultimo = elegidos[-1]
    previos = [c for c in todos if c["fecha_corte"] < ultimo["fecha_corte"]]
    anterior = previos[-1] if previos else None
    anio = next((c for fe, c in por_fecha.items() if fe.year == ultimo["fecha_corte"].year - 1 and fe.month == ultimo["fecha_corte"].month), None)

    def comp(c):
        if not c:
            return None
        ant, act = float(c["costo_mn"]), float(ultimo["costo_mn"])
        return {"fecha": c["fecha_corte"], "mn": ant, "dif": act - ant, "pct": (act / ant - 1) * 100 if ant else None}

    # Bodegas que más subieron y bajaron (C$) contra el cierre anterior en pantalla
    movimientos = sorted((f for s in secciones for f in s["filas"] if f["var"] and abs(f["var"]["mn"]) >= 0.005),
                         key=lambda f: -abs(f["var"]["mn"]))
    return {
        "fechas": fechas, "tcs": tcs, "cierres": elegidos, "secciones": secciones, "total": total, "var": variacion(total),
        "ultimo": ultimo, "tc": float(ultimo["tc"]), "vs_anterior": comp(anterior), "vs_anio": comp(anio),
        "sin_clasificar": sum(len(s["filas"]) for s in secciones if s["grupo"] is None),
        "suben": [f for f in movimientos if f["var"]["mn"] > 0][:3], "bajan": [f for f in movimientos if f["var"]["mn"] < 0][:3],
    }


def parametros(args, cierres):
    fechas = [c["fecha_corte"] for c in cierres]
    hasta = None
    try:
        hasta = datetime.strptime(args.get("hasta", ""), "%Y-%m-%d").date()
    except ValueError:
        pass
    if hasta not in fechas:
        hasta = fechas[-1] if fechas else None
    try:
        meses = int(args.get("meses", 4))
    except ValueError:
        meses = 4
    return hasta, min(max(meses, 1), 24)


@cierre_inventario_bp.route("/cierre_inventario")
@requiere_acceso("cierre_inventario")
def reporte():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cierres = lista_cierres(cursor)
        hasta, meses = parametros(request.args, cierres)
        r = armar_reporte(cursor, hasta, meses) if cierres else None
    finally:
        cursor.close()
        conn.close()
    fechas = [c["fecha_corte"] for c in cierres]
    i = fechas.index(hasta) if hasta in fechas else None
    return render_template("cierre_inventario.html", r=r, cierres=cierres, hasta=hasta, meses=meses, opciones_meses=OPCIONES_MESES,
                           anterior=fechas[i - 1] if i else None, siguiente=fechas[i + 1] if i is not None and i + 1 < len(fechas) else None,
                           mes_txt=mes_txt, nombre_grupo=NOMBRE_GRUPO)


@cierre_inventario_bp.route("/cierre_inventario/excel")
@requiere_acceso("cierre_inventario")
def excel():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        cierres = lista_cierres(cursor)
        hasta, meses = parametros(request.args, cierres)
        r = armar_reporte(cursor, hasta, meses) if cierres else None
    finally:
        cursor.close()
        conn.close()
    if not r:
        flash("⚠️ No hay cierres de inventario cargados.", "warning")
        return redirect(url_for("cierre_inventario.reporte"))

    fechas = r["fechas"]
    wb = Workbook()
    ws = wb.active
    ws.title = f"Cierre {fechas[-1]:%b %Y}"
    n_col = 2 + 3 * len(fechas)
    delgado = Side(style="thin", color="BFBFBF")
    borde = Border(left=delgado, right=delgado, top=delgado, bottom=delgado)
    fmt_u = '#,##0.00;[Red]-#,##0.00'
    fmt_c = '"C$" #,##0.00;[Red]-"C$" #,##0.00;"C$" -'
    fmt_d = '"$" #,##0.00;[Red]-"$" #,##0.00;"$" -'
    titulo = f"SALDO DE INVENTARIO POR BODEGA {mes_txt(fechas[0]).upper()} - {mes_txt(fechas[-1]).upper()}"
    if len(fechas) == 1:
        titulo = f"SALDO DE INVENTARIO POR BODEGA {mes_txt(fechas[0]).upper()}"
    ws.cell(row=1, column=1, value=titulo).font = Font(bold=True, size=13, color=VERDE)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_col)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=2, column=1, value="Costo U$ = costo C$ ÷ tipo de cambio del cierre (" +
            ", ".join(f"{mes_txt(f, True)}: {r['tcs'][f]:.2f}" for f in fechas) + ")").font = Font(size=9, color="52514E")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_col)
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

    fila = 4
    totales_filas = []

    def encabezado(fila, titulo_seccion, unidades_txt):
        if titulo_seccion:
            ws.cell(row=fila, column=2, value=titulo_seccion.upper()).font = Font(bold=True, color=VERDE)
        for j, fe in enumerate(fechas):
            col = 3 + 3 * j
            c = ws.cell(row=fila, column=col, value=mes_txt(fe, True))
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=VERDE)
            c.alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col + 2)
        for i, txt in enumerate(["LINEA", "BODEGA"] + [unidades_txt, "COSTO MN C$", "COSTO ME U$"] * len(fechas), start=1):
            c = ws.cell(row=fila + 1, column=i, value=txt)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NARANJA)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde
        return fila + 2

    for s in r["secciones"]:
        primera = s is r["secciones"][0]
        fila = encabezado(fila, None if primera else s["titulo"], "UNIDADES" if primera else "EXISTENCIAS")
        ini = fila
        for f in s["filas"]:
            ws.cell(row=fila, column=1, value=f["linea"]).alignment = Alignment(horizontal="center")
            ws.cell(row=fila, column=2, value=f["nombre"])
            for j, v in enumerate(f["celdas"]):
                col = 3 + 3 * j
                ws.cell(row=fila, column=col, value=round(v["unidades"], 4) if v else 0).number_format = fmt_u
                ws.cell(row=fila, column=col + 1, value=round(v["mn"], 2) if v else 0).number_format = fmt_c
                ws.cell(row=fila, column=col + 2, value=f"={get_column_letter(col + 1)}{fila}/{r['tcs'][fechas[j]]:g}").number_format = fmt_d
            for i in range(1, n_col + 1):
                ws.cell(row=fila, column=i).border = borde
            fila += 1
        for i in range(3, n_col + 1):
            letra = get_column_letter(i)
            c = ws.cell(row=fila, column=i, value=f"=SUM({letra}{ini}:{letra}{fila - 1})")
            c.font = Font(bold=True)
            c.number_format = [fmt_u, fmt_c, fmt_d][(i - 3) % 3]
            c.border = Border(top=Side(style="medium", color=VERDE))
        ws.cell(row=fila, column=2, value=f"Total {s['titulo'].lower()}").font = Font(bold=True)
        totales_filas.append(fila)
        fila += 2

    ws.cell(row=fila, column=2, value="TOTAL").font = Font(bold=True, size=12)
    for i in range(3, n_col + 1):
        letra = get_column_letter(i)
        c = ws.cell(row=fila, column=i, value="=" + "+".join(f"{letra}{t}" for t in totales_filas))
        c.font = Font(bold=True, size=12, color=VERDE)
        c.number_format = [fmt_u, fmt_c, fmt_d][(i - 3) % 3]
        c.fill = PatternFill("solid", fgColor="E9F4F2")

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 32
    for j in range(len(fechas)):
        for k, ancho in enumerate((13, 18, 15)):
            ws.column_dimensions[get_column_letter(3 + 3 * j + k)].width = ancho
    ws.freeze_panes = "C4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    nombre = f"cierre_inventario_{fechas[0]:%Y-%m}_a_{fechas[-1]:%Y-%m}.xlsx" if len(fechas) > 1 else f"cierre_inventario_{fechas[-1]:%Y-%m}.xlsx"
    return send_file(salida, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------
# Sección de cada bodega
# -----------------------------
@cierre_inventario_bp.route("/cierre_inventario/bodegas", methods=["GET", "POST"])
@requiere_acceso("cargar_cierre_inventario")
def bodegas():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        if request.method == "POST":
            claves = request.form.getlist("clave")
            filas = []
            for k in claves:
                grupo = request.form.get(f"grupo_{k}") or ""
                if grupo not in ("BODEGAS", "REVISAR", "LIQUIDAR"):
                    continue
                try:
                    orden = int(request.form.get(f"orden_{k}") or 999)
                except ValueError:
                    orden = 999
                nombre = (request.form.get(f"nombre_{k}") or "").strip()[:100]
                filas.append((k, nombre or k, grupo, orden))
            sin_grupo = [k for k in claves if (request.form.get(f"grupo_{k}") or "") not in ("BODEGAS", "REVISAR", "LIQUIDAR")]
            try:
                conn.start_transaction()
                cursor.executemany("REPLACE INTO inv_bodega_grupo (clave, bodega, grupo, orden) VALUES (%s, %s, %s, %s)", filas)
                if sin_grupo:
                    cursor.execute(f"DELETE FROM inv_bodega_grupo WHERE clave IN ({', '.join(['%s'] * len(sin_grupo))})", sin_grupo)
                conn.commit()
                flash(f"✅ Secciones guardadas ({len(filas)} bodegas clasificadas).", "success")
            except Exception as e:
                conn.rollback()
                flash(f"❌ No se guardaron los cambios: {e}", "danger")
            return redirect(url_for("cierre_inventario.bodegas"))

        grupos = grupos_bodega(cursor)
        cursor.execute("SELECT bodega, fecha_corte, unidades, costo_mn FROM inv_cierre_bodega ORDER BY fecha_corte")
        vistas, ultima = {}, None
        for f in cursor.fetchall():   # queda el dato del cierre más reciente de cada bodega
            vistas[clave_bodega(f["bodega"])] = {"nombre": f["bodega"], "ultima": f["fecha_corte"], "mn": float(f["costo_mn"])}
            ultima = f["fecha_corte"]
    finally:
        cursor.close()
        conn.close()

    filas = []
    for k in set(grupos) | set(vistas):
        g, v = grupos.get(k), vistas.get(k)
        filas.append({"clave": k, "nombre": g["bodega"] if g else v["nombre"], "nombre_archivo": v["nombre"] if v else None,
                      "grupo": g["grupo"] if g else "", "orden": g["orden"] if g else "",
                      "ultima": v["ultima"] if v else None, "activa": bool(v and v["ultima"] == ultima),
                      "mn": v["mn"] if v and v["ultima"] == ultima else None})
    orden_grupo = {"BODEGAS": 0, "REVISAR": 1, "LIQUIDAR": 2, "": 3}
    filas.sort(key=lambda f: (not f["activa"], orden_grupo[f["grupo"]], f["orden"] if f["orden"] != "" else 999, f["nombre"]))
    return render_template("cierre_inventario_bodegas.html", filas=filas, ultima=ultima, nombre_grupo=NOMBRE_GRUPO, mes_txt=mes_txt)


# -----------------------------
# Carga del cierre
# -----------------------------
def resumir(registros):
    """Saldo por bodega y detalle (sin filas en cero) del Reporte de Análisis de Inventario."""
    bodegas, detalle = {}, []
    for r in registros:
        u, m = r["Cant_Final"] or 0, r["Monto Final"] or 0
        b = bodegas.setdefault(r["Bodega"], [0.0, 0.0, 0])
        b[0] += u
        b[1] += m
        if abs(u) > 1e-9 or abs(m) > 0.004:
            b[2] += 1
            detalle.append([r["Bodega"], int(r["Id Producto"]) if r["Id Producto"] is not None else None, r["Codigo_producto"],
                            (r["Producto"] or "")[:255] or None, (r["Número Parte"] or "")[:100] or None, (r["Marca"] or "")[:150] or None,
                            (r["UOM"] or "")[:20] or None, round(u, 4), round(m, 4)])
    return bodegas, detalle


def _contexto_carga(cursor):
    cierres = lista_cierres(cursor)
    return {"cierres": list(reversed(cierres))[:14], "tc_defecto": float(cierres[-1]["tc"]) if cierres else float(TC_DEFECTO),
            "mes_txt": mes_txt}


@cierre_inventario_bp.route("/cargar_cierre_inventario", methods=["GET", "POST"])
@requiere_acceso("cargar_cierre_inventario")
def cargar():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    try:
        ctx = _contexto_carga(cursor)
        if request.method == "GET":
            return render_template("cargar_cierre_inventario.html", **ctx)

        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename:
            return render_template("cargar_cierre_inventario.html", errores_archivo=["Falta el Reporte de Análisis de Inventario."], **ctx)
        try:
            tc = Decimal((request.form.get("tc") or "").replace(",", "."))
            if not Decimal("1") <= tc <= Decimal("1000"):
                raise ValueError
        except (ValueError, ArithmeticError):
            return render_template("cargar_cierre_inventario.html", errores_archivo=["El tipo de cambio no es válido."], **ctx)
        try:
            registros, errores, advertencias, periodo = leer_existencias(archivo)
        except ErrorArchivo as e:
            return render_template("cargar_cierre_inventario.html",
                                   errores_archivo=[f"{archivo.filename}: {e}. ¿Seleccionó el Reporte de Análisis de Inventario?"], **ctx)
        except Exception as e:
            return render_template("cargar_cierre_inventario.html", errores_archivo=[f"{archivo.filename}: error inesperado al leer – {e}"], **ctx)

        # Fecha de corte: la escrita en el formulario o el "Al:" del encabezado del reporte
        fecha_corte = None
        for texto in (request.form.get("fecha_corte"), periodo[1] if periodo else None):
            try:
                fecha_corte = datetime.strptime((texto or "")[:10], "%Y-%m-%d").date()
                break
            except ValueError:
                continue
        if not fecha_corte:
            return render_template("cargar_cierre_inventario.html",
                                   errores_archivo=["El archivo no dice hasta qué fecha es (línea «Desde: … Al: …»). Escriba la fecha de corte."], **ctx)
        if fecha_corte != fin_de_mes(fecha_corte):
            advertencias.insert(0, f"La fecha de corte {fecha_corte:%d/%m/%Y} no es fin de mes: el reporte la mostrará como «{mes_txt(fecha_corte)}».")
        if fecha_corte > date.today():
            advertencias.insert(0, f"La fecha de corte {fecha_corte:%d/%m/%Y} es posterior a hoy.")
        if periodo and periodo[0] > "2021-01-01":
            advertencias.append(f"El reporte empieza el {periodo[0]}: para el saldo completo debe salir desde el inicio (2021-01-01).")

        bodegas, detalle = resumir(registros)
        grupos = grupos_bodega(cursor)
        cierres = lista_cierres(cursor)
        existente = next((c for c in cierres if c["fecha_corte"] == fecha_corte), None)
        previos = [c for c in cierres if c["fecha_corte"] < fecha_corte]
        anterior = previos[-1] if previos else None
        ant_bodegas = {}
        if anterior:
            cursor.execute("SELECT bodega, unidades, costo_mn FROM inv_cierre_bodega WHERE fecha_corte = %s", (anterior["fecha_corte"],))
            for f in cursor.fetchall():
                ant_bodegas[clave_bodega(f["bodega"])] = float(f["costo_mn"])
    finally:
        cursor.close()
        conn.close()

    filas = []
    for nombre, (u, m, p) in bodegas.items():
        k = clave_bodega(nombre)
        g = grupos.get(k)
        filas.append({"bodega": nombre, "grupo": g["grupo"] if g else None, "orden": g["orden"] if g else 999, "unidades": u, "mn": m,
                      "productos": p, "anterior": ant_bodegas.get(k)})
    orden_grupo = {"BODEGAS": 0, "REVISAR": 1, "LIQUIDAR": 2, None: 3}
    filas.sort(key=lambda f: (orden_grupo[f["grupo"]], f["orden"], -f["mn"]))
    nuevas = [f["bodega"] for f in filas if f["grupo"] is None]
    if nuevas:
        advertencias.append(f"{len(nuevas)} bodega(s) sin sección en el reporte: {', '.join(nuevas)}. Se mostrarán en «Otras bodegas»; "
                            "puede asignarles sección después de cargar.")
    faltan = [k for k in ant_bodegas if k not in {clave_bodega(b) for b in bodegas} and abs(ant_bodegas[k]) > 0.005]
    if faltan:
        advertencias.append(f"{len(faltan)} bodega(s) del cierre anterior no vienen en el archivo; verifique que el reporte no tenga filtro de bodega.")

    limpiar_temporales()
    token = uuid.uuid4().hex
    with open(ruta_temporal(token, PREFIJO_TEMP), "w", encoding="utf-8") as f:
        json.dump({"fecha_corte": fecha_corte.isoformat(), "tc": str(tc), "archivo": archivo.filename[:255],
                   "bodegas": {b: v for b, v in bodegas.items()}, "detalle": detalle}, f, ensure_ascii=False)
    session["carga_cierre_inv_token"] = token

    total_mn = sum(f["mn"] for f in filas)
    return render_template("cargar_cierre_inventario.html", validado=True, token=token, fecha_corte=fecha_corte, tc=float(tc),
                           filas=filas, total_mn=total_mn, total_unidades=sum(f["unidades"] for f in filas),
                           total_productos=sum(f["productos"] for f in filas), lineas=len(registros),
                           existente=existente, anterior=anterior, errores=errores, advertencias=advertencias,
                           nombre_grupo=NOMBRE_GRUPO, **ctx)


@cierre_inventario_bp.route("/cargar_cierre_inventario/confirmar", methods=["POST"])
@requiere_acceso("cargar_cierre_inventario")
def confirmar():
    token = request.form.get("token")
    ruta = ruta_temporal(token, PREFIJO_TEMP)
    if not ruta or token != session.get("carga_cierre_inv_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir el archivo.", "warning")
        return redirect(url_for("cierre_inventario.cargar"))
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    fecha_corte = date.fromisoformat(datos["fecha_corte"])

    conn = get_connection()
    try:
        conn.start_transaction()
        guardar_cierre(conn, fecha_corte, Decimal(datos["tc"]), "sistema", datos["archivo"], session.get("usuario"),
                       {b: tuple(v) for b, v in datos["bodegas"].items()}, datos["detalle"])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó el cierre. Error de base de datos: {e}", "danger")
        return redirect(url_for("cierre_inventario.cargar"))
    finally:
        conn.close()

    os.remove(ruta)
    session.pop("carga_cierre_inv_token", None)
    flash(f"✅ Cierre de {mes_txt(fecha_corte)} guardado: {len(datos['bodegas'])} bodegas y {len(datos['detalle']):,} productos.", "success")
    return redirect(url_for("cierre_inventario.reporte", hasta=fecha_corte.isoformat()))
