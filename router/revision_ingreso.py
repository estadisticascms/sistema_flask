"""Revisión de facturas de ingreso de mercadería (antes de ingresarlas al sistema).

Se sube la factura del proveedor (PROVEEDOR, FACTURA, ITEM, Nº PARTE, MARCA, ORIGEN, U/M, CANT, P.UNIT, TOTAL; si trae COM o ARANCEL
también se revisan) y cada línea se busca en el catálogo para devolver código (COM), descripción Comersa y arancel (código SAC):
1) nº de parte exacto (sin espacios, mayúsculas); si hay varios códigos se prefiere la misma marca;
2) nº de parte sin guiones, barras ni puntos (posible coincidencia, a confirmar);
3) nº de parte dentro de las equivalencias de otro producto (posible, a confirmar);
4) si no aparece: PRODUCTO NUEVO.
Se marcan duplicados (en el catálogo y dentro de la misma factura), se completa código, descripción y código SAC (arancel)
y se compara el P.UNIT nuevo contra el último precio FOB del catálogo."""
import io
import json
import os
import re
import uuid
from datetime import date, datetime

from flask import Blueprint, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils.db import get_connection
from router.accesos import requiere_acceso
from router.ventas_bodegas import en_cache

revision_ingreso_bp = Blueprint("revision_ingreso", __name__)

CARPETA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "uploads")
VERDE, NARANJA = "027266", "FE6412"
ALERTA_FOB = 10.0   # % de variación del FOB que se resalta
EXTREMO_FOB = 200.0  # más que esto casi siempre es otra unidad de medida (caja vs unidad): no entra en el total
ESTADOS = {"encontrado": "Encontrado", "posible": "Posible (confirmar)", "duplicado": "Duplicado en catálogo", "nuevo": "Producto nuevo"}
# Encabezados de la plantilla -> campo (se buscan por palabra clave, sin importar mayúsculas)
COLUMNAS = [("proveedor", ("PROVEEDOR",)), ("factura", ("FACTURA",)), ("item", ("ITEM", "TEM")), ("com", ("COM",)),
            ("parte", ("PARTE",)), ("descripcion", ("DESCRIPCION", "DESCRIPCIÓN")), ("marca", ("MARCA",)), ("origen", ("ORIGEN",)),
            ("um", ("U/M", "UM", "UNIDAD")), ("cantidad", ("CANT",)), ("precio", ("P.UNIT", "P. UNIT", "PUNIT", "PRECIO")),
            ("total", ("TOTAL",)), ("arancel", ("ARANCEL", "SAC"))]


def exacto(parte):
    return "".join(str(parte or "").upper().split())


def compacto(parte):
    return re.sub(r"[^A-Z0-9]", "", str(parte or "").upper())


def texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def numero(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("$", "").replace("U", "").replace(",", "").strip())
    except ValueError:
        return None


# ---------------------------------------------------------------- catálogo indexado
def indice_catalogo():
    def leer():
        conn = get_connection()
        cursor = conn.cursor(dictionary=True)
        try:
            cursor.execute("""SELECT `Codigo Producto` AS codigo, Nombre AS nombre, `Numero Parte` AS parte, Marca AS marca, `Codigo SAC` AS sac,
                                     `Ultimo Precio FOB` AS fob, `Fecha Ultima Compra` AS fecha_compra, `Proveedor Ultima Compra` AS proveedor,
                                     `Ultimo Costo` AS ultimo_costo, Equivalencias AS equivalencias, `Unidad Medida` AS um
                              FROM catalogo_productos""")
            productos = cursor.fetchall()
            cursor.execute("SELECT Codigo_producto AS codigo, SUM(Cant_Final) AS e FROM existencias_bodega GROUP BY Codigo_producto")
            existencia = {r["codigo"]: float(r["e"] or 0) for r in cursor.fetchall()}
        finally:
            cursor.close()
            conn.close()
        por_exacto, por_compacto, por_codigo = {}, {}, {}
        for p in productos:
            p["existencia"] = existencia.get(p["codigo"], 0.0)
            p["fob"] = float(p["fob"]) if p["fob"] is not None else None
            p["fecha_compra"] = p["fecha_compra"].isoformat() if p["fecha_compra"] else None
            p["equivalencias_c"] = compacto(p["equivalencias"]) if p["equivalencias"] else ""
            p.pop("equivalencias", None)
            por_codigo[(p["codigo"] or "").upper()] = p
            if exacto(p["parte"]):
                por_exacto.setdefault(exacto(p["parte"]), []).append(p)
                por_compacto.setdefault(compacto(p["parte"]), []).append(p)
        return {"exacto": por_exacto, "compacto": por_compacto, "codigo": por_codigo, "productos": productos}
    return en_cache("revision_ingreso_catalogo", leer, segundos=600)


def encabezado(c):
    """Texto del encabezado sin espacios, saltos de línea ni símbolos: 'M A R C A', 'N*PARTE 01' → 'MARCA', 'NPARTE01'."""
    t = texto(c).upper()
    for a_, b_ in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")):
        t = t.replace(a_, b_)
    return re.sub(r"[^A-Z0-9/.]", "", t)


def leer_plantilla(archivo):
    """Busca en las hojas la fila de encabezados (la que tiene PROVEEDOR y PARTE) y mapea columnas por palabra clave."""
    wb = load_workbook(archivo, data_only=True, read_only=True)
    hojas = [(ws.title, list(ws.iter_rows(values_only=True))) for ws in wb.worksheets]
    wb.close()
    def tiene_encabezado(fila):
        enc = [encabezado(c) for c in fila]
        return any("PROVEEDOR" in e for e in enc) and any("PARTE" in e for e in enc)
    # La hoja que tenga la fila de encabezados, no necesariamente la primera
    hoja, filas = next(((t, h) for t, h in hojas if any(tiene_encabezado(f) for f in h[:50])), hojas[0] if hojas else ("", []))
    inicio, mapa, nombres = None, {}, {}
    for i, fila in enumerate(filas[:50]):
        if not tiene_encabezado(fila):
            continue
        inicio = i
        enc = [encabezado(c) for c in fila]
        for campo, claves in COLUMNAS:
            for j, e in enumerate(enc):
                if j in mapa.values() or not e:
                    continue
                if campo == "com" and e != "COM" and not e.startswith("COD"):
                    continue   # "COM" debe ser la columna exacta (no "DESCRIPCION COMERSA")
                if campo == "total" and e != "TOTAL":
                    continue
                if any(k.replace(" ", "") in e for k in claves):
                    mapa[campo] = j
                    nombres[campo] = " ".join(texto(fila[j]).split())
                    break
        break
    if inicio is None or "parte" not in mapa:
        raise ValueError("No se encontró la fila de encabezados (debe tener PROVEEDOR y Nº PARTE).")
    lineas, anterior = [], {}
    for n, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        val = {campo: (fila[j] if j < len(fila) else None) for campo, j in mapa.items()}
        if not texto(val.get("parte")) and not texto(val.get("com")):
            continue
        # Proveedor, factura, marca y origen se arrastran de la fila anterior si vienen vacíos (celdas combinadas)
        for campo in ("proveedor", "factura", "marca", "origen"):
            if texto(val.get(campo)):
                anterior[campo] = val[campo]
            else:
                val[campo] = anterior.get(campo)
        lineas.append({"fila": n, "proveedor": texto(val.get("proveedor")), "factura": texto(val.get("factura")), "item": texto(val.get("item")),
                       "com": texto(val.get("com")).upper(), "parte": texto(val.get("parte")), "descripcion": texto(val.get("descripcion")),
                       "marca": " ".join(texto(val.get("marca")).upper().split()), "origen": texto(val.get("origen")), "um": texto(val.get("um")),
                       "cantidad": numero(val.get("cantidad")), "precio": numero(val.get("precio")), "total": numero(val.get("total")),
                       "arancel": texto(val.get("arancel"))})
    return lineas, {"hoja": hoja, "fila": inicio + 1, "columnas": nombres}


GENERICOS = {"INC", "SA", "S.A", "LLC", "LTD", "LTDA", "CO", "CORP", "CORPORATION", "DE", "CV", "SRL", "GROUP", "THE", "Y", "CIA", "COMPANY"}


def proveedor_clave(nombre):
    """'HYGLOX INC', 'HYGLOX, INC.' → 'HYGLOX' (sin puntuación ni palabras genéricas)."""
    palabras = re.sub(r"[^A-Z0-9 ]", " ", (nombre or "").upper()).split()
    return " ".join(p for p in palabras if p not in GENERICOS)


def mismo_proveedor(a, b):
    a, b = proveedor_clave(a), proveedor_clave(b)
    return bool(a and b and (a == b or a in b or b in a))


def misma_marca(a, b):
    """Misma marca, o la misma familia (SKF y SKF-GENA). "CL/SKF" no es SKF."""
    a, b = (a or "").upper().strip(), (b or "").upper().strip()
    return bool(a and b and (a == b or a.startswith(b + "-") or b.startswith(a + "-")))


def revisar(lineas):
    """Busca cada línea por Nº PARTE + MARCA del documento. Productos de otras marcas nunca se toman como resultado:
    si el número no existe en esa marca es PRODUCTO NUEVO (y solo se avisa que el número existe en otras marcas)."""
    idx = indice_catalogo()
    repetidas = {}
    for x in lineas:
        repetidas.setdefault(exacto(x["parte"]), []).append(x["item"] or str(x["fila"]))
    for x in lineas:
        clave_e, clave_c = exacto(x["parte"]), compacto(x["parte"])
        # Con marca: solo productos de esa marca. Sin marca: los del mismo proveedor (si hay), si no todos
        if x["marca"]:
            de_marca = lambda ps: [p for p in ps if misma_marca(p["marca"], x["marca"])]
        else:
            de_marca = lambda ps: [p for p in ps if mismo_proveedor(p["proveedor"], x["proveedor"])] or list(ps)
        exactos = idx["exacto"].get(clave_e, [])
        compactos = idx["compacto"].get(clave_c, []) if clave_c else []
        elegidos, via = de_marca(exactos), "exacto"
        if not elegidos:
            elegidos, via = de_marca(compactos), "compacto"     # mismo número salvo guiones / símbolos
        if not elegidos and len(clave_c) >= 5:                    # el número está en las equivalencias de un producto de la marca
            elegidos, via = de_marca([p for p in idx["productos"] if p["equivalencias_c"] and clave_c in p["equivalencias_c"]])[:5], "equivalencia"
        candidatos = list({p["codigo"]: p for p in exactos + compactos}.values())   # todas las marcas, solo para avisar
        com = idx["codigo"].get(x["com"]) if x["com"] else None
        notas = []
        if com and com not in elegidos:
            notas.append(f"El COM {x['com']} es otro producto (nº parte {com['parte']}, {com['marca']})")
        if x["com"] and not com:
            notas.append(f"El COM {x['com']} no existe en el catálogo")
        if not elegidos:
            estado, prod = "nuevo", None
        elif len(elegidos) > 1:
            del_proveedor = [p for p in elegidos if mismo_proveedor(p["proveedor"], x["proveedor"])]
            estado, prod = "duplicado", (com if com in elegidos else (del_proveedor or elegidos)[0])
            if len(del_proveedor) == 1:
                notas.append(f"De los duplicados, {del_proveedor[0]['codigo']} es el que se compró a {x['proveedor']}")
        else:
            prod = elegidos[0]
            estado = "encontrado" if via == "exacto" else "posible"
        if prod and x["proveedor"] and prod["proveedor"] and not mismo_proveedor(prod["proveedor"], x["proveedor"]):
            notas.append(f"La última compra fue a otro proveedor: {prod['proveedor']}")
        if via == "compacto" and prod:
            notas.append(f"Nº parte en catálogo: {prod['parte']} (difiere en guiones o símbolos)")
        elif via == "equivalencia" and prod:
            notas.append(f"No está como nº parte; aparece en las equivalencias de {prod['codigo']} ({prod['parte']})")
        otras_marcas = [p for p in candidatos if not misma_marca(p["marca"], x["marca"])] if x["marca"] else []
        if estado == "nuevo" and otras_marcas:
            notas.append("No existe en " + x["marca"] + "; el número existe en otras marcas: " +
                         ", ".join(f"{p['codigo']} ({p['marca'] or 'sin marca'})" for p in otras_marcas[:5]))
        otros = [p for p in elegidos if prod is None or p["codigo"] != prod["codigo"]]   # duplicados de la misma marca
        if len(repetidas.get(clave_e, [])) > 1:
            notas.append(f"Repetido en esta factura (ítems {', '.join(repetidas[clave_e])})")
        fob = prod["fob"] if prod else None
        var = (x["precio"] / fob - 1) * 100 if prod and fob and x["precio"] else None
        x.update({
            "estado": estado, "via": via if elegidos else None, "notas": notas,
            "codigo": prod["codigo"] if prod else "", "nombre": prod["nombre"] if prod else "PRODUCTO NUEVO",
            "parte_catalogo": prod["parte"] if prod else "", "marca_catalogo": prod["marca"] if prod else "",
            "sac": (prod["sac"] or "") if prod else "", "fob_anterior": fob, "fecha_compra": prod["fecha_compra"] if prod else None,
            "proveedor_anterior": prod["proveedor"] if prod else "", "existencia": prod["existencia"] if prod else None,
            "proveedor_igual": bool(prod and mismo_proveedor(prod["proveedor"], x["proveedor"])),
            "variacion_fob": var, "repetido": len(repetidas.get(clave_e, [])) > 1,
            "sac_distinto": bool(prod and x["arancel"] and prod["sac"] and compacto(x["arancel"]) != compacto(prod["sac"])),
            "otros": [{"codigo": p["codigo"], "parte": p["parte"], "marca": p["marca"], "nombre": p["nombre"], "existencia": p["existencia"]}
                      for p in otros[:6]],
        })
        if var is not None and abs(var) > EXTREMO_FOB:
            x["notas"].append(f"FOB {var:+.0f}% contra la última compra: revisar unidad de medida o precio")
        if x["sac_distinto"]:
            x["notas"].append(f"Arancel de la plantilla ({x['arancel']}) distinto al del catálogo ({prod['sac']})")
    return lineas


def resumen(lineas):
    con_fob = [x for x in lineas if x["variacion_fob"] is not None and abs(x["variacion_fob"]) <= EXTREMO_FOB]
    extremos = sum(1 for x in lineas if x["variacion_fob"] is not None and abs(x["variacion_fob"]) > EXTREMO_FOB)
    total_nuevo = sum(x["total"] or (x["cantidad"] or 0) * (x["precio"] or 0) for x in lineas)
    total_anterior = sum((x["cantidad"] or 0) * x["fob_anterior"] for x in con_fob)
    total_nuevo_comp = sum((x["cantidad"] or 0) * (x["precio"] or 0) for x in con_fob)
    return {"lineas": len(lineas), "estados": {k: sum(1 for x in lineas if x["estado"] == k) for k in ESTADOS},
            "repetidos": sum(1 for x in lineas if x["repetido"]), "total": total_nuevo,
            "extremos": extremos,
            "subieron": sum(1 for x in con_fob if x["variacion_fob"] > ALERTA_FOB) + sum(1 for x in lineas if (x["variacion_fob"] or 0) > EXTREMO_FOB),
            "bajaron": sum(1 for x in con_fob if x["variacion_fob"] < -ALERTA_FOB) + sum(1 for x in lineas if (x["variacion_fob"] or 0) < -EXTREMO_FOB),
            "variacion_total": (total_nuevo_comp / total_anterior - 1) * 100 if total_anterior else None,
            "diferencia_total": total_nuevo_comp - total_anterior, "con_fob": len(con_fob),
            "proveedores": sorted({x["proveedor"] for x in lineas if x["proveedor"]}),
            "facturas": sorted({x["factura"] for x in lineas if x["factura"]}),
            "marcas": sorted({x["marca"] for x in lineas if x["marca"]})}


def ruta(token):
    return os.path.join(CARPETA, f"revision_ingreso_{token}.json")


# ---------------------------------------------------------------- vistas
@revision_ingreso_bp.route("/revision_ingreso", methods=["GET", "POST"])
@requiere_acceso("revision_ingreso")
def vista():
    if request.method == "POST":
        archivo = request.files.get("archivo")
        if not archivo or not archivo.filename.lower().endswith((".xlsx", ".xlsm")):
            flash("Suba la plantilla en Excel (.xlsx).", "warning")
            return redirect(url_for("revision_ingreso.vista"))
        try:
            lineas, lectura = leer_plantilla(archivo)
        except Exception as e:
            flash(f"No se pudo leer el archivo: {e}", "danger")
            return redirect(url_for("revision_ingreso.vista"))
        if not lineas:
            flash("La plantilla no tiene líneas con número de parte.", "warning")
            return redirect(url_for("revision_ingreso.vista"))
        lineas = revisar(lineas)
        token = uuid.uuid4().hex
        os.makedirs(CARPETA, exist_ok=True)
        with open(ruta(token), "w", encoding="utf-8") as fh:
            json.dump({"archivo": archivo.filename, "fecha": datetime.now().isoformat(timespec="minutes"), "lectura": lectura, "lineas": lineas},
                      fh, ensure_ascii=False)
        session["revision_ingreso"] = token
        return redirect(url_for("revision_ingreso.vista", t=token))

    token = request.args.get("t") or ""
    datos = None
    if token and re.fullmatch(r"[0-9a-f]{32}", token) and os.path.exists(ruta(token)):
        with open(ruta(token), encoding="utf-8") as fh:
            datos = json.load(fh)
    filtro = request.args.get("estado", "")
    lineas = datos["lineas"] if datos else []
    visibles = [x for x in lineas if not filtro or x["estado"] == filtro or (filtro == "fob" and x["variacion_fob"] is not None
                                                                            and abs(x["variacion_fob"]) > ALERTA_FOB)
                or (filtro == "repetido" and x["repetido"])]
    return render_template("revision_ingreso.html", datos=datos, token=token, lineas=visibles, r=resumen(lineas) if datos else None,
                           estados=ESTADOS, filtro=filtro, alerta=ALERTA_FOB)


@revision_ingreso_bp.route("/revision_ingreso/excel")
@requiere_acceso("revision_ingreso")
def excel():
    token = request.args.get("t") or ""
    if not (re.fullmatch(r"[0-9a-f]{32}", token) and os.path.exists(ruta(token))):
        flash("La revisión ya no está disponible; suba el archivo otra vez.", "warning")
        return redirect(url_for("revision_ingreso.vista"))
    with open(ruta(token), encoding="utf-8") as fh:
        datos = json.load(fh)
    wb = Workbook()
    ws = wb.active
    ws.title = "Revisión"
    # Primero las columnas tal como vienen en el archivo; a la derecha lo que agrega la revisión
    columnas = [("PROVEEDOR", 18), ("FACTURA", 10), ("ITEM", 6), ("N°PARTE 01", 18), ("MARCA", 12), ("ORIGEN", 10), ("U/M", 6),
                ("CANT", 8), ("P.UNIT", 10), ("TOTAL", 11),
                ("COM", 11), ("DESCRIPCION COMERSA", 42), ("ARANCEL (SAC)", 17), ("ESTADO", 18), ("FOB ANTERIOR", 11), ("VAR. FOB %", 9),
                ("FECHA ÚLT. COMPRA", 12), ("PROVEEDOR ANTERIOR", 22), ("EXISTENCIA", 10), ("Nº PARTE CATÁLOGO", 18), ("MARCA CATÁLOGO", 14),
                ("OTROS CÓDIGOS MISMO Nº PARTE", 34), ("OBSERVACIONES", 60)]
    for i, (nombre, ancho) in enumerate(columnas, start=1):
        c = ws.cell(row=1, column=i, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NARANJA if i >= 11 else VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[1].height = 34
    colores = {"nuevo": "FFF1E8", "duplicado": "FDE8E8", "posible": "FEF7E0"}
    for n, x in enumerate(datos["lineas"], start=2):
        valores = [x["proveedor"], x["factura"], x["item"], x["parte"], x["marca"], x["origen"], x["um"], x["cantidad"], x["precio"],
                   x["total"] if x["total"] is not None else (x["cantidad"] or 0) * (x["precio"] or 0),
                   x["codigo"], x["nombre"], x["sac"] or x["arancel"], ESTADOS[x["estado"]] + (" · repetido" if x["repetido"] else ""),
                   x["fob_anterior"], x["variacion_fob"] / 100 if x["variacion_fob"] is not None else None,
                   date.fromisoformat(x["fecha_compra"]) if x["fecha_compra"] else None, x["proveedor_anterior"], x["existencia"],
                   x["parte_catalogo"], x["marca_catalogo"], ", ".join(f"{o['codigo']} ({o['marca'] or 's/m'})" for o in x["otros"]),
                   " · ".join(x["notas"])]
        for col, v in enumerate(valores, start=1):
            ws.cell(row=n, column=col, value=v)
        for col in (9, 10, 15):
            ws.cell(row=n, column=col).number_format = "#,##0.00"
        ws.cell(row=n, column=16).number_format = "0%"
        ws.cell(row=n, column=17).number_format = "dd/mm/yyyy"
        if x["estado"] in colores:
            for col in (11, 12, 14):
                ws.cell(row=n, column=col).fill = PatternFill("solid", fgColor=colores[x["estado"]])
        if x["estado"] == "nuevo":
            ws.cell(row=n, column=12).font = Font(bold=True, color="9A3412")
        if x["variacion_fob"] is not None and abs(x["variacion_fob"]) > ALERTA_FOB:
            ws.cell(row=n, column=16).font = Font(bold=True, color="B42318" if x["variacion_fob"] > 0 else VERDE)
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{max(len(datos['lineas']) + 1, 2)}"
    salida = io.BytesIO()
    wb.save(salida)
    salida.seek(0)
    base = os.path.splitext(datos["archivo"])[0]
    return send_file(salida, as_attachment=True, download_name=f"{base}_revisado.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
