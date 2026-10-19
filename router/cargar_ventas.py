import io
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser

import pandas as pd
from flask import Blueprint, request, render_template, session, flash, redirect, url_for
from utils.db import get_connection
from router.accesos import requiere_acceso

cargar_ventas_bp = Blueprint("cargar_ventas", __name__)

CARPETA_TEMP = "uploads"
TOLERANCIA_MONTO = Decimal("0.05")
DIAS_POSIBLE_DUPLICADO = 120

# Largo máximo de cada columna de texto en ventas_historicas
LARGOS = {
    "tipo factura": 50, "Credito/Contado": 50, "tipo_venta": 20, "Bodega": 100,
    "vendedor": 100, "facturador": 100, "factura": 50, "cliente": 150,
    "descripcion": 100, "estado_factura": 20, "identificador": 200,
}

COLUMNAS_SQL = [
    "tipo factura", "Credito/Contado", "tipo_venta", "Bodega", "vendedor", "facturador",
    "factura", "Fecha", "cliente", "descripcion", "SubTotal", "Descuento", "IVA", "Total",
    "Costo Total", "Margen Bruto", "neto", "utilidad", "estado_factura", "fecha_vencimiento",
    "identificador",
]


class ErrorArchivo(Exception):
    """Problema que impide leer un archivo completo."""


# -----------------------------
# Utilidades de conversión
# -----------------------------
def a_texto(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()  # el histórico conserva los espacios internos (ej. "CONECTOR  ELECTRICO")


def a_decimal(valor):
    texto = a_texto(valor).replace(",", "").replace("C$", "").strip()
    if texto == "":
        return Decimal("0.00")
    try:
        return Decimal(texto).quantize(Decimal("0.01"))
    except InvalidOperation:
        raise ValueError(f"monto inválido '{valor}'")


def a_fecha(valor):
    if isinstance(valor, datetime):  # incluye pd.Timestamp; Excel guarda fecha+hora
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = a_texto(valor)
    if not texto:
        return None
    if re.match(r"^\d{4}-\d{2}-\d{2}", texto):
        f = pd.to_datetime(texto[:10], format="%Y-%m-%d", errors="coerce")
    else:
        f = pd.to_datetime(texto, dayfirst=True, errors="coerce")
    return None if pd.isna(f) else f.date()


def a_factura(valor):
    return a_texto(valor).lstrip("#").strip()


# -----------------------------
# Lectura de archivos
# -----------------------------
class _TablaHTML(HTMLParser):
    """Extrae las filas de las tablas de un .xls que en realidad es HTML."""

    def __init__(self):
        super().__init__()
        self.filas, self._fila, self._celda, self._span = [], None, None, 1

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._fila = []
        elif tag in ("td", "th") and self._fila is not None:
            self._celda = []
            try:
                self._span = int(dict(attrs).get("colspan") or 1)
            except ValueError:
                self._span = 1

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._celda is not None:
            self._fila.append(" ".join("".join(self._celda).split()))
            self._fila.extend([""] * (self._span - 1))
            self._celda = None
        elif tag == "tr" and self._fila is not None:
            self.filas.append(self._fila)
            self._fila = None

    def handle_data(self, data):
        if self._celda is not None:
            self._celda.append(data)


def leer_filas(archivo, hoja=None):
    """Devuelve las filas del archivo como listas (acepta .xlsx, .xls y .xls-HTML)."""
    contenido = archivo.read()
    if not contenido:
        raise ErrorArchivo("el archivo está vacío")

    if contenido.lstrip(b"\xef\xbb\xbf \r\n\t")[:1] == b"<":
        try:
            texto = contenido.decode("utf-8-sig")
        except UnicodeDecodeError:
            texto = contenido.decode("cp1252", errors="replace")
        parser = _TablaHTML()
        parser.feed(texto)
        return parser.filas

    try:
        libro = pd.ExcelFile(io.BytesIO(contenido))
    except Exception as e:
        raise ErrorArchivo(f"no se pudo abrir como Excel ({e})")
    if hoja is not None and hoja not in libro.sheet_names:
        raise ErrorArchivo(f"no tiene la hoja '{hoja}' (hojas: {', '.join(libro.sheet_names)})")
    # keep_default_na=False: textos como "N/A", "NA" o "NULL" son datos reales (ej. UOM "N/A"), no vacíos
    df = pd.read_excel(libro, sheet_name=hoja or 0, header=None, dtype=object, keep_default_na=False, na_values=[])
    return df.where(pd.notna(df), None).values.tolist()


def buscar_encabezado(filas, requeridas):
    """Ubica la fila de encabezados y devuelve (indice, {columna: posicion})."""
    for i, fila in enumerate(filas[:40]):
        nombres = [a_texto(c) for c in fila]
        if all(r in nombres for r in requeridas):
            posiciones = {}
            for j, n in enumerate(nombres):
                posiciones.setdefault(n, j)  # primera aparición (C$ antes que US$)
            return i, posiciones
    raise ErrorArchivo("no se encontraron las columnas " + ", ".join(requeridas))


def leer_reporte_ventas(archivo):
    requeridas = ["Bodega", "Facturador", "Vendedor", "# Factura", "Fecha", "Nombre de Cliente",
                  "Descripcion", "Tipo Factura", "Credito/Contado", "SubTotal", "Descuento",
                  "IVA", "Total", "Costo Total", "Margen Bruto"]
    filas = leer_filas(archivo)
    inicio, col = buscar_encabezado(filas, requeridas)

    registros, total_general = [], None
    for num, fila in enumerate(filas[inicio + 1:], start=inicio + 2):
        celda = lambda nombre: fila[col[nombre]] if col[nombre] < len(fila) else None
        primera = a_texto(fila[0]) if fila else ""
        if primera.upper().startswith("TOTAL GENERAL"):
            try:
                total_general = a_decimal(celda("SubTotal"))
            except ValueError:
                pass
            continue
        if not a_factura(celda("# Factura")) or a_texto(celda("# Factura")) == "# Factura":
            continue  # filas vacías o encabezado repetido al final del reporte
        registros.append({"fila": num, **{n: celda(n) for n in requeridas}})
    if not registros:
        raise ErrorArchivo("no contiene facturas")
    return registros, total_general


def leer_facturas_credito(archivo):
    requeridas = ["Bodega", "# Factura", "Contado/Credito", "Fecha", "Vencimiento", "Total", "Estado"]
    filas = leer_filas(archivo)
    inicio, col = buscar_encabezado(filas, requeridas)
    datos = {}
    for fila in filas[inicio + 1:]:
        factura = a_factura(fila[col["# Factura"]])
        if not factura.isdigit():
            continue
        clave = (factura, a_texto(fila[col["Bodega"]]))
        datos[clave] = {
            "vencimiento": a_fecha(fila[col["Vencimiento"]]),
            "fecha": a_fecha(fila[col["Fecha"]]),
            "total": fila[col["Total"]],
            "estado": a_texto(fila[col["Estado"]]),
        }
    if not datos:
        raise ErrorArchivo("no contiene facturas de crédito")
    return datos


def leer_caja_contado(archivo):
    requeridas = ["Factura", "Bodega", "Fecha", "Total", "Estado"]
    filas = leer_filas(archivo, hoja="Facturas Contado")
    inicio, col = buscar_encabezado(filas, requeridas)
    datos = {}
    for fila in filas[inicio + 1:]:
        factura = a_factura(fila[col["Factura"]])
        if not factura.isdigit():
            continue
        clave = (factura, a_texto(fila[col["Bodega"]]))
        datos[clave] = {
            "fecha": a_fecha(fila[col["Fecha"]]),
            "total": fila[col["Total"]],
            "estado": a_texto(fila[col["Estado"]]),
        }
    return datos


# -----------------------------
# Validación y armado de registros
# -----------------------------
def diferencia_total(total_ventas, total_otro):
    try:
        return abs(total_ventas - a_decimal(total_otro)) > TOLERANCIA_MONTO
    except ValueError:
        return True


def procesar(ventas, credito, contado, total_general):
    nuevos, errores, advertencias, duplicados = [], [], [], []
    vistos = set()

    for v in ventas:
        ref = f"Factura {a_factura(v['# Factura'])} ({a_texto(v['Bodega'])}) fila {v['fila']}"
        try:
            factura = a_factura(v["# Factura"])
            bodega = a_texto(v["Bodega"])
            tipo_cc = a_texto(v["Credito/Contado"]).upper()
            fecha = a_fecha(v["Fecha"])
            montos = {n: a_decimal(v[n]) for n in
                      ["SubTotal", "Descuento", "IVA", "Total", "Costo Total", "Margen Bruto"]}
        except ValueError as e:
            errores.append(f"{ref}: {e}")
            continue

        if not bodega:
            errores.append(f"{ref}: no tiene bodega")
            continue
        if fecha is None:
            errores.append(f"{ref}: fecha inválida '{a_texto(v['Fecha'])}'")
            continue
        if tipo_cc not in ("CREDITO", "CONTADO"):
            errores.append(f"{ref}: Credito/Contado desconocido '{tipo_cc}'")
            continue

        identificador = f"{factura}_{tipo_cc}_{bodega}_{fecha.isoformat()}"
        if identificador in vistos:
            errores.append(f"{ref}: viene repetida dentro del Reporte de Ventas")
            continue
        vistos.add(identificador)

        clave = (factura, bodega)
        estado, vencimiento = None, None
        if tipo_cc == "CREDITO":
            info = credito.get(clave)
            if info is None:
                advertencias.append(f"{ref}: no aparece en Facturas de Crédito; se guarda sin fecha de vencimiento")
                estado = "Journal"
            else:
                estado, vencimiento = info["estado"] or "Journal", info["vencimiento"]
                if diferencia_total(montos["Total"], info["total"]):
                    advertencias.append(f"{ref}: el Total no coincide (Ventas {montos['Total']} vs Crédito {a_texto(info['total'])})")
                if info["fecha"] and info["fecha"] != fecha:
                    advertencias.append(f"{ref}: la fecha no coincide (Ventas {fecha} vs Crédito {info['fecha']})")
                if vencimiento is None:
                    advertencias.append(f"{ref}: fecha de vencimiento vacía en Facturas de Crédito")
        else:
            vencimiento = fecha
            estado = "Pagado"
            if contado is not None:
                info = contado.get(clave)
                if info is None:
                    advertencias.append(f"{ref}: no aparece en Facturas Contado del Informe de Caja; se guarda como Pagado")
                else:
                    estado = info["estado"] or "Pagado"
                    if diferencia_total(montos["Total"], info["total"]):
                        advertencias.append(f"{ref}: el Total no coincide (Ventas {montos['Total']} vs Caja {a_texto(info['total'])})")

        if estado.lower().startswith("anulad"):
            errores.append(f"{ref}: está ANULADA según el archivo de {'crédito' if tipo_cc == 'CREDITO' else 'caja'}; no se carga")
            continue

        neto = montos["SubTotal"] - montos["Descuento"]
        registro = {
            "tipo factura": a_texto(v["Tipo Factura"]),
            "Credito/Contado": tipo_cc,
            "tipo_venta": tipo_cc,
            "Bodega": bodega,
            "vendedor": a_texto(v["Vendedor"]),
            "facturador": a_texto(v["Facturador"]),
            "factura": factura,
            "Fecha": fecha.isoformat(),
            "cliente": a_texto(v["Nombre de Cliente"]),
            "descripcion": a_texto(v["Descripcion"]) or None,
            **{k: str(m) for k, m in montos.items()},
            "neto": str(neto),
            "utilidad": str(neto - montos["Costo Total"]),
            "estado_factura": estado,
            "fecha_vencimiento": vencimiento.isoformat() if vencimiento else None,
            "identificador": identificador,
        }

        largos = [c for c, m in LARGOS.items() if registro[c] and len(registro[c]) > m]
        if largos:
            errores.append(f"{ref}: texto demasiado largo en {', '.join(largos)}")
            continue
        nuevos.append(registro)

    # Integridad del reporte de ventas
    if total_general is not None:
        suma = Decimal("0")
        for v in ventas:
            try:
                suma += a_decimal(v["SubTotal"])
            except ValueError:
                pass  # ya reportada como error de la fila
        if abs(suma - total_general) > Decimal("1.00"):
            advertencias.insert(0, f"La suma del SubTotal de las filas ({suma:,.2f}) no cuadra con el TOTAL GENERAL del reporte ({total_general:,.2f}); revise que el archivo esté completo")

    # Facturas que están en crédito/caja pero no en ventas (anuladas = informativo)
    claves_ventas = {(a_factura(v["# Factura"]), a_texto(v["Bodega"])) for v in ventas}
    anuladas = []
    for origen, datos in (("Crédito", credito), ("Caja", contado or {})):
        for (factura, bodega), info in sorted(datos.items()):
            if (factura, bodega) in claves_ventas:
                continue
            if info["estado"].lower().startswith("anulad"):
                anuladas.append(f"{factura} ({bodega}) – {origen}")
            else:
                advertencias.append(f"Factura {factura} ({bodega}) aparece en {origen} pero no en el Reporte de Ventas; no se cargará")

    # Duplicados contra la base de datos
    if nuevos:
        nuevos, duplicados, posibles = separar_duplicados(nuevos)
        advertencias.extend(posibles)

    return nuevos, errores, advertencias, duplicados, anuladas


def separar_duplicados(registros):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        ids = [r["identificador"] for r in registros]
        marcas = ", ".join(["%s"] * len(ids))
        cursor.execute(f"SELECT identificador FROM ventas_historicas WHERE identificador IN ({marcas})", ids)
        existentes = {fila[0] for fila in cursor.fetchall()}

        fecha_min = min(date.fromisoformat(r["Fecha"]) for r in registros) - timedelta(days=DIAS_POSIBLE_DUPLICADO)
        facturas = sorted({r["factura"] for r in registros})
        marcas = ", ".join(["%s"] * len(facturas))
        cursor.execute(
            f"SELECT factura, Bodega, identificador FROM ventas_historicas "
            f"WHERE Fecha >= %s AND factura IN ({marcas})",
            [fecha_min] + facturas,
        )
        recientes = {(f, b): i for f, b, i in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    nuevos, duplicados, posibles = [], [], []
    for r in registros:
        if r["identificador"] in existentes:
            duplicados.append(r)
            continue
        previo = recientes.get((r["factura"], r["Bodega"]))
        if previo:
            r["_motivo"] = f"ya existe como {previo}"
            duplicados.append(r)
            posibles.append(f"Factura {r['factura']} ({r['Bodega']}): ya está en la base con otro identificador ({previo}); no se carga")
            continue
        nuevos.append(r)
    return nuevos, duplicados, posibles


# -----------------------------
# Archivo temporal entre "validar" y "confirmar"
# -----------------------------
def ruta_temporal(token, prefijo="carga_ventas_"):
    if not token or not re.fullmatch(r"[0-9a-f]{32}", token):
        return None
    return os.path.join(CARPETA_TEMP, f"{prefijo}{token}.json")


def limpiar_temporales():
    try:
        for nombre in os.listdir(CARPETA_TEMP):
            ruta = os.path.join(CARPETA_TEMP, nombre)
            if nombre.startswith("carga_") and nombre.endswith(".json") and pd.Timestamp.now().timestamp() - os.path.getmtime(ruta) > 86400:
                os.remove(ruta)
    except OSError:
        pass


def resumen_por_fecha(registros):
    resumen = {}
    for r in registros:
        fila = resumen.setdefault(r["Fecha"], {"fecha": r["Fecha"], "facturas": 0, "credito": 0, "contado": 0,
                                               "subtotal": Decimal("0"), "total": Decimal("0")})
        fila["facturas"] += 1
        fila["credito" if r["Credito/Contado"] == "CREDITO" else "contado"] += 1
        fila["subtotal"] += Decimal(r["SubTotal"])
        fila["total"] += Decimal(r["Total"])
    return [resumen[k] for k in sorted(resumen)]


# -----------------------------
# Vistas
# -----------------------------
@cargar_ventas_bp.route("/cargar_ventas", methods=["GET", "POST"])
@requiere_acceso("cargar_ventas")
def cargar_ventas():
    if request.method == "GET":
        return render_template("cargar_ventas.html")

    archivos = {
        "ventas": request.files.get("archivo_ventas"),
        "credito": request.files.get("archivo_credito"),
        "caja": request.files.get("archivo_caja"),
    }
    errores_archivo = []
    if not archivos["ventas"] or not archivos["ventas"].filename:
        errores_archivo.append("Falta el Reporte de Ventas.")
    if not archivos["credito"] or not archivos["credito"].filename:
        errores_archivo.append("Falta el Reporte de Facturas de Crédito.")
    if errores_archivo:
        return render_template("cargar_ventas.html", errores_archivo=errores_archivo)

    ventas = credito = contado = total_general = None
    for clave, nombre, lector in (
        ("ventas", "Reporte de Ventas", leer_reporte_ventas),
        ("credito", "Reporte de Facturas de Crédito", leer_facturas_credito),
        ("caja", "Informe de Caja", leer_caja_contado),
    ):
        archivo = archivos[clave]
        if not archivo or not archivo.filename:
            continue
        try:
            resultado = lector(archivo)
        except ErrorArchivo as e:
            errores_archivo.append(f"{nombre} ({archivo.filename}): {e}. ¿Seleccionó el archivo correcto?")
            continue
        except Exception as e:
            errores_archivo.append(f"{nombre} ({archivo.filename}): error inesperado al leer – {e}")
            continue
        if clave == "ventas":
            ventas, total_general = resultado
        elif clave == "credito":
            credito = resultado
        else:
            contado = resultado
    if errores_archivo:
        return render_template("cargar_ventas.html", errores_archivo=errores_archivo)

    try:
        nuevos, errores, advertencias, duplicados, anuladas = procesar(ventas, credito, contado, total_general)
    except Exception as e:
        return render_template("cargar_ventas.html", errores_archivo=[f"Error al validar contra la base de datos: {e}"])

    token = None
    if nuevos:
        limpiar_temporales()
        token = uuid.uuid4().hex
        with open(ruta_temporal(token), "w", encoding="utf-8") as f:
            json.dump(nuevos, f, ensure_ascii=False)
        session["carga_ventas_token"] = token

    return render_template(
        "cargar_ventas.html",
        validado=True,
        token=token,
        total_leidas=len(ventas),
        nuevos=nuevos,
        resumen=resumen_por_fecha(nuevos),
        errores=errores,
        advertencias=advertencias,
        duplicados=duplicados,
        anuladas=anuladas,
        sin_caja=contado is None,
    )


@cargar_ventas_bp.route("/cargar_ventas/confirmar", methods=["POST"])
@requiere_acceso("cargar_ventas")
def confirmar_carga_ventas():
    token = request.form.get("token")
    ruta = ruta_temporal(token)
    if not ruta or token != session.get("carga_ventas_token") or not os.path.exists(ruta):
        flash("⚠️ La validación expiró o ya fue cargada. Vuelva a subir los archivos.", "warning")
        return redirect(url_for("cargar_ventas.cargar_ventas"))

    with open(ruta, encoding="utf-8") as f:
        registros = json.load(f)

    columnas = ", ".join(f"`{c}`" for c in COLUMNAS_SQL)
    marcas = ", ".join(["%s"] * len(COLUMNAS_SQL))
    sql = f"INSERT INTO ventas_historicas ({columnas}) VALUES ({marcas})"

    conn = get_connection()
    cursor = conn.cursor()
    try:
        # Revalidar por si alguien cargó los mismos datos entre la validación y la confirmación
        ids = [r["identificador"] for r in registros]
        cursor.execute(
            f"SELECT identificador FROM ventas_historicas WHERE identificador IN ({', '.join(['%s'] * len(ids))})", ids)
        ya_existen = {fila[0] for fila in cursor.fetchall()}
        pendientes = [r for r in registros if r["identificador"] not in ya_existen]

        if pendientes:
            cursor.executemany(sql, [tuple(r[c] for c in COLUMNAS_SQL) for r in pendientes])
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f"❌ No se guardó nada. Error de base de datos: {e}", "danger")
        return redirect(url_for("cargar_ventas.cargar_ventas"))
    finally:
        cursor.close()
        conn.close()

    os.remove(ruta)
    session.pop("carga_ventas_token", None)

    mensaje = f"✅ Se guardaron {len(pendientes)} facturas en ventas_historicas."
    if ya_existen:
        mensaje += f" Se omitieron {len(ya_existen)} que ya existían."
    flash(mensaje, "success")
    return render_template("cargar_ventas.html", cargado=True, resumen=resumen_por_fecha(pendientes))
