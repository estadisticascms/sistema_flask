"""Lectura del PDF «Orden de Venta/Pedido» del sistema de ventas (sección 4.6 de la especificación).

Basado en docs/entregas/referencia/leer_orden.py (probado con el pedido 270091: 1 página, 7 líneas). Las columnas se ubican
por coordenada X; si el formato del PDF cambia, ajustar COLUMNAS. Cambios respecto a la referencia:
- Lee el PDF desde bytes (lo que sube el vendedor) y da errores claros en lugar de excepciones.
- Varias páginas: cada línea termina, a más tardar, al final de su página (no se traga el encabezado de la siguiente).
- Control obligatorio: la suma de los subtotales de las líneas debe igualar el subtotal impreso. Si no cuadra, la orden no se
  acepta. Los números de línea pueden tener saltos (el sistema de facturación no vuelve a numerar cuando se elimina un ítem,
  p. ej. 3, 4, 5…), pero deben ir de menor a mayor y sin repetirse; si no, es una lectura equivocada y se rechaza.

[PENDIENTE] Probado solo con una orden de una página: probar con 5 a 10 órdenes reales, incluida una de varias páginas.
"""
import re
from datetime import date

import pymupdf

# (campo, x desde, x hasta) en puntos PDF
COLUMNAS = [("bodega", 60, 120), ("codigo", 120, 190), ("producto", 190, 265), ("medida", 265, 330),
            ("marca", 330, 400), ("cantidad", 450, 495), ("subtotal", 575, 645)]
TOLERANCIA = 0.05     # diferencia permitida entre la suma de las líneas y el subtotal impreso (redondeo)


class ErrorLectura(Exception):
    """La orden no se pudo leer o no cuadra: se muestra al vendedor y la orden no se acepta."""


def _numero(texto, que):
    texto = (texto or "").replace(",", "").strip()
    if not texto:
        return 0.0
    if len(texto.split()) > 1:
        raise ErrorLectura(f"{que}: se leyeron varios valores ({texto}); parece que una línea no tiene su número")
    try:
        return float(texto)
    except ValueError:
        raise ErrorLectura(f"{que}: «{texto}» no es un número")


def _fecha(texto):
    try:
        return date.fromisoformat(texto.strip())
    except (ValueError, AttributeError):
        return None


def leer_orden(contenido):
    """bytes del PDF -> dict con pedido, cliente, terminos, fecha, ingresado_por, subtotal, lineas, suma_lineas, paginas.
    Lanza ErrorLectura si el archivo no es la orden esperada o si no cuadra."""
    if not contenido or not contenido.startswith(b"%PDF"):
        raise ErrorLectura("el archivo no es un PDF")
    try:
        doc = pymupdf.open(stream=contenido, filetype="pdf")
    except Exception:
        raise ErrorLectura("el PDF está dañado o no se puede abrir")

    # Todas las páginas en una sola lista; la Y se desplaza por página para mantener el orden
    palabras, fin_pagina = [], []
    for i, pagina in enumerate(doc):
        desplazamiento = i * pagina.rect.height
        for w in pagina.get_text("words"):
            t = w[4]
            if t.isprintable() and not ("" <= t[0] <= ""):    # descarta íconos (uso privado de Unicode)
                palabras.append((w[0], w[1] + desplazamiento, t))
        fin_pagina.append(desplazamiento + pagina.rect.height)
    paginas = len(fin_pagina)
    doc.close()
    if not palabras:
        raise ErrorLectura("el PDF no tiene texto (¿es una imagen escaneada?)")

    def misma_fila(y, x0, x1):
        return " ".join(t for x, yy, t in sorted(palabras) if abs(yy - y) < 4 and x0 <= x < x1)

    def y_de(texto, x0=0, x1=9999):
        ys = [y for x, y, t in palabras if t == texto and x0 <= x < x1]
        if not ys:
            raise ErrorLectura(f"no se reconoce el formato de la orden (falta «{texto}»)")
        return min(ys)

    def pagina_de(y):
        return next(n for n, fin in enumerate(fin_pagina) if y < fin)

    orden = {
        "pedido": misma_fila(y_de("Vencimiento"), 150, 290),
        "cliente": misma_fila(y_de("Cliente"), 150, 570),
        "terminos": misma_fila(y_de("Pago"), 400, 570),
        "fecha_texto": misma_fila(y_de("Fecha", 280, 320), 400, 570),
        "subtotal": _numero(misma_fila(y_de("Sub-Total"), 650, 9999), "Sub-Total de la orden"),
        "ingresado_por": misma_fila(y_de("Ingresado"), 140, 290),
        "paginas": paginas,
    }
    orden["fecha"] = _fecha(orden["fecha_texto"])
    titulo = misma_fila(y_de("Venta/Pedido"), 500, 700).replace("#", "").strip()
    if not re.fullmatch(r"\d{1,20}", orden["pedido"] or ""):
        raise ErrorLectura(f"no se encontró el número de pedido (se leyó «{orden['pedido']}»)")
    if titulo and titulo != orden["pedido"]:
        raise ErrorLectura(f"el número de pedido no coincide: título #{titulo}, campo {orden['pedido']}")
    if not orden["cliente"]:
        raise ErrorLectura("no se encontró el cliente")

    # Cada línea empieza con su número en x < 60; termina donde empieza la siguiente, en TOTALES o al final de su página
    y_encabezado = y_de("Linea")
    y_totales = y_de("TOTALES...")
    inicios = sorted((y, t) for x, y, t in palabras if x < 60 and re.fullmatch(r"\d+", t) and y_encabezado + 5 < y < y_totales)
    lineas = []
    for i, (y, n) in enumerate(inicios):
        hasta = inicios[i + 1][0] - 3 if i + 1 < len(inicios) else y_totales - 10   # los montos de TOTALES van ~6 pt arriba
        hasta = min(hasta, fin_pagina[pagina_de(y)])
        fila = sorted((round(yy), x, t) for x, yy, t in palabras if y - 3 <= yy < hasta)
        linea = {"linea": int(n)}
        for campo, a, b in COLUMNAS:
            linea[campo] = " ".join(t for _, x, t in fila if a <= x < b)
        linea["cantidad"] = _numero(linea["cantidad"], f"Línea {n}, cantidad")
        linea["subtotal"] = _numero(linea["subtotal"], f"Línea {n}, subtotal")
        lineas.append(linea)
    if not lineas:
        raise ErrorLectura("la orden no tiene líneas de productos")
    orden["lineas"] = lineas
    orden["suma_lineas"] = round(sum(l["subtotal"] for l in lineas), 2)

    # Control obligatorio
    numeros = [l["linea"] for l in lineas]
    if any(b <= a for a, b in zip(numeros, numeros[1:])):   # con saltos sí; repetidos o desordenados no
        raise ErrorLectura(f"los números de línea están repetidos o desordenados (se leyeron {', '.join(map(str, numeros))})")
    if abs(orden["suma_lineas"] - orden["subtotal"]) >= TOLERANCIA:
        raise ErrorLectura(f"no cuadra: la suma de las líneas da C$ {orden['suma_lineas']:,.2f} y el subtotal impreso es "
                           f"C$ {orden['subtotal']:,.2f}")
    for l in lineas:
        if not l["bodega"] or not l["codigo"]:
            raise ErrorLectura(f"la línea {l['linea']} no tiene bodega o código")
        if l["cantidad"] <= 0:
            raise ErrorLectura(f"la línea {l['linea']} tiene cantidad {l['cantidad']:g}")
    return orden
