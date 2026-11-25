"""Lee el PDF de una "Orden de Venta/Pedido" del sistema de ventas.

Probado solo con "Pedido de Cliente.pdf" (pedido 270091, 1 página, 7 líneas).
Las columnas se ubican por coordenada X; si el formato del PDF cambia, ajustar COLS.
Uso: python leer_orden.py "Pedido de Cliente.pdf"
"""
import re
import sys

import pymupdf  # pip install pymupdf

# (campo, x desde, x hasta) en puntos PDF
COLS = [('bodega', 60, 120), ('codigo', 120, 190), ('producto', 190, 265), ('medida', 265, 330),
        ('marca', 330, 400), ('cantidad', 450, 495), ('subtotal', 575, 645)]


def _num(s):
    return float(s.replace(',', '')) if s else 0.0


def leer_orden(ruta_pdf):
    doc = pymupdf.open(ruta_pdf)
    # todas las páginas en una sola lista; la Y se desplaza por página para mantener el orden
    W = []
    for i, p in enumerate(doc):
        for w in p.get_text('words'):
            t = w[4]
            if t.isprintable() and not ('' <= t[0] <= ''):  # descarta íconos
                W.append((w[0], w[1] + i * p.rect.height, t))

    def misma_fila(y, x0, x1):
        return ' '.join(t for x, yy, t in sorted(W) if abs(yy - y) < 4 and x0 <= x < x1)

    def y_de(txt, x0=0, x1=9999):
        return next(y for x, y, t in W if t == txt and x0 <= x < x1)

    orden = dict(
        pedido=misma_fila(y_de('Vencimiento'), 150, 290),
        cliente=misma_fila(y_de('Cliente'), 150, 570),
        terminos=misma_fila(y_de('Pago'), 400, 570),
        fecha=misma_fila(y_de('Fecha', 280, 320), 400, 570),
        subtotal=_num(misma_fila(y_de('Sub-Total'), 650, 9999)),
        ingresado_por=misma_fila(y_de('Ingresado'), 140, 290),
    )

    # cada línea empieza con su número en x<60; termina donde empieza la siguiente o en TOTALES
    y_tot = y_de('TOTALES...')
    nums = sorted((y, t) for x, y, t in W if x < 60 and re.fullmatch(r'\d+', t) and y_de('BODEGA') - 5 < y < y_tot)
    lineas = []
    for i, (y, n) in enumerate(nums):
        y1 = nums[i + 1][0] - 3 if i + 1 < len(nums) else y_tot - 10  # los montos de TOTALES van ~6pt arriba
        fila = sorted((round(yy), x, t) for x, yy, t in W if y - 3 <= yy < y1)
        f = {'linea': int(n)}
        for campo, a, b in COLS:
            f[campo] = ' '.join(t for _, x, t in fila if a <= x < b)
        f['cantidad'] = _num(f['cantidad'])
        f['subtotal'] = _num(f['subtotal'])
        lineas.append(f)
    orden['lineas'] = lineas

    # control: la suma de las líneas debe dar el subtotal impreso y los números de línea ser 1..N
    orden['lectura_ok'] = (abs(sum(l['subtotal'] for l in lineas) - orden['subtotal']) < 0.05
                           and [l['linea'] for l in lineas] == list(range(1, len(lineas) + 1)))
    return orden


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    o = leer_orden(sys.argv[1])
    for k, v in o.items():
        if k != 'lineas':
            print(f'{k}: {v}')
    for l in o['lineas']:
        print(l)
