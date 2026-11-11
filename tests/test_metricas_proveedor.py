"""Pruebas de router/metricas_proveedor.py (funciones puras, sin base de datos).
Con pytest:  python -m pytest tests
Sin pytest:  python tests/test_metricas_proveedor.py"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from router import metricas_proveedor as M  # noqa: E402


def cerca(a, b, tol=1e-6):
    return a is not None and abs(a - b) <= tol


# ---------- A1: venta mensual de productos nuevos ----------
def test_venta_mensual_producto_nuevo_divide_entre_sus_meses():
    # Ingresó hace 2 meses y vendió 100: 50 por mes, no 100 / 12
    assert cerca(M.venta_mensual(100, 2), 50)


def test_venta_mensual_producto_viejo_divide_entre_12():
    assert cerca(M.venta_mensual(120, 40), 10)
    assert cerca(M.venta_mensual(120, None), 10)       # sin fecha de primer ingreso: 12 meses


def test_venta_mensual_minimo_un_mes_y_sin_venta():
    assert cerca(M.venta_mensual(30, 0.4), 30)         # menos de un mes: se divide entre 1
    assert M.venta_mensual(0, 5) == 0


def test_meses_entre():
    assert cerca(M.meses_entre(date(2026, 8, 17), date(2026, 11, 11)), 86 / M.DIAS_MES)
    assert M.meses_entre(None, date(2026, 11, 11)) is None


# ---------- Clasificación ----------
def test_producto_nuevo_no_es_lento_ni_sin_venta():
    # Como COM42279: ingresó el 17/08/26, mucho inventario y poca venta todavía
    meses = M.meses_entre(date(2026, 8, 17), date(2026, 11, 11))
    estado, comentario = M.clasificar(meses, 10, date(2026, 11, 1), disponible=2132, mensual=50)
    assert estado == "nuevo"
    assert "sin evaluar" in comentario
    # nuevo sin ninguna venta: tampoco es «sin venta»
    assert M.clasificar(1.5, None, None, 100, 0)[0] == "nuevo"


def test_clasificacion_basica():
    assert M.clasificar(40, None, None, 10, 0)[0] == "sin_venta"                               # nunca vendido
    assert M.clasificar(40, 400, date(2025, 10, 7), 10, 0)[0] == "sin_venta"                   # más de 12 meses
    assert M.clasificar(40, 20, date(2026, 10, 22), 0, 5)[0] == "agotado"
    assert M.clasificar(40, 20, date(2026, 10, 22), 130, 10)[0] == "lento"                     # 13 meses de inventario
    assert M.clasificar(40, 20, date(2026, 10, 22), 60, 10)[0] == "rota"                       # 6 meses
    assert M.clasificar(40, 20, date(2026, 10, 22), 120, 10)[0] == "rota"                      # justo 12 meses


def test_recomendaciones():
    assert M.recomendacion("sin_venta", 10, None, 0) == M.NO_RECOMPRAR
    assert M.recomendacion("lento", 300, 30, 120) == M.NO_RECOMPRAR_12
    assert M.recomendacion("lento", 150, 15, 120) is None                  # lento pero menos de 24 meses
    assert M.recomendacion("nuevo", 500, 40, 50) == M.REVISAR_NUEVO
    assert M.recomendacion("agotado", 0, 0, 80) == M.REPONER
    assert M.recomendacion("rota", 10, 1.5, 80) == M.REPONER
    assert M.recomendacion("rota", 40, 6, 80) is None


# ---------- A2: rotación y GMROI de 12 meses ----------
def test_rotacion_y_gmroi():
    assert cerca(M.rotacion(450, 1000), 0.45)
    assert M.rotacion(450, 0) is None
    assert cerca(M.gmroi(300, 1000), 0.3)


def test_rentabilidad_usa_12_meses_para_el_veredicto():
    # Venta histórica de 6 años (2190 días) más floja que la de los últimos 12 meses
    r = M.rentabilidad(venta=600, costo=360, inventario=1000, dias=2190, venta_12m=1200, costo_12m=700)
    assert cerca(r["rotacion_hist"], 360 * 365 / 2190 / 1000)       # 0.06
    assert cerca(r["rotacion_12"], 0.7)
    assert cerca(r["gmroi_12"], 0.5)
    assert r["rotacion"] == r["rotacion_12"]
    assert cerca(r["meses_inventario"], 12 / 0.7)
    assert r["nivel"] == "malo" and "0.50" in r["veredicto"]
    r = M.rentabilidad(600, 360, 1000, 2190, 3000, 1300)            # GMROI 12m = 1.7
    assert r["nivel"] == "bueno"
    r = M.rentabilidad(600, 360, 1000, 2190, 0, 0)                  # sin venta en 12 meses
    assert r["nivel"] == "malo" and "12 meses" in r["veredicto"]


# ---------- Tendencia y lento desde ----------
def test_cambio_y_nivel_tendencia():
    assert cerca(M.cambio_pct(117, 100), 17)
    assert M.cambio_pct(50, 0) is None
    assert M.nivel_tendencia(17) == "sube"
    assert M.nivel_tendencia(-12) == "baja"
    assert M.nivel_tendencia(5) == "igual"
    assert M.nivel_tendencia(None, 10) == "nueva"


def test_lento_desde_detecta_caida_sostenida():
    serie = [100] * 24 + [40] * 12          # cae al 40% desde el mes 24 y no se recupera
    i = M.lento_desde(serie, primero=12)
    assert i is not None and 24 <= i <= 26   # el promedio móvil de 3 meses tarda en bajar del 70%


def test_lento_desde_no_cuenta_caida_que_se_recupera_o_muy_reciente():
    assert M.lento_desde([100] * 30 + [20] * 3 + [100] * 3, primero=12) is None   # se recuperó
    assert M.lento_desde([100] * 34 + [0] * 2, primero=12) is None                # solo 2 meses
    assert M.lento_desde([0] * 36, primero=12) is None                            # nunca vendió


def test_dejar_de_comprar():
    assert M.dejar_de_comprar(-35, 90)
    assert not M.dejar_de_comprar(-35, 60)          # cae pero el inventario se mueve
    assert not M.dejar_de_comprar(-15, 95)
    assert M.dejar_de_comprar(None, 100, venta_12m=0)


# ---------- Otros ----------
def test_temporada_y_mediana():
    texto, pct, inicio = M.temporada({1: 1, 2: 1, 3: 1, 4: 10, 5: 10, 6: 10, 7: 1, 8: 1, 9: 1, 10: 1, 11: 1, 12: 1})
    assert texto == "Abr–Jun" and inicio == 4 and cerca(pct, 30 / 39 * 100)
    assert M.meses_temporada(11) == [11, 12, 1]
    assert M.mediana_fecha([date(2024, 12, 1), None, date(2024, 8, 1), date(2025, 1, 1)]) == date(2024, 12, 1)
    assert M.mediana_fecha([]) is None


def test_lista_meses_sin_mes_en_curso():
    meses = M.lista_meses(date(2026, 11, 11), 12)
    assert meses[0] == "2025-11" and meses[-1] == "2026-10" and len(meses) == 12


def test_abreviar():
    assert M.abreviar(91425832) == "C$ 91.4 M"
    assert M.abreviar(850000) == "C$ 850 K"
    assert M.abreviar(320) == "C$ 320"


if __name__ == "__main__":      # sin pytest instalado
    pruebas = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    fallas = 0
    for nombre, prueba in pruebas:
        try:
            prueba()
            print(f"ok     {nombre}")
        except AssertionError as e:
            fallas += 1
            print(f"FALLA  {nombre} {e}")
    print(f"\n{len(pruebas) - fallas} de {len(pruebas)} pruebas pasaron")
    sys.exit(1 if fallas else 0)
