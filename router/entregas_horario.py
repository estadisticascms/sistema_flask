"""Entregas: horario laboral para medir tiempos (reportes de la parte 12). Se configura en Catálogos › Configuración:

- horario_lunes_viernes, horario_sabado, horario_domingo: 'HH:MM-HH:MM'; vacío = ese día no se trabaja.
- feriados: lista separada por comas; 'dd/mm' se repite cada año y 'dd/mm/aaaa' es solo esa fecha (por ejemplo Semana Santa).

Mientras no se guarde, se usa el horario que dio la empresa: lunes a viernes de 8:00 a 17:00 y sábado de 8:00 a 13:00, sin feriados.
Horario.minutos(a, b) cuenta solo los minutos entre a y b que caen dentro del horario y fuera de los feriados.
"""
import re
from datetime import date, datetime, time, timedelta

from router.entregas_formato import hora12

CLAVES = ("horario_lunes_viernes", "horario_sabado", "horario_domingo", "feriados")
INICIAL = {"horario_lunes_viernes": "08:00-17:00", "horario_sabado": "08:00-13:00", "horario_domingo": "", "feriados": ""}


def leer_rango(texto):
    """'08:00-17:00' -> (time, time); '' -> None; inválido -> 'error'."""
    texto = (texto or "").replace(" ", "")
    if not texto:
        return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})", texto)
    if not m:
        return "error"
    h1, m1, h2, m2 = map(int, m.groups())
    if h1 > 23 or h2 > 23 or m1 > 59 or m2 > 59 or (h2, m2) <= (h1, m1):
        return "error"
    return time(h1, m1), time(h2, m2)


def leer_feriados(texto):
    """'25/12, 1/1, 17/4/2026' -> ({(25, 12), (1, 1)}, {date(2026, 4, 17)}, '25/12, 01/01, 17/04/2026'); inválido -> (None, None, 'error')."""
    anuales, fechas, limpio = set(), set(), []
    for parte in re.split(r"[,;\n]+", texto or ""):
        parte = parte.strip()
        if not parte:
            continue
        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})(?:/(\d{4}))?", parte)
        try:
            if not m:
                raise ValueError
            d, mes = int(m.group(1)), int(m.group(2))
            if m.group(3):
                fechas.add(date(int(m.group(3)), mes, d))
                limpio.append(f"{d:02d}/{mes:02d}/{m.group(3)}")
            else:
                date(2024, mes, d)          # valida el día (2024 es bisiesto: acepta 29/02)
                anuales.add((d, mes))
                limpio.append(f"{d:02d}/{mes:02d}")
        except ValueError:
            return None, None, "error"
    return anuales, fechas, ", ".join(limpio)


def valores(config):
    """Valores guardados en ent_configuracion, o los iniciales si todavía no se guardó el horario."""
    return {k: config[k] if k in config else INICIAL[k] for k in CLAVES}


class Horario:
    def __init__(self, config):
        v = valores(config)
        rangos = [leer_rango(v[k]) for k in CLAVES[:3]]
        rangos = [r if r != "error" else None for r in rangos]
        self.ventanas = {d: rangos[0] for d in range(5)}
        self.ventanas[5], self.ventanas[6] = rangos[1], rangos[2]
        anuales, fechas, _ = leer_feriados(v["feriados"])
        self.anuales, self.fechas = anuales or set(), fechas or set()

    def es_feriado(self, dia):
        return dia in self.fechas or (dia.day, dia.month) in self.anuales

    def minutos(self, a, b):
        """Minutos laborales entre a y b (0 si b <= a)."""
        if not a or not b or b <= a:
            return 0.0
        total, dia = 0.0, a.date()
        while dia <= b.date():
            ventana = self.ventanas.get(dia.weekday())
            if ventana and not self.es_feriado(dia):
                ini = max(a, datetime.combine(dia, ventana[0]))
                fin = min(b, datetime.combine(dia, ventana[1]))
                if fin > ini:
                    total += (fin - ini).total_seconds() / 60
            dia += timedelta(days=1)
        return total

    def texto(self):
        def r(v):
            return f"{hora12(v[0])} a {hora12(v[1])}"
        partes = [f"L–V {r(self.ventanas[0])}" if self.ventanas[0] else None,
                  f"sáb {r(self.ventanas[5])}" if self.ventanas[5] else None,
                  f"dom {r(self.ventanas[6])}" if self.ventanas[6] else None]
        n = len(self.anuales) + len(self.fechas)
        return ", ".join(p for p in partes if p) + (f"; sin {n} feriado(s)" if n else "; sin feriados cargados")
