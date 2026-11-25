"""Entregas: formato de horas para mostrar (12 horas con a. m. / p. m.). Los campos donde se escribe la hora siguen en 24 horas,
que es lo que esperan los <input type="time"> y "datetime-local"."""
from datetime import datetime, time, timedelta

NBSP = " "


def a_time(valor):
    """datetime, time, timedelta (TIME de MySQL) o '15:00' -> time, o None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.time()
    if isinstance(valor, time):
        return valor
    if isinstance(valor, timedelta):
        segundos = int(valor.total_seconds()) % 86400
        return time(segundos // 3600, segundos % 3600 // 60)
    try:
        h, m = str(valor).split(":")[:2]
        return time(int(h), int(m))
    except (ValueError, TypeError):
        return None


def hora12(valor, fecha=None):
    """'3:05 p. m.'; con fecha (formato strftime, p. ej. '%d/%m') y un datetime: '25/11 3:05 p. m.'. Vacío si no hay hora."""
    t = a_time(valor)
    if t is None:
        return ""
    texto = f"{t.hour % 12 or 12}:{t.minute:02d}{NBSP}{'a.' if t.hour < 12 else 'p.'}{NBSP}m."
    if fecha and isinstance(valor, datetime):
        return f"{valor.strftime(fecha)} {texto}"
    return texto
