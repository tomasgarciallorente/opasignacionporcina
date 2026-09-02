# -*- coding: utf-8 -*-
"""Hora local de Argentina. Streamlit Cloud corre en UTC, así que datetime.now() /
date.today() del servidor daban ~3 h adelantadas (Tomás, 09-2026: "corregí hora" —
el header mostraba 20:45 cuando eran las 17:46). Todo lo que sea "ahora" o "hoy" en
la app tiene que pasar por acá."""
import datetime

try:
    from zoneinfo import ZoneInfo
    _TZ = ZoneInfo("America/Argentina/Buenos_Aires")
except Exception:  # sin tzdata disponible: fallback a UTC-3 fijo (Argentina no usa DST)
    _TZ = datetime.timezone(datetime.timedelta(hours=-3))


def ahora() -> datetime.datetime:
    """datetime con tz de Argentina."""
    return datetime.datetime.now(_TZ)


def hoy() -> datetime.date:
    """Fecha de hoy en Argentina."""
    return ahora().date()
