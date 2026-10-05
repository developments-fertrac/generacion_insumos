"""Utilidades compartidas de Excel COM (Windows).

- ``HAS_COM``: si win32com esta disponible en esta maquina.
- ``excel_serial_from_date``: fecha -> numero de serie de Excel.

Las conversiones por COM/xlrd y el manejo de copias temporales que vivian aqui
se retiraron el 2026-10-05: ninguna tarea las usaba (estan en el historial de git).
"""

from __future__ import annotations

try:
    import win32com.client as win32

    HAS_COM = True
except Exception:
    HAS_COM = False


def excel_serial_from_date(dt) -> int | None:
    if dt is None:
        return None
    from datetime import datetime as dt_type
    if hasattr(dt, "year") and hasattr(dt, "month") and hasattr(dt, "day"):
        return (dt_type(dt.year, dt.month, dt.day) - dt_type(1899, 12, 30)).days
    return None

