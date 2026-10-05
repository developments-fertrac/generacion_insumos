"""Titulo diario de la columna de existencia (``EXISTENCIA OCT 02``).

Lo usan los dos escritores de inventario (COM y openpyxl). El escritor plano
``EscritorInventarioSimple`` que vivia aqui se retiro el 2026-10-05: el modo
``--dry-run`` usa ``escritor_openpyxl.EscritorInventarioOpenpyxl``.
"""

from __future__ import annotations

from datetime import date

_MESES = ("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")


def titulo_existencia(hoy: date) -> str:
    """``EXISTENCIA OCT 02``: mismo formato que el legacy."""
    return f"EXISTENCIA {_MESES[hoy.month - 1]} {hoy.day:02d}"

