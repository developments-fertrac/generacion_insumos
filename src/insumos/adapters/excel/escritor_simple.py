"""Escritor de inventario sin Excel instalado (modo ``--dry-run``).

Escribe el resultado del pipeline en un ``.xlsx`` plano, con los encabezados de
la plantilla, para revisar los datos antes de tocar el archivo de produccion.
No conserva formatos, formulas ni tablas dinamicas: para eso esta
``com_inventario.EscritorInventarioCom``.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from insumos.domain.rules.inventario import columnas as C

_MESES = ("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")


def titulo_existencia(hoy: date) -> str:
    """``EXISTENCIA OCT 02``: mismo formato que el legacy."""
    return f"EXISTENCIA {_MESES[hoy.month - 1]} {hoy.day:02d}"


@dataclass
class EscritorInventarioSimple:
    carpeta: Path
    prefijo: str = "DRYRUN INVENTARIO GENERAL"

    def escribir(self, datos: pd.DataFrame, *, hoy: date) -> Path:
        self.carpeta.mkdir(parents=True, exist_ok=True)
        ruta = self.carpeta / f"{self.prefijo} {datetime.now():%Y%m%d_%H%M}.xlsx"
        cols = [c for c in C.COLUMNAS_PLANTILLA if c in datos.columns and c not in C.COLUMNAS_FORMULA]
        salida = datos[cols].rename(columns={C.EXISTENCIA: titulo_existencia(hoy)})
        salida[C.REFERENCIA] = salida[C.REFERENCIA].map(
            lambda v: str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip()
        )
        salida.to_excel(ruta, sheet_name="INVENTARIO", index=False)
        return ruta
