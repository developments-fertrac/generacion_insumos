"""Reloj del proceso en hora de Colombia, sin importar la zona del servidor.

El servidor no esta en hora de Colombia (datetime.now() da +6 h): con el reloj
local, una corrida despues de las 6 p. m. quedaria con la fecha del dia
siguiente en el titulo "EXISTENCIA", en el nombre del archivo y en el aviso de
"archivo no es de hoy". Todo lo que dependa de la fecha usa este modulo.
"""

from __future__ import annotations

from datetime import date, datetime
from zoneinfo import ZoneInfo

ZONA_COLOMBIA = ZoneInfo("America/Bogota")


def ahora() -> datetime:
    """Fecha y hora de Colombia (sin zona, para nombres de archivo y Excel)."""
    return datetime.now(ZONA_COLOMBIA).replace(tzinfo=None)


def hoy() -> date:
    return ahora().date()


def desde_timestamp(segundos: float) -> datetime:
    """Fecha de modificacion de un archivo (st_mtime) en hora de Colombia."""
    return datetime.fromtimestamp(segundos, ZONA_COLOMBIA).replace(tzinfo=None)
