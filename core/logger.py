from __future__ import annotations

import logging
import os
import sys
from datetime import date, datetime, timezone
from pathlib import Path

from zoneinfo import ZoneInfo

from core.seguridad import FiltroSecretos


def _logs_root() -> Path:
    """Raiz de logs.

    Fuera del repositorio por defecto (Phase 0: el repo no guarda estado
    operativo). Se puede fijar con la variable de entorno LOGS_DIR.
    """
    configurado = os.environ.get("LOGS_DIR", "").strip()
    if configurado:
        return Path(configurado)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "GeneracionInsumos" / "logs"


_LOGS_ROOT = _logs_root()

_TZ_COLOMBIA = ZoneInfo("America/Bogota")


def _now() -> datetime:
    """Fecha/hora actual en zona horaria de Colombia (independiente de la zona del host)."""
    return datetime.now(_TZ_COLOMBIA)


class _ColombiaFormatter(logging.Formatter):
    """Formatter que muestra la hora en la zona de Bogota."""
    def formatTime(self, record, datefmt=None):
        dt = datetime.fromtimestamp(record.created, tz=timezone.utc).astimezone(_TZ_COLOMBIA)
        return dt.strftime(datefmt or self.default_time_format)


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError, ValueError):
        pass

    day_dir = _LOGS_ROOT / _now().date().isoformat()
    day_dir.mkdir(parents=True, exist_ok=True)

    file_handler = logging.FileHandler(
        day_dir / f"{name}.log", encoding="utf-8", mode="a"
    )
    file_handler.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    fmt = _ColombiaFormatter(
        "[%(asctime)s] [%(levelname)-7s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(fmt)
    console_handler.setFormatter(fmt)

    # Ninguna contrasena de .env llega a consola ni a archivo (se ve su huella).
    filtro = FiltroSecretos()
    file_handler.addFilter(filtro)
    console_handler.addFilter(filtro)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
