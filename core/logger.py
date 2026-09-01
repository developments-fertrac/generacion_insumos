from __future__ import annotations

import logging
import sys
from datetime import date, datetime, timezone
from pathlib import Path

from zoneinfo import ZoneInfo


_LOGS_ROOT = Path(__file__).resolve().parent.parent / "logs"

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

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
