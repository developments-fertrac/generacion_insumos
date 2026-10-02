"""Busqueda de archivos de entrada por prefijo, ignorando temporales de Excel."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from pathlib import Path

EXTENSIONES_EXCEL = (".xlsx", ".xlsm", ".xls")


def _norm(texto: str) -> str:
    s = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"^\$+", "", s.strip())
    return re.sub(r"\s+", " ", s).strip().lower()


def candidatos(carpeta: Path, prefijo: str, extensiones: Iterable[str] = EXTENSIONES_EXCEL) -> list[Path]:
    """Archivos de ``carpeta`` cuyo nombre (sin ``$`` inicial) empieza por ``prefijo``."""
    if not carpeta.is_dir():
        return []
    pref = _norm(prefijo)
    exts = {e.lower() for e in extensiones}
    salida = []
    for f in carpeta.iterdir():
        if not f.is_file() or f.suffix.lower() not in exts:
            continue
        if f.name.startswith(("~$", "TEMP_")):
            continue
        if _norm(f.stem).startswith(pref):
            salida.append(f)
    return salida


def mas_reciente(carpeta: Path, prefijo: str) -> Path | None:
    """El candidato modificado mas recientemente, o None."""
    encontrados = candidatos(carpeta, prefijo)
    return max(encontrados, key=lambda p: p.stat().st_mtime) if encontrados else None


def requerido(carpeta: Path, prefijo: str, que_es: str) -> Path:
    archivo = mas_reciente(carpeta, prefijo)
    if archivo is None:
        raise FileNotFoundError(f"No se encontro {que_es}: ningun archivo '{prefijo}*' en {carpeta}")
    return archivo
