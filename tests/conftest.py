"""Fixtures compartidas de la suite de pruebas.

Reglas de las pruebas (ver plan, seccion 7):
- ``unit/``      DataFrames sinteticos, sin disco, sin Excel, < 1 s.
- ``contract/``  verifica que un adaptador cumple su puerto.
- ``golden/``    equivalencia vieja vs nueva sobre archivos reales.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_RAIZ = Path(__file__).resolve().parent.parent
for _ruta in (str(_RAIZ / "src"), str(_RAIZ)):
    if _ruta not in sys.path:
        sys.path.insert(0, _ruta)


@pytest.fixture(autouse=True)
def _entorno_limpio(monkeypatch: pytest.MonkeyPatch) -> None:
    """Cada prueba arranca sin estado heredado de la maquina."""
    for clave in ("LOGS_DIR", "STATE_DIR"):
        monkeypatch.delenv(clave, raising=False)


@pytest.fixture
def raiz_repo() -> Path:
    return _RAIZ


@pytest.fixture
def directorio_temporal(tmp_path: Path) -> Path:
    """Directorio aislado por prueba; nunca el repo."""
    destino = tmp_path / "datos"
    destino.mkdir(parents=True, exist_ok=True)
    return destino
