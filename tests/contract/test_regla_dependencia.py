"""Pruebas de la regla de dependencia hexagonal.

No usan import-linter (mas lento): son una red rapida que corre en cada commit.
import-linter queda como el contrato autoritativo para CI.
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

_RAIZ = Path(__file__).resolve().parents[2]
_SRC = _RAIZ / "src" / "insumos"

MODULOS_PROHIBIDOS = {
    "win32com",
    "win32clipboard",
    "pythoncom",
    "selenium",
    "openpyxl",
    "msoffcrypto",
    "xlsxwriter",
    "insumos.adapters",
    "insumos.application",
    "tasks",
    "core",
    "config",
    "orchestrator",
}

CAPAS = ("domain", "application", "adapters")


def _modulos_python(capa: str) -> list[Path]:
    return sorted((_SRC / capa).rglob("*.py"))


def _imports_de(ruta: Path) -> set[str]:
    arbol = ast.parse(ruta.read_text(encoding="utf-8"), filename=str(ruta))
    encontrados: set[str] = set()
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Import):
            encontrados.update(alias.name for alias in nodo.names)
        elif isinstance(nodo, ast.ImportFrom) and nodo.module and nodo.level == 0:
            encontrados.add(nodo.module)
    return encontrados


def test_estructura_de_capas_existe() -> None:
    for capa in CAPAS:
        assert (_SRC / capa / "__init__.py").is_file(), f"falta src/insumos/{capa}/__init__.py"


@pytest.mark.parametrize("ruta", _modulos_python("domain"), ids=lambda p: p.name)
def test_dominio_no_importa_infraestructura(ruta: Path) -> None:
    intrusos = sorted(
        m for m in _imports_de(ruta)
        if m.split(".")[0] in MODULOS_PROHIBIDOS or m in MODULOS_PROHIBIDOS
    )
    assert not intrusos, f"{ruta.name} importa modulo prohibido: {intrusos}"


@pytest.mark.parametrize("ruta", _modulos_python("application"), ids=lambda p: p.name)
def test_aplicacion_no_importa_infraestructura(ruta: Path) -> None:
    intrusos = sorted(
        m for m in _imports_de(ruta)
        if m.split(".")[0] in MODULOS_PROHIBIDOS and m != "insumos.adapters"
    )
    assert not intrusos, f"{ruta.name} importa modulo prohibido: {intrusos}"
