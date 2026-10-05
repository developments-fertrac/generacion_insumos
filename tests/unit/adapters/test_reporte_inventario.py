"""Nombres de hoja del reporte de eliminaciones (hallazgo 10, ADR 0010)."""

from __future__ import annotations

import pytest

from insumos.adapters.excel.reporte_inventario import nombres_de_hoja

pytestmark = pytest.mark.unit


def test_nombres_respetan_el_limite_y_quitan_caracteres_invalidos() -> None:
    m = nombres_de_hoja(["PASO 1: REFERENCIA DUPLICADA", "BD: EXCLUIDA/POR [MOTIVO]?"])
    assert m["PASO 1: REFERENCIA DUPLICADA"] == "PASO 1 REFERENCIA DUPLICADA"
    assert all(len(n) <= 31 for n in m.values())
    assert not any(c in n for n in m.values() for c in "[]:*?/\\")


def test_pasos_con_el_mismo_prefijo_no_chocan() -> None:
    a = "PASO 2: NO LLEGA COMO INVENTARIO EN BASE DE DATOS"
    b = "PASO 2: NO LLEGA COMO INVENTARIO EN BASE DE DATOS (OTRO)"
    m = nombres_de_hoja([a, b])
    assert m[a] != m[b]
    assert len({n.upper() for n in m.values()}) == 2
    assert all(len(n) <= 31 for n in m.values())


def test_no_choca_con_las_hojas_fijas() -> None:
    m = nombres_de_hoja(["RESUMEN POR PASO"])
    assert m["RESUMEN POR PASO"].upper() != "RESUMEN POR PASO"
