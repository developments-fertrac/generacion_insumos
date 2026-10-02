"""Partes puras del escritor COM: se prueban sin Excel."""

from __future__ import annotations

import pandas as pd
import pytest

from insumos.adapters.excel.com_inventario import (
    ajustar_formula_subtotal,
    construir_matriz,
    letra_columna,
    mapear_columnas,
)
from insumos.adapters.excel.escritor_simple import titulo_existencia
from insumos.domain.rules.inventario import columnas as C

pytestmark = pytest.mark.unit

ENCABEZADOS = [
    "REFERENCIA", "NOMBRE LISTA", "NOMBRE ODOO", "NOMBRE MYR ", "MARCA copia", "INV BODEGA GERENCIA ",
    "EXISTENCIA SEP 10", "COSTO PROMEDIO", "TOTAL INV", "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA",
    "CLASIFICACION", "Marca sistema ", "Dif marca", "Linea sistema", "Dif linea", "Sub- linea sistema",
    "Dif sub-linea", None, None,
]


def test_letra_columna() -> None:
    assert [letra_columna(i) for i in (1, 9, 26, 27, 52)] == ["A", "I", "Z", "AA", "AZ"]


def test_mapear_columnas_de_la_plantilla_real() -> None:
    m = mapear_columnas(ENCABEZADOS)
    assert m[C.EXISTENCIA] == 7 and m[C.TOTAL_INV] == 9
    assert m[C.MARCA_SISTEMA] == 14 and m[C.SUBLINEA_SISTEMA] == 18 and m[C.DIF_SUBLINEA] == 19


def test_mapear_columnas_falla_sin_requeridas() -> None:
    with pytest.raises(ValueError, match="EXISTENCIA"):
        mapear_columnas(["REFERENCIA", "COSTO PROMEDIO", "TOTAL INV"])


def test_matriz_escribe_formulas_dif_y_referencia_texto() -> None:
    m = mapear_columnas(ENCABEZADOS)
    datos = pd.DataFrame({
        C.REFERENCIA: [1060.0, "AB-1"], C.EXISTENCIA: [3, 0], C.COSTO_PROMEDIO: [2.5, 0.0],
        C.TOTAL_INV: [7.5, 0.0], C.MARCA_COPIA: ["X", None], C.MARCA_SISTEMA: ["x ", "Y"], "_NUEVA": [False, True],
    })
    filas = construir_matriz(datos, m, len(ENCABEZADOS), primera_fila=3)
    assert len(filas) == 2 and len(filas[0]) == len(ENCABEZADOS)
    assert filas[0][0] == "1060"
    assert filas[0][6] == 3 and filas[0][8] == "=G3*H3"
    assert filas[0][14] is True and filas[1][14] is False  # "x " = "X"; "Y" <> vacio
    assert filas[1][4] is None
    copia = construir_matriz(datos, m, len(ENCABEZADOS), primera_fila=3, formulas=False)
    assert copia[0][8] == 7.5
    con_formulas = construir_matriz(datos, m, len(ENCABEZADOS), primera_fila=3, formulas_dif=True)
    assert con_formulas[0][14] == "=+N3=E3" and con_formulas[1][16] == "=+P4=J4"


def test_ajustar_formula_subtotal() -> None:
    assert ajustar_formula_subtotal("=SUBTOTAL(109,G3:G15557)", 3, 15647) == "=SUBTOTAL(109,G3:G15647)"
    assert ajustar_formula_subtotal("=+IFERROR(A1,0)", 3, 10) == "=+IFERROR(A1,0)"


def test_titulo_existencia() -> None:
    from datetime import date

    assert titulo_existencia(date(2026, 10, 2)) == "EXISTENCIA OCT 02"
