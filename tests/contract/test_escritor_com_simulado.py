"""El escritor COM contra una hoja de Excel simulada.

No reemplaza la corrida real en Windows, pero fija el contrato con el modelo de
objetos de Excel: que celdas se escriben, en que orden y con que forma
(``Range.Value`` como tupla de tuplas, formulas Dif, subtotales, texto en
REFERENCIA, limpieza de las filas sobrantes de ayer).
"""

from __future__ import annotations

from datetime import date
from typing import Any

import pandas as pd
import pytest

from insumos.adapters.excel.com_inventario import EscritorInventarioCom
from insumos.domain.rules.inventario import columnas as C

pytestmark = pytest.mark.contract


class _Celda:
    def __init__(self, hoja: _Hoja, fila: int, col: int) -> None:
        self.hoja, self.Row, self.Column = hoja, fila, col

    @property
    def Value(self) -> Any:  # noqa: N802 - API de Excel
        return self.hoja.valores.get((self.Row, self.Column))

    @Value.setter
    def Value(self, v: Any) -> None:  # noqa: N802
        self.hoja.valores[(self.Row, self.Column)] = v

    Formula = Value


class _Rango:
    def __init__(self, hoja: _Hoja, a: _Celda, b: _Celda) -> None:
        self.hoja = hoja
        self.f1, self.c1, self.f2, self.c2 = a.Row, a.Column, b.Row, b.Column

    def _celdas(self) -> list[tuple[int, int]]:
        return [(f, c) for f in range(self.f1, self.f2 + 1) for c in range(self.c1, self.c2 + 1)]

    @property
    def Value(self) -> Any:  # noqa: N802
        return tuple(
            tuple(self.hoja.valores.get((f, c)) for c in range(self.c1, self.c2 + 1))
            for f in range(self.f1, self.f2 + 1)
        )

    @Value.setter
    def Value(self, filas: Any) -> None:  # noqa: N802
        filas = list(filas)
        assert len(filas) == self.f2 - self.f1 + 1, "filas != alto del rango"
        for i, fila in enumerate(filas):
            assert len(fila) == self.c2 - self.c1 + 1, "columnas != ancho del rango"
            for j, v in enumerate(fila):
                self.hoja.valores[(self.f1 + i, self.c1 + j)] = v

    @property
    def NumberFormat(self) -> str:  # noqa: N802
        return "General"

    @NumberFormat.setter
    def NumberFormat(self, fmt: str) -> None:  # noqa: N802
        for k in self._celdas():
            self.hoja.formatos[k] = fmt

    def ClearContents(self) -> None:  # noqa: N802
        for k in self._celdas():
            self.hoja.valores.pop(k, None)

    def ClearFormats(self) -> None:  # noqa: N802
        for k in self._celdas():
            self.hoja.formatos.pop(k, None)

    def Copy(self) -> None:  # noqa: N802
        self.hoja.portapapeles = self

    def PasteSpecial(self, Paste: int) -> None:  # noqa: N802, N803
        self.hoja.pegados.append((Paste, self.f1, self.f2))

    def AutoFilter(self) -> None:  # noqa: N802
        self.hoja.autofiltro = (self.f1, self.f2, self.c2)


class _Usado:
    def __init__(self, filas: int, cols: int) -> None:
        self.Row, self.Column = 1, 1
        self.Rows = type("R", (), {"Count": filas})()
        self.Columns = type("C", (), {"Count": cols})()


class _Hoja:
    def __init__(self, encabezados: list[Any], datos: list[list[Any]]) -> None:
        self.valores: dict[tuple[int, int], Any] = {}
        self.formatos: dict[tuple[int, int], str] = {}
        self.pegados: list[tuple[int, int, int]] = []
        self.autofiltro: tuple[int, int, int] | None = None
        self.AutoFilterMode = True
        self.Application = type("App", (), {"CutCopyMode": True})()
        self.ancho = len(encabezados)
        self.valores[(1, 7)] = "=SUBTOTAL(109,G3:G999)"
        for j, h in enumerate(encabezados, 1):
            self.valores[(2, j)] = h
        for i, fila in enumerate(datos, 3):
            for j, v in enumerate(fila, 1):
                self.valores[(i, j)] = v
        self.alto = 2 + len(datos) + 1  # + fila de subtotales de ayer
        self.valores[(self.alto, 7)] = "=SUBTOTAL(109,G3:G999)"

    @property
    def UsedRange(self) -> _Usado:  # noqa: N802
        return _Usado(self.alto, self.ancho)

    def Cells(self, f: int, c: int) -> _Celda:  # noqa: N802
        return _Celda(self, f, c)

    def Range(self, a: _Celda, b: _Celda) -> _Rango:  # noqa: N802
        return _Rango(self, a, b)


ENC = [
    "REFERENCIA", "NOMBRE LISTA", "NOMBRE ODOO", "NOMBRE MYR ", "MARCA copia", "INV BODEGA GERENCIA ",
    "EXISTENCIA SEP 10", "COSTO PROMEDIO", "TOTAL INV", "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA",
    "CLASIFICACION", "Marca sistema ", "Dif marca", "Linea sistema", "Dif linea", "Sub- linea sistema",
    "Dif sub-linea",
]


def test_proyecta_sobre_la_hoja() -> None:
    ayer = [["OLD"] + [None] * 18 for _ in range(5)]  # 5 filas de ayer
    hoja = _Hoja(ENC, ayer)
    datos = pd.DataFrame({
        C.REFERENCIA: [1060.0, "AB-1"], C.EXISTENCIA: [3.0, 0.0], C.COSTO_PROMEDIO: [2.5, 0.0],
        C.TOTAL_INV: [7.5, 0.0], C.MARCA_COPIA: ["X", "Y"], C.MARCA_SISTEMA: ["X", "Z"],
    })
    escritor = EscritorInventarioCom(plantilla=None, carpeta_salida=None, password=None)  # type: ignore[arg-type]

    ultima = escritor._proyectar(hoja, datos, date(2026, 10, 2))

    v = hoja.valores
    assert ultima == 4
    assert v[(2, 7)] == "EXISTENCIA OCT 02"
    assert v[(3, 1)] == "1060" and v[(4, 1)] == "AB-1"
    assert v[(3, 9)] == "=G3*H3"
    assert v[(3, 15)] is True and v[(4, 15)] is False  # Dif marca: X=X, Z<>Y
    # fila de subtotales nueva y la de ayer borrada
    assert v[(5, 7)] == "=SUBTOTAL(109,G3:G4)" and v[(5, 9)] == "=SUBTOTAL(109,I3:I4)"
    assert (6, 1) not in v and (8, 7) not in v
    # SUBTOTAL de la fila 1 ajustado
    assert v[(1, 7)] == "=SUBTOTAL(109,G3:G4)"
    # REFERENCIA en texto
    assert hoja.formatos[(3, 1)] == "@" and hoja.formatos[(4, 1)] == "@"
    # el autofiltro no se toca (la salida de produccion no lo tiene)
    assert hoja.autofiltro is None
