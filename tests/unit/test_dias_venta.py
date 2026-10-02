"""Dias de venta del mes en hora Colombia (hoja 'Resumen Meta 2026', columna D)."""

from __future__ import annotations

from datetime import date, datetime, timezone
from unittest import mock

import pytest

import tasks.actualizacion_ventas as ventas
from tasks.actualizacion_ventas import dias_venta_transcurridos, festivos_colombia

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("hoy, esperado", [
    (date(2026, 10, 1), 1),      # jueves
    (date(2026, 10, 2), 2),      # viernes: caso reportado por control
    (date(2026, 10, 3), 2.5),    # sabado = medio dia
    (date(2026, 10, 4), 2.5),    # domingo = 0
    (date(2026, 10, 12), 8.0),   # lunes festivo (Dia de la Raza) no cuenta
    (date(2026, 10, 31), 23.5),  # mes completo = DIAS VENTA MES de octubre
])
def test_dias_venta_octubre_2026(hoy: date, esperado: float) -> None:
    assert dias_venta_transcurridos(hoy) == esperado


def test_festivos_2026_conocidos() -> None:
    f = festivos_colombia(2026)
    for d in (date(2026, 1, 1), date(2026, 1, 12), date(2026, 4, 2), date(2026, 4, 3),
              date(2026, 10, 12), date(2026, 12, 25)):
        assert d in f, d


def test_hoy_es_el_de_bogota_aunque_el_servidor_este_en_utc() -> None:
    """2 de octubre 8:30 p. m. en Bogota = 3 de octubre 01:30 UTC."""
    utc = datetime(2026, 10, 3, 1, 30, tzinfo=timezone.utc)

    class _Reloj(datetime):
        @classmethod
        def now(cls, tz=None):  # noqa: ANN001, ANN206
            return utc.astimezone(tz) if tz else utc.replace(tzinfo=None)

    with mock.patch.object(ventas, "datetime", _Reloj):
        hoy = ventas._hoy_bogota()
    assert hoy == date(2026, 10, 2)
    assert dias_venta_transcurridos(hoy) == 2
