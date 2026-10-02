"""'$2026 VENTAS_Actualizacion.xlsx' usa su propia contrasena (area autorizada)."""

from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest

from tasks.actualizacion_ventas import ActualizacionVentas

pytestmark = pytest.mark.unit


def _tarea(clave_area: str) -> ActualizacionVentas:
    t = ActualizacionVentas.__new__(ActualizacionVentas)
    t.log = logging.getLogger("prueba")
    t.excel_cfg = SimpleNamespace(password="General1", passwords_try=(),
                                  password_ventas_actualizacion=clave_area)
    return t


def test_usa_la_clave_del_area() -> None:
    t = _tarea("ClaveArea")
    assert t.PASSWORD_ACTUALIZACION == "ClaveArea"
    assert t.PASSWORD_VENTAS == "General1"


def test_sin_clave_configurada_usa_la_general() -> None:
    assert _tarea("").PASSWORD_ACTUALIZACION == "General1"


def test_transicion_abre_con_la_anterior(monkeypatch: pytest.MonkeyPatch) -> None:
    t = _tarea("ClaveArea")
    intentos: list[str] = []

    def _descifrar(ruta, password=None):  # noqa: ANN001, ANN202
        intentos.append(password)
        if password != "General1":
            raise RuntimeError("clave incorrecta")
        return "contenido"

    monkeypatch.setattr(t, "_decrypt_to_stream", _descifrar)
    from pathlib import Path

    assert t._abrir_actualizacion(Path("$2026 VENTAS_Actualizacion.xlsx")) == "contenido"
    assert intentos == ["ClaveArea", "General1"]
