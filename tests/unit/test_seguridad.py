"""Ninguna contrasena de .env debe quedar escrita en un log."""

from __future__ import annotations

import logging

import pytest

from core.seguridad import FiltroSecretos, huella, ocultar

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _secretos(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("EXCEL_PASSWORD", "Principal2026")
    monkeypatch.setenv("EXCEL_PASSWORDS_TRY", "Compras2027, Compras2028")
    monkeypatch.setenv("SMTP_PASSWORD", "abcd efgh ijkl")


def test_huella_es_estable_y_no_revela() -> None:
    h = huella("Compras2028")
    assert h.startswith("sha256:") and len(h) == 17
    assert "Compras2028" not in h
    assert h == huella("Compras2028")
    assert huella(None) == "(sin contrasena)"


def test_ocultar_reemplaza_todas_las_contrasenas() -> None:
    texto = "abri con Compras2028, antes probe Principal2026 y smtp abcd efgh ijkl"
    limpio = ocultar(texto)
    for secreto in ("Compras2028", "Principal2026", "abcd efgh ijkl"):
        assert secreto not in limpio
    assert huella("Compras2028") in limpio


def test_filtro_limpia_mensaje_con_argumentos_y_traceback() -> None:
    registros: list[str] = []

    class _Captura(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            registros.append(self.format(record))

    log = logging.getLogger("prueba_seguridad")
    handler = _Captura()
    handler.addFilter(FiltroSecretos())
    log.addHandler(handler)
    try:
        log.warning("'%s' desencriptado con contrasena '%s'", "VENTAS.xlsx", "Compras2028")
        try:
            raise ValueError("clave invalida: Principal2026")
        except ValueError:
            log.exception("fallo")
    finally:
        log.removeHandler(handler)

    salida = "\n".join(registros)
    assert "Compras2028" not in salida and "Principal2026" not in salida
    assert huella("Compras2028") in salida


def test_clave_de_actualizacion_tambien_se_oculta(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("VENTAS_ACTUALIZACION_PASSWORD", "ClaveSoloArea1")
    assert "ClaveSoloArea1" not in ocultar("abri con ClaveSoloArea1")
