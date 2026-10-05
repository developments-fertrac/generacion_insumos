"""Regresion de los hallazgos de la auditoria del 2026-10-05 (ADR 0010).

- Un solo correo por corrida, y el de exito solo cuando todo termino.
- El envio usa la hora de Colombia.
- El envio solo cierra el Chrome del perfil de WhatsApp y el Excel propio.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pytest

import tasks.actualizacion_ventas as ventas
import tasks.envio_informe_ventas as envio
from config.settings import Settings

pytestmark = pytest.mark.unit


class _Notificador:
    instancias: list[_Notificador] = []

    def __init__(self, *_: Any, **__: Any) -> None:
        self.exitos: list[dict[str, Any]] = []
        self.fallos: list[dict[str, Any]] = []
        _Notificador.instancias.append(self)

    def notify_success(self, detail: str, attachment: Path | None = None) -> bool:
        self.exitos.append({"detail": detail, "attachment": attachment})
        return True

    def notify_failure(self, error: str, attachment: Path | None = None) -> bool:
        self.fallos.append({"error": error, "attachment": attachment})
        return True


@pytest.fixture
def settings(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Settings:
    for k, v in {"SMTP_SENDER_EMAIL": "a@b.co", "SMTP_PASSWORD": "x" * 8, "EXCEL_PASSWORD": "y" * 8,
                 "BASE_PATH": str(tmp_path)}.items():
        monkeypatch.setenv(k, v)
    _Notificador.instancias.clear()
    monkeypatch.setattr(ventas, "EmailNotifier", _Notificador)
    monkeypatch.setattr(envio, "EmailNotifier", _Notificador)
    return Settings()


# --------------------------------------------------------------- hallazgo 1
def test_ventas_envia_un_solo_correo_de_exito(settings: Settings, monkeypatch: pytest.MonkeyPatch) -> None:
    def _ejecutar(self: ventas.ActualizacionVentas) -> None:
        self._detalle_exito = "Archivo: X — 10 registros"

    monkeypatch.setattr(ventas.ActualizacionVentas, "_ejecutar", _ejecutar)
    assert ventas.ActualizacionVentas(settings).run() is True
    (n,) = _Notificador.instancias
    assert len(n.exitos) == 1 and not n.fallos
    assert n.exitos[0]["detail"] == "Archivo: X — 10 registros"


def test_ventas_envia_un_solo_correo_de_fallo_con_traceback(settings: Settings, monkeypatch: pytest.MonkeyPatch) -> None:
    def _ejecutar(self: ventas.ActualizacionVentas) -> None:
        self._traceback_error = "Traceback (simulado)"
        raise RuntimeError("no se pudo reemplazar")

    monkeypatch.setattr(ventas.ActualizacionVentas, "_ejecutar", _ejecutar)
    assert ventas.ActualizacionVentas(settings).run() is False
    (n,) = _Notificador.instancias
    assert not n.exitos and len(n.fallos) == 1
    assert "no se pudo reemplazar" in n.fallos[0]["error"] and "Traceback (simulado)" in n.fallos[0]["error"]


# --------------------------------------------------------------- hallazgo 2
class _WhatsApp:
    def __init__(self, exito: bool) -> None:
        self.exito = exito

    def iniciar_sesion(self) -> bool:
        return True

    def enviar_reporte_completo(self, **_: Any) -> bool:
        return self.exito

    def cerrar(self) -> None:
        pass


def _preparar_envio(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, exito: bool) -> None:
    archivo = tmp_path / "Pruebas" / "$2026 VENTAS_Actualizacion.xlsx"
    archivo.parent.mkdir(parents=True)
    archivo.write_bytes(b"x")
    monkeypatch.setattr(envio, "capturar_multiples_rangos", lambda **_: (["img.png"], "23,4%"))
    monkeypatch.setattr(envio, "WhatsAppWeb", lambda **_: _WhatsApp(exito))
    monkeypatch.setattr(envio.time, "sleep", lambda *_: None)
    monkeypatch.setenv("WHATSAPP_CHATS", "Chat 1")


def test_envio_sin_destinatarios_no_manda_correo_de_exito(settings: Settings, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _preparar_envio(monkeypatch, tmp_path, exito=False)
    assert envio.EnvioInformeVentas(settings).run() is False
    (n,) = _Notificador.instancias
    assert not n.exitos, "llego un [OK] aunque nadie recibio el informe"
    assert len(n.fallos) == 1 and "Ningún destinatario" in n.fallos[0]["error"]


def test_envio_exitoso_manda_un_solo_correo(settings: Settings, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _preparar_envio(monkeypatch, tmp_path, exito=True)
    assert envio.EnvioInformeVentas(settings).run() is True
    (n,) = _Notificador.instancias
    assert len(n.exitos) == 1 and not n.fallos
    assert "Envíos exitosos: 1/1" in n.exitos[0]["detail"]


# --------------------------------------------------------------- hallazgo 3
def test_envio_usa_la_hora_de_colombia(settings: Settings, monkeypatch: pytest.MonkeyPatch) -> None:
    # 31-oct 20:30 en Bogota (servidor en UTC ya estaria en noviembre)
    monkeypatch.setattr(envio, "_ahora_colombia", lambda: datetime(2026, 10, 31, 20, 30))
    monkeypatch.setattr(envio, "_hoy_colombia", lambda: datetime(2026, 10, 31).date())
    assert envio.obtener_mes_actual_espanol() == "OCTUBRE"
    tarea = envio.EnvioInformeVentas(settings)
    tarea.setup()
    assert tarea.log_file.parent.name == "2026-10-31"


# --------------------------------------------------------------- hallazgo 5
class _Proc:
    def __init__(self, pid: int, nombre: str, cmd: list[str]) -> None:
        self.info = {"pid": pid, "name": nombre, "cmdline": cmd}


def test_solo_se_cierra_el_chrome_del_perfil(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    perfil = str(tmp_path / "chrome_profile_whatsapp_session")
    procesos = [
        _Proc(10, "chrome.exe", ["chrome.exe", f"--user-data-dir={perfil}"]),
        _Proc(11, "chrome.exe", ["chrome.exe", "--type=renderer", f"--user-data-dir={perfil}"]),
        _Proc(20, "chrome.exe", ["chrome.exe", "--profile-directory=Default"]),  # Chrome del usuario
        _Proc(30, "EXCEL.EXE", ["EXCEL.EXE"]),
    ]

    class _Psutil:
        @staticmethod
        def process_iter(_: Any) -> list[_Proc]:
            return [p for p in procesos if p.info["pid"] not in matados]

    matados: list[int] = []
    comandos: list[list[str]] = []

    def _run(args: list[str], **_: Any) -> Any:
        comandos.append(list(args))
        if "/PID" in args:
            matados.append(int(args[args.index("/PID") + 1]))

        class _R:
            stdout = ""
        return _R()

    monkeypatch.setattr(envio, "PSUTIL_DISPONIBLE", True)
    monkeypatch.setattr(envio, "psutil", _Psutil)
    monkeypatch.setattr(envio.subprocess, "run", _run)
    monkeypatch.setattr(envio.time, "sleep", lambda *_: None)

    envio.cerrar_chrome_del_perfil(perfil)

    assert sorted(matados) == [10, 11]
    planos = [" ".join(c).upper() for c in comandos]
    assert not any("/IM CHROME.EXE" in c for c in planos)
    assert not any("EXCEL.EXE" in c for c in planos)


def test_no_quedan_taskkill_globales_de_excel_ni_chrome() -> None:
    raiz = Path(__file__).resolve().parents[2]
    for ruta in [*(raiz / "tasks").glob("*.py"), *(raiz / "core").glob("*.py")]:
        texto = ruta.read_text(encoding="utf-8").upper().replace('"', "'")
        assert "'/IM', 'EXCEL.EXE'" not in texto, ruta.name
        assert "'/IM', 'CHROME.EXE'" not in texto, ruta.name


def test_las_tareas_no_notifican_por_su_cuenta() -> None:
    """Solo BaseTask.run notifica: ni _ejecutar ni execute llaman al notificador."""
    import inspect

    for metodo in (ventas.ActualizacionVentas._ejecutar, ventas.ActualizacionVentas.execute,
                   envio.EnvioInformeVentas.execute):
        fuente = inspect.getsource(metodo)
        assert "notifier.notify_" not in fuente, metodo.__qualname__
