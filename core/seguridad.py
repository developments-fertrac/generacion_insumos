"""Proteccion de secretos en logs.

- ``huella(secreto)``: identifica una contrasena sin revelarla
  (``sha256:3f1c9a0b2e``). Sirve para saber CUAL de las contrasenas de
  ``EXCEL_PASSWORDS_TRY`` abrio un archivo sin escribirla en el log.
- ``FiltroSecretos``: red de seguridad en los handlers del logger. Si algun
  mensaje (o traceback) llega a contener una contrasena de ``.env``, se
  reemplaza por su huella antes de escribirse en consola o archivo.
"""

from __future__ import annotations

import hashlib
import logging
import os

# Variables de .env cuyo valor nunca debe aparecer en un log.
VARIABLES_SECRETAS = (
    "EXCEL_PASSWORD", "EXCEL_PASSWORDS_TRY", "VENTAS_ACTUALIZACION_PASSWORD", "SMTP_PASSWORD", "FERTRAC_PASS",
)
_LARGO_MINIMO = 4  # evita reemplazar fragmentos triviales


def huella(secreto: str | None) -> str:
    """``sha256:<10 hex>`` del secreto, o ``(sin contrasena)``."""
    if not secreto:
        return "(sin contrasena)"
    return "sha256:" + hashlib.sha256(secreto.encode("utf-8")).hexdigest()[:10]


def secretos_activos() -> list[str]:
    """Valores secretos presentes en el entorno, del mas largo al mas corto."""
    valores: set[str] = set()
    for variable in VARIABLES_SECRETAS:
        crudo = os.environ.get(variable, "")
        for parte in [crudo, *crudo.split(",")]:
            parte = parte.strip()
            if len(parte) >= _LARGO_MINIMO:
                valores.add(parte)
    return sorted(valores, key=len, reverse=True)


def ocultar(texto: str, secretos: list[str] | None = None) -> str:
    for secreto in secretos if secretos is not None else secretos_activos():
        if secreto in texto:
            texto = texto.replace(secreto, huella(secreto))
    return texto


class FiltroSecretos(logging.Filter):
    """Reemplaza cualquier contrasena conocida por su huella en mensajes y tracebacks."""

    def filter(self, record: logging.LogRecord) -> bool:
        secretos = secretos_activos()
        if not secretos:
            return True
        mensaje = record.getMessage()
        limpio = ocultar(mensaje, secretos)
        if limpio != mensaje:
            record.msg, record.args = limpio, ()
        if record.exc_info and not record.exc_text:
            record.exc_text = logging.Formatter().formatException(record.exc_info)
        if record.exc_text:
            record.exc_text = ocultar(record.exc_text, secretos)
        return True
