"""Contrato de los pipelines declarativos (config/rules/*.yaml).

Este modulo existe desde Fase 0 para que la proteccion contra pipelines vacios
esté **codificada y probada**, no apoyada en un comentario que alguien puede no
leer. El loader de Fase 1 se apoya en estas mismas funciones.

Por que importa
---------------
``config/rules/inventario.yaml`` arranca con ``rules: []`` y asi debe permanecer
hasta la Fase 3. Si alguien lo conectara antes de tiempo, el inventario pasaria
completo y sin transformar: el archivo de salida tendria el tamano correcto y
las cifras de ayer. Es el peor modo de fallo posible en un pipeline: silencioso
y con apariencia de exito.

Por eso un pipeline vacio es un **error explicito**, no un no-op.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


class ConfiguracionInvalida(ValueError):
    """El YAML es legible pero no describe un pipeline ejecutable."""


@dataclass(frozen=True)
class EntradaRegla:
    """Una regla tal como aparece en el YAML, todavia sin resolver."""

    id: str
    enabled: bool = True
    params: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def desde_datos(cls, datos: Any, donde: str) -> EntradaRegla:
        if isinstance(datos, str):
            return cls(id=datos)
        if not isinstance(datos, dict):
            raise ConfiguracionInvalida(f"{donde}: se esperaba str o mapa, hay {type(datos).__name__}")

        identificador = datos.get("id")
        if not identificador or not isinstance(identificador, str):
            raise ConfiguracionInvalida(f"{donde}: falta 'id' o no es texto")

        enabled = datos.get("enabled", True)
        if not isinstance(enabled, bool):
            raise ConfiguracionInvalida(f"{donde}: 'enabled' debe ser true/false, hay {enabled!r}")

        params = datos.get("params", {})
        if not isinstance(params, dict):
            raise ConfiguracionInvalida(f"{donde}: 'params' debe ser un mapa")

        return cls(id=identificador, enabled=enabled, params=params)


@dataclass(frozen=True)
class DeclaracionPipeline:
    """Lo que declara el YAML, ya validado. Sin reglas resueltas todavia."""

    nombre: str
    version: str
    reglas: tuple[EntradaRegla, ...]
    allow_empty: bool

    @property
    def reglas_activas(self) -> tuple[EntradaRegla, ...]:
        return tuple(r for r in self.reglas if r.enabled)


def _leer(ruta: Path) -> dict[str, Any]:
    if not ruta.is_file():
        raise ConfiguracionInvalida(f"No existe el archivo de reglas: {ruta}")
    datos = yaml.safe_load(ruta.read_text(encoding="utf-8"))
    if datos is None:
        datos = {}
    if not isinstance(datos, dict):
        raise ConfiguracionInvalida(f"{ruta.name}: el archivo debe ser un mapeo")
    return datos


def leer_declaracion(ruta: Path) -> DeclaracionPipeline:
    """Valida la forma del YAML. NO ejecuta nada.

    Falla si el pipeline esta vacio y no se declaro ``allow_empty: true``.
    """
    datos = _leer(ruta)

    nombre = datos.get("pipeline")
    if not nombre or not isinstance(nombre, str):
        raise ConfiguracionInvalida(f"{ruta.name}: falta 'pipeline'")

    version = datos.get("version")
    if version is None or not isinstance(version, (str, int, float)):
        raise ConfiguracionInvalida(f"{ruta.name}: falta 'version'")

    allow_empty = datos.get("allow_empty", False)
    if not isinstance(allow_empty, bool):
        raise ConfiguracionInvalida(f"{ruta.name}: 'allow_empty' debe ser true/false")

    reglas_crudas = datos.get("rules")
    if not isinstance(reglas_crudas, list):
        raise ConfiguracionInvalida(f"{ruta.name}: 'rules' debe ser una lista")

    reglas = tuple(
        EntradaRegla.desde_datos(dato, f"{ruta.name}: rules[{i}]")
        for i, dato in enumerate(reglas_crudas)
    )

    declaracion = DeclaracionPipeline(
        nombre=nombre,
        version=str(version),
        reglas=reglas,
        allow_empty=allow_empty,
    )

    if not declaracion.reglas_activas and not allow_empty:
        raise ConfiguracionInvalida(
            f"{ruta.name}: el pipeline '{nombre}' no tiene reglas activas. "
            "Ejecutarlo pasaria los datos sin transformar, en silencio. "
            "O bien declara las reglas, o bien pon 'allow_empty: true' "
            "si de verdad ese pipeline debe ser un no-op."
        )

    return declaracion
