"""Reglas de inventario general.

Desde 2026-10 la fuente es la exportacion de base de datos (ver ADR 0008):

- ``base_datos``    reglas sobre ``Inventario.xlsx`` (pipeline ``inventario_bd``)
- ``actualizacion`` reglas sobre la plantilla (pipeline ``inventario_general``)

Importar este paquete registra todas las reglas, que el YAML refiere por id.
"""

from __future__ import annotations

from insumos.domain.rules.inventario import actualizacion, base_datos

__all__ = ["actualizacion", "base_datos"]
