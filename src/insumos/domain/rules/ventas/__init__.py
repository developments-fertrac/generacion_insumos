"""Reglas de ventas (``$2026 VENTAS``), pipeline ``config/rules/ventas.yaml``.

- ``informe``        limpieza del informe de facturas (_268)
- ``integraciones``  cruces con inventario, MYR, matriz de clientes y licitados
- ``salida``         forma final de la hoja VENTAS 2026 y validacion

Importar este paquete registra todas las reglas.
"""

from __future__ import annotations

from insumos.domain.rules.ventas import informe, integraciones, salida

__all__ = ["informe", "integraciones", "salida"]
