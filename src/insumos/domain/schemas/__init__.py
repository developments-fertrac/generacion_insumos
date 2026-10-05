"""Reservado para esquemas de datos del dominio (contratos de entrada y salida).

Hoy los contratos se verifican con reglas del propio pipeline:

- ``inv.validar_columnas_bd``  columnas de la exportacion de base de datos
- ``inv.validar_salida``       cuadre del inventario contra la base de datos
- ``ven.validar_salida``       columnas, filas y anio de la hoja VENTAS 2026

La dependencia ``pandera`` se retiro porque ningun modulo la usaba (ADR 0010).
Si se implementa el primer esquema, se vuelve a declarar en ``pyproject.toml``.
"""
