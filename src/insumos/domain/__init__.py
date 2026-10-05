"""Generacion de Insumos - nucleo de dominio (arquitectura hexagonal).

REGLA DE DEPENDENCIA (verificada por import-linter):
    adapters -> application -> domain

Este paquete y sus subpaquetes solo pueden importar pandas, PyYAML, unidecode y
la biblioteca estandar. Queda prohibido importar win32com, selenium, openpyxl,
msoffcrypto o cualquier modulo de ``insumos.adapters`` / ``insumos.application``.
"""
