"""Adaptadores de Excel: lectores, escritores COM/openpyxl y optimizador.

Aqui vive el unico codigo que abre y escribe archivos de Excel:
    - readers.py:          msoffcrypto + pandas (descifrado, descubrimiento por prefijo)
    - com_writer.py:       win32com (plantilla, formulas, pivots, subtotales)
    - openpyxl_writer.py:  alternativa sin COM
    - optimizer.py:        xlsx_cleaner + copia .xlsb
"""
