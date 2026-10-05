"""Adaptadores conducidos (driven): implementaciones reales de los puertos.

- ``excel``:  lectura (msoffcrypto + pandas), escritura de la plantilla
  (COM en produccion, openpyxl en ``--dry-run``), reporte de auditoria y
  comparadores de salidas.
- ``system``: busqueda de archivos por prefijo y reloj en hora de Colombia.
- ``notify``: reservado para SMTP y WhatsApp Web (Fase 5).

Aqui SI se permite importar win32com, openpyxl y msoffcrypto.
"""
