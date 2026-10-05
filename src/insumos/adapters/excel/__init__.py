"""Adaptadores de Excel: el unico codigo del nucleo que abre y escribe libros.

- ``lector``              ExcelReader: lee .xlsx cifrados o no, descifrando en memoria.
- ``maestros``            Matriz USD y Distribucion de matrices.
- ``fuentes_inventario``  puerto FuenteInventario sobre las carpetas de Fertrac.
- ``fuentes_ventas``      entradas de ventas (plantilla, inventario, MYR, matriz, licitados).
- ``com_inventario``      escritor de produccion via Excel COM (+ funciones puras).
- ``escritor_openpyxl``   escritor sin Excel sobre la plantilla real (``--dry-run``).
- ``escritor_simple``     titulo diario de la columna de existencia.
- ``reporte_inventario``  REPORTE_ELIMINACIONES_*.xlsx.
- ``comparador``          comparacion golden (por clave) y de paridad (por posicion).
"""
