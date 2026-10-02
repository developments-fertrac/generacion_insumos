"""Esquemas pandera: contratos de datos en la entrada y salida del pipeline.

- ``InventarioSchema``: DataFrame de inventario general
- ``ValorizadoSchema``: DataFrame de valorizados por almacen
- ``VentasSchema``: DataFrame de ventas

Un esquema fallido NO detiene el pipeline: se usa ``lazy=True`` para reportar
todos los errores juntos y las reglas se encargan de aislar las filas malas
devolviendo ``RuleResult.removed`` con su motivo.
"""
