# Golden files

El escenario `inventario_normal` (capturado del legacy con ERP + valorizados)
se retiro el 2026-10-02: desde el ADR 0008 la fuente de inventario es la
exportacion de base de datos, asi que la salida del legacy dejo de ser la
referencia correcta.

La equivalencia de inventario se verifica ahora contra la base de datos en cada
corrida (`inv.validar_salida`): mismas referencias que `MOTIVO = INVENTARIO*` y
mismo TOTAL INV. Si no cuadra, no se escribe el archivo.

Este directorio queda para los golden de **ventas** (Fase 4): inputs reales
congelados en `<escenario>/inputs`, salida del legacy en `<escenario>/expected`
(parquet), comparacion por clave de negocio con `adapters/excel/comparador.py`.
No subir datos reales al repositorio sin anonimizar.
