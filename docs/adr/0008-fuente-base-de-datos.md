# 0008. La base de datos reemplaza la descarga del ERP por Selenium

## Estado
Aceptada (2026-10-02). Reemplaza la fuente de datos asumida en el plan
(secciones 4.4, 5.1 y 5.3) para inventario y ventas.

## Contexto

Las tareas `descarga_inv_general`, `descarga_valorizados` y `descarga_ventas`
automatizaban el ERP con Selenium para bajar el inventario general, cuatro
valorizados por almacen y el informe de ventas. `actualizacion_inv` luego
reconstruia la existencia y el costo cruzando esos archivos (5.300 lineas,
la mayoria parches).

La base de datos ahora exporta directamente:

- `Inventario.xlsx`: por referencia, `EXISTENCIA BRUTA`, faltantes, impo,
  aforo, `EXISTENCIA NETA`, `COSTO`, `TOTAL INV` y un `MOTIVO` que dice si entra
  al informe (`INVENTARIO...`) o por que no (`EXCLUIDO: ...`).
- `InformesDeVentas(Facturas)_268*.xlsx`: mismo nombre y estructura que el
  informe que se descargaba.

## Decision

1. Se retiran las tres descargas por Selenium y su infraestructura
   (`core/browser.py`, `core/erp_navigation.py`, `config/erp_selectors.py`).
   `chromedriver_utils.py` se conserva: lo usa el envio por WhatsApp.
2. Inventario se reescribe sobre la arquitectura hexagonal con dos pipelines:
   - `inventario_bd.yaml`: valida columnas, filtra `MOTIVO` que inicia con
     `INVENTARIO`, quita duplicados y pone costo 0 donde la existencia es 0.
   - `inventario.yaml`: la base de datos decide que referencias forman el
     informe (las ausentes salen con su motivo, las nuevas entran);
     EXISTENCIA = EXISTENCIA NETA, COSTO PROMEDIO = COSTO, TOTAL INV = TOTAL INV;
     columnas *sistema* desde la BD; NOMBRE LISTA/MYR, COPIA, lider y
     clasificacion como en el legacy; orden por TOTAL INV; cuadre obligatorio.
3. Valorizados, marcas propias, eliminaciones por linea y criterios de negocio
   ya no se aplican en Python: los resuelve la base de datos.
4. `LLEGA AL INFORME` no se usa como filtro: el criterio acordado es `MOTIVO`
   (las referencias con existencia 0 llegan con `NO` y deben quedar en 0).
5. Ventas conserva su codigo; solo cambia la carpeta de origen
   (`DB_EXPORT_DIR`) y la seleccion del archivo: si no hay uno fechado de hoy,
   se toma el ultimo modificado (el vigente de la BD no trae fecha).
6. Ventas toma LINEA / SUB-LINEA / LIDER LINEA del ultimo
   `$2026 INVENTARIO GENERAL ACTUALIZADO *.xlsx` (salida de `actualizacion_inv`,
   fuente base de datos), ya no del maestro `$2026 INVENTARIO GENERAL.xlsx`
   (decision de negocio, 2026-10-02). Por eso inventario debe correr antes que
   ventas (el workflow `completo` ya lo hace en ese orden).

## Consecuencias

- El golden capturado del legacy (`tests/golden/inventario_normal`) deja de
  ser una referencia valida: sus entradas eran ERP + valorizados. La
  equivalencia ahora se verifica contra la base de datos misma
  (`inv.validar_salida`: mismas referencias y mismo TOTAL INV, si no, no se
  escribe). Primera corrida real (2026-10-02, datos de prueba): 15.645
  referencias, TOTAL INV 51.618.683.352,21 = BD.
- El escritor COM (`adapters/excel/com_inventario.py`) solo se puede certificar
  en Windows con Excel: primera corrida con `--dry-run` y luego real,
  comparando contra el archivo de revision.
- `DB_EXPORT_DIR` pasa a ser obligatorio para inventario.
- Se corrige de paso el cifrado de la salida: el legacy llamaba
  `OfficeFile.encrypt(fout)` con la firma equivocada y dejaba el archivo sin
  contrasena (solo lo registraba en el log).
