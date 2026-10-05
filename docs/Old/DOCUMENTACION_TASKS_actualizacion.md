# Documentacion - Proceso de Actualizacion de Inventario

Task: `actualizacion_inventario`
Modulo: `tasks/actualizacion_inventario.py`
Clase: `ActualizacionInventario(BaseTask)`

---

## 1. Que hace el proceso

Consolida el inventario general tomando como base la salida del dia
anterior (o la plantilla maestra), le cruza los datos del ERP y de los
valorizados, completa los campos de negocio que falten y elimina los
registros que no deben aparecer en el reporte. Genera un archivo nuevo
`$2026 INVENTARIO GENERAL ACTUALIZADO <AAAAMMDD_HHMM>.xlsx`, un respaldo
de la hoja `INVENTARIO` en la hoja `INVENTARIO COPIA`, un reporte de
eliminaciones y un correo de notificacion con estadisticas.

### Como se ejecuta

```bash
python run.py --task actualizacion_inv
```

Tambien corre dentro del pipeline completo (`python run.py`) en la
Fase 2, y dentro del workflow de inventario:

```bash
python run.py --workflow inventario
```

### Patron de diseno

Sigue el patron **Template Method** documentado en `README.md`: `BaseTask.run()`
ejecuta `setup() -> execute() -> teardown()` y, al terminar, llama a
`_notify_success()` (exito) o `_notify_failure()` (fallo) para enviar el
correo. `execute()` nunca envia correos a mano ni lanza `SystemExit`: si
algo falla, la excepcion se propaga para que `run()` la atrape una sola vez.

---

## 2. Rutas e insumos usados

Tomadas de `config/settings.py` (variables de entorno del `.env`):

| Insumo | Ruta | Prefijo / patron de busqueda |
|---|---|---|
| Inventario actualizado (ERP) | `BASE_PATH\INFORMES\INVENTARIO GENERAL ACTUALIZADO\<MM. MES>\` | `INVENTARIO GENERAL ACTUALIZADO` |
| Valorizados diarios | `BASE_PATH\Pruebas Inv General\Valorizados\` | `VALORIZADO GENERAL`, `VALORIZADO FALTANTES IMPO`, `VALORIZADO FALTANTES`, `VALORIZADO TOBERIN` |
| Matriz USD | `BASE_PATH\` (carpeta raiz de insumos) | `2026 MATRIZ USD` |
| Marcas propias | `BASE_PATH\` | `MARCAS` |
| Distribucion de matrices | `BASE_PATH\` | `DISTRIBUCION DE MATRICES` |
| Consolidado de remisiones | raiz / remisiones del mes / carpeta ERP | `CONSOLIDADO REMISIONES 2019-2026` |
| Plantilla (salida previa o maestra) | `BASE_PATH\Pruebas Inv General\` (salida previa) o raiz (maestra) | `$2026 INVENTARIO GENERAL ACTUALIZADO` |
| Salida del proceso | `BASE_PATH\Pruebas Inv General\` | `$2026 INVENTARIO GENERAL ACTUALIZADO <AAAAMMDD_HHMM>.xlsx` |

El archivo ERP se elige con `find_by_prefix()`: el mas reciente por fecha
de modificacion. Si ninguno de HOY esta disponible se usa igual el mas
reciente y se deja **evidencia en el log, el correo y el reporte de
eliminaciones** (a traves de `_marcar_fuente_erp` y el `EliminacionTracker`).

### Hojas y encabezados del archivo de salida

- Hoja de trabajo: `INVENTARIO` (fila de encabezados = 2).
- Hoja de respaldo: `INVENTARIO COPIA` (clon completo, con formatos).
- Columnas relevantes: `REFERENCIA` (col A), `NOMBRE LISTA`,
  `NOMBRE ODOO` (col C), `NOMBRE MYR`, `MARCA copia`,
  `INV BODEGA GERENCIA`, `EXISTENCIA <MES DIA>` (encabezado se
  actualiza a la fecha de hoy), `COSTO PROMEDIO`, `LINEA COPIA`,
  `SUB-LINEA COPIA`, `LIDER LINEA`, `CLASIFICACION`,
  `Marca sistema`, `Linea sistema`, `Sub- linea sistema`, `TOTAL INV`,
  `Dif marca`, `Dif linea`, `Dif sub-linea`.

---

## 3. Flujo por fases

El log registra cada fase con `>>> INICIO ...` / `<<< OK ...` y, si
falla, `<<< FALLO ...`. Las fases son:

1. `SETUP` - Configuracion de rutas, se crean directorios y se instancia el
   notificador de correo.
2. `FASE 1` - Cargar insumos.
3. `FASE 2` - Leer plantilla de inventario general.
4. `FASE 3` - Actualizar referencias.
5. `FASE 4` - Aplicar reglas de marcas propias (hoy: exclusion explicita).
6. `FASE 4b` - Referencias nuevas sin clasificar (LINEA COPIA / LIDER LINEA).
7. `FASE 5` - Eliminaciones y limpieza (duplicados, descontinuadas y
   existencias negativas).
8. `FASE 5b` - Escribir resultados calculados en la hoja `INVENTARIO`.
9. `FASE 5c` - Completar campos en `INVENTARIO`, eliminar lineas invalidas
   y criterios de negocio, generar backup `INVENTARIO COPIA`, subtotales.
10. `FASE 6` - Guardar resultado (y proteger con contrasena si aplica).
11. `FASE 7` - Preparar notificacion.
12. `TEARDOWN` - Cierre de Excel/COM y limpieza de temporales.

### FASE 1: Cargar insumos

- **1.1 Inventario actualizado (ERP):** `cargar_inventario_actualizado()`
  busca el archivo `INVENTARIO GENERAL ACTUALIZADO` mas reciente, lo lee y
  normaliza como `__REFERENCIA__`, `__NOMBRE__`, `__MARCA_SYS__`,
  `__LINEA_SYS__`, `__SUBLINEA_SYS__`, `__COSTO__`. Se registra la fuente
  (archivo + fecha de modificacion) en el `EliminacionTracker`. Si no hay
  descarga del ERP se usa la plantilla propia como fallback (queda marcado).
- **1.2 Valorizados:** se cargan todos los `VALORIZADO*` encontrados para
  el conteo general (`df_val_list`).
- **1.2b Valorizados por categoria:** se cargan por separado los 4
  valorizados identificados por prefijo para calcular la existencia real:
  `EXISTENCIA = GENERAL - FALTANTES IMPO - FALTANTES - TOBERIN`
  (`calcular_existencia_valorizados()`). Si alguno falta se usa vacio (0).
- **1.3 Matriz USD:** `cargar_matriz_usd()` -> `__REF_MATRIZ__`,
  `__DESC_LISTA__`, `__REF_LISTA_PRECIOS__`. Se usa para NOMBRE LISTA.
- **1.4 Marcas:** `cargar_marcas()` -> `set` de nombres de marca
  (columna `MARCAS` del archivo `MARCAS.xlsx`).
- **1.5 Distribucion:** `cargar_distribucion()` -> `{'gestor': {LINEA: gestor},
  'clasificacion': {LINEA: categoria}}` desde `DISTRIBUCION DE MATRICES`.

  Nota: replica el comportamiento del script original: si una LINEA aparece
  en varias filas del archivo, gana la ULTIMA; a las claves se les quita el
  sufijo entre parentesis (ej. `REDAT (EN DESARROLLO)` -> `REDAT`).
- **1.6 Consolidado de remisiones:** `cargar_consolidado_remisiones()` ->
  `__REF_REM__`, `__DESC_REM__`. Se busca en raiz, remisiones del mes y
  carpeta del ERP (la primera donde aparezca).

### FASE 2: Leer plantilla de inventario general

- Se toma la **salida previa** de `Pruebas Inv General\` (`$2026 INVENTARIO
  GENERAL ACTUALIZADO*` mas reciente). Si no existe, se busca la **plantilla
  maestra** en la raiz de insumos (`2026 INVENTARIO GENERAL*`); como ultimo
  recurso, la descarga diaria del ERP.
- Se lee la hoja `INVENTARIO` con fila de encabezados en 2.
- Se crean `__REFERENCIA__` (referencia limpia) y `__REFERENCIA_ACT__`.
- Red de seguridad: si `INVENTARIO` llega sin datos (archivo corrupto/corte),
  se lee el respaldo del dia anterior desde `INVENTARIO COPIA`.

### FASE 3: Actualizar referencias

- **3.0** `actualizar_referencias_inventario_original()` resuelve cadenas de
  aliases (`resolver_cadena_referencias`) y llama a
  `fusionar_datos_erp_en_inventario()`, que cruza por referencia y agrega:
  `__NOMBRE_ERP__`, `__MARCA_ERP__`, `__LINEA_ERP__`, `__SUBLINEA_ERP__`,
  `__COSTO_ERP__`.
- **3.1** `agregar_referencias_nuevas_desde_erp()`: detecta las referencias
  que estan en el ERP pero NO en la plantilla y las inserta como filas
  nuevas (con NOMBRE ODOO del ERP, NOMBRE LISTA desde Matriz USD o `"0"`,
  Marca/Linea/Sub-linea sistema y Costo). Este fue el cierre del hueco
  reportado: antes ~97 referencias del ERP no aparecian en el informe.
- **3.1b** `calcular_nombre_lista_y_myr()`: recalcula para TODAS las filas:
  - `NOMBRE LISTA` = descripcion en Matriz USD; `"0"` si hay Matriz USD
    pero la referencia no aparece.
  - `NOMBRE MYR` = NOMBRE LISTA si tiene contenido, si no NOMBRE ODOO
    (el `"0"` de NOMBRE LISTA cuenta como vacio; nunca se queda en `"0"`).
- **3.2** `aplicar_existencia_calculada()`: escribe `__EXISTENCIA_CALC__`
  (EXISTENCIA de hoy calculada de los valorizados). Referencias sin
  valorizado quedan en 0.

### FASE 4: Aplicar reglas marcas propias

- `aplicar_reglas_marcas_propias()`: hoy solo hace la **exclusion explicita**
  de la lista `REFERENCIAS_A_ELIMINAR` (ver condiciones de eliminacion).
  La reclasificacion forzada de marcas propias se retiro: el llenado de
  `MARCA copia`/`LINEA COPIA`/`SUB-LINEA COPIA`/`LIDER LINEA`/
  `CLASIFICACION` se hace ahora en FASE 5c con la LINEA real del ERP.

### FASE 4b: Referencias nuevas sin clasificar

- `eliminar_referencias_nuevas_sin_clasificar()`: replica el criterio del
  script original (nunca dejar en el reporte una fila sin `LINEA COPIA` ni
  sin `LIDER LINEA`), pero **solo para referencias nuevas** de esta corrida.
- Por defecto **no elimina**: solo reporta candidatas (flag
  `HABILITAR_ELIMINACION_NUEVAS_SIN_CLASIFICAR=False`).

### FASE 5: Eliminaciones y limpieza (sobre el DataFrame y la hoja abierta)

- Se abre la plantilla con Excel COM (`excel_open`), en modo editable; si
  esta protegida se descifra a un temporal `TEMP_INVENTARIO_*.xlsx`.
- `eliminar_registros_linea_copia_indeterminada()`:
  - **Paso 1 (borra siempre):** registros con referencia duplicada dentro
    de `INVENTARIO COPIA`.
  - **Paso 2 (solo reporta por defecto):** referencias del inventario que
    ya NO existen en el ERP de hoy (candidatas a "referencia
    descontinuada"). Para eliminarlas automaticamente hay que activar
    `HABILITAR_ELIMINACION_REFERENCIA_DESCONTINUADA=True`.
- `procesar_existencias_negativas_y_cero()`: lleva a 0 las existencias
  negativas (no es una eliminacion; es una correccion que se reporta).
  Las existencias en 0 no se tocan.

### FASE 5b: Escribir resultados calculados en `INVENTARIO`

- `escribir_resultados_en_hoja_inventario()` vuelca en la hoja real lo que
  calcularon FASE 1/3/4/5:
  - Actualiza el encabezado de existencia al dia de hoy
    (`actualizar_encabezado_existencia`).
  - Para filas ya existentes (por `__REFERENCIA__`): NOMBRE ODOO,
    NOMBRE LISTA, NOMBRE MYR, Marca/Linea/Sub-linea sistema, Costo
    promedio y EXISTENCIA de hoy.
  - Para filas nuevas: las agrega fisicamente al final de la hoja
    (`agregar_fila_nueva_en_hoja`), copiando las formulas de
    `TOTAL INV`, `Dif marca`, `Dif linea`, `Dif sub-linea` desde la fila
    anterior.
  - Red de seguridad: si la hoja llega sin filas de datos, vuelca todo el
    DataFrame de una vez (`escribir_hoja_completa_desde_df`).

### FASE 5c: Completar campos, eliminar lineas invalidas y generar backup

1. `calcular_completado_inventario_copia()` + `escribir_completado_en_hoja_inventario()`
   completan en `INVENTARIO` los 7 campos pedidos, con la regla
   "si ya hay un valor util se conserva; si no, se completa":
   - `NOMBRE MYR`: existente, o el mas largo entre NOMBRE ODOO y NOMBRE
     LISTA (con respaldo a los valores ya arrastrados en la plantilla).
   - `MARCA copia` / `LINEA COPIA` / `SUB-LINEA COPIA`: existente, o el
     valor real del ERP (`Marca/Linea/Sub-linea sistema`).
   - `INV BODEGA GERENCIA`: existente, o `"0"`.
   - `LIDER LINEA` / `CLASIFICACION`: existente, o cruce de `LINEA COPIA`
     ya resuelta contra `DISTRIBUCION DE MATRICES`.
2. `completar_nombre_myr_en_hoja()`: barrido final sobre TODAS las filas
   fisicas de la hoja (incluidas las que el DataFrame no capture) para que
   NOMBRE MYR nunca quede vacio/`"0"` cuando hay un nombre disponible.
3. `eliminar_registros_linea_invalida()`: **borra fisicamente** las filas
   cuya `LINEA COPIA` quedo invalida (ver condiciones de eliminacion).
4. `eliminar_registros_por_criterios_negocio()`: **borra fisicamente** las
   filas que cumplen los criterios de negocio de NOMBRE ODOO / REFERENCIA
   (ver condiciones de eliminacion).
5. Se genera el reporte de eliminaciones
   (`EliminacionTracker.mostrar_resumen()` + `generar_reporte_excel()`).
6. `columna_referencia_como_texto()`: deja la columna A/`REFERENCIA` con
   formato TEXTO.
7. `agregar_subtotales_finales()`: agrega una fila de subtotales al final
   con `=SUBTOTAL(109;...)` en las columnas de EXISTENCIA y TOTAL INV
   (reutiliza la fila de subtotales del dia anterior si la hubiera).
8. `clonar_hoja_completa()`: clona `INVENTARIO` a `INVENTARIO COPIA`
   (valores + formatos + anchos de columna), para que el respaldo sea
   identico al reporte ya terminado y sin las filas eliminadas.

Nota: si la hoja `INVENTARIO COPIA` no se puede crear/encontrar, se
completa `INVENTARIO` sin generar respaldo (queda registrado en el log).

### FASE 6: Guardar resultado

- Guarda como `$2026 INVENTARIO GENERAL ACTUALIZADO <AAAAMMDD_HHMM>.xlsx`
  en `Pruebas Inv General\`.
- Si `APPLY_PASSWORD_TO_OUTPUT` y hay contrasena configurada, se protege
  el archivo con msoffcrypto.
- Cierra Excel y elimina el temporal de plantilla.

### FASE 7: Preparar notificacion

- Arma el detalle del correo con estadisticas (inventario original,
  referencias actualizadas, referencias nuevas agregadas, valorizados,
  Matriz USD, Marcas, Distribucion, Remisiones, eliminaciones, archivos
  generados) y la evidencia de la **fuente ERP** (con advertencia
  destacada si el archivo no era de hoy o si se uso fallback).
- El envio real lo hace `BaseTask.run()` -> `_notify_success()`, adjuntando
  el archivo generado como `Path` (no como `str`).

---

## 4. Condiciones de eliminacion (hasta el momento)

Todas las eliminaciones se registran en el `EliminacionTracker`
(`TRACKER_ELIMINACIONES`) y aparecen en el **REPORTE_ELIMINACIONES_<fecha>.xlsx**
y en el resumen del log. El reporte incluye: hoja `FUENTES DE DATOS`,
`TODAS LAS ELIMINACIONES`, `RESUMEN POR PASO` y una hoja por paso.

### 4.1 FASE 4 - Exclusion explicita (se borra SIEMPRE)

Lista `REFERENCIAS_A_ELIMINAR` (hardcodeada en el modulo, literales):

```
0041R, 43, 42, 41, 44, 45, 47, 107,
5566507 FULLER ARM NO FACTURABLE, FSB5406B/TAC1711 DESARME,
0041DESARME CAJA, 104502-2 NO FACT, 4842 NO FACTURABLE,
6-4-7771-1X NO FACT
```

Si la REFERENCIA coincide (normalizada) con alguno de estos valores, la
fila se elimina. Paso en el reporte: `EXCLUSION EXPLICITA (no facturable)`.

### 4.2 FASE 5 - Paso 1: Referencias duplicadas (se borra SIEMPRE)

Si una misma REFERENCIA aparece mas de una vez (duplicado en la hoja
`INVENTARIO COPIA`), se eliminan las filas duplicadas (se conserva la
primera). Paso en el reporte: `PASO 1: REFERENCIA DUPLICADA`.

### 4.3 FASE 5 - Paso 2: Referencias descontinuadas (solo candidatas por defecto)

Referencias del inventario que **ya no existen** en el inventario
actualizado del ERP de HOY. Por defecto `HABILITAR_ELIMINACION_REFERENCIA_DESCONTINUADA=False`:
solo se reportan como candidatas. Si se pone en `True`, se eliminan.
Paso en el reporte: `PASO 2: REFERENCIA DESCONTINUADA (candidata, no eliminada)`
o `... (eliminada)` con el flag activado.

### 4.4 FASE 4b - Referencias nuevas sin clasificar (solo candidatas por defecto)

Solo para referencias NUEVAS de esta corrida: si su `LINEA COPIA` quedo
vacia/`INDETERMINADO`/`#N/D`/`#N/A`/`N/A`/`NA`/`NONE`, **o** su
`LIDER LINEA` quedo vacio, se consideran candidatas. Por defecto
`HABILITAR_ELIMINACION_NUEVAS_SIN_CLASIFICAR=False`: solo se reportan.
Con el flag en `True` se excluyen antes de agregarse a la hoja.

### 4.5 FASE 5c - LINEA COPIA invalida (se borra SIEMPRE)

`eliminar_registros_linea_invalida()` borra fisicamente las filas cuya
`LINEA COPIA` (columna con ese nombre en la hoja) sea cualquiera de:

| Valor | Significado |
|---|---|
| (vacia) | Sin linea asignada |
| `VARIOS` | Linea sin clasificar |
| `0` / `0.0` | Placeholder de linea no resuelta |
| `CERO` / `CERO.` | Placeholder de linea no resuelta |

Paso en el reporte: `REFERENCIA SIN LINEA VALIDA (LINEA COPIA)`.
Se ejecuta DESPUES de completar LINEA COPIA (FASE 5c) y ANTES de clonar
a `INVENTARIO COPIA`, para que el respaldo NO contenga estas filas.

### 4.6 FASE 5c - Criterios de negocio NOMBRE ODOO / REFERENCIA (se borra SIEMPRE) - NUEVO

`eliminar_registros_por_criterios_negocio()` borra fisicamente las filas
que cumplan **cualquiera** de estos 4 criterios sobre la hoja `INVENTARIO`
del archivo `$2026 INVENTARIO GENERAL ACTUALIZADO`:

| # | Condicion | Columna | Paso en reporte |
|---|---|---|---|
| 1 | `NOMBRE ODOO` es **exactamente** `PUBLICIDAD DANA` | C (`NOMBRE ODOO`) | `NOMBRE ODOO = PUBLICIDAD DANA` |
| 2 | `NOMBRE ODOO` **empieza con** `MANO DE OBRA` | C (`NOMBRE ODOO`) | `NOMBRE ODOO INICIA CON MANO DE OBRA` |
| 3 | `REFERENCIAS` **termina en** `DF` | A (`REFERENCIAS`) | `REFERENCIA TERMINA EN DF` |
| 4 | `REFERENCIAS` **empieza con** `004` **y** largo **<= 5** | A (`REFERENCIAS`) | `REFERENCIA INICIA 004 LARGO <= 5` |

Detalles de implementacion:

- La comparacion es **case-insensitive** y sobre el valor limpio (se quita
  el sufijo `.0` que Excel pone a numeros, preservando ceros iniciales).
- Cada fila se registra con su motivo como
  `Criterio de negocio: <paso>` en el reporte.
- Se ejecuta junto con la eliminacion de LINEA invalida (FASE 5c), ANTES
  de clonar a `INVENTARIO COPIA`, asi el respaldo no contiene las filas
  borradas. Tambien se filtran las referencias eliminadas del DataFrame en
  memoria y de la lista de referencias nuevas para mantener estadisticas
  coherentes.

---

## 5. Flags de seguridad (eliminacion automatica)

| Flag | Valor por defecto | Efecto al ponerlo en `True` |
|---|---|---|
| `HABILITAR_ELIMINACION_REFERENCIA_DESCONTINUADA` | `False` | Elimina las referencias que ya no existen en el ERP de hoy |
| `HABILITAR_ELIMINACION_NUEVAS_SIN_CLASIFICAR` | `False` | Excluye referencias nuevas sin LINEA COPIA / LIDER LINEA |

Las eliminaciones de las secciones 4.1, 4.2, 4.5 y 4.6 **no** tienen flag:
se ejecutan siempre.

---

## 6. Reporte de eliminaciones

Archivo: `REPORTE_ELIMINACIONES_<AAAAMMDD_HHMMSS>.xlsx` en
`Pruebas Inv General\`. Se genera aunque no haya eliminaciones (si hay
fuentes registradas). Hojas:

- `FUENTES DE DATOS`: que archivo se uso como fuente de cada insumo y su
  fecha de modificacion; se resalta en rojo cuando el archivo no es de hoy.
- `TODAS LAS ELIMINACIONES`: cada registro con TIMESTAMP, PASO, FILA_EXCEL,
  REFERENCIA, NOMBRE, MARCA, LINEA y MOTIVO; coloreado por paso.
- `RESUMEN POR PASO`: total por paso, porcentaje y ejemplos.
- Una hoja por cada paso (con eliminaciones).

---

## 7. Salidas del proceso

1. `$2026 INVENTARIO GENERAL ACTUALIZADO <AAAAMMDD_HHMM>.xlsx`
   (hojas `INVENTARIO` e `INVENTARIO COPIA`, con la fila de subtotales y
   la columna A en formato TEXTO).
2. `REPORTE_ELIMINACIONES_<AAAAMMDD_HHMMSS>.xlsx`.
3. Correo SMTP con estadisticas + archivo adjunto (si el notificador esta
   habilitado en `config/settings.py`).

---

## 8. Notas de mantenimiento

- Para agregar/quitar una referencia de la exclusion explicita editar
  `REFERENCIAS_A_ELIMINAR` (seccion 4.1).
- Para agregar/ajustar un criterio de negocio editar
  `eliminar_registros_por_criterios_negocio()` (seccion 4.6).
- Los nombres de pasos se controlan con constantes:
  `NOMBRE_PASO_LINEA_INVALIDA`, `NOMBRE_PASO_PUBLICIDAD_DANA`,
  `NOMBRE_PASO_MANO_DE_OBRA`, `NOMBRE_PASO_REF_TERMINA_DF`,
  `NOMBRE_PASO_REF_004_SI_NO`.
- La seleccion del ERP depende de que el archivo del dia ya este
  descargado; si no, el proceso usa el mas reciente disponible y lo
  advertira (revisar la seccion "Fuente ERP" del correo/reporte).