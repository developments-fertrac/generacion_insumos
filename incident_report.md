# Informe de Incidencias - Migración a Arquitectura Hexagonal

**Fecha:** 2026-10-02  
**Rama:** master  
**Baseline git:** f3f2ec5  
**Estado:** Fases 0-2 (infraestructura) completadas. Fases 3-6 no iniciadas (bloqueadas por certificación golden incompleta).

## Resumen

La migración ha avanzado siguiendo el enfoque strangler fig. Se completaron las fundaciones y el kernel de dominio (0-1), y se construyó la infraestructura de lectura/comparación golden (2), incluida verificación contra archivos reales. Se encontró y corrigió un bug preexistente en `tasks/actualizacion_inventario.py` que bloqueaba la captura de golden. La captura de golden para **inventario_normal** fue exitosa. Sin embargo, la certificación completa (8 escenarios) y el switch-over a shadow mode (5 días) no son posibles en este entorno porque faltan los inputs para los escenarios restantes, tal como se documenta más adelante.

## Incidencias relevantes

### 1. Envío de correo durante ejecución de prueba (controlado)
- **Qué ocurrió:** Al ejecutar la tarea `actualizacion_inv` contra el escenario montado, `BaseTask.run()` envió correos a las direcciones de `.env` (`ctorres@fertrac.com`, `asistentecompras@fertrac.com`, `analistacompras5@fertrac.com`) notificando éxito/fallo.
- **Impacto:** Contacto externo no intencionado durante pruebas.
- **Mitigación aplicada:** Se creó `scripts/capturar_golden.py` que parchea `EmailNotifier.send` para suprimir el envío (solo log), y se inyecta la variable de entorno o se evita llamar a SMTP. Para ejecuciones futuras, se recomienda desactivar SMTP (`enabled=False`) o redirigir a un sink local.
- **Estado:** Corregido (suprimido en script de captura).

### 2. Bug preexistente: `pd.NA` vs `NaN` abortaba FASE 5b
- **Síntoma:** `actualizacion_inv` fallaba en FASE 5b con `TypeError: boolean value of NA is ambiguous` al comparar valores.
- **Causa raíz:** Las guardas existentes solo reconocían `NaN` de numpy (`isinstance(v, float) and pd.isna(v)`). Los dtypes nullables (`string`, `Int64`, `boolean`) usan `pd.NA`, que no es `float`, pasaba el filtro y llegaba a `if old_val != new_val` donde Python evaluaba `bool(pd.NA)`.
- **Ubicación:** `tasks/actualizacion_inventario.py` — `agregar_fila_nueva_en_hoja()` (línea ~3293), `_values_map()` en dos sitios (~3497 y ~4302).
- **Corrección aplicada:** Se añadió `valor_es_vacio(valor)` que trata `None`, `pd.NA`, `NaT`, `NaN` de forma segura (envolviendo `pd.isna` con try/except). Se reemplazaron las tres guardas. Pruebas unitarias: `tests/unit/test_valor_es_vacio.py` (22 casos).
- **Verificación:** Tras el arreglo, `actualizacion_inv` completó las 7 fases con éxito (5m 42s) en el escenario montado.

### 3. Falso negativo al detectar archivos maestros (informativo)
- **Qué ocurrió:** En una verificación previa se afirmó que faltaban `MARCAS.xlsx`, `DISTRIBUCION DE MATRICES.xlsx` y `$2026 MATRIZ USD.xlsx` en `Archivos Validacion/`. En realidad **sí existían** (fechas 2026-10-01).
- **Causa:** Error de afirmación en mensaje anterior (no un fallo de herramientas). Los archivos estaban presentes al re-listar.
- **Estado:** Corregido (se retracta explícitamente).

### 4. Lectura de archivos reales con encabezados desplazados
- **Hallazgo:** Ninguno de los archivos reales tiene encabezado en fila 0. El legacy usa convenciones específicas:
  - INVENTARIO: fila de encabezado 2 (visible), auto-detección si no encuentra referencia.
  - INV LISTA PRECIOS: fila 1.
  - VALORIZADOS: fila 9 (header=None, slice, primera fila como columnas).
- **Solución:** `src/insumos/adapters/excel/lector_inventario.py` implementa estas convenciones fielmente (copiadas del legacy), con normalización de nombres, búsqueda difusa de columnas (cascada) y validación estricta (lanza `ColumnaNoEncontrada` si falta clave). Verificado con datos reales: 15,707 filas inventario, 4 valorizados correctos, 229 marcas.

## Estado por fases (plan)

| Fase | Alcance | Estado | Evidencias |
|---|---|---|---|
| 0 | Fundaciones | **Hecho** | Commits f3f2ec5 → 4a71068. `pyproject.toml`, uv.lock, `.importlinter` (INI), skeleton hexagonal, ADR 0005/0006/0007, ratchet 117. |
| 1 | Dominio (reglas + pipeline) | **Hecho** | `3ad954c`. `domain/rules/*` (11 reglas inventario), `domain/pipeline.py`, `domain/audit.py`, `registry` con `@rule`. Tests: 108/108. |
| 2 | Infra golden + baseline | **Infra hecha. Baseline parcial.** | `comparador.py`, `lector.py`, `lector_inventario.py`, `tests/unit/test_golden.py` (31). **Golden capturado:** `tests/golden/inventario_normal/` (inputs + expected parquet). Bug legacy corregido + tests. **Bloqueo:** faltan 7 escenarios más (3 inventario + 4 ventas). |
| 3 | Migrar tasks clave (strangler) | **No iniciado** | Requiere certificar golden completo (0 diferencias no explicadas) por escenario. |
| 4 | Reglas COM-coupled a DataFrame | **No iniciado** | Depende de golden completo. |
| 5 | Adapters Excel/Selenium | **No iniciado** | Depende de fases anteriores. |
| 6 | Switch-over + shadow mode 5 días | **No iniciado** | Criterio: 0 diferencias + 5 días shadow en producción. No factible sin todos los golden. |

## Bloqueos para completar fases 3-6

**Bloqueante (técnico):** El protocolo golden exige 4 escenarios por inventario y 4 por ventas (8 totales), con sus inputs congelados bajo `tests/golden/<escenario>/inputs`. En este entorno solo disponemos de los archivos necesarios para **1 escenario** (`inventario_normal`). Los inputs para los escenarios restantes **no existen** en `Archivos Validacion/` (o no fueron proporcionados):

Escenarios pendientes (según plan §7/§8):
- Inventario: "start of month", "ERP with new references", "negative stock" (además de normal)
- Ventas: 4 escenarios con sus inputs correspondientes

**Bloqueante (operativo):** El criterio de switch-over incluye **5 días consecutivos de shadow mode en producción** (ambas versiones en paralelo). Eso requiere acceso a entorno productivo (ERP Fertrac, credenciales, datos diarios reales) y no puede realizarse en este entorno de desarrollo aislado.

## Entregables verificados

- Infra lector: verifica descifrado (MATRIZ USD 21MB cifrado), hojas, columnas reales. Scripts: `scripts/probar_lector_real.py`, `scripts/probar_lector_inventario.py`.
- Lector inventario con convenciones legacy: pasa verificación contra datos reales.
- Golden baseline parcial: `inventario_normal` capturado desde implementación actual (exitoso). Parquets de inventario + 4 hojas de reporte de eliminaciones.
- Test skeleton golden: parametrizado, skips explícitos (razón documentada). Escenario detectado correctamente.
- Fix legacy: `valor_es_vacio` + 22 tests. No introduce regresiones (suite existente intacta).
- Herramientas de captura/escenario: `scripts/montar_escenario.py`, `scripts/capturar_golden.py` (sin correo).

## Recomendación

1. **Entregar este informe** y los archivos de infraestructura + golden parcial + fix, **sin reclamar finalización de fases 3–6**.
2. Solicitar al usuario que aporte los inputs para los 7 escenarios restantes (o confirme que no están disponibles). Con ellos se puede completar la matriz golden.
3. Solo proceder con 3–6 cuando: (a) todos los golden pasan con 0 diferencias no explicadas; (b) exista autorización y entorno para shadow mode de 5 días.

## Conclusión

**Fases 0–1:** completas y verificadas.  
**Fase 2:** infraestructura completa y funcional; baseline **parcial** (1/8 escenarios).  
**Fases 3–6:** no iniciadas, bloqueadas por requisitos de certificación golden + shadow mode que no pueden satisfacerse con los datos presentes.

El código entregado no rompe contrato externo (`run.py`, Orchestrator, BaseTask, TASK_REGISTRY). La migración avanza por strangler fig, con golden como red de seguridad y ratchet de tipos mantenido.
