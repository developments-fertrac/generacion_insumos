# 06 · Decisiones de arquitectura (ADR)

> Fuente: texto de los ADR recuperado de `5e249ec`. **Los ADR no forman parte del grafo**: el grafo se genera sobre `HEAD` (`088f6f0`) y ahí `docs/adr/` no existe. Lo que sí está en el grafo es la implementación de cada decisión (nodos de código y conceptos re-anclados al código, citados en cada ADR).
> **Advertencia de trazabilidad**: en la rama `main` (`088f6f0`) **no existe** `docs/adr/`. El commit `86f7093` ("Eliminacion de documentacion obsoleta") borró ADR 0005–0010 y su índice; el texto se recuperó de `5e249ec`. ADR 0001–0004 nunca tuvieron archivo: solo una fila en el índice. README, `tasks/__init__.py`, `tasks/actualizacion_inventario.py` (docstring) y `scripts/publicar_a_produccion.ps1` siguen apuntando a `docs/adr/` (nodo `readme_pendiente_docs_adr_eliminados`).

## 1. Índice

| ADR | Título | Estado | Fecha | Archivo en la rama | Nodo en el grafo |
|---|---|---|---|---|---|
| 0001 | Migración incremental (strangler fig), no reescritura | Aceptada | **[NO VERIFICADO]** | No (solo índice) | — (no está en `HEAD`) |
| 0002 | pandas (+pandera) como portador de datos del dominio | Aceptada; pandera retirado por 0010 | **[NO VERIFICADO]** | No (solo índice) | — (no está en `HEAD`) |
| 0003 | Orden de reglas en YAML, no en Python | Aceptada | **[NO VERIFICADO]** | No (solo índice) | — (no está en `HEAD`) |
| 0004 | COM se queda en fase 1; openpyxl se evalúa en fase 6 | Aceptada | **[NO VERIFICADO]** | No (solo índice) | — (no está en `HEAD`) |
| 0005 | Git como único historial | Aceptada (Fase 0) | 2026-10-01 | No (historial) | — (no está en `HEAD`) |
| 0006 | Estado operativo fuera del repositorio | Aceptada (Fase 0) | 2026-10-01 | No (historial) | — (no está en `HEAD`) |
| 0007 | Ratchet de tipos sobre el legacy; mypy estricto en código nuevo | Aceptada (Fase 0) | 2026-10-01 (rev.) | No (historial) | — (no está en `HEAD`) |
| 0008 | La base de datos reemplaza la descarga Selenium del ERP | Aceptada | 2026-10-02 | No (historial) | — (no está en `HEAD`) |
| 0009 | Ventas: pipeline de reglas con motor dual hasta probar paridad | Aceptada | 2026-10-02 | No (historial) | — (no está en `HEAD`) |
| 0010 | Cierre de hallazgos de la auditoría (modifica 0002) | Aceptada | 2026-10-05 | No (historial) | — (no está en `HEAD`) |

Decisiones del plan sin ADR (índice, sección final): composition root manual en `bootstrap.py` (no existe en el grafo: `src_insumos_domain_rules_registry_deuda_bootstrap_inexistente`), pandera `lazy=True` (obsoleta por 0010), eliminaciones por filtrado del DataFrame, evaluación de Prefect/Dagster.

## 2. Detalle (contexto · decisión · consecuencias · implementación en el grafo)

### ADR 0001 — Strangler fig
- **Contexto**: reescribir de una vez arriesga el informe diario. *(Reconstruido del docx Parte 2 §1 y de los docstrings.)*
- **Decisión**: conservar el contrato `CLI → Orchestrator → BaseTask` y reemplazar el interior tarea por tarea.
- **Consecuencias**: coexistencia legacy/núcleo; necesidad de ratchet (0007) y de paridad (0009).
- **Implementación**: `tasks_actualizacion_inventario_actualizacioninventario` (`implements`, INFERRED 0.75).

### ADR 0002 — pandas en el dominio
- **Decisión**: DataFrames como portador de datos; reglas vectorizadas.
- **Consecuencias**: el dominio depende de pandas (`pandas>=2.2,<3`; pandas 3 bloqueado por el legacy de ventas — docx Parte 2 §1.3). pandera retirado (0010).
- **Implementación**: `src_insumos_domain_rules_base_rulecontext` (INFERRED 0.75); `src_insumos_domain_schemas_init` documenta el retiro.

### ADR 0003 — Orden en YAML
- **Decisión**: el orden y los parámetros de las reglas viven en `config/rules/*.yaml`.
- **Implementación**: `src_insumos_domain_pipeline_rulepipeline`, `src_insumos_domain_rules_registry_rule` (INFERRED 0.95); protección contra pipeline vacío en `src_insumos_domain_reglas_declarativas_leer_declaracion`.

### ADR 0004 — COM en fase 1
- **Decisión**: Excel COM sigue escribiendo plantillas de producción; openpyxl se evalúa en Fase 6.
- **Consecuencias**: dependencia de Windows + Excel + sesión interactiva; el `--dry-run` usa openpyxl con la misma estructura.
- **Implementación**: `src_insumos_adapters_excel_com_inventario_escritorinventariocom`, `tasks_actualizacion_ventas_actualizacionventas_com_write_df_into_template` (INFERRED 0.85).

### ADR 0005 — Git como único historial
- **Contexto**: versionado por copias `*.BACKUP_*.py`, `backup.7z`.
- **Decisión**: commit baseline con todo y luego `git rm` de respaldos.
- **Consecuencias**: 7.188 líneas fuera del árbol; reversibles desde el baseline. **Tensión actual**: el mismo patrón (borrar del árbol y confiar en git) se aplicó a los ADR en `86f7093`, dejando referencias rotas.

### ADR 0006 — Estado operativo fuera del repo
- **Decisión**: `STATE_DIR` / `LOGS_DIR` (por defecto `%LOCALAPPDATA%\GeneracionInsumos`); `.gitignore` cubre estado y datos.
- **Consecuencias**: repo liviano; sesión de WhatsApp fuera de git. Pendiente: `chromedriver.exe` sigue en la raíz (`core_chromedriver_utils_deuda_chromedriver_raiz`).
- **Implementación**: `core_logger` (INFERRED 0.95), `core_chromedriver_utils` (INFERRED 0.85); variables `env_state_dir`, `env_logs_dir`.

### ADR 0007 — Ratchet de tipos
- **Decisión**: mypy estricto en `src/` y `tests/`; legacy medido con `scripts/ratchet_types.py` contra `quality-baseline.json` (hoy **12** errores, re-lineado 2026-10-05).
- **Consecuencias**: dos configuraciones mypy (`pyproject.toml`, `mypy-legacy.toml`); funciones legacy sin anotar no se revisan.
- **Implementación**: `scripts_ratchet_types`, `quality_baseline`, `mypy_legacy` (comunidad «Ratchet de tipos»).

### ADR 0008 — Fuente base de datos
- **Decisión**: retirar descargas Selenium; inventario reescrito con `inventario_bd.yaml` + `inventario.yaml`; ventas cambia solo carpeta y selección del `_268`; ventas toma líneas del inventario actualizado.
- **Consecuencias**: golden legacy de inventario inválido (`tests/golden/README.md`); `DB_EXPORT_DIR` obligatorio; se corrigió el cifrado de salida que el legacy no aplicaba.
- **Implementación**: `tasks_actualizacion_inventario_actualizacioninventario`, `src_insumos_domain_rules_inventario_base_datos_filtrarmotivoinventario`, `src_insumos_domain_rules_inventario_actualizacion_validarsalida`, `tasks_actualizacion_ventas_actualizacionventas_find_inventario_actualizado` (EXTRACTED).

### ADR 0009 — Motor dual de ventas
- **Decisión**: 21 reglas `ven.*`; `TransformarVentas` + `EntradasVentas`; `_transformar_legacy` como oráculo; `VENTAS_MOTOR`; Paso 10 sin cambios; ajustes de negocio N1–N3.
- **Consecuencias**: dos implementaciones mientras dure el motor dual; eliminaciones auditadas por regla.
- **Implementación**: hiperarista `he_motor_dual_paridad`; ver `04_motor_de_reglas_ventas.md`.

### ADR 0010 — Cierre de hallazgos de auditoría
- **Decisión**: un correo por ejecución de tarea (`BaseTask.run`); hora Colombia; solo procesos propios; retiro de código y dependencias sin uso (pandera, PyMuPDF, `HEADLESS`); nombres de hoja únicos; logs saneados con huella.
- **Consecuencias**: el envío depende de leer el command line de procesos (psutil o PowerShell/CIM). Pendientes: re-linear ratchet (hecho, base 12) y `uv lock` con red.
- **Implementación**: `tasks_base_task_basetask_run`, `tasks_envio_informe_ventas_cerrar_chrome_del_perfil`, `src_insumos_adapters_excel_reporte_inventario_nombres_de_hoja`, `tests_unit_test_hallazgos_auditoria` (EXTRACTED).

## 3. Recomendación

Restaurar `docs/adr/` en la rama (`git checkout 5e249ec -- docs/adr`) y redactar 0001–0004 con su contexto real, o registrar un ADR 0011 que explique dónde vive ahora el registro de decisiones. Es un prerequisito de gobierno para la Fase 5.
