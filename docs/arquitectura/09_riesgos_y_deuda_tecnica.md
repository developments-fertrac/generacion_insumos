# 09 · Riesgos y deuda técnica

> Fuente: nodos `concept` de pendientes/riesgos (GRAPH_REPORT §7), god nodes (§2), conexiones sorprendentes (§3.1) y verificación de la regla de dependencia (§6).
> Escala: Probabilidad (P) e Impacto (I) de 1 a 5; Exposición = P × I.

## 1. Matriz de riesgos

| ID | Riesgo | Evidencia (archivo · nodo) | P | I | Exp. | Acción | Criterio de cierre |
|---|---|---|---|---|---|---|---|
| R1 | **Oráculo de paridad no independiente**: el legacy usa `texto_licitado`, `normalizar_referencia`, `descuento_a_decimal` del dominio; un defecto común pasa la paridad | `tasks/actualizacion_ventas.py:L2474`, `L27` · `tasks_actualizacion_ventas_actualizacionventas_transformar_legacy` → `src_insumos_domain_rules_ventas_columnas_texto_licitado` | 3 | 4 | **12** | Casos de negocio firmados por Control de Ventas como pruebas tabulares; golden real anonimizado | Tabla de casos aprobada + golden verde |
| R2 | **Éxito silencioso sin COM**: si `HAS_COM=False`, Paso 10 y post-proceso se omiten y la tarea envía `[OK]` | `tasks/actualizacion_ventas.py` (`_ejecutar`, `if HAS_COM:`) · `tasks_actualizacion_ventas_riesgo_paso10_omitido_sin_com`; `core/excel_processing.py` · `core_excel_processing` | 2 | 5 | **10** | `raise RuntimeError` si no hay COM o si `tmp_out_path` no existe tras el Paso 10 | Prueba que exige fallo con `HAS_COM=False` |
| R3 | **`uv.lock` eliminado** con `.bat` en `uv run --frozen` y publicación que copia `uv.lock` | commit `088f6f0`; `scripts/publicar_a_produccion.ps1` · `scripts_publicar_a_produccion_ps1`; `readme_pendiente_uv_lock` | 4 | 4 | **16** | `uv lock` con red, versionar, publicar | `uv run --frozen python run.py --list` OK en producción |
| R4 | **Dependencia de Excel COM y sesión interactiva** (Paso 10 ventas, escritor de inventario, capturas del envío) | `src_insumos_adapters_excel_com_inventario_escritorinventariocom`; `tasks_actualizacion_ventas_actualizacionventas_com_write_df_into_template`; `tasks_envio_informe_ventas_capturar_multiples_rangos`; `readme_sesion_interactiva` | 3 | 4 | **12** | Corto plazo: monitoreo de sesión y reintentos COM existentes. Mediano: Fase 6 openpyxl para escritura (ADR 0004) donde no haya tablas dinámicas | Escritura de inventario sin COM certificada con `comparar_estructura.py` |
| R5 | **Contraseñas en historial git** (logs, documentos y pruebas antiguas) | `docs_documentacion_generacion_de_insumos_pendiente_rotar_contrasenas`; ADR 0010 §7 | 3 | 5 | **15** | Rotar `EXCEL_PASSWORD`, `VENTAS_ACTUALIZACION_PASSWORD`, `SMTP_PASSWORD`; evaluar `git filter-repo` | Contraseñas rotadas y verificadas |
| R6 | **Scripts de validación con contraseñas en código** (`Archivos Validacion/`, fuera de git) | `docs_documentacion_generacion_de_insumos_pendiente_scripts_validacion_contrasenas` | 3 | 4 | **12** | Migrar a `.env` o retirar si la tarea programada ya no existe | Búsqueda sin coincidencias |
| R7 | **Correo de fallo con traceback sin filtro de secretos** | `tasks/actualizacion_ventas.py` (`_notify_failure` usa `_traceback_error`); `core_seguridad_filtrosecretos` actúa solo en handlers de log | 2 | 4 | 8 | Pasar el cuerpo por `core.seguridad.ocultar()` en `EmailNotifier.send` | Prueba en `test_seguridad` |
| R8 | **Registro de decisiones eliminado** (`docs/adr/`, `docs/Old/`) pero referenciado | commit `86f7093`; `readme_pendiente_docs_adr_eliminados` | 5 | 2 | 10 | Restaurar desde `5e249ec` o ADR 0011 | README sin enlaces rotos |
| R9 | **`VENTAS_ACTUALIZACION_PASSWORD` ausente en producción** → el libro del área se guarda con la clave general | `tasks/actualizacion_ventas.py:L187` · `readme_pendiente_password_env_produccion` | 3 | 3 | 9 | Configurar variable | Log sin advertencia "no configurada" |
| R10 | **Lógica de licitados duplicada** (adaptador vs legacy) durante el motor dual | `src_insumos_adapters_excel_fuentes_ventas_leer_precios_licitados`; `_transformar_legacy` L2432–L2518 | 3 | 3 | 9 | Retirar legacy tras corte (doc 04 §4) | `_transformar_legacy` eliminado |
| R11 | **`BASE_PATH` con ruta de red por defecto** en el código | `config/settings.py` · `config_settings_deuda_base_path_por_defecto` | 2 | 3 | 6 | Hacer `BASE_PATH` obligatorio | Error claro si falta |
| R12 | **Carpeta `_obsoleto_2026-10-02/`** pendiente de borrar | `readme_pendiente_carpeta_obsoleto` | 2 | 2 | 4 | Borrar tras confirmar | Carpeta inexistente |
| R13 | **`.bat` y n8n fuera de control de versiones** | `docs_documentacion_generacion_de_insumos_pendiente_bat_sin_uv` | 3 | 3 | 9 | Versionar `.bat` en `ops/` (sin secretos) | `.bat` en git |
| R14 | **`test_regla_dependencia` sin marcador** → `pytest -m contract` no lo ejecuta; además no detecta `insumos.adapters.*` desde el dominio | `tests/contract/test_regla_dependencia.py` · `tests_contract_test_regla_dependencia` | 2 | 3 | 6 | `pytestmark = pytest.mark.contract` y `startswith` | Prueba roja ante import prohibido sintético |

## 2. Deuda técnica por acoplamiento (god nodes y aristas del grafo)

| Elemento | Métrica | Problema | Acción |
|---|---|---|---|
| `ActualizacionVentas` (`tasks_actualizacion_ventas_actualizacionventas`) | Grado 79; 72 callables; ~2.660 líneas; mayor betweenness de archivo (`tasks_actualizacion_ventas`, 0,20) | Lectura, motor legacy, Paso 10, post-proceso y notificación en una clase | Fase 4b: `FuenteVentasArchivos`; Paso 10 a `adapters/excel/com_ventas.py` (puerto `EscritorVentas`) |
| `RuleContext` (`src_insumos_domain_rules_base_rulecontext`) | Grado 112 (máximo); betweenness 0,12 | Contexto "bolsa" con campos del modelo ERP retirado | Retirar `marcas_propias`, `remisiones`, `valorizados` (`src_insumos_domain_rules_base_deuda_campos_erp_retirados`); evaluar contextos tipados por dominio |
| `RuleResult` / `ReglaBase` / `rule()` | Grados 88 / 52 / 44 | Contrato central (esperable) | Mantener estable; cambios solo con ADR |
| `ExcelReader` (`src_insumos_adapters_excel_lector_excelreader`) | Grado 40; puerto `LectorExcel` en adaptadores | Puerto mal ubicado | Mover el Protocol a `domain/ports.py` |
| `_log()` del envío (`tasks_envio_informe_ventas_log`) | Grado 29 | Logger ad hoc paralelo a `core.logger` | Fase 5 |
| `Settings` (`config_settings_settings`) | Grado 28; `core` → `config` (8 aristas) | Infraestructura acoplada a configuración global | Inyectar configuración en adaptadores de `notify` (Fase 5) |
| `escritor_openpyxl` → `com_inventario._cifrar` | Arista lateral a privado | Dry-run arrastra el módulo COM | Extraer `cifrado.py` |
| `scripts/comparar_ventas.py` → `ActualizacionVentas` (privado) | Arista a privado | El script se rompe al retirar el legacy | Re-apuntar a `FuenteVentas` en Fase 4b |
| `_cargar_modulos_de_reglas` | 0 llamadores | Código muerto / composición implícita por import | Eliminar o invocar desde una raíz de composición explícita |

## 3. Código legacy pendiente

| Módulo | Estado | Errores de tipo en ratchet | Fase |
|---|---|---|---|
| `tasks/actualizacion_ventas.py` | Motor dual; Paso 10 legacy | Sin errores en `quality-baseline.json` | 4 / 4b |
| `tasks/envio_informe_ventas.py` | Legacy completo | 4 (`assignment` ×2, `import-untyped`, `union-attr`) | 5 |
| `core/xlsx_cleaner.py` | Legacy compartido | 4 | 4b/5 |
| `core/logger.py`, `run.py`, `tasks/actualizacion_inventario.py` | Legacy de orquestación | 1 + 2 + 1 | transversal |

Fuente: `quality-baseline.json` (total 12).
