# 08 · Calidad y pruebas

> Fuente: nodos de `tests/` (comunidades 5, 8, 9, 10, 11, 12, 14), `pyproject.toml` (`pyproject`), `.importlinter` (`importlinter_contract_1..5`), `scripts_ratchet_types`, `quality_baseline`.
> No hay reporte de cobertura en el repositorio: la cobertura de este documento es **estructural** (aristas estáticas pruebas → código). El porcentaje de líneas **[NO VERIFICADO]**.

## 1. Estrategia

| Red | Herramienta | Alcance | Evidencia |
|---|---|---|---|
| Tipos estrictos | `mypy` (`disallow_untyped_defs`, `warn_return_any` en dominio) | `src/`, `tests/` — cero errores | `pyproject.toml` `[tool.mypy]` |
| Ratchet del legacy | `scripts/ratchet_types.py` + `mypy-legacy.toml` (`platform=win32`) | `tasks/`, `core/`, `config/`, `run.py`, `orchestrator.py`; base **12** errores | `scripts/ratchet_types.py:L125` · `scripts_ratchet_types_main`; `quality-baseline.json` · `quality_baseline` |
| Arquitectura | `lint-imports` (5 contratos) + `test_regla_dependencia` (AST, rápido) | Regla de dependencia | `importlinter`; `tests_contract_test_regla_dependencia` |
| Lint | `ruff` (E, F, I, B, UP, SIM, C4, RET, ARG, PTH); legacy con ignores amplios | Todo | `pyproject.toml` `[tool.ruff]` |
| Unitarias | `pytest -m unit` | Reglas, motor, lector, comparador, seguridad, utilidades | marcador `unit` |
| Contrato | `pytest -m contract` | Escritor COM contra hoja simulada; regla de dependencia | `tests/contract/*` |
| Paridad | `test_paridad_ventas` (sintético) + `scripts/comparar_ventas.py` (real) | Motor dual de ventas | hiperarista `he_motor_dual_paridad` |
| Golden / e2e | Marcadores declarados | **Sin pruebas activas** (golden de inventario retirado por ADR 0008; ventas pendiente) | `tests/golden/README.md` · `tests_golden_readme` |

Política de pruebas (`tests/conftest.py` · `tests_conftest`; `tests/unit/rules/__init__.py`): cada prueba arranca sin estado heredado (`_entorno_limpio`), usa directorio temporal aislado, sin disco del repo, sin Excel, sin Selenium. Fixtures versionados solo sintéticos (`tests/fixtures/README.md`).

## 2. Inventario de pruebas

| Archivo | Funciones `test_*` | Marcador | Qué protege | Comunidad |
|---|---|---|---|---|
| `tests/unit/rules/test_motor.py` | 33 | — (sin marcador de módulo) | Registro, `RulePipeline`, `AuditTrail`, `RuleContext`, `RuleResult` | 5 |
| `tests/unit/test_golden.py` | 32 | unit | Comparador por clave y `ExcelReader` (cifrado en memoria) | 9 / 0 |
| `tests/unit/rules/test_reglas_inventario_bd.py` | 19 | unit | 11 reglas `inv.*` y pipelines BD/plantilla | 5 |
| `tests/unit/rules/test_reglas_declarativas.py` | 12 | unit | Contrato YAML, pipeline vacío, convención de ids, YAML reales resueltos | 8 |
| `tests/unit/rules/test_paridad_ventas.py` | 11 | unit | Paridad legacy↔reglas, ajustes N1–N3, auditoría de eliminaciones | 6 |
| `tests/unit/test_hallazgos_auditoria.py` | 8 | unit | Regresión ADR 0010 (un correo, hora Colombia, procesos propios) | 10 |
| `tests/unit/adapters/test_com_inventario_puro.py` | 6 | unit | Funciones puras del escritor COM | 0 |
| `tests/unit/test_pipeline_legacy.py` | 5 | unit | `TASK_REGISTRY`, workflows, `Orchestrator` | 2 |
| `tests/unit/test_seguridad.py` | 4 | unit | Huella y `FiltroSecretos` | 2 |
| `tests/contract/test_regla_dependencia.py` | 3 (parametrizadas por módulo) | — | Imports prohibidos en dominio/aplicación | 11 |
| `tests/unit/test_dias_venta.py` | 3 | unit | Festivos y días de venta (hora Colombia) | 3 |
| `tests/unit/test_clave_actualizacion.py` | 3 | unit | Clave del área y transición | 3 |
| `tests/unit/adapters/test_reporte_inventario.py` | 3 | unit | Nombres de hoja únicos (hallazgo 10) | 7 |
| `tests/contract/test_escritor_com_simulado.py` | 1 | contract | Proyección COM sobre hoja simulada | 12 |
| **Total** | **143 funciones** (220 casos ejecutados con parametrización, según docx/ADR 0010) | | | |

## 3. Cobertura estructural por módulo de producción

"Callables tocados" = funciones/clases/métodos del módulo con al menos una arista directa desde `tests/`. Las reglas ejecutadas vía YAML se ejercitan **indirectamente** (p. ej. las 20 `ven.*` corren dentro de `test_paridad_ventas`), por eso su conteo directo es bajo.

| Módulo | Callables | Tocados directo | Pruebas que lo alcanzan | Lectura |
|---|---|---|---|---|
| `src/insumos/domain/pipeline.py` | 11 | 5 | test_motor, test_reglas_inventario_bd | Bien cubierto |
| `src/insumos/domain/reglas_declarativas.py` | 7 | 4 | test_reglas_declarativas | Bien cubierto |
| `src/insumos/domain/rules/registry.py` | 11 | 4 | test_motor, test_reglas_declarativas | Bien cubierto |
| `src/insumos/domain/rules/base.py` | 8 | 4 | test_motor, test_paridad_ventas | Bien cubierto |
| `src/insumos/domain/audit.py` | 8 | 2 | test_motor | Cubierto vía pipeline |
| `src/insumos/domain/rules/inventario/*` | 31 | 7 | test_reglas_inventario_bd, test_motor | Cubierto vía YAML |
| `src/insumos/domain/rules/ventas/*` | 53 | 4 | test_paridad_ventas | Indirecto (paridad); `salida.py` sin prueba directa |
| `src/insumos/application/transformar_ventas.py` | 4 | 1 | test_paridad_ventas | Cubierto |
| `src/insumos/application/actualizar_inventario.py` | 4 | **0** | — | **Brecha**: el caso de uso completo no tiene prueba con puertos falsos |
| `src/insumos/adapters/excel/lector.py` | 15 | 3 | test_golden | Cubierto |
| `src/insumos/adapters/excel/comparador.py` | 13 | 2 | test_golden, test_paridad_ventas | Cubierto |
| `src/insumos/adapters/excel/com_inventario.py` | 15 | 5 | test_com_inventario_puro, test_escritor_com_simulado | Partes puras + simulación; COM real solo en Windows |
| `src/insumos/adapters/excel/reporte_inventario.py` | 8 | 1 | test_reporte_inventario | Parcial |
| `src/insumos/adapters/excel/fuentes_ventas.py` | 9 | 1 | test_paridad_ventas | Parcial |
| `src/insumos/adapters/excel/fuentes_inventario.py` | 13 | **0** | — | **Brecha** |
| `src/insumos/adapters/excel/maestros.py` | 4 | **0** | — | **Brecha** (Matriz USD, Distribución) |
| `src/insumos/adapters/excel/escritor_openpyxl.py` | 5 | **0** | — | **Brecha** (dry-run) |
| `src/insumos/adapters/system/reloj.py`, `archivos.py` | 6 | **0** | — | Brecha menor |
| `core/xlsx_cleaner.py` | 7 | **0** | — | **Brecha crítica**: es la funcionalidad de reducción de tamaño (objetivo del proyecto) |
| `core/seguridad.py` | 5 | 3 | test_seguridad | Cubierto |
| `core/email_notifier.py`, `core/logger.py`, `core/excel_utils.py`, `core/chromedriver_utils.py` | 34 | **0** | — | Brecha (legacy) |
| `tasks/actualizacion_ventas.py` | 72 | 4 | test_clave_actualizacion, test_dias_venta, test_hallazgos_auditoria | Paso 10 y post-proceso sin prueba |
| `tasks/envio_informe_ventas.py` | 41 | 3 | test_hallazgos_auditoria | Legacy; Fase 5 |
| `tasks/actualizacion_inventario.py` | 7 | **0** | — | Brecha (raíz de composición) |
| `orchestrator.py` / `tasks/base_task.py` | 17 | 2 | test_pipeline_legacy | Humo |
| `scripts/*` | 10 | **0** | — | Herramientas sin prueba |

## 4. Pruebas de paridad

| Nivel | Qué compara | Criterio | Limitación |
|---|---|---|---|
| Sintético | `_transformar_legacy` vs `TransformarVentas` sobre entradas construidas con todas las ramas (`test_paridad_con_proceso_anterior`, `test_paridad_con_filas_de_titulo`, `test_sin_hoja_licitados_sigue_sin_la_columna`) | Igualdad de DataFrames | Solo ramas imaginadas por quien escribió el fixture |
| Real | `scripts/comparar_ventas.py` sobre los archivos del día | `comparar_por_posicion` vacío → exit 0 | Lógica compartida entre motores no se valida (ver `09`, R1) |
| Golden ventas | Inputs reales congelados + salida legacy en parquet | Comparación por clave de negocio (`comparar`) | **No implementado** (`tests/golden/README.md`) |

## 5. Recomendaciones (orden de valor)

1. Prueba unitaria de `core/xlsx_cleaner.reducir_tamano_xlsx` con un `.xlsx` sintético con dibujos huérfanos (cifrado y sin cifrar) — objetivo declarado del proyecto.
2. Prueba del caso de uso `ActualizarInventario` con puertos falsos (fuente, escritor y reporte en memoria) que verifique que **no** se escribe si `inv.validar_salida` falla.
3. Prueba de `_ejecutar` con `HAS_COM=False` que exija fallo explícito (hoy termina en éxito).
4. Golden de ventas anonimizado y marcador `golden` activo antes de retirar `_transformar_legacy`.
5. Publicar cobertura (`pytest --cov=src/insumos --cov-report=xml`) y fijar umbral por capa (dominio ≥ 90 %).
