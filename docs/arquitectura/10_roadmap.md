# 10 · Roadmap

> Fuente: nodos `readme_pendiente_fase_4b`, `readme_pendiente_fase_5`, `readme_pendiente_paridad_real`, pendientes operativos (nodos `*_pendiente_*` del grafo), ADR 0004/0007/0009 y brechas de `08_calidad_y_pruebas.md`.
> Las fechas objetivo no están definidas en el repositorio **[NO VERIFICADO]**; se proponen horizontes relativos.

## 0. Higiene operativa (inmediato, prerequisito de todo lo demás)

| # | Acción | Evidencia | Criterio de salida medible |
|---|---|---|---|
| H1 | Regenerar y versionar `uv.lock` (`uv lock` con red) | `readme_pendiente_uv_lock` (commit `088f6f0`) | `uv run --frozen python run.py --list` termina con exit 0 en producción |
| H2 | Configurar `VENTAS_ACTUALIZACION_PASSWORD` en el `.env` de producción | `readme_pendiente_password_env_produccion` | 5 corridas consecutivas sin la advertencia "VENTAS_ACTUALIZACION_PASSWORD no configurada" |
| H3 | Rotar contraseñas expuestas en el historial y retirar las de `Archivos Validacion/` | `docs_documentacion_generacion_de_insumos_pendiente_rotar_contrasenas`, `docs_documentacion_generacion_de_insumos_pendiente_scripts_validacion_contrasenas` | Todas las claves rotadas; búsqueda de las antiguas = 0 coincidencias |
| H4 | Restaurar `docs/adr/` (o ADR 0011) y corregir enlaces de README/docstrings/publicación | `readme_pendiente_docs_adr_eliminados` | 0 referencias a rutas inexistentes (`grep docs/adr` resuelve) |
| H5 | Borrar `_obsoleto_2026-10-02/` | `readme_pendiente_carpeta_obsoleto` | Carpeta inexistente en desarrollo y producción |
| H6 | Versionar los `.bat` de n8n (sin secretos) con `PATH_UV` explícito | `docs_documentacion_generacion_de_insumos_pendiente_bat_sin_uv` | `.bat` en el repositorio; corrida n8n con `resultado_<tarea>.txt = 0` |
| H7 | Fallo explícito sin COM en ventas | `tasks_actualizacion_ventas_riesgo_paso10_omitido_sin_com` | Prueba unitaria roja→verde; correo `[ERROR]` en ejecución sin Excel |

## 1. Fase 4 (en curso) — corte del motor de ventas

| # | Actividad | Criterio de salida |
|---|---|---|
| 4.1 | Ejecutar `scripts/comparar_ventas.py` con datos reales | **exit 0** (hoja `DIFERENCIAS` vacía) |
| 4.2 | Activar `VENTAS_MOTOR=reglas` | Variable en `.env` de producción; log "Motor de transformacion: reglas" |
| 4.3 | Observación | **10 días hábiles / 2 semanas** de corridas con comparación diaria en 0 diferencias y sin incidentes de Control de Ventas |
| 4.4 | Casos de negocio firmados para lógica compartida (R1) | Tabla de casos aprobada y convertida en prueba parametrizada |

## 2. Fase 4b — ventas hexagonal completa

| # | Actividad | Evidencia del grafo | Criterio de salida |
|---|---|---|---|
| 4b.1 | Adaptador `FuenteVentasArchivos` que implemente `FuenteVentas` (localización, descifrado, `EntradasVentas`) | `src_insumos_domain_ports_fuenteventas` sin `implements`; `tasks_actualizacion_ventas_actualizacionventas_preparar_entradas` | Arista `implements` en el grafo; `preparar_entradas`/`localizar_archivos` eliminados del legacy |
| 4b.2 | Retirar `_transformar_legacy` y la duplicación de licitados | `tasks_actualizacion_ventas_actualizacionventas_transformar_legacy` | Nodo inexistente; `ActualizacionVentas` < 1.000 líneas |
| 4b.3 | Puerto `EscritorVentas` + adaptador COM para Paso 10 y post-proceso | hiperarista `he_paso10_postproceso` | `tasks/actualizacion_ventas.py` reducido a raíz de composición (como inventario) |
| 4b.4 | Golden de ventas anonimizado | `tests_golden_readme` | Marcador `golden` con ≥ 1 escenario verde |
| 4b.5 | Re-apuntar `scripts/comparar_ventas.py` (o retirarlo) | arista a privado (`scripts_comparar_ventas_uso_privado`) | Sin imports de `tasks/` en `scripts/` |
| 4b.6 | Ratchet | `quality_baseline` | Base re-lineada; `scripts/ratchet_types.py` OK |

## 3. Fase 5 — envío del informe

| # | Actividad | Evidencia | Criterio de salida |
|---|---|---|---|
| 5.1 | Puertos `CapturadorInforme` y `Mensajeria`; adaptadores COM/portapapeles y WhatsApp Web en `adapters/notify` | `src_insumos_adapters_notify_init` (vacío); comunidades «WhatsApp Web (Selenium)» y «Envío: procesos y puertos» | `tasks/envio_informe_ventas.py` como raíz de composición; caso de uso `EnviarInforme` en `application/` |
| 5.2 | Notificación SMTP como adaptador | `core_email_notifier_emailnotifier` | `core/email_notifier.py` retirado o envuelto; filtro de secretos aplicado al cuerpo (R7) |
| 5.3 | Mover `chromedriver.exe` al estado operativo | `core_chromedriver_utils_deuda_chromedriver_raiz` | Binario fuera de la raíz del proyecto |
| 5.4 | Retirar `mypy-legacy.toml` y el ratchet | ADR 0007 | `quality-baseline.json` en 0 o eliminado |
| 5.5 | Pruebas de contrato de los nuevos adaptadores | — | `pytest -m contract` cubre puertos de envío |

## 4. Fase 6 (evaluación, ADR 0004) — reducir dependencia de COM

| Actividad | Criterio de salida |
|---|---|
| Evaluar escritura openpyxl para inventario en producción (ya existe para `--dry-run`) | `scripts/comparar_estructura.py` sin diferencias de hojas/formatos/fórmulas vs. salida COM en 5 corridas |
| Evaluar orquestador (Prefect/Dagster, índice ADR) vs. n8n + schtasks | ADR con decisión y costo operativo |

## 5. Mejoras transversales de calidad

| Actividad | Criterio de salida |
|---|---|
| Pruebas de `core/xlsx_cleaner.py`, `ActualizarInventario`, `fuentes_inventario`, `maestros`, `escritor_openpyxl` | Cada módulo con ≥ 1 arista directa desde `tests/` en el grafo regenerado |
| Cobertura publicada | Reporte `coverage.xml`; dominio ≥ 90 %, aplicación ≥ 80 % |
| Contrato import-linter 6 (`insumos.adapters` no importa `tasks/core/config/orchestrator`) | `lint-imports` con 6 contratos en verde |
| Retirar campos ERP de `RuleContext` | Grado de `RuleContext` reducido y sin campos `valorizados/remisiones/marcas_propias` |
| CI (GitHub Actions o equivalente) con ruff, mypy, pytest, lint-imports, ratchet | Pipeline obligatorio en PR **[NO VERIFICADO: no existe hoy]** |
