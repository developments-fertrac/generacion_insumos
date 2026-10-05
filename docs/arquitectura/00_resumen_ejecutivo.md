# 00 · Resumen ejecutivo — Generación de Insumos

> Fuente única de verdad: `graphify-out/graph.json` y `graphify-out/GRAPH_REPORT.md` (rama `main`, commit `088f6f0`, 1.389 nodos, 3.263 aristas, 69 comunidades; graphify 0.9.77 oficial).
> Convención de citas: `archivo:Lnn` · nodo `id_del_grafo`. Lo que el grafo no confirma se marca **[NO VERIFICADO]**.

## 1. Propósito de negocio

Generación de Insumos es el pipeline del Área de Datos de Fertrac que convierte las **exportaciones diarias de la base de datos corporativa** en los dos libros consolidados con los que opera el negocio: el **Inventario General** y las **Ventas 2026**. Además distribuye el **informe diario de ventas por WhatsApp** a la gerencia comercial.

- Contrato de ejecución: `run.py` → `orchestrator.py` → `BaseTask.run()` (nodos `run`, `orchestrator_orchestrator`, `tasks_base_task_basetask_run`; `tasks/base_task.py:L21`).
- Fuente de datos: carpeta `DB_EXPORT_DIR` con `Inventario.xlsx` e `InformesDeVentas(Facturas)_268*.xlsx`, que reemplazó las descargas Selenium del ERP (nodo ADR 0008 (archivo eliminado en `86f7093`); `config/settings.py:L83` · `config_settings_pathsconfig`).
- Restricción de producto: la salida debe **preservar hojas, formatos, títulos y fórmulas** del proceso anterior (nodo `tasks_actualizacion_inventario_preservacion_salida`), por eso la escritura final se hace con **Excel COM** (nodos `src_insumos_adapters_excel_com_inventario_escritorinventariocom`, `tasks_actualizacion_ventas_actualizacionventas_com_write_df_into_template`).

## 2. Usuarios y áreas consumidoras

| Área | Qué consume | Evidencia |
|---|---|---|
| Compras / Inventarios | `$2026 INVENTARIO GENERAL ACTUALIZADO <fecha>.xlsx` + `REPORTE_ELIMINACIONES_*.xlsx` | `docs/Documentacion_Generacion_de_Insumos.docx` (Parte 1 §1, §2.5) · nodo `docs_documentacion_generacion_de_insumos`; destinatarios `SMTP_RECIPIENT_EMAILS_INV_GENERAL` (`env_smtp_recipient_emails_inv_general`) |
| Control de Ventas (dueña de `$2026 VENTAS_Actualizacion.xlsx`, con contraseña propia) | Libro de ventas protegido | `tasks/actualizacion_ventas.py:L187` · `tasks_actualizacion_ventas_actualizacionventas_password_actualizacion` |
| Áreas comerciales / usuarios finales | Copia liviana `$2026 VENTAS_<fecha>.xlsb` | `tasks/actualizacion_ventas.py:L2095` · `tasks_actualizacion_ventas_actualizacionventas_guardar_copia_xlsb` |
| Gerencia comercial | 4 imágenes + mensaje "VTAS <MES> MG NETO PONDERADO" por WhatsApp | `tasks/envio_informe_ventas.py:L1626` · `tasks_envio_informe_ventas_envioinformeventas_execute` |
| Área de Datos (operación) | Correos `[OK]`/`[ERROR]` con log adjunto, logs por tarea | `tasks/base_task.py:L21` · nodo `tasks_base_task_notificacion_por_tarea` |

## 3. Valor operativo

| Palanca | Mecanismo | Evidencia |
|---|---|---|
| Elimina consolidación manual | 31 reglas puras declaradas en 3 YAML (11 `inv.*`, 20 `ven.*` con 21 entradas) | `config/rules/*.yaml` · nodos `config_rules_inventario`, `config_rules_inventario_bd`, `config_rules_ventas` |
| Integridad de datos | El inventario **no se escribe** si no cuadra contra la BD (referencias y TOTAL INV) | `src/insumos/domain/rules/inventario/actualizacion.py:L358` · `src_insumos_domain_rules_inventario_actualizacion_validarsalida` |
| Trazabilidad | Cada eliminación queda con motivo y regla (`AuditTrail`) y en el reporte | `src/insumos/domain/audit.py:L39` · `src_insumos_domain_audit_audittrail` |
| Libros más livianos | PivotCache compartido, limpieza de dibujos, copia `.xlsb` | `core/xlsx_cleaner.py:L160` · `core_xlsx_cleaner_reducir_tamano_xlsx` |
| Seguridad | Contraseñas reemplazadas por huella `sha256:` en logs | `core/seguridad.py:L50` · `core_seguridad_filtrosecretos` |
| Riesgo de migración controlado | Motor dual en ventas con oráculo legacy | `tasks/actualizacion_ventas.py:L2171` · `tasks_actualizacion_ventas_actualizacionventas_motor` |

## 4. Estado de la migración hexagonal

| Componente | Estado | Evidencia en el grafo | Próximo hito |
|---|---|---|---|
| `actualizacion_inv` | **Migrada** (raíz de composición + caso de uso + adaptadores). Corrida real 2026-10-02 | Hiperarista `he_composicion_inventario`; `tasks/actualizacion_inventario.py:L63` · `tasks_actualizacion_inventario_actualizacioninventario_construir` | Certificar en producción (operativo) |
| `actualizacion_ventas` | **Fase 4 — motor dual**; defecto `VENTAS_MOTOR=legacy` | Hiperarista `he_motor_dual_paridad`; `env_ventas_motor` | Paridad real → `reglas` → Fase 4b (`readme_pendiente_fase_4b`) |
| `envio_informe_ventas` | **Legacy** (Selenium + COM + portapapeles) | comunidades «WhatsApp Web (Selenium)» y «Envío: procesos y puertos»; `readme_pendiente_fase_5` | Fase 5 |
| Descargas Selenium del ERP | **Retiradas** (ADR 0008) | `scripts/publicar_a_produccion.ps1` borra los módulos viejos en destino | — |
| Regla de dependencia | **0 violaciones** en `src/insumos` | aristas `imports` de `graph.json` (0 desde `domain`/`application` hacia capas externas) | Mantener `lint-imports` en CI **[NO VERIFICADO: no hay pipeline de CI en el repo]** |

## 5. Riesgos que la dirección debe conocer

1. **Oráculo de paridad no independiente**: el legacy de ventas usa funciones del dominio nuevo (`texto_licitado`, `normalizar_referencia`, `descuento_a_decimal`), por lo que la paridad no detecta defectos en esa lógica compartida (`tasks/actualizacion_ventas.py:L2474`; GRAPH_REPORT §3.1).
2. **Éxito silencioso sin Excel**: si `HAS_COM` es falso, el Paso 10 y el post-proceso se omiten y la tarea termina en `[OK]` (`tasks/actualizacion_ventas.py:L2603` · nodo `tasks_actualizacion_ventas_riesgo_paso10_omitido_sin_com`).
3. **Lockfile eliminado** (`088f6f0`) mientras los `.bat` de producción ejecutan `uv run --frozen` (nodo `readme_pendiente_uv_lock`).
4. **Documentación de decisiones eliminada del árbol** (`86f7093`: `docs/adr/`, `docs/Old/`) pero referenciada por README, docstrings y el script de publicación (nodo `readme_pendiente_docs_adr_eliminados`).
5. **Contraseñas expuestas en el historial git** pendientes de rotación (nodo `docs_documentacion_generacion_de_insumos_pendiente_rotar_contrasenas`).

Detalle y plan de acción: `09_riesgos_y_deuda_tecnica.md` y `10_roadmap.md`.
