# 01 · Contexto C4 (niveles 1 y 2)

> Fuente: `graphify-out/graph.json` (comunidades 0–4, hiperaristas `he_composicion_inventario`, `he_paso10_postproceso`, `he_motor_dual_paridad`) y nodos semánticos `docs_documentacion_generacion_de_insumos_orquestacion_n8n`, `readme_sesion_interactiva`.
> Los `.bat` de producción y la configuración de n8n **no están en el repositorio**: su existencia se toma del docx (Parte 1 §3.3) y se marca **[NO VERIFICADO]** en el código.

## 1. Nivel 1 — Contexto del sistema

```mermaid
flowchart LR
    subgraph Fertrac
      BD[(Base de datos corporativa<br/>exporta a DB_EXPORT_DIR)]
      MAE[[Maestros en BASE_PATH<br/>Matriz USD · Distribución · MYR · Matriz de clientes]]
      N8N([n8n<br/>orquestador de automatizaciones])
      OPS((Área de Datos<br/>operación y soporte))
    end
    SYS[["Generación de Insumos<br/>(pipeline Python + Excel COM)"]]
    COMPRAS((Compras / Inventarios))
    CTRL((Control de Ventas))
    COM((Áreas comerciales))
    GER((Gerencia comercial))
    SMTP[(SMTP Gmail)]
    WA[(WhatsApp Web)]

    BD -- "Inventario.xlsx<br/>InformesDeVentas(Facturas)_268*.xlsx" --> SYS
    MAE -- "libros cifrados" --> SYS
    N8N -- "schtasks /Run → .bat → uv run" --> SYS
    SYS -- "INVENTARIO GENERAL ACTUALIZADO<br/>+ REPORTE_ELIMINACIONES" --> COMPRAS
    SYS -- "$2026 VENTAS_Actualizacion.xlsx (clave del área)" --> CTRL
    SYS -- "$2026 VENTAS_«fecha».xlsb" --> COM
    SYS -- "imágenes del informe" --> WA --> GER
    SYS -- "[OK]/[ERROR] + log" --> SMTP --> OPS
```

| Actor / sistema | Relación | Evidencia |
|---|---|---|
| Base de datos (exportación) | Entrada única de inventario y ventas | `docs_adr_0008_fuente_base_de_datos`; `config/settings.py:L83` · `config_settings_pathsconfig`; `env_db_export_dir` |
| Maestros en `BASE_PATH` | Enriquecimiento (Matriz USD, Distribución, MYR, Matriz de clientes) | `src/insumos/adapters/excel/maestros.py:L35` · `src_insumos_adapters_excel_maestros_leer_matriz_usd`; `src/insumos/adapters/excel/fuentes_ventas.py:L86` · `src_insumos_adapters_excel_fuentes_ventas_leer_myr` |
| n8n + Programador de tareas | Disparo diario; lee `resultado_<tarea>.txt` | `docs_documentacion_generacion_de_insumos_orquestacion_n8n` **[NO VERIFICADO en código: .bat fuera del repo]** |
| SMTP | Un correo por corrida | `core/email_notifier.py:L16` · `core_email_notifier_emailnotifier` |
| WhatsApp Web | Envío del informe | `tasks/envio_informe_ventas.py:L360` · `tasks_envio_informe_ventas_whatsappweb` |

## 2. Nivel 2 — Contenedores

```mermaid
flowchart TB
    subgraph Host["Servidor Windows con sesión interactiva"]
      BAT["*.bat de tarea<br/>uv run --frozen python run.py --workflow X<br/>[NO VERIFICADO: fuera del repo]"]
      CLI["CLI run.py<br/>(argparse, códigos 0/1)"]
      ORQ["orchestrator.py<br/>PIPELINES y fases"]
      subgraph Tareas["tasks/ (BaseTask)"]
        TINV["ActualizacionInventario<br/>raíz de composición (migrada)"]
        TVEN["ActualizacionVentas<br/>motor dual + Paso 10 COM (legacy)"]
        TENV["EnvioInformeVentas<br/>legacy Selenium + COM"]
      end
      subgraph Nucleo["src/insumos (hexagonal)"]
        APP["application<br/>ActualizarInventario · TransformarVentas"]
        DOM["domain<br/>RulePipeline · reglas inv/ven · puertos · AuditTrail"]
        ADP["adapters<br/>excel (lector, fuentes, COM, openpyxl, reporte) · system"]
      end
      CORE["core/<br/>logger · seguridad · email · excel_utils · xlsx_cleaner · chromedriver"]
      CFG["config/<br/>settings.py (.env) · rules/*.yaml"]
      EXCEL[["Microsoft Excel (COM)"]]
      CHROME[["Google Chrome + ChromeDriver"]]
      STATE[("%LOCALAPPDATA%/GeneracionInsumos<br/>logs · perfil WhatsApp · caches")]
    end
    N8N([n8n]) --> BAT --> CLI --> ORQ --> Tareas
    TINV --> APP
    TVEN --> APP
    TVEN --> ADP
    APP --> DOM
    ADP --> DOM
    TINV --> ADP
    ADP --> EXCEL
    TVEN --> EXCEL
    TENV --> EXCEL
    TENV --> CHROME
    Tareas --> CORE
    Tareas --> CFG
    APP -. "lee YAML" .-> CFG
    CORE --> STATE
    DBX[(DB_EXPORT_DIR)] --> ADP
    DBX --> TVEN
    MAE[(BASE_PATH maestros)] --> ADP
    MAE --> TVEN
    Tareas -- "salidas" --> OUT[("BASE_PATH: Pruebas · Pruebas Inv General · Backups_Ventas")]
```

| Contenedor | Responsabilidad | Comunidad del grafo | Evidencia |
|---|---|---|---|
| CLI `run.py` | Parseo de argumentos, `--dry-run` vía `INSUMOS_DRY_RUN`, código de salida | 2 · Núcleo legacy y configuración | `run.py:L1` · `run`; `env_insumos_dry_run` |
| Orquestador | Workflows `completo`, `ventas`, `inventario`, `parcialVentas`, `envio`; un fallo no detiene las fases siguientes | 2 | `orchestrator.py:L66` · `orchestrator_orchestrator` |
| Tarea inventario | Raíz de composición: fuente → caso de uso → escritor COM/openpyxl → reporte | 0 · Adaptadores Excel de inventario | `tasks/actualizacion_inventario.py:L63` · `tasks_actualizacion_inventario_actualizacioninventario_construir` |
| Tarea ventas | Localiza insumos, respaldo, selector de motor, Paso 10 COM, post-proceso | 3 · Ventas legacy y Paso 10 | `tasks/actualizacion_ventas.py:L2553` · `tasks_actualizacion_ventas_actualizacionventas_ejecutar` |
| Tarea envío | Capturas Excel → portapapeles → WhatsApp Web | 4 · Envío WhatsApp y ChromeDriver | `tasks/envio_informe_ventas.py:L1626` · `tasks_envio_informe_ventas_envioinformeventas_execute` |
| Núcleo `src/insumos` | Dominio puro + casos de uso + adaptadores | 1, 5, 6, 7, 8 | `src_insumos_domain_ports`, `src_insumos_domain_pipeline_rulepipeline` |
| `core/` | Infraestructura compartida del legacy | 2, 3, 4 | `core_logger_get_logger`, `core_seguridad_filtrosecretos`, `core_xlsx_cleaner_reducir_tamano_xlsx` |
| `config/` | `.env` → dataclasses inmutables; YAML de reglas | 2, 6, 7 | `config/settings.py:L137` · `config_settings_settings` |
| Excel COM | Escritura con fidelidad de formato, tablas dinámicas, cifrado, capturas | externo | imports `win32com` en 4 archivos (GRAPH_REPORT §9) |
| Chrome/ChromeDriver | Sesión WhatsApp persistente; driver compatible | externo | `core/chromedriver_utils.py:L259` · `core_chromedriver_utils_obtener_chromedriver_path` |
| Estado operativo | Logs y perfiles fuera del repo | — | `core/logger.py:L14` · `core_logger_logs_root`; `docs_adr_0006_estado_operativo_fuera` |

## 3. Restricciones de despliegue que condicionan la arquitectura

- **Sesión interactiva obligatoria**: Excel COM y el portapapeles no funcionan en Session 0 ni con escritorio remoto desconectado (`readme_sesion_interactiva`; `tasks_envio_informe_ventas_diagnosticar_portapapeles`).
- **Windows + Excel de escritorio + Chrome**: dependencias de `pywin32`, `selenium`, `webdriver-manager` declaradas con marcador `sys_platform == 'win32'` donde aplica (`pyproject.toml`).
- **Hora de negocio America/Bogota** independiente del servidor (`docs_adr_0010_cierre_hallazgos_hora_colombia`; `src/insumos/adapters/system/reloj.py:L22` · `src_insumos_adapters_system_reloj_hoy`).
- **Orden**: inventario antes que ventas (`docs_adr_0008_fuente_base_de_datos_orden_inventario_antes_ventas`; `tasks/actualizacion_ventas.py:L500` · `tasks_actualizacion_ventas_actualizacionventas_find_inventario_actualizado`).
