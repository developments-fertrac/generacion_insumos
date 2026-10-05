# 07 · Operación y despliegue

> Fuente: nodos `env_*` (extraídos de `.env.example` y de las llamadas `_env`/`os.getenv` del código), `scripts_publicar_a_produccion_ps1`, `core_logger_*`, `core_email_notifier_*`, `docs_documentacion_generacion_de_insumos_orquestacion_n8n` y FAQ del docx (Parte 1 §4).
> Sin valores sensibles. Rutas expresadas como `<BASE_PATH>`, `<DB_EXPORT_DIR>`, `<carpeta de producción>`.

## 1. Variables de entorno (`.env`)

`config/settings.py` carga `.env` desde la raíz del proyecto (`load_dotenv(_ENV_PATH)`, `config/settings.py`) y expone dataclasses inmutables (`config_settings_settings`, L137).

| Variable | Obligatoria | Uso | Leída en | Nodo |
|---|---|---|---|---|
| `DB_EXPORT_DIR` | Sí para inventario | Carpeta de `Inventario.xlsx` y del `_268`; si falta, ventas busca el `_268` en la carpeta mensual de VENTAS 2026 | `config/settings.py` (`PathsConfig.db_export`) | `env_db_export_dir` |
| `INVENTARIO_BD_FILE` | No (`Inventario.xlsx`) | Nombre de la exportación de inventario | idem | `env_inventario_bd_file` |
| `INVENTARIO_TMP_DIR` | No | Carpeta de la copia temporal descifrada que abre Excel | idem | `env_inventario_tmp_dir` |
| `BASE_PATH` | Recomendada | Raíz de producción (`ARCHIVOS DIARIOS 2026`). **Tiene un valor por defecto con ruta de red en el código** | idem | `env_base_path` · `config_settings_deuda_base_path_por_defecto` |
| `EXCEL_PASSWORD` | Sí | Descifrar/cifrar libros (inventario, MYR, matriz, `.xlsb`) | `ExcelConfig` | `env_excel_password` |
| `EXCEL_PASSWORDS_TRY` | No | Contraseñas alternativas (CSV) | `ExcelConfig` | `env_excel_passwords_try` |
| `VENTAS_ACTUALIZACION_PASSWORD` | Sí en producción | Clave exclusiva de `$2026 VENTAS_Actualizacion.xlsx`; si falta, se usa `EXCEL_PASSWORD` con advertencia | `ExcelConfig`; `tasks/actualizacion_ventas.py:L187` | `env_ventas_actualizacion_password` |
| `VENTAS_MOTOR` | No (`legacy`) | `legacy` \| `reglas` | `tasks/actualizacion_ventas.py:L2171` | `env_ventas_motor` |
| `SMTP_SENDER_EMAIL`, `SMTP_PASSWORD` | Sí | Remitente Gmail (SSL 465, fallback STARTTLS 587) | `SmtpConfig`; `core/email_notifier.py:L51` | `env_smtp_sender_email`, `env_smtp_password` |
| `SMTP_RECIPIENT_EMAILS` / `_VENTAS` / `_INV_GENERAL` | No | Destinatarios por tarea | `SmtpConfig*` | `env_smtp_recipient_emails`, `env_smtp_recipient_emails_ventas`, `env_smtp_recipient_emails_inv_general` |
| `WHATSAPP_CHATS` | Para envío | Chats destino (CSV) | `tasks/envio_informe_ventas.py:L111` | `env_whatsapp_chats` |
| `STATE_DIR` / `LOGS_DIR` | No | Estado operativo (por defecto `%LOCALAPPDATA%\GeneracionInsumos`) | `core/logger.py:L14`, `tasks/envio_informe_ventas.py`, `core/chromedriver_utils.py` | `env_state_dir`, `env_logs_dir` |
| `INSUMOS_DRY_RUN` | Interna | La fija `run.py --dry-run` | `tasks/actualizacion_inventario.py` | `env_insumos_dry_run` (no está en `.env.example`) |
| `LOCALAPPDATA` | Sistema | Base del estado operativo | `core/logger.py` | `env_localappdata` |

Nota: `core/seguridad.py` también trata `FERTRAC_PASS` como secreto (variable del ERP retirado); ningún módulo la lee **[NO VERIFICADO en producción]**.

## 2. Comandos (`uv`)

| Objetivo | Comando |
|---|---|
| Instalar dependencias | `uv sync` · desarrollo: `uv sync --extra dev` |
| Completo (inventario + ventas, sin envío) | `uv run python run.py` |
| Inventario | `uv run python run.py --workflow inventario` |
| Inventario de revisión (sin tocar plantilla ni Excel) | `uv run python run.py --task actualizacion_inv --dry-run` |
| Ventas + envío | `uv run python run.py --workflow ventas` |
| Solo ventas | `uv run python run.py --workflow parcialVentas` |
| Solo envío | `uv run python run.py --workflow envio` |
| Listar | `uv run python run.py --list` (requiere `.env` válido: `Settings()` se instancia antes de listar — `run.py`) |
| Paridad de ventas | `uv run python scripts/comparar_ventas.py` (exit 0 = paridad) |
| Catálogo de reglas | `uv run python scripts/ver_reglas.py` |
| Estructura de dos inventarios | `uv run python scripts/comparar_estructura.py` |
| Calidad | `uv run ruff check .` · `uv run mypy` · `uv run pytest` · `uv run lint-imports` · `uv run python scripts/ratchet_types.py` |

Códigos de salida de `run.py`: `0` todo OK, `1` alguna tarea falló (`orchestrator.py`: `PipelineResult.all_success`).

## 3. Publicación (`scripts/publicar_a_produccion.ps1`)

```powershell
.\scripts\publicar_a_produccion.ps1 -Destino '<carpeta de producción>'
```

| Paso | Detalle | Observación del grafo |
|---|---|---|
| Copia de carpetas | `robocopy /E` de `src`, `tasks`, `core`, `config`, `scripts`, `docs\adr`, `tests` (excluye cachés) | `docs\adr` ya no existe en la rama: se omite en silencio (`if (Test-Path $o)`) |
| Archivos raíz | `run.py`, `orchestrator.py`, `pyproject.toml`, `uv.lock`, `requirements.txt`, `.importlinter`, `.env.example`, `README.md` | **`uv.lock` no existe desde `088f6f0`**: no se copia y el destino conserva un lock antiguo o ninguno (`readme_pendiente_uv_lock`) |
| Limpieza ADR 0008 | Borra en destino `descarga_*`, `core/browser.py`, `core/erp_navigation.py`, `config/erp_selectors.py` | aristas `references` del script hacia módulos (`scripts_publicar_a_produccion_ps1`) |
| Dependencias | `uv sync` en destino | Sin lock, `uv sync` resuelve versiones nuevas (requiere red) |
| No toca | `.env`, `.venv`, logs, datos | Verificar manualmente `DB_EXPORT_DIR` y `VENTAS_ACTUALIZACION_PASSWORD` en el `.env` de destino |

## 4. Integración con n8n

Flujo documentado (docx Parte 1 §3.3; nodo `docs_documentacion_generacion_de_insumos_orquestacion_n8n`): n8n ejecuta `schtasks /Run /TN "<tarea>"` → la tarea programada lanza `<tarea>.bat nopause` → el `.bat` localiza `uv`, ejecuta `uv run --frozen python run.py --workflow <x>`, escribe su bitácora y deja `resultado_<tarea>.txt` con el código de salida (0 = OK) → n8n decide el siguiente paso.

| Requisito | Motivo | Evidencia |
|---|---|---|
| Tarea en modo "Solo interactivo" con usuario logueado (bloqueado sirve, desconectado no) | Excel COM y portapapeles requieren escritorio | `readme_sesion_interactiva`; `tasks_envio_informe_ventas_diagnosticar_portapapeles` |
| `uv` en el `PATH` del usuario de la tarea o `PATH_UV` en el `.bat` | FAQ: "ERROR: no se encontró uv" | `docs_documentacion_generacion_de_insumos_pendiente_bat_sin_uv` |
| `uv.lock` presente | `--frozen` exige lockfile | `readme_pendiente_uv_lock` |
| Inventario antes que ventas | Ventas lee el inventario del día | `docs_adr_0008_fuente_base_de_datos_orden_inventario_antes_ventas` |

Los `.bat` y el flujo de n8n **no están versionados** en este repositorio (solo `test.bat`, diagnóstico de usuario/directorio): **[NO VERIFICADO]** su contenido exacto.

## 5. Logs y notificaciones

| Aspecto | Comportamiento | Evidencia |
|---|---|---|
| Ubicación | `<LOGS_DIR>/<AAAA-MM-DD>/<tarea>.log` (hora Colombia); por defecto `%LOCALAPPDATA%\GeneracionInsumos\logs` | `core/logger.py:L14` · `core_logger_logs_root`; `core/logger.py:L44` · `core_logger_get_logger` |
| Niveles | Archivo DEBUG, consola INFO | idem |
| Secretos | `FiltroSecretos` en ambos handlers reemplaza contraseñas por `sha256:<10 hex>` (mensajes y tracebacks) | `core/seguridad.py:L50` · `core_seguridad_filtrosecretos` |
| Correo | Uno por corrida desde `BaseTask.run`: `[OK]` al terminar todo, `[ERROR]` con traceback y log adjunto | `tasks/base_task.py:L21`; overrides en ventas/envío/inventario |
| Brecha | El cuerpo del correo de fallo de ventas incluye `traceback.format_exc()` **sin pasar por `FiltroSecretos`** (el filtro actúa solo sobre registros de log) | `tasks/actualizacion_ventas.py` (`_notify_failure`, `_ejecutar`) |
| Indicadores | `PROCESANDO.txt` en `BASE_PATH` durante ventas; `resultado_<tarea>.txt` del `.bat` | `tasks_actualizacion_ventas_actualizacionventas_crear_indicador_progreso` |

## 6. Troubleshooting

| Síntoma | Causa probable | Acción | Evidencia |
|---|---|---|---|
| `Variable de entorno requerida no encontrada: EXCEL_PASSWORD` (o `SMTP_*`) | `.env` incompleto o ejecución desde otra carpeta | Completar `.env` en la carpeta del proyecto | `config/settings.py` (`_env(required=True)`) |
| `Falta DB_EXPORT_DIR en .env` | Inventario sin carpeta de exportación | Definir `DB_EXPORT_DIR` | `PathsConfig.inventario_bd` |
| `La salida no cuadra; no se escribe el archivo: …` | Diferencias de referencias/TOTAL INV o filas `TOTAL ≠ EXISTENCIA × COSTO` | Revisar `Inventario.xlsx`; re-exportar | `src_insumos_domain_rules_inventario_actualizacion_validarsalida` |
| `No hay plantilla: ni salida previa … ni maestro …` | Primera corrida sin maestro | Copiar el maestro a `<BASE_PATH>` | `src_insumos_adapters_excel_fuentes_inventario_fuenteinventarioarchivos` |
| `No hay $2026 INVENTARIO GENERAL ACTUALIZADO *.xlsx` | Ventas sin inventario previo | Correr `--workflow inventario` | `tasks_actualizacion_ventas_actualizacionventas_find_inventario_actualizado` |
| `No hay datos del año actual en el informe` | `_268` sin ventas del año | Verificar exportación | `src_insumos_domain_rules_ventas_informe_filtraranioencurso` |
| `No se pudo reemplazar … el archivo está abierto` | Usuario con el libro abierto | Cerrar; el resultado quedó en `_tmp_actualizacion_ventas.xlsx` | `tasks_actualizacion_ventas_actualizacionventas_ejecutar` |
| `No se pudo crear el backup…` | Permisos/disco en `Backups_Ventas` | Liberar; el proceso se detuvo a propósito | idem |
| Ventas "OK" pero el libro no cambió | Ejecución sin COM (`HAS_COM=False`) | Ejecutar en Windows con Excel; ver riesgo R2 | `tasks_actualizacion_ventas_riesgo_paso10_omitido_sin_com` |
| Excel no abre / `grabclipboard() devolvió None` | Tarea sin escritorio | Modo "Solo interactivo"; `INVENTARIO_TMP_DIR` si hay problemas de rutas | `readme_sesion_interactiva` |
| `No se pudo iniciar WhatsApp Web` | Sesión expirada / Chrome actualizado | Escanear QR con el perfil; el sistema revalida ChromeDriver | `tasks_envio_informe_ventas_whatsappweb_iniciar_sesion`; `core_chromedriver_utils_obtener_chromedriver_path` |
| Quedó `PROCESANDO.txt` | Corte durante ventas | Eliminarlo y re-ejecutar | `tasks_actualizacion_ventas_actualizacionventas_crear_indicador_progreso` |
| `uv run --frozen` falla por lockfile | `uv.lock` ausente | `uv lock` con red y versionar | `readme_pendiente_uv_lock` |
