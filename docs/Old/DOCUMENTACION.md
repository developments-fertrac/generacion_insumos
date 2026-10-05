# Documentacion por Modulo - Generacion de Insumos

---

## Modulo: Config

### Archivo: `config/settings.py`

**Imports:**
- `os` — acceder a variables de entorno via `os.getenv()`
- `dataclasses` — `dataclass`, `field` — clases de datos con valores por defecto
- `Path` (de `pathlib`) — manejo de rutas del sistema de archivos
- `date` (de `datetime`) — obtener fecha actual para carpetas mensuales
- `dotenv` / `load_dotenv` — cargar variables desde `.env` al inicio de la aplicacion

**Funcion auxiliar: `_env(key, default=None, required=False)`**
- Requiere una **llave** (nombre de la variable de entorno)
- Valores por **default**: `None` (si no se especifica)
- Si `required=True` y la llave no existe, lanza `EnvironmentError`
- Retorna `str` vacio `""` si no encuentra el valor y no es requerida
- Se usa para **leer credenciales, rutas y configuraciones** desde `.env` sin hardcodear valores

**Diccionarios de meses:**
- `MESES_ES` — mapa `{int -> str}`: `{1: "ENERO", 2: "FEBRERO", ...}`
- `MESES_ES_NOMBRE` — mapa invertido `{str -> int}`: `{"enero": 1, "febrero": 2, ...}`
- `MESES_ES_INVERTIDO` — alias similar a `MESES_ES_NOMBRE`
- `MESES_REMISIONES` — lista con prefijos numericos: `["1. ENERO", "2. FEBRERO", ...]`
- Se usan para **generar nombres de carpetas mensuales** automaticamente

**Funciones de fecha:**
- `get_month_folder()` — retorna `"MM. MES"` (ej: `"08. AGOSTO"`) segun el mes actual
- `get_month_folder_remisiones()` — retorna `"M. MES"` (ej: `"8. AGOSTO"`) segun el mes actual

**Dataclasses (configuracion tipada):**

| Clase | Campos principales | Valores por defecto |
|-------|-------------------|---------------------|
| `ErpConfig` | `usuario`, `clave`, `url_login`, `url_inventario`, `url_productos`, `url_crm` | `usuario="consultas"`, `clave` desde `.env` (requerida) |
| `SmtpConfig` | `server`, `port`, `sender_email`, `sender_password`, `recipient_emails`, `enabled` | `server="smtp.gmail.com"`, `port=587`, `enabled=True` |
| `SmtpConfigVentas` | Hereda de `SmtpConfig` | Sobreescribe `sender_email="ctorrese@fertrac.com"` |
| `SmtpConfigInvGeneral` | Hereda de `SmtpConfig` | Agrega destinatarios: `asistentecompras`, `analistacompras5`, `ctorres` |
| `ExcelConfig` | `password`, `passwords_try` | `password` desde `.env` (default `"sha256:74a57e3d78"`), `passwords_try` desde `.env` (default `"sha256:918f1f5c3c"`) |
| `PathsConfig` | `base`, `remisiones` + properties | `base` y `remisiones` desde `.env` con defaults hardcodeados |
| `BrowserConfig` | `headless`, timeouts, reintentos | `headless` desde `.env` (default `True`), `max_intentos_driver=5`, `max_intentos_login=3` |
| `Settings` | Composicion de todas las anteriores | Instancia todo: `erp`, `smtp`, `smtp_ventas`, `smtp_inv_general`, `excel`, `paths`, `browser` |

**Propiedades de `PathsConfig`:**
- `informes` → `base / "INFORMES"`
- `inventario_general` → `informes / "INVENTARIO GENERAL ACTUALIZADO"`
- `inventario_general_mes` → `inventario_general / get_month_folder()`
- `ventas` → `informes / "VENTAS 2026"`
- `ventas_mes` → `ventas / get_month_folder()`
- `valorizados` → `base / "Pruebas Inv General" / "Valorizados"`
- `output_inv_general` → `base / "Pruebas Inv General"`
- `remisiones_mes` → `remisiones / get_month_folder_remisiones()`
- `ensure_dirs()` — crea todas las carpetas necesarias si no existen

**Uso tipico:**
```python
from config.settings import Settings
settings = Settings()
# settings.erp.usuario          -> "consultas"
# settings.browser.headless     -> True/False desde .env
# settings.paths.ventas_mes     -> Path(".../08. AGOSTO")
```

---

## Modulo: Core

### Archivo: `core/logger.py`

**Imports:**
- `logging` — modulo estandar de Python para logging
- `sys` — acceder a `sys.stdout` para el handler de consola
- `date` (de `datetime`) — obtener fecha actual para nombrar carpetas de logs
- `Path` (de `pathlib`) — manejo de rutas

**Constantes:**
- `_LOGS_ROOT` — ruta raiz: `project_root / "logs"`

**Funcion: `get_logger(name)`**
- Requiere un `name: str` (nombre del logger, generalmente coincide con el nombre de la task)
- Crea un logger con **nivel DEBUG**
- Crea automaticamente la carpeta del dia: `logs/YYYY-MM-DD/`
- Escribe a archivo: `logs/YYYY-MM-DD/{name}.log` (encoding UTF-8, modo append)
- Escribe a consola: solo nivel INFO y superior
- Formato: `[YYYY-MM-DD HH:MM:SS] [LEVEL ] [name] mensaje`
- Si el logger ya tiene handlers, retorna el existente (evita duplicados)
- Se usa para **unificar logging** en todo el proyecto con un formato consistente

**Uso tipico:**
```python
from core.logger import get_logger
log = get_logger("descarga_ventas")
log.info("Iniciando proceso...")
```

---

### Archivo: `core/email_notifier.py`

**Imports:**
- `socket` — obtener hostname del servidor para incluir en el email
- `smtplib` — enviar emails via SMTP
- `ssl` — contexto SSL para conexion segura
- `datetime` — timestamp para el asunto del email
- `MIMEBase`, `MIMEMultipart`, `MIMEText` (de `email.mime`) — construir emails con formato HTML
- `encoders` (de `email`) — codificar archivos adjuntos en base64
- `Path` (de `pathlib`) — manejo de rutas para adjuntos
- `SmtpConfig` (de `config.settings`) — configuracion SMTP tipada
- `get_logger` (de `core.logger`) — logging unificado

**Clase: `EmailNotifier`**

| Metodo | Parametros | Descripcion |
|--------|-----------|-------------|
| `__init__` | `config: SmtpConfig`, `task_name: str` | Inicializa con configuracion SMTP y nombre de la task |
| `send` | `subject: str`, `html_body: str`, `attachment: Path\|None` | Envia un email. Retorna `True` si fue exitoso |
| `notify_success` | `detail: str`, `attachment: Path\|None` | Envia email de exito con asunto `[OK] task_name - Completado fecha` |
| `notify_failure` | `error: str`, `attachment: Path\|None` | Envia email de error con asunto `[ERROR] task_name - Error fecha` |

**Metodo interno: `_send_with_fallback(msg)`**
- Intenta primero **SSL directo** (puerto 465)
- Si falla, usa **STARTTLS** (puerto 587) como fallback
- Timeout de 30 segundos por conexion

**Uso tipico:**
```python
from config.settings import Settings
from core.email_notifier import EmailNotifier

settings = Settings()
notifier = EmailNotifier(settings.smtp, "descarga_ventas")
notifier.notify_success("Descarga completada - 150 archivos procesados")
notifier.notify_failure("Timeout esperando descarga")
```

---

### Archivo: `core/browser.py`

**Imports:**
- `os` — manejo de directorios, variables de entorno
- `tempfile` — crear directorios temporales para perfiles de Chrome
- `time` — sleep entre reintentos
- `webdriver` (de `selenium`) — control del navegador Chrome
- `Options`, `Service` (de `selenium.webdriver.chrome`) — configuracion del driver
- `By` (de `selenium.webdriver.common`) — localizadores XPath/CSS
- `EC`, `WebDriverWait` (de `selenium.webdriver.support`) — esperas explicitas
- `ChromeDriverManager` (de `webdriver_manager.chrome`) — descarga automatica de ChromeDriver
- `BrowserConfig`, `ErpConfig` (de `config.settings`) — configuracion tipada
- `get_logger` (de `core.logger`) — logging unificado

**Funciones principales:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `create_driver` | `download_dir: str`, `config: BrowserConfig\|None` | Crea instancia de Chrome con configuracion. Reintenta hasta `max_intentos_driver` veces |
| `do_login` | `driver`, `erp: ErpConfig`, `config: BrowserConfig\|None` | Realiza login en ERP Fertrac. Reintenta hasta `max_intentos_login` veces. Retorna `True/False` |
| `accept_cookies_if_present` | `driver` | Acepta banner de cookies si aparece (boton OK) |
| `wait_for_download` | `carpeta: str`, `timeout: int`, `snapshot: dict\|None` | Espera a que aparezca un archivo nuevo o modificado en la carpeta |
| `take_xlsx_snapshot` | `carpeta: str` | Toma fotografia de archivos de hoja de calculo existentes (.xlsx, .xlsm, .xls, .csv) con su mtime |
| `wait_for_download_complete` | `carpeta`, `snapshot`, `timeout`, `grace`, `margin` | Espera a que la descarga se complete (sin archivos temporales .crdownload/.tmp) |
| `_glob_hojas` | `carpeta: str` | Interna. Lista archivos .xlsx, .xlsm, .xls, .csv de un directorio |

**Funcion interna: `_build_options(download_dir, config)`**
- Configura directorio de descarga, popup deshabilitado
- Si `headless=True`: agrega `--headless=new`, `--window-size=1920,1080`
- Siempre agrega: `--disable-gpu`, `--no-sandbox`, `--disable-dev-shm-usage`, `--ignore-certificate-errors`, `--disable-extensions`, `--disable-background-networking`, `--disable-default-apps`, `--disable-sync`, `--metrics-recording-only`
- Crea directorio temporal para perfil de Chrome (`--user-data-dir`)
- **Nota:** ya no agrega `--remote-debugging-port=0`

**Funciones internas de gestion del driver:**
- `_chrome_version_instalada()` — lee la version instalada de Chrome desde el binario (BLBeacon). Retorna `str | None`
- `_driver_path(version_dir)` — devuelve la ruta del ChromeDriver usando `ChromeDriverManager(driver_version=<mymayor>)`, alineandolo con la version mayor del Chrome instalado (evita el error `SessionNotCreatedException` por descargar un driver mas nuevo que el Chrome real)

**Soporte de formato de descarga:**
- `take_xlsx_snapshot` y `wait_for_download_complete` detectan **.xlsx, .xlsm, .xls y .csv** (el ERP Fertrac a veces exporta `product.template.xls` en lugar de `.xlsx`; sin esta extension el proceso quedaba esperando el timeout).

**Uso tipico:**
```python
from config.settings import Settings
from core.browser import create_driver, do_login

settings = Settings()
driver = create_driver("/path/to/downloads", settings.browser)
if do_login(driver, settings.erp, settings.browser):
    # navegar y descargar...
driver.quit()
```

---

### Archivo: `core/erp_navigation.py`

**Imports:**
- `re` — expresiones regulares (normalizacion de texto)
- `time` — esperas entre acciones Selenium
- `ActionChains`, `By`, `Keys` (de `selenium`) — interaccion con el navegador
- `EC`, `WebDriverWait` (de `selenium.webdriver.support`) — esperas explicitas
- `Select` (de `selenium.webdriver.support.ui`) — dropdowns HTML
- `get_logger` (de `core.logger`)
- Selectores XPath centralizados desde `config/erp_selectors`

**Funciones principales:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `click_with_fallback` | `driver`, `element`, `description` | Hace click; si falla, reintenta con click via JS (`arguments[0].click()`) |
| `find_visible_element` | `driver`, `selectors: list[str]`, `timeout` | Espera y retorna el primer elemento visible/clickable de una lista de selectores. Retorna `None` si no aparece |
| `find_all_visible` | `driver`, `selector` | Retorna todos los elementos visibles que coinciden con un selector |
| `open_menu_by_selectors` | `driver`, `selectors`, `menu_name`, `timeout` | Abre un menu usando selectores (con scroll y click con fallback) |
| `close_modal` | `driver` | Cierra un modal abierto |
| `get_pager_info` | `driver` | Lee el paginador Odoo (ej: `1-80 / 15941`). Soporta `<span class="o_pager_value">` y `<span class="o_pager_limit">` |
| `open_action_menu`, `select_export_option`, `select_export_file_option` | `driver`, ... | Flujo de exportacion del ERP |
| `detect_and_accept_empty_alert` | `driver`, `timeout` | Detecta y acepta la alerta ERP "No hay registros que coincidan" (UI, no JS) |
| `wait_for_erp_section` | `driver`, `url`, `section_check`, `timeout` | Navega a una URL y espera la seccion |
| `descartar_alert_js` | `driver` | Acepta una **alerta JavaScript** abierta (`window.alert`). Evita el "congelamiento" de Selenium cuando una alerta nativa del navegador bloquea el bucle de comandos. Retorna `True` si habia una alerta |

---

### Archivo: `core/excel_utils.py`

**Imports:**
- `io` — manejo de streams en memoria (`BytesIO`)
- `os` — bloqueo de archivos, tamano
- `re` — expresiones regulares para normalizacion
- `difflib` — coincidencia difusa de columnas (`get_close_matches`)
- `Path` (de `pathlib`) — manejo de rutas
- `msoffcrypto` — desencriptar archivos Excel protegidos
- `pandas` — lectura/escritura de DataFrames
- `unidecode` (de `unidecode`) — convertir caracteres especiales a ASCII
- `ExcelConfig`, `MESES_ES`, `MESES_ES_NOMBRE`, `MESES_ES_INVERTIDO` (de `config.settings`)
- `get_logger` (de `core.logger`)

**Funciones de normalizacion:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `norm` | `s: str` | Normaliza texto: lowercase, sin acentos, sin caracteres especiales. Usa cache interno |
| `norm_simple` | `s: str` | Normaliza texto: lowercase, sin acentos, conserva espacios y caracteres |
| `norm_colname` | `s: str` | Normaliza nombres de columnas: lowercase, sin acentos, sin caracteres especiales |
| `norm_label` | `s: str\|None` | Normaliza etiquetas: acepta `None` retornando `""`, lowercase, sin acentos |
| `norm_sheet` | `s: str` | Normaliza nombres de hojas: lowercase, sin acentos, sin espacios extra |
| `norm_base_filename` | `name: str` | Extrae stem del archivo, elimina `$` y `~$`, normaliza |
| `strip_dolares_temporales` | `name: str` | Elimina `~$` y `$` iniciales del nombre del archivo |

**Funciones de desencriptacion:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `decrypt_to_stream` | `xlsx_path: Path`, `password: str\|None`, `config: ExcelConfig\|None` | Desencripta archivo Excel. Prueba contrasenas en orden. Retorna `BytesIO` |
| `_decrypt_with_retry` | `xlsx_path: Path`, `passwords: list[str]` | Intenta desencriptar con cada contrasena. Si no esta encriptado, retorna directamente |
| `_read_file_raw` | `path: Path` | Lee archivo bytes con bloqueo Windows (`msvcrt.locking`) para evitar conflictos |

**Funciones de lectura Excel:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `read_excel_any` | `xlsx: Path\|BytesIO`, `**kwargs` | Lee Excel (.xlsx, .xls, .csv) detectando formato automaticamente via magic bytes |
| `find_sheet_name` | `xlsx_stream_or_path`, `targets: tuple[str,...]` | Busca hoja por nombres normalizados. Retorna primera coincidencia o primera hoja |
| `find_sheet_by_pattern` | `workbook`, `patron: str`, `ignorar_dolares: bool` | Busca hoja por patron en nombre. Opcionalmente ignora prefijos `$` |
| `find_file_by_pattern` | `directorio: str\|Path`, `patron: str`, `ignorar_dolares: bool` | Busca archivo .xlsx por patron en nombre dentro de un directorio |

**Funciones de columnas:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `resolve_cols` | `df: DataFrame`, `cols_map: dict` | Resuelve nombres de columnas con normalizacion, sinonomos y coincidencia difusa |

**Funcion de escritura:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `write_excel` | `df_or_dict`, `path_out: Path`, `sheetname: str\|None` | Escribe DataFrame o dict de DataFrames a Excel |

**Constantes:**
- `COL_SYNONYMS` — diccionario de sinonomos de columnas (vacio por defecto, extensible)
- `_NORM_CACHE` — cache para la funcion `norm` (evita recalcular normalizaciones)

**Uso tipico:**
```python
from core.excel_utils import decrypt_to_stream, read_excel_any, norm

# Desencriptar y leer
stream = decrypt_to_stream(Path("archivo.xlsx"), config=settings.excel)
df = read_excel_any(stream, sheet_name="Inventario")

# Normalizar texto para comparaciones
norm("PROVEEDOR S.A.") == norm("proveedor s.a.")  # True
```

---

### Archivo: `core/chromedriver_utils.py` (NUEVO - auto-recuperacion del ChromeDriver)

**Problema que resuelve:**
- Chrome se actualiza automaticamente y deja el `chromedriver.exe` local obsoleto. Cuando la version mayor del driver no coincide con la del Chrome instalado, Selenium lanza:
  `session not created: This version of ChromeDriver only supports Chrome version X / Current browser version is Y`.
- Ese error tiraba abajo el envio de informes por WhatsApp Web (`envio_informe_ventas`). Este modulo lo detecta y se recupera solo.

**Imports:**
- `os`, `re`, `shutil`, `subprocess` — rutas, parseo, copias y consulta de versiones
- `Path` (de `pathlib`) — manejo de rutas
- `get_logger` (de `core.logger`) — logging unificado

**Constantes:**
- `PROJECT_ROOT` — raiz del proyecto
- `_RAICES_CACHE` — carpetas donde buscar drivers cacheados por webdriver-manager: `.wdm`, `.wdm_cache`, `~/.wdm`
- `_RUTAS_CHROME` — rutas tipicas del binario `chrome.exe`
- `_CLAVES_REGISTRO` — claves de registro de Windows con la version de Chrome (HKLM/HKCU BLBeacon)

**Funciones publicas:**

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `chrome_version_instalada` | — | Version completa de Chrome instalado (registro de Windows primero, binario como respaldo). Retorna `str\|None` |
| `chrome_version_mayor` | — | Version mayor de Chrome instalada (ej. `"153"`) |
| `chromedriver_version` | `ruta` | Version completa del ChromeDriver en esa ruta (ej. `153.0.8010.52`) |
| `chromedriver_mayor` | `ruta` | Version mayor del ChromeDriver en esa ruta |
| `es_compatible` | `ruta`, `chrome_mayor` | `True` si el driver es compatible con la mayor de Chrome. Si no se puede leer la version de Chrome, asume compatible |
| `obtener_chromedriver_path` | `reemplazar_local: bool=True` | **Entrada principal.** Si el local ya es compatible lo usa; si no, busca en cache/descarga y lo copia encima del local (con backup `.exe.bak`), dejando la recuperacion hecha para las proximas corridas |
| `asegurar_chromedriver_local` | — | Revalida y, si hace falta, reemplaza el ChromeDriver local. Devuelve su ruta |

**Flujo de recuperacion (`obtener_chromedriver_path`):**
1. Detecta la version mayor de Chrome instalado (registro de Windows / binario).
2. Valida el `chromedriver.exe` local de la raiz del proyecto contra esa version.
3. Si el local no sirve, busca un driver compatible en las caches de webdriver-manager (`.wdm` / `.wdm_cache`), ya descargado en corridas previas (sin internet).
4. Si no hay cache compatible, descarga el ChromeDriver correcto con `ChromeDriverManager` (`webdriver-manager`), primero por version mayor y luego auto-detect.
5. Reemplaza el ChromeDriver local obsoleto por el compatible, copiando el anterior a `chromedriver.exe.bak` (por si se quiere volver atras). Si el archivo local esta bloqueado por un proceso activo, devuelve la ruta del driver bueno de cache/descarga para no fallar igualmente.

**Uso tipico:**
```python
from core.chromedriver_utils import obtener_chromedriver_path

# Devuelve la ruta de un ChromeDriver compatible y deja el local actualizado
path = obtener_chromedriver_path(reemplazar_local=True)
service = Service(path)  # selenium.webdriver.chrome.service.Service
```

**Nota:** `core/browser.py` no usa este modulo (ya fuerza la version via `ChromeDriverManager`); el modulo se creo para la task `envio_informe_ventas`, que lanza Chrome con el driver local de la raiz.

---

## Modulo: Tasks

### Archivo: `tasks/base_task.py` (Clase abstracta base)

**Imports:**
- `time` — medir tiempo de ejecucion
- `traceback` — formatear excepciones completas
- `ABC`, `abstractmethod` (de `abc`) — clase abstracta y metodo abstracto
- `Path` (de `pathlib`) — manejo de rutas
- `Settings` (de `config.settings`) — configuracion centralizada
- `get_logger` (de `core.logger`) — logging unificado
- `EmailNotifier` (de `core.email_notifier`) — notificaciones por email

**Clase: `BaseTask(ABC)`**

| Propiedad/Metodo | Tipo | Descripcion |
|-----------------|------|-------------|
| `name` | `str` | Nombre de la task (sobreescribir en subclases) |
| `settings` | `Settings` | Configuracion inyectada via constructor |
| `log` | `Logger` | Logger unificado para esta task |
| `notifier` | `EmailNotifier\|None` | Notificador de email (opcional) |
| `run()` | `bool` | **Template Method**: ejecuta `setup()` → `execute()` → `teardown()` y maneja notificaciones |
| `setup()` | `None` | Hook opcional antes de execute (sobreescribir si necesita preparacion) |
| `execute()` | `None` | **Abstracto** — implementar la logica de la task |
| `teardown()` | `None` | Hook opcional despues de execute (limpieza, cerrar drivers, etc.) |

**Flujo de `run()`:**
1. Log de inicio con separador visual
2. Ejecuta `setup()`
3. Ejecuta `execute()`
4. Si exito: calcula tiempo transcurrido, log de exito, notifica email
5. Si falla: captura traceback completo, log de error, notifica email con error
6. Finalmente ejecuta `teardown()`

**Uso tipico:**
```python
from tasks.base_task import BaseTask

class MiTask(BaseTask):
    name = "mi_task"

    def setup(self):
        self.driver = create_driver(...)

    def execute(self):
        # logica principal
        pass

    def teardown(self):
        if hasattr(self, 'driver'):
            self.driver.quit()
```

---

### Archivo: `tasks/__init__.py` (Registro de tasks)

**Imports:**
- `BaseTask` (de `tasks.base_task`)
- `DescargaValorizados`, `DescargaInventarioGeneral`, `DescargaInformeVentas` — tasks de scraping
- `ActualizacionInventario`, `ActualizacionVentas` — tasks de procesamiento local

**Constante: `TASK_REGISTRY: dict[str, type[BaseTask]]`**
- Mapea nombre de task → clase
- `"descarga_valorizados"` → `DescargaValorizados`
- `"descarga_inv_general"` → `DescargaInventarioGeneral`
- `"descarga_ventas"` → `DescargaInformeVentas`
- `"actualizacion_inv"` → `ActualizacionInventario`
- `"actualizacion_ventas"` → `ActualizacionVentas`

**Uso tipico:**
```python
from tasks import TASK_REGISTRY
task_class = TASK_REGISTRY["descarga_ventas"]  # retorna DescargaInformeVentas
task = task_class(settings)
```

---

### Archivos de tasks individuales

| Archivo | Clase | Tipo | Descripcion |
|---------|-------|------|-------------|
| `tasks/descarga_inventario_general.py` | `DescargaInventarioGeneral` | Scraping | Descarga informe de inventario general desde ERP Fertrac via Selenium |
| `tasks/descarga_informe_ventas.py` | `DescargaInformeVentas` | Scraping | Descarga informe de ventas desde ERP Fertrac via Selenium |
| `tasks/descarga_valorizados.py` | `DescargaValorizados` | Scraping | Descarga valorizados desde ERP Fertrac via Selenium |
| `tasks/actualizacion_inventario.py` | `ActualizacionInventario` | Procesamiento | Procesa archivos Excel de inventario y genera consolidado actualizado |
| `tasks/actualizacion_ventas.py` | `ActualizacionVentas` | Procesamiento | Procesa archivos Excel de ventas y genera consolidado actualizado |

---

## Modulo: Orchestrator

### Archivo: `orchestrator.py`

**Imports:**
- `time` — medir tiempo total del pipeline
- `dataclasses` — `dataclass`, `field` — estructuras de resultado
- `Settings` (de `config.settings`) — configuracion
- `get_logger` (de `core.logger`)
- `TASK_REGISTRY` (de `tasks`) — registro de tasks disponibles
- `BaseTask` (de `tasks.base_task`)

**Dataclasses de resultado:**

| Clase | Campos | Descripcion |
|-------|--------|-------------|
| `PhaseResult` | `task_name: str`, `success: bool`, `elapsed_seconds: int` | Resultado de una sola task |
| `PipelineResult` | `phases: list[list[PhaseResult]]`, `total_elapsed: int` | Resultado completo del pipeline |

**Pipelines predefinidos:**

| Pipeline | Fases | Descripcion |
|----------|-------|-------------|
| `PIPELINE_COMPLETO` | Fase 1: `[descarga_inv_general, descarga_ventas, descarga_valorizados]` (secuencial) → Fase 2: `[actualizacion_inv, actualizacion_ventas]` (paralelo) | Ejecucion completa |
| `PIPELINE_VENTAS` | Fase 1: `[descarga_ventas]` → Fase 2: `[actualizacion_ventas]` | Solo ventas |
| `PIPELINE_INVENTARIO` | Fase 1: `[descarga_inv_general, descarga_valorizados]` → Fase 2: `[actualizacion_inv]` | Solo inventario |
| `PIPELINES` | Diccionario `{"completo": ..., "ventas": ..., "inventario": ...}` | Lookup por nombre |

**Clase: `Orchestrator`**

| Metodo | Parametros | Descripcion |
|--------|-----------|-------------|
| `__init__` | `settings: Settings`, `pipeline: list[list[str]]\|None` | Inicializa con configuracion y pipeline (default: completo) |
| `run_all()` | — | Ejecuta todas las fases secuencialmente. Retorna `PipelineResult` |
| `run_task(task_name)` | `task_name: str` | Ejecuta una sola task por nombre |
| `_run_phase(phase_idx, task_names)` | `int`, `list[str]` | Ejecuta todas las tasks de una fase (actualmente secuencial) |
| `_execute_task(task_class)` | `type[BaseTask]` | Instancia y ejecuta una task, retorna `PhaseResult` |

**Comportamiento:**
- Fases se ejecutan **secuencialmente** (fase 1 completa antes de fase 2)
- Si una fase falla, **continua** con la siguiente fase (no aborta el pipeline)
- Al final genera **resumen** con estado de cada task y tiempo total

---

## Modulo: Run (Entry Point)

### Archivo: `run.py`

**Imports:**
- `argparse` — parsing de argumentos de linea de comandos
- `sys` — acceso a `sys.path` y `sys.exit()`
- `Path` (de `pathlib`) — para agregar root del proyecto al `sys.path`
- `Settings` (de `config.settings`)
- `get_logger` (de `core.logger`)
- `Orchestrator`, `PIPELINES` (de `orchestrator`)
- `TASK_REGISTRY` (de `tasks`)

**Argumentos de CLI:**

| Argumento | Alias | Tipo | Descripcion |
|-----------|-------|------|-------------|
| `--workflow` | `-w` | `str` (choices: `completo`, `ventas`, `inventario`) | Ejecutar un workflow completo |
| `--task` | `-t` | `str` | Ejecutar solo una task especifica |
| `--phase` | `-p` | `int` (1 o 2) | Ejecutar solo una fase especifica |
| `--list` | `-l` | `bool` (flag) | Listar workflows y tasks disponibles |

**Ejemplos de uso:**
```bash
python run.py                          # Pipeline completo
python run.py --workflow ventas        # Solo ventas
python run.py --workflow inventario    # Solo inventario
python run.py --task descarga_ventas   # Solo una task
python run.py --phase 1                # Solo fase 1 (scraping)
python run.py --phase 2                # Solo fase 2 (procesamiento)
python run.py --list                   # Listar opciones
```

---

## Modulo: Entorno

### Archivo: `.env`

**Variables de entorno:**

| Variable | Descripcion | Default |
|----------|-------------|---------|
| `FERTRAC_USER` | Usuario ERP Fertrac | `consultas` |
| `FERTRAC_PASS` | Contrasena ERP Fertrac | *(requerida)* |
| `SMTP_PASSWORD` | Contrasena App Gmail para notificaciones | *(requerida para email)* |
| `EXCEL_PASSWORD` | Contrasena para desencriptar archivos Excel | `sha256:74a57e3d78` |
| `EXCEL_PASSWORDS_TRY` | Lista separada por comas de contrasenas alternas | `sha256:918f1f5c3c` |
| `BASE_PATH` | Ruta base del servidor de archivos | `D:\Fertrac\...\ARCHIVOS DIARIOS 2026` |
| `REMISIONES_BASE` | Ruta base de remisiones | `D:\Fertrac\...\REMISIONES` |
| `HEADLESS` | Modo sin ventana del navegador | `true` |
