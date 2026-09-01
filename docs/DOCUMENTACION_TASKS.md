# Documentacion Detallada de Tasks - Generacion de Insumos

---

## 1. `descarga_inventario_general.py` (538 lineas)

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Permite usar anotaciones de tipo modernas (list[str] en vez de List[str]) |
| `os` | modulo completo | Acceder a variables de entorno, rutas del sistema |
| `re` | modulo completo | Expresiones regulares para detectar versiones de Chrome |
| `subprocess` | modulo completo | Ejecutar comandos del sistema (verificar version Chrome) |
| `time` | modulo completo | Esperas entre acciones Selenium (sleep) |
| `datetime` | `datetime` | Obtener fecha actual para renombrar archivos descargados |
| `selenium.webdriver.common.action_chains` | `ActionChains` | Simular acciones complejas del mouse/teclado |
| `selenium.webdriver.common.keys` | `Keys` | Enviar teclas especiales (Enter, Escape, etc.) |
| `selenium.webdriver.support` | `expected_conditions`, `WebDriverWait` | Esperas explicitas hasta que un elemento este listo |
| `selenium.webdriver.support.ui` | `Select` | Interactuar con dropdowns HTML |
| `config.erp_selectors` | `Checkbox`, `Menu`, `Modal`, `Navegacion`, `Paginador`, `VistaLista` | Selectores XPath centralizados para el ERP |
| `config.settings` | `Settings` | Configuracion centralizada del proyecto |
| `core.browser` | `create_driver`, `do_login`, `take_xlsx_snapshot`, `wait_for_download_complete` | Crear navegador Chrome, hacer login, gestionar descargas |
| `core.erp_navigation` | `close_modal`, `click_with_fallback`, `find_all_visible`, `find_visible_element`, `open_action_menu`, `open_menu_by_selectors`, `select_export_file_option`, `select_export_option`, `select_from_modal_dropdown`, `wait_for_erp_section` | Funciones compartidas de navegacion en el ERP |
| `core.email_notifier` | `EmailNotifier` | Enviar notificaciones por email |
| `tasks.base_task` | `BaseTask` | Clase abstracta base para todas las tasks |

### Clase: `DescargaInventarioGeneral(BaseTask)`

**Atributo de clase:**
- `name = "descarga_inv_general"` — nombre identificador de la task

#### Metodo: `setup(self) -> None`
- **Que hace:** Inicializa el notificador de email con la configuracion SMTP para inventario general.
- **Por que existe:** Se ejecuta antes de `execute()` para preparar los recursos necesarios. Es el hook de inicializacion del patron Template Method.

#### Metodo: `execute(self) -> None`
- **Que hace:** Orquesta todo el proceso de descarga: crea el navegador, hace login, navega al inventario, selecciona productos, cambia a vista lista, detecta total de registros, modifica el rango, espera carga, selecciona todos, exporta y renombra el archivo.
- **Por que existe:** Es el metodo principal que contiene la logica de negocio de la descarga. Es el metodo abstracto heredado de `BaseTask` que debe ser implementado por cada subclase.

#### Metodo: `teardown(self) -> None`
- **Que hace:** Cierra el navegador Chrome si esta abierto.
- **Por que existe:** Libera recursos del sistema despues de ejecutar la task. Se ejecuta siempre, tanto si hubo exito como si hubo error.

#### Metodo: `_verificar_versiones_chrome() -> None` (estatico)
- **Que hace:** Verifica que la version de Chrome y ChromeDriver sean compatibles ejecutando comandos del sistema.
- **Por que existe:** Evita errores de compatibilidad entre versiones de navegador y driver Selenium.

#### Metodo: `_navegar_a_inventario(self) -> None`
- **Que hace:** Navega a la seccion de Inventario del ERP Fertrac usando los selectores de navegacion centralizados.
- **Por que existe:** Accede a la pagina donde se encuentra el informe de inventario general.

#### Metodo: `_seleccionar_productos(self) -> bool`
- **Que hace:** Selecciona la opcion "Productos" del menu de datos principales del ERP.
- **Por que existe:** Es el primer paso para acceder al listado de productos del inventario.

#### Metodo: `_cambiar_a_vista_lista(self) -> bool`
- **Que hace:** Cambia la vista del ERP de Kanban/Grid a Lista para poder seleccionar todos los registros.
- **Por que existe:** La vista lista es necesaria para poder seleccionar registros y exportarlos.

#### Metodo: `_detectar_total_registros(self) -> tuple`
- **Que hace:** Lee el paginador del ERP para detectar cuantos registros hay en total (ej: "1-80 / 1500").
- **Por que existe:** Necesita saber el total para poder modificar el rango y cargar todos los registros de una vez.

#### Metodo: `_modificar_rango_registros(self, total: int) -> bool`
- **Que hace:** Modifica el campo de rango del paginador para mostrar todos los registros de una vez.
- **Por que existe:** Por defecto el ERP muestra 80 registros por pagina. Este metodo cambia el limite para mostrar todos.

#### Metodo: `_esperar_carga_completa(self, total: int) -> bool`
- **Que hace:** Espera a que todos los registros se carguen en la pagina, verificando el paginador periodicamente.
- **Por que existe:** Despues de modificar el rango, la pagina tarda en cargar todos los registros. Este metodo espera a que termine.

#### Metodo: `_seleccionar_todos_registros(self) -> bool`
- **Que hace:** Hace click en el checkbox del header para seleccionar todos los registros visibles.
- **Por que existe:** Para exportar todos los registros, primero deben estar seleccionados.

#### Metodo: `_esperar_procesamiento_post_exportacion(self) -> None`
- **Que hace:** Espera a que el ERP termine de procesar la exportacion antes de continuar.
- **Por que existe:** Despues de hacer click en exportar, el ERP tarda en generar el archivo.

#### Metodo: `_renombrar_archivo_con_fecha(self, archivo_original: str) -> str`
- **Que hace:** Renombra el archivo descargado agregando la fecha actual al nombre.
- **Por que existe:** Para mantener un historial organizado de archivos descargados.
- **Detalle (mejora):** el `os.rename` reintenta hasta **60 veces x 2 s (~2 min)** si el archivo recien descargado esta aun en uso (Chrome/antivirus aun lo tiene abierto), evitando el error `PermissionError: [WinError 32]` intermitente.

---

## 2. `descarga_informe_ventas.py` (577 lineas)

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Anotaciones de tipo modernas |
| `glob` | modulo completo | Buscar archivos por patron en directorios |
| `os` | modulo completo | Variables de entorno, rutas |
| `time` | modulo completo | Esperas entre acciones Selenium |
| `datetime` | `datetime` | Fecha actual para configurar rangos de fechas |
| `selenium.webdriver.common.by` | `By` | Localizadores XPath/CSS |
| `selenium.webdriver.common.keys` | `Keys` | Teclas especiales |
| `selenium.webdriver.support` | `expected_conditions`, `WebDriverWait` | Esperas explicitas |
| `selenium.webdriver.support.ui` | `Select` | Dropdowns HTML |
| `config.erp_selectors` | `InformeVentas`, `Menu`, `Navegacion`, `Valorizado` | Selectores XPath para informes de ventas |
| `core.browser` | `create_driver`, `do_login`, `take_xlsx_snapshot`, `wait_for_download_complete` | Navegador y gestion de descargas |
| `core.email_notifier` | `EmailNotifier` | Notificaciones email |
| `core.erp_navigation` | `click_with_fallback`, `find_all_visible`, `find_visible_element`, `open_menu_by_selectors`, `wait_for_erp_section` | Navegacion compartida en ERP |
| `tasks.base_task` | `BaseTask` | Clase abstracta base |

### Clase: `DescargaInformeVentas(BaseTask)`

**Atributo de clase:**
- `name = "descarga_ventas"` — nombre identificador de la task

#### Metodo: `setup(self) -> None`
- **Que hace:** Inicializa el notificador de email con la configuracion SMTP para ventas.
- **Por que existe:** Prepara recursos antes de ejecutar la descarga.

#### Metodo: `execute(self) -> None`
- **Que hace:** Orquesta toda la descarga del informe de ventas: crea navegador, hace login, navega a CRM, abre menu de informes, selecciona ventas y remisiones, selecciona informe de facturas, configura fechas del ano completo, marca mostrar costo, configura tipo detallado, genera XLSX y espera descarga.
- **Por que existe:** Es la tarea principal que contiene toda la logica de negocio.

#### Metodo: `teardown(self) -> None`
- **Que hace:** Cierra el navegador Chrome.
- **Por que existe:** Libera recursos del sistema.

#### Metodo: `_limpiar_crdownload_huerfanos(self, carpeta: str) -> None`
- **Que hace:** Elimina archivos `.crdownload` que quedaron de descargas anteriores incompletas.
- **Por que existe:** Limpia el directorio de descargas para evitar confusiones con archivos nuevos.

#### Metodo: `_navegar_a_crm(self) -> None`
- **Que hace:** Navega a la seccion CRM del ERP Fertrac.
- **Por que existe:** Los informes de ventas se encuentran en la seccion CRM, no en Inventario.

#### Metodo: `_abrir_menu_informes(self) -> bool`
- **Que hace:** Abre el menu "Informes" dentro de CRM.
- **Por que existe:** Accede al submenu donde estan los reportes de ventas.

#### Metodo: `_seleccionar_ventas_y_remisiones(self) -> bool`
- **Que hace:** Selecciona la opcion "Ventas y remisiones" del menu de informes.
- **Por que existe:** Es el tipo de informe que se necesita descargar.

#### Metodo: `_seleccionar_informe_facturas(self, max_intentos: int = 3) -> bool`
- **Que hace:** Selecciona el informe especifico de "Facturas" del campo de seleccion de informe.
- **Por que existe:** Dentro de ventas y remisiones, se necesita el reporte de facturas.

#### Metodo: `_configurar_fechas_ano_completo(self) -> bool`
- **Que hace:** Configura las fechas "Desde" y "Hasta" para cubrir todo el ano actual.
- **Por que existe:** Se necesita descargar todas las ventas del ano, no solo las de un mes.

#### Metodo: `_marcar_mostrar_costo(self) -> bool`
- **Que hace:** Marca el checkbox "Mostrar costo" para incluir costos en el informe.
- **Por que existe:** El informe debe incluir costos para poder hacer el procesamiento posterior.

#### Metodo: `_configurar_tipo_detallado(self) -> bool`
- **Que hace:** Selecciona el tipo de informe "Detallado" en el dropdown.
- **Por que existe:** Se necesita el formato detallado para tener toda la informacion necesaria.

#### Metodo: `_generar_xlsx(self) -> bool`
- **Que hace:** Hace click en el boton "Generar XLSX" para iniciar la exportacion.
- **Por que existe:** Es el paso final que genera el archivo Excel descargado.

---

## 3. `descarga_valorizados.py` (790 lineas)

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Anotaciones de tipo modernas |
| `glob` | modulo completo | Buscar archivos por patron |
| `os` | modulo completo | Variables de entorno, rutas |
| `shutil` | modulo completo | Copiar/mover archivos |
| `time` | modulo completo | Esperas entre acciones |
| `datetime` | `datetime` | Fecha para renombrar archivos |
| `selenium.webdriver.common.action_chains` | `ActionChains` | Acciones complejas del mouse |
| `selenium.webdriver.common.keys` | `Keys` | Teclas especiales |
| `selenium.webdriver.support` | `expected_conditions`, `WebDriverWait` | Esperas explicitas |
| `selenium.webdriver.support.ui` | `Select` | Dropdowns HTML |
| `config.erp_selectors` | `AlertaERP`, `Menu`, `Modal`, `Navegacion`, `Valorizado` | Selectores XPath para valorizados |
| `config.settings` | `Settings` | Configuracion centralizada |
| `core.browser` | `create_driver`, `do_login`, `wait_for_download` | Navegador y gestion de descargas |
| `core.email_notifier` | `EmailNotifier` | Notificaciones email |
| `core.erp_navigation` | `click_with_fallback`, `close_modal`, `descartar_alert_js`, `detect_and_accept_empty_alert`, `find_all_visible`, `find_visible_element`, `open_menu_by_selectors`, `wait_for_erp_section` | Navegacion compartida |
| `tasks.base_task` | `BaseTask` | Clase abstracta base |

### Clase: `DescargaValorizados(BaseTask)`

**Atributo de clase:**
- `name = "descarga_valorizados"` — nombre identificador de la task
- `ALMACENES_CONFIG` — lista de diccionarios con la configuracion de almacenes a descargar. Cada diccionario contiene: `nombre`, `tipo_consulta`, `ubicacion`, `almacen` (ej: TOBERIN → `3/Aforo Impo`, FALTANTES → `4/Faltantes`, FALTANTES IMPO → `7/Faltantes_Impo`, FALTANTES_2 → `4/Faltantes`)

**Selectores utilizados (clase `Valorizado` de `config/erp_selectors.py`):**
- `UBICACION_INPUT_NOMBRE` — input `<div name="location_id">` del modal (ubicacion a seleccionar)
- `AUTOCOMPLETE_OPCIONES` — items del autocompletado jQuery-UI (`<li id="ui-id-...">` y su `<a>`)
- `BUSCAR_MAS` — item "Buscar más..." del dropdown de ubicacion (abre el modal de busqueda)
- `MODAL_FILA_UBICACION` — fila del modal cuyo `<td>` contiene el valor de la ubicacion (usado por `_seleccionar_ubicacion_en_modal`)
- `MODAL_ABRIR` — tabla de lista del modal cuando esta visible ("Buscar más...")

#### Metodo: `__init__(self, settings: Settings)`
- **Que hace:** Inicializa la clase padre y configura las rutas de descarga para cada almacen.
- **Por que existe:** Necesita configurar rutas especificas para cada almacen antes de ejecutar.

#### Metodo: `setup(self) -> None`
- **Que hace:** Inicializa el notificador de email y crea las carpetas de destino para cada almacen.
- **Por que existe:** Prepara el entorno antes de iniciar las descargas.

#### Metodo: `execute(self) -> None`
- **Que hace:** Itera sobre cada almacen en `ALMACENES_CONFIG` y descarga su valorizado correspondiente. Para cada almacen: crea navegador, hace login, navega a inventario, abre menu de informes, selecciona valorizado, configura tipo consulta y ubicacion, genera XLSX y mueve el archivo.
- **Por que existe:** Descarga los valorizados de todos los almacenes configurados.

#### Metodo: `teardown(self) -> None`
- **Que hace:** Cierra el navegador Chrome si esta abierto.
- **Por que existe:** Libera recursos del sistema.

#### Metodo: `_enviar_email_exito(self, archivos_descargados: list[str], archivos_fallidos: list[str] | None = None, tiempo_total: str = "") -> None`
- **Que hace:** Envia un email de exito con el resumen de archivos descargados y fallidos.
- **Por que existe:** Notifica al usuario sobre el resultado de la descarga de valorizados.

#### Metodo: `navegar_a_inventario(self, driver) -> None`
- **Que hace:** Navega a la seccion de Inventario del ERP.
- **Por que existe:** Accede a la pagina donde se encuentran los valorizados.

#### Metodo: `abrir_menu_informes(self, driver) -> bool`
- **Que hace:** Abre el menu de "Informes" dentro de Inventario.
- **Por que existe:** Accede al submenu de reportes.

#### Metodo: `seleccionar_valorizado(self, driver, max_intentos: int = 3) -> bool`
- **Que hace:** Selecciona la opcion "Valorizado" del menu de informes.
- **Por que existe:** Es el tipo de reporte que se necesita.

#### Metodo: `seleccionar_tipo_consulta(self, driver, tipo_consulta: str) -> bool`
- **Que hace:** Selecciona el tipo de consulta (ej: "General", "Por almacen") en el dropdown del modal.
- **Por que existe:** Configura el tipo de valorizado a generar.

#### Metodo: `seleccionar_ubicacion_dropdown(self, driver, nombre_ubicacion: str) -> bool`
- **Que hace:** Selecciona la ubicacion (ej: "3/Aforo Impo", "7/Faltantes_Impo") en el modal de tipo de consulta. Prioriza el input especifico `location_id` (`//div[@name="location_id"]//input[contains(@id, "o_field_input")]`); si el autocomplete no ofrece la opcion exacta, busca el item "Buscar más..." y abre el modal para seleccionar la fila. Acepta alertas JS antes de interactuar para evitar el cuelgue de Selenium.
- **Por que existe:** Filtra los datos por ubicacion. El autocomplete del ERP a veces solo ofrece "Buscar más..." en lugar de la opcion directa, requiriendo seleccion via modal.
- **Detalle:** usa `_texto_opcion` para leer el texto del `<li>` o su `<a>` hijo normalizado (espacios/acentos), y `_seleccionar_ubicacion_en_modal` para localizar la fila en el modal por el valor.

#### Metodo: `_texto_opcion(self, opcion) -> str`
- **Que hace:** Retorna el texto visible normalizado de una opcion del dropdown de autocompletado. Si el `<li>` no tiene texto directo, lee su hijo `<a>`. Normaliza multiples espacios.
- **Por que existe:** Las opciones jQuery-UI del ERP guardan el texto en un `<li>` con un `<a>` hijo; leer solo `element.text` devolvia cadenas vacias y nunca coincidian.

#### Metodo: `_seleccionar_ubicacion_en_modal(self, driver, nombre_ubicacion: str) -> bool`
- **Que hace:** Espera a que la tabla del modal "Buscar más..." este visible y hace click en la fila cuyo `<td>` contiene el valor de la ubicacion buscada (ej: `<td class="o_data_cell o_readonly_modifier">3/Aforo Impo</td>`), usando `click_with_fallback`.
- **Por que existe:** El autocomplete no siempre ofrece la opcion directa; para ubicaciones como "7/Faltantes_Impo" solo aparece "Buscar más...", que abre un modal de seleccion. Este metodo resuelve la seleccion dentro del modal.

#### Metodo: `seleccionar_almacen_dropdown(self, driver, nombre_almacen: str) -> bool`
- **Que hace:** Selecciona el almacen especifico en el dropdown del modal.
- **Por que existe:** Filtra los datos por almacen.

#### Metodo: `_crear_valorizado_vacio(self, carpeta_destino: str, nombre_archivo: str) -> str | None`
- **Que hace:** Crea un archivo Excel vacio cuando no hay datos para un almacen.
- **Por que existe:** Mantiene la consistencia del proceso aunque no haya datos.

#### Metodo: `generar_xlsx(self, driver) -> bool`
- **Que hace:** Hace click en el boton "Generar XLSX" para iniciar la exportacion.
- **Por que existe:** Genera el archivo Excel con los datos del valorizado.

#### Metodo: `renombrar_y_mover_archivo(self, archivo_original: str, nuevo_nombre: str, carpeta_destino: str) -> str`
- **Que hace:** Renombra el archivo descargado y lo mueve a la carpeta de destino del almacen.
- **Por que existe:** Organiza los archivos descargados por almacen.

---

## 4. `actualizacion_inventario.py` (1604 lineas)

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Anotaciones de tipo modernas |
| `contextlib` | modulo completo | Context managers (with statements) |
| `gc` | modulo completo | Garbage collection para liberar memoria |
| `io` | modulo completo | Streams en memoria (BytesIO) |
| `os` | modulo completo | Variables de entorno, rutas |
| `re` | modulo completo | Expresiones regulares para limpiar datos |
| `shutil` | modulo completo | Copiar/mover/eliminar archivos |
| `sys` | modulo completo | Acceso al sistema |
| `tempfile` | modulo completo | Archivos temporales |
| `time` | modulo completo | Medir tiempos, esperas |
| `warnings` | modulo completo | Suprimir advertencias |
| `datetime` | `date`, `datetime` | Fechas para nombres de archivos |
| `pathlib` | `Path` | Manejo moderno de rutas |
| `msoffcrypto` | modulo completo | Desencriptar archivos Excel |
| `numpy` | `np` | Operaciones numericas |
| `pandas` | `pd` | Manipulacion de DataFrames |
| `unidecode` | `unidecode` | Convertir caracteres especiales a ASCII |
| `config.settings` | `Settings` | Configuracion centralizada |
| `core.email_notifier` | `EmailNotifier` | Notificaciones email |
| `core.excel_processing` | `HAS_COM`, `com_convert_to_xlsx`, `decrypt_to_stream_local`, `is_encrypted_xlsx`, `limpiar_archivo_temporal`, `limpiar_copias_temporales_antiguas`, `obtener_archivo_trabajo`, `safe_close_workbook`, `safe_quit_excel`, `save_bytesio_to_temp`, `verificar_archivo_disponible` (y `convert_xls_legacy`/`convert_xls_via_xlrd` via import perezoso) | Funciones compartidas de COM/msoffcrypto |
| `core.excel_utils` | `decrypt_to_stream`, `find_file_by_pattern`, `find_sheet_by_pattern`, `norm`, `norm_colname`, `norm_label`, `norm_sheet`, `read_excel_any`, `strip_dolares_temporales`, `write_excel` | Utilidades Excel (decrypt, read, write) |
| `core.logger` | `get_logger` | Logging unificado |
| `tasks.base_task` | `BaseTask` | Clase abstracta base |

### Clase auxiliar: `EliminacionTracker`

**Proposito:** Rastrea y registra todos los registros eliminados durante el proceso de actualizacion de inventario, generando un reporte detallado.

#### Metodo: `__init__(self)`
- **Que hace:** Inicializa las listas de eliminaciones por cada paso del proceso.
- **Por que existe:** Mantener un registro completo de que datos se eliminaron y por que.

#### Metodo: `registrar(self, paso, fila_excel, referencia, nombre, marca, linea, motivo, datos_extra)`
- **Que hace:** Registra una eliminacion con todos sus detalles (paso, fila, referencia, motivo, etc.).
- **Por que existe:** Para poder generar un reporte completo al final del proceso.

#### Metodo: `log_eliminacion(self, paso, fila_excel, referencia, motivo, mostrar_en_consola)`
- **Que hace:** Registra una eliminacion en el log y opcionalmente la muestra en consola.
- **Por que existe:** Para mantener un registro en el archivo de log.

#### Metodo: `mostrar_resumen(self)`
- **Que hace:** Muestra un resumen de todas las eliminaciones realizadas.
- **Por que existe:** Para que el usuario vea un resumen rapido de lo que se elimino.

#### Metodo: `generar_reporte_excel(self, base_path: Path) -> Path`
- **Que hace:** Genera un archivo Excel con el reporte detallado de eliminaciones.
- **Por que existe:** Para documentar formalmente que datos se eliminaron y por que.

### Clase principal: `ActualizacionInventario(BaseTask)`

**Atributos de clase:**
- `name = "actualizacion_inventario"` — nombre identificador
- `description = ...` — descripcion de la tarea
- `TRACKER_ELIMINACIONES` — instancia global de `EliminacionTracker`

#### Metodo: `__init__(self, settings: Settings)`
- **Que hace:** Inicializa la clase padre y configura las rutas base del proyecto.
- **Por que existe:** Configura las rutas necesarias para encontrar los archivos de inventario.

#### Metodo: `setup(self)`
- **Que hace:** Inicializa el notificador de email y configura las variables globales de rutas.
- **Por que existe:** Prepara el entorno antes de ejecutar el procesamiento.

#### Metodo: `execute(self)`
- **Que hace:** Ejecuta todo el proceso de actualizacion de inventario: carga archivos Excel, procesa datos, aplica reglas de negocio, genera el consolidado actualizado.
- **Por que existe:** Es la tarea principal que contiene toda la logica de procesamiento.
- **Notas de log:** el error "No se encontro columna REFERENCIA" ahora incluye la **ruta completa** (y `resolve()`) del archivo para facilitar el diagnostico; lo mismo la linea "Plantilla de inventario".

### Funciones standalone principales:

| Funcion | Parametros | Descripcion |
|---------|-----------|-------------|
| `log(msg)` | `msg: str` | Funcion de logging global |
| `month_abbr_es(dt: date) -> str` | Fecha | Retorna abreviatura del mes en espanol |
| `to_num_str(x)` | Valor | Convierte un valor a string numerico |
| `limpiar_referencia(valor)` | str | Limpia y normaliza referencias de productos |
| `open_as_excel_source(path: Path, passwords=None)` | Ruta, contrasenas | Abre un archivo Excel como fuente de datos. Para `.xls` legacy convierte el archivo con **xlrd** (puro Python, sin Excel COM) y, si es una tabla HTML con extension `.xls`, usa `pd.read_html`; si el archivo es un OLE2 protegido (salida previa cifrada con `Elementry`/msoffcrypto) lo descifra con msoffcrypto y lo materializa como **archivo temporal** (devuelve siempre un Path, nunca un BytesIO, para que las lecturas repetidas de pandas no falle); solo usa COM (`com_convert_to_xlsx`) como ultimo recurso. Evita el error `Microsoft Excel no puede obtener acceso al archivo Temp\` cuando la tarea corre como servicio de Windows |
| `read_excel_header_at(path, sheet, header_row_visible, passwords)` | Path, hoja, fila, contrasenas | Lee un Excel usando una fila de cabecera. **Auto-detecta** la fila real de la cabecera: si la fila fija no contiene columnas tipo cabecera (referencia/codigo/nombre/descripcion), escanea las primeras 25 filas buscando la celda "Referencia" (habitualmente columna B) y re-lee con esa fila. Corrige el error "no encuentro columna REFERENCIA" cuando la cabecera no esta en la fila 2 |
| `cargar_inventario_actualizado(base_dir: Path) -> pd.DataFrame` | Directorio base | Carga el archivo de inventario actualizado |
| `cargar_valorizado(base_dir: Path, prefix: str) -> pd.DataFrame` | Directorio, prefijo | Carga archivos de valorizados |
| `cargar_matriz_usd(base_dir: Path) -> pd.DataFrame` | Directorio base | Carga la matriz de precios USD |
| `cargar_marcas(base_dir: Path) -> pd.DataFrame` | Directorio base | Carga el archivo de marcas propias. **Tolerante**: si no existe columna "REFERENCIA" (p. ej. `MARCAS.xlsx` que solo tiene la columna "MARCAS"), usa la columna cuyo nombre contenga "marca" como lista de referencias (`__REF_LISTA__`) |
| `cargar_distribucion(base_dir: Path) -> pd.DataFrame` | Directorio base | Carga el archivo de distribucion de matrices |
| `cargar_consolidado_remisiones(base_dir: Path) -> pd.DataFrame` | Directorio base | Carga el consolidado de remisiones |
| `actualizar_referencias_inventario_original(...)` | DataFrames | Actualiza las referencias del inventario original |
| `aplicar_reglas_marcas_propias(...)` | DataFrames | Aplica reglas especificas para marcas propias |
| `eliminar_registros_linea_copia_indeterminada(...)` | DataFrame, hoja, workbook | Elimina registros con linea indeterminada |
| `procesar_existencias_negativas_y_cero(...)` | DataFrame, worksheet | Procesa y ajusta existencias negativas o en cero |
| `excel_open(ruta, *, visible, alerts, editable)` | Ruta, opciones | Abre Excel via COM con configuracion especifica. Usa `win32.DispatchEx("Excel.Application")`; el import de `win32com.client as win32` esta **protegido** (si falla, `win32 = None`), evitando el error `NameError: name 'win32' is not defined` en equipos sin la libreria |
| `excel_close(excel, wb, *, save)` | Excel, workbook, guardar | Cierra Excel via COM de forma segura |
| `ws_headers(ws, *, start)` | Worksheet, fila inicio | Obtiene los headers de una hoja Excel |
| `ws_fill_column_values(...)` | Worksheet, opciones | Rellena valores en una columna especifica |

### Constantes principales:

| Constante | Valor | Descripcion |
|-----------|-------|-------------|
| `OUTPUT_BASENAME` | `"$2026 INVENTARIO GENERAL ACTUALIZADO"` | Nombre base del archivo de salida |
| `SHEET_INV_ORIG` | `"INVENTARIO"` | Nombre de la hoja original |
| `SHEET_INV_COPIA` | `"INVENTARIO COPIA"` | Nombre de la hoja de copia |
| `HEADER_ROW_INV` | `2` | Fila de headers en el inventario |
| `TRACKER_ELIMINACIONES` | `EliminacionTracker()` | Instancia global de tracking |

### Flujo de plantilla y guardado (FASE 2, 5 y 6):

- **FASE 2 (plantilla)**: `p_inv` se busca PRIMERO en `OUTPUT_PATH` (Pruebas Inv General) con `find_by_prefix(OUTPUT_PATH, OUTPUT_BASENAME)` → usa la salida previa `$2026 INVENTARIO GENERAL ACTUALIZADO {fecha}.xlsx` mas reciente como plantilla. Solo si no existe alguna, cae a `BASE_PATH` (carpeta del ERP). Esto evita usar el `.xls` crudo del ERP (hoja `Sheet 1`, 8 columnas) como plantilla, que producia una salida con estructura equivocada.
- **FASE 5 (COM)**: la plantilla previa suele estar **protegida con contrasena** (OLE2 cifrado). Se descifra con msoffcrypto (`[PASS_INV] + PASSWORDS_TRY`) a un archivo temporal `TEMP_INVENTARIO_*.xlsx` en `OUTPUT_PATH` (ruta accesible para Excel COM, NO `%TEMP%`) y se abre ESE temporal. Ademas `ws` se selecciona a la hoja `INVENTARIO` (no `Sheets(1)`), para que la copia de headers de `INVENTARIO COPIA` use los headers correctos (fila 2). El temporal se elimina al cerrar.
- **FASE 6 (guardado)**: el nombre de salida es `f"{OUTPUT_BASENAME} {AAAAMMDD_HHMM}.xlsx"`, p.ej. `$2026 INVENTARIO GENERAL ACTUALIZADO 20260828_2353.xlsx`. El reporte `REPORTE_ELIMINACIONES_{AAAAMMDD_HHMMSS}.xlsx` con segundos.

---

## 5. `actualizacion_ventas.py` (2434 lineas)

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Anotaciones de tipo modernas |
| `os` | modulo completo | Variables de entorno, rutas |
| `io` | modulo completo | Streams en memoria (BytesIO) |
| `re` | modulo completo | Expresiones regulares para normalizar datos |
| `difflib` | modulo completo | Coincidencia difusa de columnas |
| `warnings` | modulo completo | Suprimir advertencias |
| `contextlib` | modulo completo | Context managers |
| `time` | modulo completo | Medir tiempos, esperas |
| `pathlib` | `Path` | Manejo moderno de rutas |
| `datetime` | `datetime`, `date`, `timezone` | Fechas para configurar periodos |
| `dateutil.relativedelta` | `relativedelta` | Calculo de fechas relativas |
| `shutil` | modulo completo | Copiar/mover/eliminar archivos |
| `tempfile` | modulo completo | Archivos temporales |
| `pandas` | `pd` | Manipulacion de DataFrames |
| `numpy` | `np` | Operaciones numericas |
| `msoffcrypto` | modulo completo | Desencriptar archivos Excel |
| `unidecode` | `unidecode` | Convertir caracteres especiales a ASCII |
| `tasks.base_task` | `BaseTask` | Clase abstracta base |
| `core.logger` | `get_logger` | Logging unificado |
| `core.email_notifier` | `EmailNotifier` | Notificaciones email |
| `core.excel_processing` | `HAS_COM`, `safe_close_workbook`, `safe_quit_excel`, `excel_serial_from_date` | Funciones COM compartidas |
| `core.excel_utils` | `decrypt_to_stream`, `read_excel_any`, `find_sheet_name`, `norm`, `norm_simple`, `norm_colname`, `norm_label`, `norm_sheet`, `norm_base_filename`, `strip_dolares_temporales`, `extract_fecha_es`, `resolve_cols`, `write_excel` | Utilidades Excel compartidas |
| `config.settings` | `Settings`, `MESES_ES`, `MESES_ES_NOMBRE` | Configuracion y diccionarios de meses |

### Clase: `ActualizacionVentas(BaseTask)`

**Atributo de clase:**
- `name = "actualizacion_ventas"` — nombre identificador
- `COL_SYNONYMS` — diccionario de sinonimos de columnas para normalizar nombres

#### Metodo: `setup(self)`
- **Que hace:** Inicializa el notificador de email y configura las rutas del proyecto.
- **Por que existe:** Prepara los recursos antes de ejecutar el procesamiento.

#### Metodo: `execute(self)`
- **Que hace:** Ejecuta todo el proceso de actualizacion de ventas: carga archivos Excel de ventas, informe de ventas y matriz de clientes, procesa datos, aplica formulas, genera el consolidado actualizado.
- **Por que existe:** Es la tarea principal con toda la logica de negocio.

#### Propiedades principales:

| Propiedad | Que retorna | Para que sirve |
|-----------|-------------|----------------|
| `BASE_PATH` | `Path` | Ruta base del servidor |
| `PASSWORD_VENTAS` | `str` | Contrasena de archivos de ventas |
| `PASSWORD_INV_GENERAL` | `str` | Contrasena de archivos de inventario |
| `PASSWORD_MYR` | `str` | Contrasena de archivos MYR |
| `FN_VENTAS` | `str` | Nombre del archivo de ventas |
| `FN_INV_GENERAL` | `str` | Nombre del archivo de inventario |
| `FN_INFORME_VENTAS` | `str` | Nombre del informe de ventas |
| `FN_MATRIZ_CLIENTES` | `str` | Nombre de la matriz de clientes |
| `DIR_INFORME_VENTAS_MES` | `Path` | Directorio del informe de ventas del mes |
| `SHEET_VENTAS_OPTIONS` | `list` | Opciones de nombres de hojas de ventas |
| `SHEET_INV_GENERAL` | `str` | Nombre de la hoja de inventario |
| `SHEET_COSTOS_INVFINAL` | `str` | Nombre de la hoja de costos |
| `SHEET_MATRIZ` | `str` | Nombre de la hoja de matriz |
| `COLS_CLAVE_VENTAS` | `list` | Columnas clave para ventas |
| `COLS_INFORME` | `list` | Columnas del informe |
| `COLS_FORMULAS_PRE` | `list` | Columnas con formulas (antes) |
| `COLS_FORMULAS_POST` | `list` | Columnas con formulas (despues) |
| `COLS_MATRIZ_MAP` | `dict` | Mapeo de columnas de matriz |

#### Metodos principales:

| Metodo | Que hace | Para que sirve |
|--------|----------|----------------|
| `_crear_indicador_progreso()` | Crea un archivo temporal que indica que el proceso esta ejecutandose | Para que otros procesos sepan que esta corriendo |
| `_eliminar_indicador_progreso(indicador_path)` | Elimina el archivo de indicador | Limpieza al finalizar |
| `_extract_date_yyyyMMdd_from_tail(name_no_ext)` | Extrae fecha del nombre del archivo | Identificar archivos por fecha |
| `find_matriz_clientes_by_prefix(base_dir, prefixes)` | Busca el archivo de matriz de clientes | Encontrar el archivo correcto |
| `find_informe_facturas_by_prefix(base_dir, prefix, only_today)` | Busca el informe de facturas | Encontrar el archivo correcto |
| `find_myr_existencia_by_fecha(base_dir, prefix_expected)` | Busca el archivo MYR | Encontrar el archivo correcto |
| `_decrypt_to_stream(xlsx_path, password)` | Desencripta un archivo Excel | Leer archivos protegidos |
| `_read_excel_any(xlsx, **kwargs)` | Lee archivos Excel en cualquier formato | Leer datos de archivos |
| `_write_excel(df_or_dict, path_out, sheetname)` | Escribe datos a Excel | Guardar resultados |
| `transformar_informe_ventas(df)` | Transforma el informe de ventas al formato requerido | Preparar datos para procesamiento |
| `normalizar_nombre_columna(nombre)` | Normaliza nombres de columnas | Estandarizar encabezados |
| `normalizar_columnas_df(df)` | Normaliza todas las columnas del DataFrame | Preparar datos para merge |
| `integrar_linea_sublinea(df_ventas, df_inv, COLS_NORM)` | Integra informacion de linea y sublinea | Enriquecer datos de ventas |
| `integrar_costo_factor_hoy(df_ventas, df_myr, COLS_NORM)` | Integra costos del MYR | Agregar costos a ventas |
| `limpiar_linea_referencias_invalidas(df_ventas, COLS_NORM)` | Limpia referencias invalidas | Eliminar datos corruptos |
| `normalizar_dctos(df_ventas)` | Normaliza descuentos | Estandarizar valores |
| `ordenar_y_fechas(df_ventas, COLS_NORM)` | Ordena y configura fechas | Organizar datos temporalmente |
| `actualizar_periodo_mes(wb, primer_dia, ultimo_dia)` | Actualiza el periodo del mes en el Excel | Configurar fechas en el reporte |
| `calcular_y_escribir_subtotales(...)` | Calcula y escribe subtotales | Generar resumen por grupo |
| `convertir_tablas_a_rango(wb)` | Convierte tablas dinamicas a rangos | Facilitar manipulacion |
| `com_write_df_into_template(...)` | Escribe DataFrame en una plantilla Excel | Generar archivo final formateado |
| `actualizar_tabla_dinamica_resumen_dia(...)` | Actualiza la tabla dinamica de resumen | Refrescar datos consolidados |

---

## 6. `envio_informe_ventas.py`

### Imports

| Modulo | Que importa | Para que sirve |
|--------|-------------|----------------|
| `__future__` | `annotations` | Anotaciones de tipo modernas |
| `glob` | modulo completo | Buscar el archivo `$2026 VENTAS*.xlsx` mas reciente |
| `os`, `re`, `shutil`, `socket`, `subprocess`, `threading`, `time` | modulos completos | Rutas, limpieza de procesos, puertos, esperas y lanzamiento de Chrome con timeout |
| `datetime` | `date`, `datetime` | Fecha/mes actual para mensajes y logs |
| `pathlib` | `Path` | Manejo moderno de rutas |
| `selenium.*` | `webdriver`, `By`, `Keys`, `WebDriverWait`, `expected_conditions`, `Service`, `TimeoutException` | Automatizacion de WhatsApp Web |
| `win32com.client` | `win32` | Abrir el Excel de ventas (COM) con contrasena |
| `pywintypes` | modulo completo | Detectar errores COM reintentables |
| `win32clipboard` | modulo completo | Copiar imagen al portapapeles (CF_DIB) para WhatsApp |
| `PIL` | `Image`, `ImageGrab` | Capturar el clipboard y convertir a BMP |
| `psutil` | modulo completo | Matar procesos Chrome residuales (opcional) |
| `config.settings` | `Settings` | Configuracion centralizada |
| `core.email_notifier` | `EmailNotifier` | Notificaciones por email |
| `core.logger` | `get_logger` | Logging unificado del proyecto (`logs/{fecha}/envio_informe_ventas.log`) |
| `tasks.base_task` | `BaseTask` | Clase abstracta base |

### Constantes principales

| Constante | Valor | Descripcion |
|-----------|-------|-------------|
| `PROJECT_ROOT` | raiz del repo | Carpeta raiz del proyecto |
| `DIR_IMAGENES` | `PROJECT_ROOT/Img informe` | Carpeta destino de las imagenes recortadas |
| `PROFILE_DIR` | `PROJECT_ROOT/chrome_profile_whatsapp_session` | Perfil persistente de Chrome (mantiene la sesion de WhatsApp, evita re-escanear QR) |
| `CAPTURAS` | lista de 4 dicts | Configuracion de capturas (ver abajo) |
| `_DEFAULT_CHATS` | `["Informe ventas diarias"]` | Destinatarios por defecto (override: variable de entorno `WHATSAPP_CHATS` separada por comas) |

**Configuracion de capturas (`CAPTURAS`):**

| Nombre | Hoja | Tipo | Rango/parametros | Mensaje (caption) |
|--------|------|------|-------------------|-------------------|
| `foto1_vta_dia` | `Resumen Dia` | dinamico | fila_inicio 9, A→D (filtra por mes actual, refresca pivot, oculta columna de margen) | VTA DIA |
| `foto2_resumen_mes` | `Resum Mes` | dinamico | fila_inicio 11, A→G | RESUMEN MES |
| `foto3_resumen_meta` | `Resumen Meta 2026` | fijo | `A1:R14` | RESUMEN META |
| `foto4_resumen_dia_vendedor` | `Resumen XDiaVend INF` | dia_dinamico | fila_fechas 8, fila_fin 29 (oculta dias anteriores al actual) | RESUMEN DIA POR VENDEDOR |

### Clase: `EnvioInformeVentas(BaseTask)`

**Atributo de clase:**
- `name = "envio_informe_ventas"` — nombre identificador (usa el logger del proyecto)

**Metodo: `setup(self)`**
- Inicializa `notifier` (SMTP ventas), rutas, `excel_cfg` (contrasena `Compras2028`), `ruta_base = paths.base/Pruebas` (donde `actualizacion_ventas` guarda su salida), destinatarios y `DIR_IMAGENES`.

**Metodo: `execute(self)`**
1. Busca con `glob("$2026 VENTAS*.xlsx")` el archivo mas reciente en `<base>/Pruebas`.
2. Llama a `capturar_multiples_rangos(...)` (Excel COM con contrasena) y guarda las imagenes en `Img informe` (**no las borra**; el proceso heredado las eliminaba tras el envio, ahora quedan persistidas).
3. Arma el mensaje `Cordial saludo. Se remite VTAS {MES} MG NETO PONDERADO {factor}`.
4. Lanza WhatsApp Web (Selenium + perfil persistente + chromedriver local/cache/webdriver-manager), envia mensaje + imagenes a cada destinatario y cuenta exitosos.
5. Si ningun destinatario recibio el informe → lanza excepcion (la task falla y notifica por email).

**Metodo: `teardown(self)`**
- Cierra el navegador de WhatsApp si quedo abierto (se ejecuta siempre, exito o error).

### Funciones standalone principales

| Funcion | Que hace |
|---------|----------|
| `WhatsAppWeb` (clase) | Cliente Selenium de WhatsApp Web (`iniciar_sesion`, `buscar_chat`, `cerrar_dialogos`, `enviar_texto`, `enviar_imagen`, `enviar_reporte_completo`, `cerrar`) con selectores resilientes y limpieza de procesos |
| `com_con_reintentos(...)` | Ejecuta llamadas COM con reintentos frente a los errores `RPC_E_CALL_REJECTED` (-2147418111) y `RPC_E_SERVERCALL_RETRYLATER` (-2147417848) |
| `matar_todos_chrome()` | Mata Chrome/ChromeDriver y espera a que terminen (taskkill + psutil) |
| `capturar_multiples_rangos(ruta, capturas_config, password, dir_imagenes)` | Abre el Excel con `Workbooks.Open(..., password, password)`, espera 30 s de carga, aplica filtro/refresco de pivots en `Resumen Dia`, extrae el factor MG NETO de la primera captura, describe el rango por tipo (`dinamico`/`fijo`/`dia_dinamico`), copia al portapapeles (`CopyPicture`) y lo guarda en `dir_imagenes` via `ImageGrab.grabclipboard()`. Retorna `(imagenes, factor)` |
| `filtrar_pivot_por_mes_actual(ws)` | Filtra la tabla dinamica de `Resumen Dia` por el mes actual (PageFields con fallback a PivotItems) |
| `ocultar_dias_anteriores_y_obtener_rango(ws, fila_fechas, fila_fin)` | Oculta columnas de dias anteriores al actual y devuelve el rango `A{fila_fechas}:{total}{fila_fin}`; `None` si no esta el dia actual (omite la captura) |
| `extraer_factor_mg_neto(ws, fila_inicio, columna_ref, columna_factor)` | Lee el factor MG NETO PONDERADO de la fila "Total general" |
| `_ocultar_columna_margen(ws)` | Oculta temporalmente la columna `Promedio de MARGEN NETO (DCTO PIE FACT)` de `Resumen Dia` antes del pantallazo y la restaura |

### Notas de integracion

- **Depende de** `descarga_ventas` + `actualizacion_ventas`: se ejecuta con el workflow `ventas` (`python run.py --workflow ventas`, ahora = descarga → actualizacion → envio) o standalone (`python run.py --task envio_informe_ventas`).
- **Contrasena**: usa `settings.excel.password` (default `Compras2028`), igual que `actualizacion_ventas`.
- **WhatsApp Web**: requiere sesion iniciada en el perfil persistente (primera vez hay que escanear el QR). El perfil y el `chromedriver.exe` se movieron de la carpeta `Enviar Informe` a la raiz del proyecto; la carpeta `Enviar Informe` (proceso heredado) fue eliminada.
- **Imagenes**: se guardan en `Img informe` y se conservan (no se eliminan tras el envio).

---

## Resumen de dependencias entre tasks

```
base_task.py (clase abstracta)
    ├── descarga_inventario_general.py (scraping - Selenium)
    ├── descarga_informe_ventas.py (scraping - Selenium)
    ├── descarga_valorizados.py (scraping - Selenium)
    ├── actualizacion_inventario.py (procesamiento - COM/pandas)
    ├── actualizacion_ventas.py (procesamiento - COM/pandas)
    └── envio_informe_ventas.py (envio - COM/PIL + Selenium WhatsApp Web)
```

### Modulos compartidos utilizados:

| Modulo | Utilizado por | Funciones principales |
|--------|---------------|----------------------|
| `core/browser.py` | 3 tasks de scraping | `create_driver`, `do_login`, `wait_for_download` |
| `core/erp_navigation.py` | 3 tasks de scraping | `open_menu_by_selectors`, `find_visible_element`, `click_with_fallback` |
| `config/erp_selectors.py` | 3 tasks de scraping | `Navegacion`, `Menu`, `Paginador`, `Exportar`, `Modal` |
| `core/excel_processing.py` | 2 tasks de procesamiento | `safe_close_workbook`, `safe_quit_excel`, `com_convert_to_xlsx` (con reintentos, copia si origen bloqueado y `taskkill` de EXCEL.EXE huerfano como ultimo recurso), `convert_xls_legacy`/`convert_xls_via_xlrd` (conversion de `.xls` legacy con xlrd/read_html sin depender de Excel COM) |
| `core/excel_utils.py` | 2 tasks de procesamiento | `decrypt_to_stream`, `read_excel_any`, `norm`, `write_excel` |
