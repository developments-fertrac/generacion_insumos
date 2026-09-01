# Generacion de Insumos

Pipeline automatizado para descargar informes del ERP Fertrac y generar insumos consolidados de Inventario y Ventas.

## Arquitectura

```
Generacion de Insumos/
├── .env                          # Credenciales y rutas (no se commitea)
├── .env.example                  # Template de variables de entorno
├── requirements.txt              # Dependencias Python
├── run.py                        # Punto de entrada CLI
├── orchestrator.py               # Orquestador de pipelines
├── config/
│   ├── settings.py               # Configuracion centralizada (lee .env)
│   └── erp_selectors.py          # Selectores XPath centralizados (14 clases)
├── core/
│   ├── logger.py                 # Logger unificado con rotacion diaria
│   ├── email_notifier.py         # Servicio de email (SMTP)
│   ├── browser.py                # ChromeDriver compartido (Selenium)
│   ├── excel_utils.py            # Utilidades Excel (decrypt, read, write)
│   ├── erp_navigation.py         # Navegacion Selenium compartida (12 funciones)
│   └── excel_processing.py       # COM/msoffcrypto compartido (12 funciones)
├── tasks/
│   ├── base_task.py              # Clase base abstracta (Template Method)
│   ├── descarga_valorizados.py   # Descarga valorizados por almacen
│   ├── descarga_inventario_general.py  # Descarga inventario general
│   ├── descarga_informe_ventas.py      # Descarga informe de ventas
│   ├── actualizacion_inventario.py     # Procesa y consolida inventario
│   ├── actualizacion_ventas.py         # Procesa y consolida ventas
│   └── envio_informe_ventas.py         # Recorta resumenes y envia por WhatsApp Web
├── Img informe/                  # Imagenes recortadas para el envio por WhatsApp
├── chromedriver.exe             # ChromeDriver para WhatsApp Web
├── chrome_profile_whatsapp_session/  # Sesion persistente de WhatsApp (evita re-escanear QR)
└── logs/                         # Logs unificados por dia
    └── YYYY-MM-DD/
        ├── descarga_valorizados.log
        ├── descarga_inv_general.log
        ├── descarga_ventas.log
        ├── actualizacion_inventario.log
        ├── actualizacion_ventas.log
        └── envio_informe_ventas.log
```

### Patron de diseno: Template Method

Cada tarea extiende `BaseTask` y solo implementa su logica de negocio:

```python
class MiTarea(BaseTask):
    name = "mi_tarea"

    def setup(self):
        # Inicializar notificador, rutas, etc.
        self.notifier = EmailNotifier(self.settings.smtp, self.name)

    def execute(self):
        # Logica principal de la tarea

    def teardown(self):
        # Limpieza (cerrar navegador, etc.)
```

El ciclo `run()` de `BaseTask` se encarga automaticamente de:
1. Configurar el logger
2. Ejecutar `setup()` -> `execute()` -> `teardown()`
3. Enviar notificacion de exito o fallo por email
4. Registrar tiempos en el log

### Capas

| Capa | Responsabilidad |
|------|-----------------|
| `config/` | Carga de `.env`, estructura de configuracion tipada |
| `core/` | Funcionalidad compartida: logging, email, Selenium, Excel |
| `tasks/` | Logica de negocio de cada tarea |
| `orchestrator.py` | Orquestacion de pipelines y fases |
| `run.py` | CLI para ejecutar pipelines o tareas individuales |

## Ejecucion

### Requisitos previos

```bash
pip install -r requirements.txt
```

Se requiere Python 3.10+ y Chrome instalado en el servidor.

### Pipeline completo

Ejecuta todas las tareas en orden:

```bash
python run.py
```

**Flujo:**
```
Fase 1 (secuencial - scraping web):
  descarga_inv_general -> descarga_ventas -> descarga_valorizados

Fase 2 (paralelo - procesamiento local):
  actualizacion_inv + actualizacion_ventas
```

### Workflows por area

Solo ventas (descarga + actualizacion + envio del informe por WhatsApp Web):
```bash
python run.py --workflow ventas
```

**Flujo del workflow ventas:**
```
Fase 1: descarga_ventas (scraping ERP -> informe de ventas)
Fase 2: actualizacion_ventas (consolida -> $2026 VENTAS_{fecha}.xlsx en <base>/Pruebas)
Fase 3: envio_informe_ventas (recorta resumenes -> Img informe -> WhatsApp Web)
```

Solo inventario (descarga inv general + descarga valorizados + actualizacion):
```bash
python run.py --workflow inventario
```

### Tarea individual

Ejecutar una sola tarea:
```bash
python run.py --task descarga_ventas
python run.py --task descarga_inv_general
python run.py --task descarga_valorizados
python run.py --task actualizacion_inv
python run.py --task actualizacion_ventas
python run.py --task envio_informe_ventas
```

### Listar opciones

```bash
python run.py --list
```

## Configuracion

### Variables de entorno (`.env`)

Editar el archivo `.env` en la raiz del proyecto. Todos los valores se cargan desde ahi:

```env
# Credenciales ERP Fertrac
FERTRAC_USER=consultas
FERTRAC_PASS=tu_password

# Credenciales SMTP (Gmail App Password)
SMTP_PASSWORD=tu_app_password

# Passwords de archivos Excel protegidos
EXCEL_PASSWORD=Compras2028
EXCEL_PASSWORDS_TRY=Compras2027,Compras2028

# Rutas base del servidor
BASE_PATH=D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026
REMISIONES_BASE=D:\Fertrac\Usuarios\infocompras\$CARPETA COMPRAS 2026\REMISIONES
```

### Cambiar password del ERP

Actualizar `FERTRAC_PASS` en `.env`. El cambio aplica a las 3 tareas de descarga automaticamente.

### Cambiar password SMTP

Actualizar `SMTP_PASSWORD` en `.env`. Aplica a todas las tareas que envian emails.

### Cambiar destinatarios de email

Editar en `config/settings.py` las clases `SmtpConfigVentas` o `SmtpConfigInvGeneral`:

```python
class SmtpConfigVentas(SmtpConfig):
    sender_email = "ctorrese@fertrac.com"
    recipient_emails = (
        "analista_automatizacion@fertrac.com",
        "data_science@fertrac.com",
    )
```

### Cambiar timeouts del navegador

Editar en `config/settings.py` la clase `BrowserConfig`:

```python
class BrowserConfig:
    headless = True                # False para ver el navegador
    timeout_descarga = 60          # Segundos esperando un archivo
    timeout_carga_maxima = 3600    # 60 min para carga de inventario
    timeout_descarga_completa = 1500  # 25 min para informe de ventas
    periodo_gracia = 180           # 3 min adicionales tras timeout
    max_intentos_driver = 5        # Reintentos al iniciar Chrome
    max_intentos_login = 3         # Reintentos de login
```

### Cambiar rutas de archivos

Las rutas se derivan de `BASE_PATH` en `.env`. Para cambiar la estructura de carpetas, editar las propiedades en `PathsConfig` (ver `config/settings.py`).

### Desactivar emails

No hay flag global. Cada notificador usa `SmtpConfig.enabled`. Para desactivar, cambiar a `False` en la clase correspondiente de `config/settings.py`.

## Envio del informe por WhatsApp Web

La tarea `envio_informe_ventas` recorta los resumenes del archivo `$2026 VENTAS_{fecha}.xlsx` (salida de `actualizacion_ventas`), los guarda en `Img informe/` y los envia por WhatsApp Web con Selenium.

### Primera ejecucion (escaneo QR)

La sesion se guarda en `chrome_profile_whatsapp_session/`; solo la primera vez (o si se borra el perfil) hay que escanear el codigo QR con el celular en la ventana de Chrome que se abre.

```bash
python run.py --task envio_informe_ventas
```

### Configurar destinatarios de WhatsApp

Por defecto envia a `Informe ventas diarias`. Para cambiarlos sin tocar codigo, definir en el entorno la variable `WHATSAPP_CHATS` separada por comas (p. ej. en `.env`):

```env
WHATSAPP_CHATS=Informe ventas diarias,Pruebas
```

### Notas

- Depende de la libreria `pillow` (ya incluida en `requirements.txt`) y de `pywin32` (Windows).
- El archivo de ventas se busca en `<base>/Pruebas` con el patron `$2026 VENTAS*.xlsx`.
- Las imagenes capturadas se conservan en `Img informe/` (no se eliminan tras el envio).
- El envío usa el perfil `chrome_profile_whatsapp_session/` y el `chromedriver.exe` del raiz del proyecto.

## Logs

Los logs se guardan en `logs/YYYY-MM-DD/{nombre_tarea}.log` con formato:

```
[2026-08-25 08:00:12] [INFO   ] [descarga_ventas] ChromeDriver iniciado
[2026-08-25 08:05:30] [ERROR  ] [descarga_inv_general] Timeout al cargar productos
[2026-08-25 09:12:00] [SUCCESS] [actualizacion_ventas] Proceso completado
```

- Se crea una carpeta por dia automaticamente
- Los logs de cada tarea van en un archivo separado
- El log de consola muestra INFO+, el archivo guarda DEBUG+

## Agregar una nueva tarea

1. Crear `tasks/mi_nueva_tarea.py`
2. Extender `BaseTask`:

```python
from tasks.base_task import BaseTask
from core.email_notifier import EmailNotifier

class MiNuevaTarea(BaseTask):
    name = "mi_nueva_tarea"

    def setup(self):
        self.notifier = EmailNotifier(self.settings.smtp, self.name)

    def execute(self):
        self.log.info("Ejecutando mi tarea...")
        # logica aqui
```

3. Registrar en `tasks/__init__.py`:

```python
from tasks.mi_nueva_tarea import MiNuevaTarea

TASK_REGISTRY = {
    ...,
    "mi_nueva_tarea": MiNuevaTarea,
}
```

4. Ejecutar: `python run.py --task mi_nueva_tarea`
