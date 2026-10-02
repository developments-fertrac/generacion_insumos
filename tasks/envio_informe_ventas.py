"""
Envio del informe de ventas diario por WhatsApp Web.

Toma el archivo mas reciente '$2026 VENTAS*.xlsx' generado por
`actualizacion_ventas` (guardado en <base>/Pruebas), recorta las imagenes
de los resumenes (Resumen Dia, Resum Mes, Resumen Meta 2026 y
Resumen XDiaVend INF), las guarda en la carpeta 'Img informe' y las envia
por WhatsApp Web a los destinatarios configurados.

Adaptado del proceso heredado 'Enviar Informe' (enviar_excel_whatsapp.py
VERSION 5.6) a la estructura de tareas del proyecto.
"""
from __future__ import annotations

import csv
import glob
import io
import os
import re
import shutil
import socket
import subprocess
import threading
import time
import traceback
from datetime import date, datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException

from core.email_notifier import EmailNotifier
from core.logger import get_logger, _LOGS_ROOT
from tasks.base_task import BaseTask

try:
    import win32com.client as win32
    import pywintypes
    HAS_PYWINTYPES = True
except ImportError:
    win32 = None
    pywintypes = None
    HAS_PYWINTYPES = False

try:
    import psutil
    PSUTIL_DISPONIBLE = True
except ImportError:
    psutil = None
    PSUTIL_DISPONIBLE = False

try:
    import win32clipboard
    from PIL import Image, ImageGrab
    HAS_PIL = True
except ImportError:
    win32clipboard = None
    Image = None
    ImageGrab = None
    HAS_PIL = False

_LOG = get_logger("envio_informe_ventas")


def _log(msg: str) -> None:
    _LOG.info("%s", msg)


# ===== CONFIGURACION =====
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _state_dir() -> Path:
    """Estado operativo fuera del repositorio (Phase 0).

    Configurable con la variable de entorno STATE_DIR.
    """
    configurado = os.environ.get("STATE_DIR", "").strip()
    if configurado:
        return Path(configurado)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "GeneracionInsumos"


STATE_DIR = _state_dir()
DIR_IMAGENES = STATE_DIR / "Img informe"
PROFILE_DIR = STATE_DIR / "chrome_profile_whatsapp_session"

_DEFAULT_CHATS = ["Informe ventas diarias"]


def _get_destinatarios() -> list[str]:
    raw = os.environ.get("WHATSAPP_CHATS", "").strip()
    if raw:
        return [c.strip() for c in raw.split(",") if c.strip()]
    return list(_DEFAULT_CHATS)


CAPTURAS = [
    {
        "nombre": "foto1_vta_dia",
        "hoja": "Resumen Dia",
        "tipo": "dinamico",
        "fila_inicio": 9,
        "columna_inicio": "A",
        "columna_fin": "D",
        "mensaje": "VTA DIA",
    },
    {
        "nombre": "foto2_resumen_mes",
        "hoja": "Resum Mes",
        "tipo": "dinamico",
        "fila_inicio": 11,
        "columna_inicio": "A",
        "columna_fin": "G",
        "mensaje": "RESUMEN MES",
    },
    {
        "nombre": "foto3_resumen_meta",
        "hoja": "Resumen Meta 2026",
        "tipo": "fijo",
        "rango": "A1:R14",
        "mensaje": "RESUMEN META",
    },
    {
        "nombre": "foto4_resumen_dia_vendedor",
        "hoja": "Resumen XDiaVend INF",
        "tipo": "dia_dinamico",
        "fila_fechas": 8,
        "fila_fin": 29,
        "mensaje": "RESUMEN DIA POR VENDEDOR",
    },
]


# ===== HELPERS DE LOGGING Y PORTAPAPELES =====
def _log_traceback(mensaje: str, error: BaseException | None = None) -> None:
    """Registra el error en INFO y el traceback completo en DEBUG."""
    _log(f"{mensaje}: {error}")
    _LOG.debug(traceback.format_exc())


def _procesos_por_nombre(nombre: str) -> list[dict]:
    """Lista los procesos con ese nombre usando tasklist (CSV): pid/sesión/memoria."""
    resultados = []
    try:
        salida = subprocess.run(
            ['tasklist', '/FI', f'IMAGENAME eq {nombre}', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=15,
        ).stdout
        for fila in csv.DictReader(io.StringIO(salida)):
            resultados.append({
                'pid': fila.get('PID', '?'),
                'sesion': fila.get('Session Name', '?'),
                'sesion_id': fila.get('Session#', '?'),
                'memoria': fila.get('Mem Usage', '?'),
            })
    except Exception:
        pass
    return resultados


def diagnosticar_portapapeles(momento: str = "") -> None:
    """Intenta enumerar los formatos del portapapeles para detectar si es accesible."""
    if not HAS_PIL:
        _log("   ⚠️ PIL no disponible; no se puede leer el portapapeles")
        return
    try:
        import win32clipboard as wc
        wc.OpenClipboard()
        try:
            formatos = []
            fmt = 0
            while True:
                fmt = wc.EnumClipboardFormats(fmt)
                if not fmt:
                    break
                formatos.append(fmt)
            if formatos:
                _log(f"   📋 Portapapeles accesible{momento} - formatos: {formatos}")
            else:
                _log(f"   ⚠️ Portapapeles accesible{momento} pero vacío (sin formatos)")
        finally:
            wc.CloseClipboard()
    except Exception as e:
        _log(f"   ❌ Portapapeles NO accesible{momento}: {e}")
        _log("   (Suele pasar cuando el proceso corre en una sesión no interactiva)")


# ===== REINTENTOS PARA LLAMADAS COM =====
def com_con_reintentos(funcion, *args, max_intentos=5, espera_base=2, **kwargs):
    """
    Ejecuta una funcion COM con reintentos automaticos.
    El error -2147418111 (RPC_E_CALL_REJECTED) significa que Excel esta ocupado.
    """
    errores_reintentables = [-2147418111, -2147417848]

    for intento in range(1, max_intentos + 1):
        try:
            return funcion(*args, **kwargs)
        except Exception as e:
            es_com_error = False
            if HAS_PYWINTYPES and isinstance(e, pywintypes.com_error):
                es_com_error = e.args[0] in errores_reintentables
            if es_com_error and intento < max_intentos:
                espera = espera_base * intento
                _log(f"   ⚠️ Excel COM ocupado (intento {intento}/{max_intentos}). "
                     f"Reintentando en {espera}s...")
                time.sleep(espera)
            else:
                raise


# ===== FUNCIONES DE LIMPIEZA Y VERIFICACIÓN =====
def matar_todos_chrome():
    """Mata TODOS los procesos de Chrome/ChromeDriver y espera a que terminen."""
    _log("   🔪 Matando procesos Chrome/ChromeDriver...")

    try:
        subprocess.run(['taskkill', '/F', '/IM', 'chromedriver.exe', '/T'],
                       capture_output=True, timeout=15)
        subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe', '/T'],
                       capture_output=True, timeout=15)
    except Exception as e:
        _log(f"   ⚠️ Error con taskkill: {e}")

    if PSUTIL_DISPONIBLE:
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    nombre = (proc.info['name'] or '').lower()
                    if 'chrome' in nombre or 'chromedriver' in nombre:
                        proc.kill()
                except Exception:
                    pass
        except Exception:
            pass

    _log("   ⏳ Esperando que procesos terminen...")
    for segundo in range(1, 11):
        time.sleep(1)
        hay_chrome = False
        if PSUTIL_DISPONIBLE:
            for proc in psutil.process_iter(['name']):
                try:
                    nombre = (proc.info['name'] or '').lower()
                    if 'chrome' in nombre:
                        hay_chrome = True
                        break
                except Exception:
                    pass
        else:
            try:
                result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq chrome.exe'],
                                        capture_output=True, text=True, timeout=5)
                if 'chrome.exe' in result.stdout.lower():
                    hay_chrome = True
            except Exception:
                pass
        if not hay_chrome:
            _log(f"   ✓ Procesos terminados ({segundo}s)")
            return
        if segundo % 3 == 0:
            _log(f"   ⏳ Aún hay procesos Chrome... ({segundo}s)")

    _log("   ⚠️ Algunos procesos pueden seguir activos")


def verificar_puerto_disponible(puerto: int) -> bool:
    """Verifica si un puerto está disponible."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            return s.connect_ex(('127.0.0.1', puerto)) != 0
    except Exception:
        return True


def limpiar_perfil_selectivo(ruta_perfil: str):
    """Limpia archivos problematicos SIN eliminar la sesión de WhatsApp."""
    if not os.path.exists(ruta_perfil):
        return

    _log("   🧹 Limpieza selectiva (preservando sesión WA)...")

    eliminar = [
        'lockfile', 'SingletonLock', 'SingletonSocket', 'SingletonCookie',
        'Cache', 'Code Cache', 'GPUCache', 'ShaderCache', 'GrShaderCache',
        'BrowserMetrics', 'Crashpad', 'FileTypePolicies',
        os.path.join('Default', 'Cache'),
        os.path.join('Default', 'Code Cache'),
        os.path.join('Default', 'GPUCache'),
        os.path.join('Default', 'Service Worker'),
        os.path.join('Default', 'LOCK'),
        os.path.join('Default', 'LOG'),
        os.path.join('Default', 'LOG.old'),
        os.path.join('Default', 'Network'),
        os.path.join('Default', 'blob_storage'),
    ]

    eliminados = 0
    for item in eliminar:
        ruta_item = os.path.join(ruta_perfil, item)
        try:
            if os.path.isfile(ruta_item):
                os.remove(ruta_item)
                eliminados += 1
            elif os.path.isdir(ruta_item):
                shutil.rmtree(ruta_item)
                eliminados += 1
        except Exception:
            pass

    if eliminados > 0:
        _log(f"   ✓ {eliminados} elementos eliminados (sesión WA preservada)")


# ===== CLIENTE WHATSAPP WEB =====
class WhatsAppWeb:
    """Cliente para WhatsApp Web vía Selenium (adaptado de VERSION 5.5/5.6)."""

    def __init__(self, sesion_nombre="whatsapp_session", max_reintentos=3):
        self.sesion_nombre = sesion_nombre
        self.max_reintentos = max_reintentos
        self.driver = None
        self.wait = None
        self.puerto_debug = None
        self.perfil_dir = str(STATE_DIR / f"chrome_profile_{sesion_nombre}")

    def _obtener_chromedriver(self):
        """
        Obtiene un ChromeDriver compatible con el Chrome instalado.

        AUTO-RECUPERACIÓN: si el ChromeDriver local (raíz del proyecto) no
        coincide con la versión mayor de Chrome instalado, se busca/descarga
        el driver correcto y se reemplaza el local obsoleto, evitando el error:
          "session not created: This version of ChromeDriver only supports
           Chrome version X / Current browser version is Y".
        """
        from core.chromedriver_utils import (
            chromedriver_mayor,
            obtener_chromedriver_path,
        )

        path = obtener_chromedriver_path(reemplazar_local=True)
        if path and os.path.exists(path):
            mayor = chromedriver_mayor(path)
            _log(f"   ✓ ChromeDriver: {os.path.basename(path)}"
                 + (f" ({mayor})" if mayor else ""))
            return Service(path)

        # Ultima opcion: ChromeDriver del PATH
        _log("   ⚠️ Usando ChromeDriver del PATH")
        return None

    def _configurar_opciones_chrome(self):
        """Configura las opciones de Chrome para WhatsApp Web."""
        from selenium.webdriver.chrome.options import Options

        chrome_options = Options()
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument(f"--remote-debugging-port={self.puerto_debug}")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-background-networking")
        chrome_options.add_argument("--no-first-run")
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--disable-session-crashed-bubble")
        chrome_options.add_argument("--disable-restore-session-state")
        chrome_options.add_argument("--hide-crash-restore-bubble")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        prefs = {
            "profile.default_content_setting_values.notifications": 2,
            "credentials_enable_service": False,
        }
        chrome_options.add_experimental_option("prefs", prefs)

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
        for path in chrome_paths:
            if os.path.exists(path):
                chrome_options.binary_location = path
                _log(f"   ✓ Chrome: {path}")
                break

        # Perfil persistente
        chrome_options.add_argument(f"user-data-dir={self.perfil_dir}")
        return chrome_options

    def iniciar_sesion(self):
        """Inicia WhatsApp Web con reintentos."""
        ultimo_error = None

        for intento in range(1, self.max_reintentos + 1):
            try:
                _log(f"\n{'='*60}")
                _log(f"⏳ Iniciando WhatsApp Web... (Intento {intento}/{self.max_reintentos})")
                _log(f"{'='*60}")

                # 1. LIMPIEZA
                _log("🔧 Paso 1: Limpieza previa...")
                matar_todos_chrome()
                if intento >= 2:
                    limpiar_perfil_selectivo(self.perfil_dir)

                # 2. PUERTO
                _log("🔧 Paso 2: Verificando puerto...")
                self.puerto_debug = None
                for puerto in range(9222, 9250):
                    if verificar_puerto_disponible(puerto):
                        self.puerto_debug = puerto
                        break
                if self.puerto_debug is None:
                    self.puerto_debug = 9222
                    _log(f"   ⚠️ Usando puerto por defecto: {self.puerto_debug}")
                else:
                    _log(f"   ✓ Puerto libre: {self.puerto_debug}")

                # 3. DIRECTORIO DE PERFIL
                if not os.path.exists(self.perfil_dir):
                    os.makedirs(self.perfil_dir)
                    _log("   📁 Perfil creado")

                # 4. CHROMEDRIVER
                _log("🔧 Paso 3: Obteniendo ChromeDriver...")
                service = self._obtener_chromedriver()
                if service is None:
                    _log("   ❌ No se encontró ChromeDriver")
                    raise FileNotFoundError("ChromeDriver no encontrado")

                # 5. OPCIONES
                _log("🔧 Paso 4: Configurando Chrome...")
                chrome_options = self._configurar_opciones_chrome()

                # 6. INICIAR CHROME CON TIMEOUT
                _log("🚀 Paso 5: Iniciando Chrome...")
                driver_iniciado = [False]
                driver_error = [None]

                def iniciar_driver():
                    try:
                        self.driver = webdriver.Chrome(service=service, options=chrome_options)
                        driver_iniciado[0] = True
                    except Exception as e:
                        driver_error[0] = e

                thread = threading.Thread(target=iniciar_driver)
                thread.start()
                thread.join(timeout=300)

                if not driver_iniciado[0]:
                    if driver_error[0]:
                        raise driver_error[0]
                    raise TimeoutError("Chrome no respondió en 90 segundos")

                _log("   ✓ Chrome iniciado correctamente")
                self.wait = WebDriverWait(self.driver, 60)

                # 7. NAVEGAR A WHATSAPP
                _log("🌐 Paso 6: Navegando a WhatsApp Web...")
                self.driver.get("https://web.whatsapp.com")

                _log("🔧 Cerrando popups de Chrome...")
                time.sleep(2)
                try:
                    self.driver.execute_script("""
                        let buttons = document.querySelectorAll('button');
                        for (let btn of buttons) {
                            let text = btn.textContent || btn.innerText || '';
                            if (text.includes('Cerrar') || text.includes('Close') ||
                                text.includes('Dismiss') || text.includes('No gracias') ||
                                text.includes('Restaurar')) {
                                try { btn.click(); } catch(e) {}
                            }
                        }
                    """)
                    _log("   ✓ Popups cerrados")
                except Exception as e:
                    _log(f"   ⚠️ Error cerrando popups: {e}")

                # 8. ESPERAR CARGA
                _log("🔍 Paso 7: Esperando carga...")
                _log("   (Si es primera vez, escanea el código QR)")

                selectores_wa_listo = [
                    (By.CSS_SELECTOR, 'input[role="textbox"][data-tab="3"]'),
                    (By.XPATH, '//input[@role="textbox"]'),
                    (By.CSS_SELECTOR, '#side'),
                    (By.CSS_SELECTOR, '#app div[data-tab]'),
                    (By.XPATH, '//*[@role="textbox"][@data-tab="3"]'),
                    (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'),
                ]

                def _detectar_wa_cargado(driver):
                    for by, selector in selectores_wa_listo:
                        try:
                            el = driver.find_element(by, selector)
                            if el and el.is_displayed():
                                return el
                        except Exception:
                            continue
                    return False

                try:
                    self.wait.until(_detectar_wa_cargado)
                    _log("✅ WhatsApp Web listo")
                    time.sleep(3)
                    return True
                except TimeoutException:
                    _log("⏳ Extendiendo espera para QR...")
                    extended_wait = WebDriverWait(self.driver, 120)
                    try:
                        extended_wait.until(_detectar_wa_cargado)
                        _log("✅ WhatsApp Web listo")
                        time.sleep(3)
                        return True
                    except Exception:
                        _log("   📋 Diagnóstico: "
                             f"title='{self.driver.title}', url='{self.driver.current_url}'")
                        raise TimeoutException("No se cargó WhatsApp Web")

            except Exception as e:
                ultimo_error = e
                _log(f"❌ Error: {str(e)[:200]}")
                if self.driver:
                    try:
                        self.driver.quit()
                    except Exception:
                        pass
                    self.driver = None
                _log("   🧹 Limpieza post-error...")
                matar_todos_chrome()
                espera = 10 * intento
                _log(f"   ⏳ Esperando {espera}s antes de reintentar...")
                time.sleep(espera)

        _log(f"\n{'='*60}")
        _log("❌ NO SE PUDO INICIAR WHATSAPP WEB")
        _log(f"{'='*60}")
        _log(f"Error: {ultimo_error}")
        return False

    def cerrar_dialogos(self):
        """Detecta y cierra dialogos/popups abiertos en WhatsApp Web."""
        try:
            dialogos = self.driver.find_elements(By.CSS_SELECTOR, '[role="dialog"]')
            dialogos_visibles = [d for d in dialogos if d.is_displayed()]
            if not dialogos_visibles:
                return False

            _log("   ⚠️ Diálogo/popup detectado, intentando cerrar...")
            from selenium.webdriver.common.action_chains import ActionChains

            selectores_cancelar = [
                '//button[contains(text(), "Cancelar")]',
                '//div[contains(text(), "Cancelar")][@role="button"]',
                '//button[contains(text(), "Cancel")]',
                '//div[@role="button"][contains(., "Cancelar")]',
            ]
            for selector in selectores_cancelar:
                try:
                    boton = self.driver.find_element(By.XPATH, selector)
                    if boton.is_displayed():
                        boton.click()
                        time.sleep(1)
                        dialogos_post = self.driver.find_elements(By.CSS_SELECTOR, '[role="dialog"]')
                        if not [d for d in dialogos_post if d.is_displayed()]:
                            _log("   ✅ Diálogo cerrado con botón Cancelar")
                            return True
                except Exception:
                    continue

            for _intento in range(5):
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)
                dialogos_post = self.driver.find_elements(By.CSS_SELECTOR, '[role="dialog"]')
                if not [d for d in dialogos_post if d.is_displayed()]:
                    _log("   ✅ Diálogo cerrado con Escape")
                    return True

            try:
                ActionChains(self.driver).move_by_offset(10, 10).click().perform()
                ActionChains(self.driver).move_by_offset(-10, -10).perform()
                time.sleep(1)
                dialogos_post = self.driver.find_elements(By.CSS_SELECTOR, '[role="dialog"]')
                if not [d for d in dialogos_post if d.is_displayed()]:
                    _log("   ✅ Diálogo cerrado con click fuera")
                    return True
            except Exception:
                pass

            _log("   ⚠️ No se pudo cerrar el diálogo, continuando de todas formas...")
            return False
        except Exception:
            return False

    def buscar_chat(self, nombre_chat):
        """Busca un chat o grupo por nombre."""
        try:
            _log(f"🔍 Buscando chat: {nombre_chat}")
            self.cerrar_dialogos()
            time.sleep(0.5)

            search_box = None
            selectores_busqueda = [
                '//input[@role="textbox"][@data-tab="3"]',
                '//input[contains(@aria-label, "Buscar")]',
                '//input[contains(@aria-label, "Search")]',
                '//input[contains(@placeholder, "Buscar")]',
                '//input[@role="textbox"]',
                '//div[@contenteditable="true"][@data-tab="3"]',
                '//div[@role="textbox"][@data-tab="3"]',
            ]
            for selector in selectores_busqueda:
                try:
                    search_box = self.wait.until(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                    if search_box:
                        _log(f"   ✓ Caja de búsqueda encontrada con: {selector}")
                        break
                except Exception:
                    continue

            if not search_box:
                _log("❌ No se encontró la caja de búsqueda")
                return False

            search_box.click()
            time.sleep(0.5)
            search_box.send_keys(Keys.CONTROL + "a")
            time.sleep(0.2)
            search_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.3)
            try:
                search_box.clear()
            except Exception:
                pass
            time.sleep(0.3)

            search_box.send_keys(nombre_chat)
            time.sleep(3)

            selectores_chat = [
                f'//span[@title="{nombre_chat}"]',
                f'//span[contains(@title, "{nombre_chat}")]',
                f'//span[text()="{nombre_chat}"]',
            ]
            chat_result = None
            for selector in selectores_chat:
                try:
                    chat_result = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except Exception:
                    continue

            if chat_result:
                chat_result.click()
                _log(f"✅ Chat '{nombre_chat}' encontrado")
                time.sleep(2)
                return True
            else:
                _log(f"⚠️ No se encontró el chat '{nombre_chat}'")
                return False
        except Exception as e:
            _log(f"❌ Error al buscar chat: {e}")
            return False

    def enviar_texto(self, nombre_chat, mensaje):
        """Envía mensaje de texto a un chat o grupo."""
        try:
            if not self.buscar_chat(nombre_chat):
                return False

            selectores_mensaje = [
                '//div[@role="textbox"][@contenteditable="true"][@data-tab="10"]',
                '//div[contains(@aria-label, "Escribir un mensaje")][@contenteditable="true"]',
                '//div[contains(@aria-label, "Type a message")][@contenteditable="true"]',
                '//div[@contenteditable="true"][@data-tab="10"]',
                '//footer//div[@contenteditable="true"]',
                '//div[@id="main"]//footer//div[@contenteditable="true"]',
                '//div[@id="main"]//div[@role="textbox"][@contenteditable="true"]',
            ]
            message_box = None
            for selector in selectores_mensaje:
                try:
                    message_box = self.wait.until(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                    break
                except Exception:
                    continue

            if not message_box:
                _log("❌ No se encontró la caja de mensajes")
                return False

            lineas = mensaje.split('\n')
            for i, linea in enumerate(lineas):
                message_box.send_keys(linea)
                if i < len(lineas) - 1:
                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
            message_box.send_keys(Keys.ENTER)
            _log(f"✅ Mensaje enviado a '{nombre_chat}'")
            time.sleep(2)
            return True
        except Exception as e:
            _log(f"❌ Error al enviar texto: {e}")
            return False

    def enviar_imagen(self, nombre_chat, ruta_imagen, caption=""):
        """Envía imagen usando clipboard + Ctrl+V."""
        try:
            if not self.buscar_chat(nombre_chat):
                return False

            _log(f"   📎 Preparando imagen: {os.path.basename(ruta_imagen)}")
            ruta_absoluta = os.path.abspath(ruta_imagen)
            if not os.path.exists(ruta_absoluta):
                _log(f"   ❌ Archivo no existe: {ruta_absoluta}")
                return False

            import io
            imagen = Image.open(ruta_absoluta)
            _log(f"   📐 Tamaño: {imagen.size[0]}x{imagen.size[1]} px")

            output = io.BytesIO()
            imagen.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            imagen.close()

            _log("   📋 Copiando al clipboard...")
            clipboard_ok = False
            for intento_cb in range(3):
                try:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    time.sleep(0.3)
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                    win32clipboard.CloseClipboard()
                    clipboard_ok = True
                    break
                except Exception as e:
                    _log(f"   ⚠️ Intento clipboard {intento_cb + 1}/3: {e}")
                    try:
                        win32clipboard.CloseClipboard()
                    except Exception:
                        pass
                    time.sleep(1)

            if not clipboard_ok:
                _log("   ❌ No se pudo copiar al clipboard")
                return False

            time.sleep(0.5)
            _log("   ⌨️ Pegando imagen...")

            selectores_mensaje = [
                '//div[@role="textbox"][@contenteditable="true"][@data-tab="10"]',
                '//div[contains(@aria-label, "Escribir un mensaje")][@contenteditable="true"]',
                '//div[contains(@aria-label, "Type a message")][@contenteditable="true"]',
                '//div[@contenteditable="true"][@data-tab="10"]',
                '//footer//div[@contenteditable="true"]',
                '//div[@id="main"]//footer//div[@contenteditable="true"]',
                '//div[@role="textbox"][@contenteditable="true"]',
            ]
            message_box = None
            for selector in selectores_mensaje:
                try:
                    message_box = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, selector))
                    )
                    if message_box and message_box.is_displayed():
                        _log(f"   ✓ Caja de mensajes encontrada: {selector}")
                        break
                    message_box = None
                except Exception:
                    continue

            if not message_box:
                _log("   ❌ No se encontró caja de mensajes")
                return False

            self.cerrar_dialogos()
            message_box.click()
            time.sleep(0.5)

            actions = ActionChains(self.driver)
            actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

            _log("   ⌨️ Imagen pegada, esperando carga...")
            time.sleep(5)

            if caption:
                _log(f"   ✏️ Escribiendo caption: {caption}")
                caption_escrito = False
                try:
                    selectores_caption = [
                        '//div[contains(@data-testid,"media-caption")]',
                        '//div[contains(@aria-label,"Añade una leyenda")]',
                        '//div[contains(@aria-label,"Add a caption")]',
                        '//div[contains(@aria-label,"caption")]',
                        '//div[contains(@aria-label,"Caption")]',
                        '//div[@contenteditable="true"][@data-testid="media-caption"]',
                        '//div[@contenteditable="true"]//following::div[@data-testid="media-caption"]',
                        '//footer//div[@contenteditable="true"]',
                        '//div[@role="textbox"][@contenteditable="true"][@data-tab="10"]',
                    ]
                    caption_box = None
                    for selector in selectores_caption:
                        try:
                            caption_box = WebDriverWait(self.driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, selector))
                            )
                            if caption_box and caption_box.is_displayed():
                                _log(f"   ✓ Caja de caption encontrada: {selector}")
                                break
                            caption_box = None
                        except Exception:
                            continue

                    if caption_box:
                        caption_box.click()
                        time.sleep(0.5)
                        caption_box.send_keys(caption)
                        caption_escrito = True
                        _log("   ✓ Caption escrito")
                        time.sleep(1)
                except Exception as e:
                    _log(f"   ⚠️ Error buscando caja de caption: {e}")

                if not caption_escrito:
                    _log("   ⚠️ No se encontró caja de caption, escribiendo con Ctrl+Shift...")
                    try:
                        message_box = self.driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
                        message_box.click()
                        time.sleep(0.3)
                        actions = ActionChains(self.driver)
                        actions.send_keys(caption).perform()
                        caption_escrito = True
                        _log("   ✓ Caption escrito (fallback)")
                        time.sleep(1)
                    except Exception as e:
                        _log(f"   ⚠️ Error en fallback de caption: {e}")
            else:
                _log("   ℹ️ No hay caption para esta imagen")

            _log("   📨 Buscando botón de enviar...")
            enviado = False
            selectores_send = [
                '//span[@data-icon="send"]',
                '//div[@aria-label="Enviar"]',
                '//button[@aria-label="Enviar"]',
                '//span[@data-testid="send"]',
                '//div[@data-testid="send"]',
                '//button[@data-testid="send"]',
            ]
            for selector in selectores_send:
                try:
                    boton_send = self.driver.find_element(By.XPATH, selector)
                    if boton_send.is_displayed():
                        boton_send.click()
                        enviado = True
                        _log("   ✓ Click en botón Enviar ejecutado")
                        break
                except Exception:
                    continue

            if not enviado:
                _log("   ⚠️ No se encontró botón, intentando con Enter...")
                try:
                    actions = ActionChains(self.driver)
                    actions.send_keys(Keys.ENTER).perform()
                    _log("   ✓ Enter enviado")
                except Exception as e:
                    _log(f"   ⚠️ Error con Enter: {e}")

            time.sleep(2)
            try:
                dialog = self.driver.find_element(By.XPATH, '//div[@role="dialog"]')
                if dialog.is_displayed():
                    _log("   ⚠️ Diálogo aún abierto, enviando de nuevo...")
                    actions = ActionChains(self.driver)
                    actions.send_keys(Keys.ENTER).perform()
                    time.sleep(2)
            except Exception:
                pass

            _log("✅ Imagen enviada correctamente")
            time.sleep(1)
            self.cerrar_dialogos()
            return True
        except Exception as e:
            _log(f"❌ Error al enviar imagen: {e}")
            try:
                actions = ActionChains(self.driver)
                actions.send_keys(Keys.ESCAPE).perform()
            except Exception:
                pass
            return False

    def enviar_reporte_completo(self, nombre_chat, mensaje_inicial, imagenes_config):
        """Envía reporte completo: mensaje inicial + imágenes."""
        try:
            _log(f"\n{'='*70}")
            _log(f"📤 ENVIANDO REPORTE A: {nombre_chat}")
            _log(f"{'='*70}\n")

            _log("📤 Enviando mensaje inicial...")
            if not self.enviar_texto(nombre_chat, mensaje_inicial):
                _log("⚠️ Error en mensaje inicial, continuando...")

            time.sleep(3)

            exitosos = 0
            for idx, img_config in enumerate(imagenes_config, 1):
                _log(f"\n📤 Imagen {idx}/{len(imagenes_config)}: {img_config['caption']}")
                if img_config['ruta'] is None:
                    _log("   ⚠️ No disponible, saltando...")
                    continue
                for intento in range(1, 4):
                    if intento > 1:
                        _log(f"   🔄 Reintento {intento}/3...")
                        time.sleep(5)
                    if self.enviar_imagen(nombre_chat, img_config['ruta'], img_config['caption']):
                        exitosos += 1
                        break
                self.cerrar_dialogos()
                time.sleep(3)

            _log(f"\n📊 Resultado: {exitosos}/{len(imagenes_config)} imágenes")
            return exitosos > 0
        except Exception as e:
            _log(f"❌ Error: {e}")
            return False

    def cerrar(self):
        """Cierra el navegador de forma segura."""
        if self.driver:
            _log("🔒 Cerrando WhatsApp Web...")
            try:
                self.driver.quit()
                _log("✅ Navegador cerrado")
            except Exception as e:
                _log(f"⚠️ Error al cerrar: {e}")
            finally:
                self.driver = None
                self.wait = None
        time.sleep(1)


# ===== FUNCIONES AUXILIARES =====
def obtener_mes_actual_espanol():
    """Retorna el mes actual en español."""
    meses = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }
    return meses[datetime.now().month]


def filtrar_pivot_por_mes_actual(ws):
    """
    Filtra la tabla dinamica de la hoja por el mes actual.
    Busca en PageFields (campos de filtro de pagina) y fallbacks.
    """
    try:
        mes_actual = obtener_mes_actual_espanol()
        _log(f"   🗓️ Filtrando tabla dinámica por mes: {mes_actual}")

        pivot_count = ws.PivotTables().Count
        if pivot_count == 0:
            _log("   ⚠️ No se encontraron tablas dinámicas en la hoja")
            return False

        pivot_table = ws.PivotTables(1)
        _log(f"   📊 Tabla dinámica encontrada: {pivot_table.Name}")

        campo_mes = None
        nombres_campo = ["MES", "Mes", "mes", "MONTH", "Month", "MES NOMBRE", "Mes Nombre"]

        # METODO 1: PageFields
        try:
            page_fields = pivot_table.PageFields
            for i in range(1, page_fields.Count + 1):
                field = page_fields.Item(i)
                field_name = str(field.Name).strip()
                if field_name.upper() in [n.upper() for n in nombres_campo]:
                    campo_mes = field
                    _log(f"   ✓ Campo MES encontrado en PageFields: '{field_name}'")
                    break
        except Exception as e:
            _log(f"   ⚠️ Error buscando en PageFields: {e}")

        # METODO 2: PivotFields
        if campo_mes is None:
            try:
                all_fields = pivot_table.PivotFields()
                for i in range(1, all_fields.Count + 1):
                    try:
                        field = all_fields.Item(i)
                        field_name = str(field.Name).strip()
                        if field_name.upper() in [n.upper() for n in nombres_campo]:
                            campo_mes = field
                            _log(f"   ✓ Campo MES encontrado: '{field_name}'")
                            break
                    except Exception:
                        continue
            except Exception:
                pass

        # METODO 3: Acceso directo por nombre
        if campo_mes is None:
            for field_name in nombres_campo:
                try:
                    campo_mes = pivot_table.PivotFields(field_name)
                    _log(f"   ✓ Campo encontrado con acceso directo: '{field_name}'")
                    break
                except Exception:
                    continue

        if campo_mes is None:
            _log("   ❌ No se encontró el campo 'MES' en la tabla dinámica")
            return False

        _log(f"   🎯 Aplicando filtro: '{mes_actual}'")

        # METODO A: CurrentPage
        try:
            campo_mes.CurrentPage = mes_actual
            _log(f"   ✓ Filtro aplicado con CurrentPage: {mes_actual}")
            time.sleep(2)
            return True
        except Exception as e:
            _log(f"   ⚠️ CurrentPage falló: {e}")

        # METODO B: PivotItems
        try:
            items_disponibles = []
            for item in campo_mes.PivotItems():
                items_disponibles.append(str(item.Name))

            mes_encontrado = None
            for item_name in items_disponibles:
                item_upper = str(item_name).upper().strip()
                if mes_actual.upper() == item_upper or mes_actual.upper() in item_upper:
                    mes_encontrado = item_name
                    break

            if mes_encontrado is None:
                _log(f"   ❌ No se encontró '{mes_actual}' en los items: {items_disponibles}")
                return False

            for item in campo_mes.PivotItems():
                if str(item.Name) == mes_encontrado:
                    item.Visible = True
            for item in campo_mes.PivotItems():
                item_name = str(item.Name)
                if item_name != mes_encontrado:
                    item_upper = item_name.upper()
                    if "TODAS" not in item_upper and "ALL" not in item_upper and "BLANK" not in item_upper:
                        try:
                            item.Visible = False
                        except Exception:
                            pass

            _log(f"   ✓ Filtro aplicado con PivotItems: {mes_encontrado}")
            time.sleep(2)
            return True
        except Exception as e:
            _log(f"   ❌ Error con PivotItems: {e}")
            return False
    except Exception as e:
        _log(f"   ❌ Error al filtrar por mes: {e}")
        return False


def encontrar_ultima_fila(ws, columna_referencia, fila_inicio):
    """Encuentra la ultima fila con datos."""
    ultima_fila = fila_inicio
    while True:
        celda_valor = ws.Range(f"{columna_referencia}{ultima_fila}").Value
        if celda_valor is None:
            return ultima_fila - 1
        if isinstance(celda_valor, str) and "Total general" in celda_valor:
            return ultima_fila
        ultima_fila += 1
        if ultima_fila > 1000:
            return fila_inicio


def ocultar_dias_anteriores_y_obtener_rango(ws, fila_fechas=9, fila_fin=29):
    """
    Oculta columnas de dias anteriores al dia actual y retorna el rango.
    Mantiene visible: columna A + dia actual + Total general.
    """
    try:
        dia_actual = datetime.now().day
        _log(f"   📅 Día actual: {dia_actual}")

        columna_dia_actual = None
        columna_total = None
        columnas_a_ocultar = []

        for col_num in range(2, 27):
            col_letter = chr(64 + col_num)
            try:
                celda_valor = ws.Range(f"{col_letter}{fila_fechas}").Value
                if celda_valor is None:
                    continue

                celda_str = str(celda_valor).strip()
                if "Total general" in celda_str or "Total" in celda_str:
                    columna_total = col_letter
                    break

                dia_celda = None
                if isinstance(celda_valor, datetime):
                    dia_celda = celda_valor.day
                elif HAS_PYWINTYPES and isinstance(celda_valor, pywintypes.TimeType):
                    dia_celda = celda_valor.day
                elif hasattr(celda_valor, "day"):
                    dia_celda = celda_valor.day
                elif "/" in celda_str:
                    try:
                        dia_celda = int(celda_str.split("/")[0].strip())
                    except ValueError:
                        continue
                elif "-" in celda_str:
                    try:
                        dia_celda = int(celda_str.split("-")[0].strip())
                    except ValueError:
                        continue
                elif isinstance(celda_valor, (int, float)):
                    try:
                        from datetime import timedelta
                        fecha_excel = datetime(1899, 12, 30) + timedelta(days=celda_valor)
                        dia_celda = fecha_excel.day
                    except Exception:
                        continue

                if dia_celda is None:
                    continue

                if dia_celda == dia_actual:
                    columna_dia_actual = col_letter
                    _log(f"   🎯 ¡ENCONTRADO! Columna día actual ({dia_actual}): {col_letter}")
                elif dia_celda < dia_actual:
                    columnas_a_ocultar.append(col_letter)
            except Exception as e:
                _log(f"      ⚠️ Error procesando columna {col_letter}: {e}")
                continue

        if columna_dia_actual is None:
            _log(f"   ❌ No se encontró columna del día actual ({dia_actual})")
            return None

        if columnas_a_ocultar:
            _log(f"   🙈 Ocultando {len(columnas_a_ocultar)} columnas de días anteriores...")
            for col in columnas_a_ocultar:
                try:
                    ws.Columns(col).Hidden = True
                except Exception as e:
                    _log(f"      ⚠️ Error ocultando columna {col}: {e}")

        if columna_total is None:
            _log("   ⚠️ No se encontró columna 'Total general'")
            return None

        rango = f"A{fila_fechas}:{columna_total}{fila_fin}"
        _log(f"   ✓ Rango de captura: {rango}")
        return rango
    except Exception as e:
        _log(f"   ❌ Error en ocultar_dias_anteriores_y_obtener_rango: {e}")
        return None


def extraer_factor_mg_neto(ws, fila_inicio, columna_referencia, columna_factor):
    """Extrae el factor MG NETO PONDERADO."""
    try:
        fila_total = encontrar_ultima_fila(ws, columna_referencia, fila_inicio)
        valor = ws.Range(f"{columna_factor}{fila_total}").Value
        if valor is not None:
            if isinstance(valor, (int, float)):
                porcentaje = valor * 100
                return f"{porcentaje:.2f}%".replace(".", ",")
            return str(valor)
        return "N/A"
    except Exception as e:
        _log(f"⚠️ Error al extraer factor: {e}")
        return "N/A"


def _ocultar_columna_margen(ws):
    """Oculta temporalmente la columna 'Promedio de MARGEN NETO' en Resumen Dia."""
    from unidecode import unidecode

    col_oculta = None
    try:
        _log("   🙈 Buscando columna 'Promedio de MARGEN NETO' para ocultar...")

        def _norm_pivot_col(s):
            t = unidecode(str(s)).lower()
            t = re.sub(r"[^a-z0-9 ]", " ", t)
            t = re.sub(r"\s+", " ", t).strip()
            return t

        nombre_buscar = _norm_pivot_col("Promedio de MARGEN NETO (DCTO PIE FACT)")
        used_cols = ws.UsedRange.Columns.Count
        used_rows = min(20, ws.UsedRange.Rows.Count)

        for r in range(1, used_rows + 1):
            for c in range(1, used_cols + 1):
                val = ws.Cells(r, c).Value
                if val and nombre_buscar in _norm_pivot_col(str(val)):
                    col_oculta = c
                    ws.Columns(c).Hidden = True
                    _log(f"   ✅ Columna {c} ('{val}') ocultada temporalmente")
                    break
            if col_oculta:
                break

        if not col_oculta:
            _log("   ⚠️ No se encontró la columna 'Promedio de MARGEN NETO'")
    except Exception as e:
        _log(f"   ⚠️ Error al ocultar columna de margen: {e}")
    return col_oculta


# ===== CAPTURA DE RANGOS (solo portapapeles) =====

def _capturar_rango(excel, ws, rango: str, captura: dict, dir_imagenes: Path):
    """
    Captura un rango como PNG usando CopyPicture + portapapeles.
    Requiere sesión interactiva con escritorio visible.
    Devuelve la ruta PNG o None.
    """
    nombre = captura['nombre']
    dir_imagenes.mkdir(parents=True, exist_ok=True)
    ruta_png = dir_imagenes / f"{nombre}.png"
    _log(f"   📂 Destino imagen: {ruta_png}")

    rango_obj = ws.Range(rango)
    rango_obj.Select()
    time.sleep(1)

    if not HAS_PIL:
        _log("   ❌ Sin PIL no hay método de portapapeles disponible")
        return None

    _log("   📸 CopyPicture + portapapeles...")
    diagnosticar_portapapeles(" (antes de copiar)")
    copiado = False
    try:
        rango_obj.CopyPicture(Appearance=1, Format=2)
        _log("   ✓ CopyPicture ejecutado correctamente")
        copiado = True
    except Exception as e:
        _log_traceback("   ⚠️ Error con CopyPicture (rango)", e)
        try:
            excel.Selection.CopyPicture(Appearance=1, Format=2)
            _log("   ✓ CopyPicture ejecutado con Selection (fallback)")
            copiado = True
        except Exception as e2:
            _log_traceback("   ❌ CopyPicture falló con ambos métodos", e2)

    if not copiado:
        return None

    time.sleep(1.5)

    imagen = ImageGrab.grabclipboard()
    if imagen:
        try:
            imagen.save(str(ruta_png))
            _log(f"   ✓ Captura guardada (portapapeles): {ruta_png}")
            return str(ruta_png)
        except Exception as e:
            _log_traceback(f"   ❌ No se pudo guardar la imagen del portapapeles", e)
    else:
        _log("   ❌ grabclipboard() devolvió None (portapapeles vacío o inaccesible "
             "en sesión no interactiva)")
        diagnosticar_portapapeles(" (después de CopyPicture)")

    return None


def _resolver_hoja(wb, nombre_esperado):
    """Obtiene la hoja de calculo por nombre exacto o, si no existe, intenta
    recuperar la correcta por coincidencia insensible a mayusculas o por
    similitud en el nombre. Devuelve la hoja o None si no hay candidata.

    Comportamiento anterior: si la hoja no se encontraba (nombre cambiado en el
    libro fuente), la captura fallaba y quedaba como imagen perdida. Ahora se
    registran las hojas disponibles y se resuelve la mas parecida.
    """
    # 1) Nombre exacto (comportamiento original, con reintentos COM)
    try:
        return com_con_reintentos(wb.Worksheets, nombre_esperado)
    except Exception:
        pass

    # 2) Enumerar hojas disponibles para recuperar la correcta
    coleccion = None
    nombres = []
    try:
        coleccion = com_con_reintentos(wb.Worksheets)
        total = com_con_reintentos(coleccion.Count)
        for i in range(1, total + 1):
            try:
                ws_item = com_con_reintentos(coleccion.Item, i)
                nombres.append(str(com_con_reintentos(ws_item.Name)))
            except Exception:
                continue
    except Exception as e:
        _log(f"   ❌ No se pudo enumerar las hojas del libro: {e}")
        return None

    _log(f"   ⚠️ Hoja '{nombre_esperado}' no encontrada por nombre exacto.")
    _log(f"   📄 Hojas disponibles: {', '.join(nombres) if nombres else '(ninguna)'}")

    def _norm(s: str) -> str:
        return re.sub(r"[^a-z0-9]", "", (s or "").lower())

    objetivo = _norm(nombre_esperado)

    # 3) Coincidencia exacta sin importar mayusculas/acentos
    for i, n in enumerate(nombres, 1):
        if _norm(n) == objetivo:
            _log(f"   ✅ Hoja resuelta (sin distinguir mayúsculas): '{n}'")
            return com_con_reintentos(coleccion.Item, i)

    # 4) Similitud por subcadena o caracteres comunes
    candidatas = []
    for i, n in enumerate(nombres, 1):
        nn = _norm(n)
        if not objetivo or not nn:
            continue
        if objetivo in nn or nn in objetivo:
            candidatas.append((max(len(objetivo), len(nn)), i, n))
        else:
            comunes = sum(1 for ch in set(objetivo) if ch in nn)
            ratio = comunes / max(len(set(objetivo)), 1)
            if ratio >= 0.6:
                candidatas.append((comunes, i, n))

    if candidatas:
        _, idx, mejor_nombre = max(candidatas, key=lambda x: (x[0], -x[1]))
        _log(f"   ✅ Hoja resuelta por similitud: '{mejor_nombre}' "
             f"(pedida: '{nombre_esperado}')")
        try:
            return com_con_reintentos(coleccion.Item, idx)
        except Exception as e:
            _log(f"   ❌ No se pudo obtener la hoja '{mejor_nombre}': {e}")
    return None


def capturar_multiples_rangos(ruta_archivo, capturas_config, password=None, dir_imagenes=None):
    """Captura multiples rangos de Excel. Retorna (imagenes, factor)."""
    excel = None
    wb = None
    imagenes = []
    factor = None
    dir_imagenes = Path(dir_imagenes) if dir_imagenes else DIR_IMAGENES

    try:
        _log(f"⏳ Abriendo Excel: {os.path.basename(ruta_archivo)}")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = True
        excel.DisplayAlerts = False

        if password:
            wb = excel.Workbooks.Open(ruta_archivo, False, False, None, password, password)
        else:
            wb = excel.Workbooks.Open(ruta_archivo)

        _log("✓ Excel abierto correctamente")
        try:
            _log(f"   📊 Versión Excel: {excel.Version}")
        except Exception:
            pass
        for p in _procesos_por_nombre("EXCEL.EXE"):
            _log(f"   📊 Proceso EXCEL.EXE: PID={p['pid']} | "
                 f"Sesión={p['sesion']} (##{p['sesion_id']}) | {p['memoria']}")
        _log("⏳ Esperando 30 segundos para carga completa...")
        time.sleep(30)
        _log("✓ Espera completada")
        _log("")

        for idx, captura in enumerate(capturas_config, 1):
            try:
                _log(f"{'='*70}")
                _log(f"📸 CAPTURA {idx}/{len(capturas_config)}: {captura['mensaje']}")
                _log(f"{'='*70}")

                ws = _resolver_hoja(wb, captura['hoja'])
                if ws is None:
                    _log(f"✗ Captura OMITIDA: no se encontró hoja '{captura['hoja']}' "
                         f"en el libro. Revisar 'Hojas disponibles' en el log.")
                    imagenes.append(None)
                    _log("")
                    continue
                com_con_reintentos(ws.Activate)
                time.sleep(0.3)

                # Resumen Dia: filtrar por mes actual y refrescar pivot
                if captura['hoja'] == "Resumen Dia":
                    try:
                        _log("   🗓️ Aplicando filtro de mes actual...")
                        if filtrar_pivot_por_mes_actual(ws):
                            _log("   ✓ Filtro de mes aplicado correctamente")
                        else:
                            _log("   ⚠️ No se pudo aplicar filtro de mes, continuando...")

                        _log(f"   🔄 Actualizando tabla dinámica de '{captura['hoja']}'...")
                        pivot_count = ws.PivotTables().Count
                        if pivot_count > 0:
                            for i in range(1, pivot_count + 1):
                                try:
                                    pivot_table = ws.PivotTables(i)
                                    pivot_table.RefreshTable()
                                except Exception as e:
                                    _log(f"      ⚠️ Error actualizando pivot {i}: {e}")
                            time.sleep(5)
                            _log("   ✓ Tabla(s) dinámica(s) actualizada(s) correctamente")
                    except Exception as e:
                        _log(f"   ⚠️ Error al procesar tabla dinámica: {e}")

                # Extraer factor MG NETO de la primera captura tipo 'dinamico'
                if idx == 1 and captura['tipo'] == 'dinamico':
                    factor = extraer_factor_mg_neto(
                        ws, captura['fila_inicio'], captura['columna_inicio'], "C"
                    )
                    _log(f"✓ Factor MG NETO extraído: {factor}")

                # Ocultar columna de margen en Resumen Dia
                col_margen_oculta = None
                if captura['hoja'] == "Resumen Dia":
                    col_margen_oculta = _ocultar_columna_margen(ws)

                # Determinar rango segun tipo
                if captura['tipo'] == 'dinamico':
                    ultima_fila = encontrar_ultima_fila(
                        ws, captura['columna_inicio'], captura['fila_inicio']
                    )
                    rango = f"{captura['columna_inicio']}{captura['fila_inicio']}:{captura['columna_fin']}{ultima_fila}"
                elif captura['tipo'] == 'dia_dinamico':
                    _log("   🔧 Procesando captura con días dinámicos...")
                    rango = ocultar_dias_anteriores_y_obtener_rango(
                        ws,
                        fila_fechas=captura.get('fila_fechas', 9),
                        fila_fin=captura.get('fila_fin', 29),
                    )
                    if rango is None:
                        _log("✗ Captura OMITIDA: No se encontró el día actual")
                        imagenes.append(None)
                        _log("")
                        continue
                else:
                    rango = captura['rango']

                _log(f"⏳ Capturando rango: {rango} de hoja '{captura['hoja']}'")

                ruta_captura = _capturar_rango(
                    excel=excel,
                    ws=ws,
                    rango=rango,
                    captura=captura,
                    dir_imagenes=dir_imagenes,
                )
                if ruta_captura:
                    imagenes.append(ruta_captura)
                    _log(f"✓ Captura guardada: {ruta_captura}")
                else:
                    _log(f"✗ Error: No se pudo capturar {captura['mensaje']} "
                         f"(falló el método de portapapeles). "
                         f"Revisar el diagnóstico de portapapeles arriba.")
                    imagenes.append(None)

                if col_margen_oculta:
                    try:
                        ws.Columns(col_margen_oculta).Hidden = False
                    except Exception:
                        pass
                _log("")
            except Exception as e:
                _log_traceback(f"✗ Error al capturar {captura['mensaje']}", e)
                imagenes.append(None)
                _log("")

        _log("✓ Cerrando Excel...")
        try:
            com_con_reintentos(wb.Close, False, max_intentos=3)
        except Exception:
            pass
        try:
            com_con_reintentos(excel.Quit, max_intentos=3)
        except Exception as e:
            _log(f"⚠️ No se pudo cerrar Excel normalmente, forzando: {e}")
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                               capture_output=True, timeout=10)
            except Exception:
                pass

        return imagenes, factor
    except Exception as e:
        _log(f"✗ Error general en Excel: {e}")
        if wb:
            try:
                com_con_reintentos(wb.Close, False, max_intentos=2)
            except Exception:
                pass
        if excel:
            try:
                com_con_reintentos(excel.Quit, max_intentos=2)
            except Exception:
                try:
                    subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                                   capture_output=True, timeout=10)
                except Exception:
                    pass
        return [None] * len(capturas_config), None
    finally:
        ws = None
        wb = None
        excel = None


# ===== TAREA =====
class EnvioInformeVentas(BaseTask):
    name = "envio_informe_ventas"

    def setup(self):
        self.notifier = EmailNotifier(self.settings.smtp_ventas, self.name)
        self.paths = self.settings.paths
        self.excel_cfg = self.settings.excel
        self.log_file = (_LOGS_ROOT / date.today().isoformat() / f"{self.name}.log")

        self.ruta_base = self.paths.base / "Pruebas"
        self.password_excel = self.excel_cfg.password
        self.destinatarios = _get_destinatarios()
        self.dir_imagenes = DIR_IMAGENES
        self.whatsapp = None

        _log(f"Ruta base (ventas): {self.ruta_base}")
        _log(f"Destinatarios: {', '.join(self.destinatarios)}")
        _log(f"Carpeta imagenes: {self.dir_imagenes}")

    def _buscar_archivo_ventas(self) -> Path:
        preferido = Path(self.ruta_base) / "$2026 VENTAS_Actualizacion.xlsx"
        if preferido.exists():
            return preferido
        patron = os.path.join(self.ruta_base, "$2026 VENTAS*.xlsx")
        _log(f"🔍 Buscando archivos en: {patron}")
        archivos = glob.glob(patron)
        if not archivos:
            raise FileNotFoundError(f"No se encontró archivo $2026 VENTAS en {self.ruta_base}")
        return Path(max(archivos, key=lambda x: Path(x).stat().st_mtime))

    def execute(self):
        archivo = self._buscar_archivo_ventas()
        _log(f"✓ Usando: {archivo.name}")

        _log("=" * 70)
        _log("📸 CAPTURANDO IMÁGENES DE EXCEL")
        _log("=" * 70)
        # '$2026 VENTAS_Actualizacion.xlsx' tiene su propia contrasena (area autorizada).
        clave = self.password_excel
        if archivo.name == "$2026 VENTAS_Actualizacion.xlsx":
            clave = getattr(self.excel_cfg, "password_ventas_actualizacion", "") or self.password_excel
        imagenes, factor = capturar_multiples_rangos(
            ruta_archivo=str(archivo),
            capturas_config=CAPTURAS,
            password=clave,
            dir_imagenes=self.dir_imagenes,
        )

        imagenes_validas = [img for img in imagenes if img is not None]
        if not imagenes_validas:
            raise ValueError("No se capturaron imágenes de Excel")
        _log(f"✅ {len(imagenes_validas)}/{len(imagenes)} imágenes capturadas")
        _log(f"✅ Imágenes guardadas en: {self.dir_imagenes}")
        for ruta_imagen in imagenes_validas:
            _log(f"   - {ruta_imagen}")

        mes_actual = obtener_mes_actual_espanol()
        mensaje_inicial = f"Cordial saludo.\n\nSe remite VTAS {mes_actual} MG NETO PONDERADO {factor if factor else 'N/A'}"

        imagenes_config = []
        for captura, ruta_imagen in zip(CAPTURAS, imagenes):
            imagenes_config.append({
                'ruta': ruta_imagen,
                'caption': captura['mensaje'],
            })

        _log("=" * 70)
        _log("🌐 INICIANDO WHATSAPP WEB")
        _log("=" * 70)
        self.whatsapp = WhatsAppWeb(max_reintentos=3)
        sesion_ok = self.whatsapp.iniciar_sesion()
        if not sesion_ok:
            # Recuperación automática: re-validar ChromeDriver (p. ej. Chrome se
            # actualizó entre intentos) y reintentar una vez más la sesión.
            _log("\n🔁 Recuperación automática: re-validando ChromeDriver y "
                 "reintentando sesión de WhatsApp Web...")
            try:
                from core.chromedriver_utils import obtener_chromedriver_path
                path = obtener_chromedriver_path(reemplazar_local=True)
                _log(f"   ChromeDriver según revalidación: {path}")
            except Exception as e:
                _log(f"   ⚠️ No se pudo re-validar ChromeDriver: {e}")
            self.whatsapp = WhatsAppWeb(max_reintentos=3)
            sesion_ok = self.whatsapp.iniciar_sesion()
        if not sesion_ok:
            raise RuntimeError("No se pudo iniciar WhatsApp Web")

        resultados = []
        total = len(self.destinatarios)
        exitosos = 0

        for idx, nombre_chat in enumerate(self.destinatarios, 1):
            _log(f"\n{'='*70}")
            _log(f"📱 DESTINATARIO {idx}/{total}: {nombre_chat}")
            _log(f"{'='*70}")

            exito = self.whatsapp.enviar_reporte_completo(
                nombre_chat=nombre_chat,
                mensaje_inicial=mensaje_inicial,
                imagenes_config=imagenes_config,
            )
            resultados.append({'nombre': nombre_chat, 'exito': exito})
            if exito:
                exitosos += 1
            if idx < total:
                _log("\n⏳ Esperando 5 segundos...")
                time.sleep(5)

        _log("\n" + "=" * 70)
        _log("📊 RESUMEN FINAL")
        _log("=" * 70)
        _log(f"📄 Archivo: {archivo.name}")
        _log(f"📊 Factor MG NETO: {factor if factor else 'N/A'}")
        for resultado in resultados:
            estado = "✅" if resultado['exito'] else "⚠️"
            _log(f"{estado} {resultado['nombre']}")
        _log(f"{'✅' if exitosos == total else '⚠️'} Envíos exitosos: {exitosos}/{total}")

        if total > 0:
            self.notifier.notify_success(
                detail=f"Archivo: {archivo.name} — Envíos exitosos: {exitosos}/{total}"
                + f" — Factor MG: {factor if factor else 'N/A'}",
                attachment=self.log_file if self.log_file.exists() else None,
            )

        if exitosos == 0:
            raise RuntimeError("Ningún destinatario recibió el informe")

    def teardown(self):
        if self.whatsapp:
            try:
                self.whatsapp.cerrar()
            except Exception as e:
                _log(f"⚠️ Error cerrando WhatsApp: {e}")
            self.whatsapp = None