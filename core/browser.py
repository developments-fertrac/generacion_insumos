from __future__ import annotations

import os
import tempfile
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from config.settings import BrowserConfig, ErpConfig
from core.logger import get_logger

log = get_logger("browser")


def _chrome_version_instalada() -> str | None:
    """Devuelve la version mayor de Chrome instalada (ej. '152') o None si no se puede leer."""
    try:
        import subprocess
        for clave in [
            r"HKLM\SOFTWARE\Google\Chrome\BLBeacon",
            r"HKCU\SOFTWARE\Google\Chrome\BLBeacon",
        ]:
            try:
                res = subprocess.run(
                    ["reg", "query", clave, "/v", "version"],
                    capture_output=True, text=True, timeout=5,
                )
                if res.returncode == 0 and res.stdout.strip():
                    ver = res.stdout.strip().split()[-1]
                    return ver.split(".")[0] if ver else None
            except Exception:
                continue
    except Exception:
        pass
    return None


def _driver_path(version_dir: str | None) -> str:
    """Resuelve el ChromeDriver. Si version_dir viene (mayor de Chrome), fuerza
    esa version para evitar el mismatch Chrome <-> ChromeDriver."""
    if version_dir:
        try:
            return ChromeDriverManager(driver_version=version_dir).install()
        except Exception as e:
            log.warning("No se pudo instalar ChromeDriver %s: %s", version_dir, e)
    return ChromeDriverManager().install()


def create_driver(download_dir: str, config: BrowserConfig | None = None) -> webdriver.Chrome:
    if config is None:
        config = BrowserConfig()

    cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".wdm_cache")
    os.makedirs(cache_dir, exist_ok=True)

    chrome_mayor = _chrome_version_instalada()
    log.info("Version mayor de Chrome detectada: %s", chrome_mayor or "desconocida")

    last_error = None
    for intento in range(1, config.max_intentos_driver + 1):
        if intento > 1:
            log.info("Reintentando arranque de Chrome (intento %d/%d)...", intento, config.max_intentos_driver)
            time.sleep(30)

        try:
            chrome_options = _build_options(download_dir, config)
            os.environ["WDM_CACHE_PATH"] = cache_dir
            os.environ["WDM_LOCAL"] = "1"
            driver_path = _driver_path(chrome_mayor)
            log.info("ChromeDriver: %s", driver_path)
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.set_page_load_timeout(120)
            driver.set_script_timeout(120)
            if not config.headless:
                driver.maximize_window()
            log.info("Driver configurado correctamente")
            return driver
        except Exception as e:
            last_error = e
            log.warning("Error arrancando Chrome en intento %d: %s", intento, e)
            if intento == config.max_intentos_driver:
                raise RuntimeError(
                    f"Chrome no pudo arrancar despues de {config.max_intentos_driver} intentos. "
                    f"Ultimo error: {last_error}"
                )
    return None


def _build_options(download_dir: str, config: BrowserConfig) -> Options:
    opts = Options()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }
    opts.add_experimental_option("prefs", prefs)

    if config.headless:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
        # El modo sin ventana descarta la descarga por el dialogo "Guardar como":
        opts.add_argument("--safebrowsing-disable-download-protection")
        opts.add_argument("--enable-features=DownloadBubble,downloadBubbleEnabled")
        log.info("Modo sin ventana activado")

    for arg in ["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage",
                "--ignore-certificate-errors", "--disable-extensions",
                "--disable-background-networking",
                "--disable-default-apps", "--disable-sync", "--metrics-recording-only"]:
        opts.add_argument(arg)

    # Perfil EFIMERO por corrida: uno persistente compartido queda bloqueado si
    # una instancia previa de Chrome queda colgada (proceso zombi), impidiendo
    # arrancar la siguiente sesion.
    temp_dir = tempfile.mkdtemp()
    opts.add_argument(f"--user-data-dir={temp_dir}")
    return opts


def accept_cookies_if_present(driver: webdriver.Chrome) -> None:
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,
                "//a[contains(@class,'btn') and contains(translate(., 'ok', 'OK'), 'OK')] | "
                "//button[contains(translate(., 'ok', 'OK'), 'OK') and string-length(normalize-space(.)) <= 5]"
            ))
        )
        btn.click()
        log.info("Banner de cookies aceptado")
    except Exception:
        pass


def do_login(driver: webdriver.Chrome, erp: ErpConfig, config: BrowserConfig | None = None) -> bool:
    if config is None:
        config = BrowserConfig()
    max_intentos = config.max_intentos_login

    log.info("Iniciando sesion en ERP...")

    for intento in range(1, max_intentos + 1):
        if intento > 1:
            log.info("Reintentando login (intento %d/%d)...", intento, max_intentos)

        driver.get(erp.url_login)
        wait = WebDriverWait(driver, 20)

        try:
            campo_usuario = wait.until(EC.presence_of_element_located((By.NAME, "login")))
            accept_cookies_if_present(driver)

            campo_usuario.clear()
            campo_usuario.send_keys(erp.usuario)
            log.info("Usuario ingresado")

            campo_clave = driver.find_element(By.NAME, "password")
            campo_clave.clear()
            campo_clave.send_keys(erp.clave)
            log.info("Contrasena ingresada")

            boton_login = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
            )
            try:
                boton_login.click()
            except Exception:
                driver.execute_script("arguments[0].click();", boton_login)
            log.info("Boton de login clickeado")

            try:
                WebDriverWait(driver, 20).until(
                    lambda d: "web/login" not in d.current_url
                )
                log.info("Sesion iniciada correctamente, URL: %s", driver.current_url)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//nav | //div[contains(@class,'o_main_navbar')] | //div[contains(@class,'o_menu')]"
                    ))
                )
                return True
            except Exception:
                current_url = driver.current_url
                login_error = driver.execute_script("""
                    var alerts = document.querySelectorAll('.alert-danger, .alert-warning, .o_login_error, .oe_login_error');
                    var msgs = [];
                    alerts.forEach(function(a) { if (a.textContent.trim()) msgs.push(a.textContent.trim()); });
                    return msgs.join(' | ');
                """) or ""
                page_title = driver.title
                log.warning(
                    "Login no completado en intento %d | URL: %s | Titulo: %s | Error ERP: %s",
                    intento, current_url, page_title, login_error,
                )
                if intento < max_intentos:
                    time.sleep(3)
                continue

        except Exception as e:
            log.error("Error en login intento %d: %s", intento, e)
            if intento < max_intentos:
                time.sleep(3)
            continue

    log.error("Login fallido despues de todos los intentos")
    return False


def wait_for_download(carpeta: str, timeout: int = 300, snapshot: dict | None = None) -> str | None:
    import glob

    log.info("Esperando descarga (maximo %ds)...", timeout)
    tiempo_inicio = time.time()
    ultimo_reporte = tiempo_inicio

    if snapshot is not None:
        archivos_existentes = snapshot
    else:
        archivos_existentes = {}
        for ext in ("*.xlsx", "*.xls"):
            for f in glob.glob(os.path.join(carpeta, ext)):
                if not os.path.basename(f).startswith("~$"):
                    archivos_existentes[f] = os.path.getmtime(f)

    while time.time() - tiempo_inicio < timeout:
        if time.time() - ultimo_reporte >= 5:
            log.info("Esperando... %ds transcurridos", int(time.time() - tiempo_inicio))
            ultimo_reporte = time.time()

        temporales = glob.glob(os.path.join(carpeta, "*.crdownload")) + glob.glob(os.path.join(carpeta, "*.tmp"))
        if temporales:
            time.sleep(1)
            continue

        archivos_actuales = {}
        for ext in ("*.xlsx", "*.xls"):
            for f in glob.glob(os.path.join(carpeta, ext)):
                if not os.path.basename(f).startswith("~$"):
                    archivos_actuales[f] = os.path.getmtime(f)

        for archivo, mtime_actual in archivos_actuales.items():
            if archivo not in archivos_existentes:
                if time.time() - mtime_actual < 90:
                    log.info("Archivo descargado: %s", os.path.basename(archivo))
                    return archivo
            elif mtime_actual > archivos_existentes[archivo]:
                if time.time() - mtime_actual < 90:
                    log.info("Archivo descargado (modificado): %s", os.path.basename(archivo))
                    return archivo

        time.sleep(1)

    log.warning("Timeout esperando la descarga")
    return None


def _glob_hojas(carpeta: str) -> list:
    import glob
    archivos = []
    for ext in ("*.xlsx", "*.xlsm", "*.xls", "*.csv"):
        archivos.extend(glob.glob(os.path.join(carpeta, ext)))
    return archivos


def take_xlsx_snapshot(carpeta: str) -> dict:
    import glob
    snapshot = {}
    for f in _glob_hojas(carpeta):
        if os.path.basename(f).startswith("~$"):
            continue
        try:
            snapshot[f] = os.path.getmtime(f)
        except OSError:
            pass
    log.info("Snapshot: %d archivos de hoja de calculo preexistentes", len(snapshot))
    return snapshot


def wait_for_download_complete(carpeta: str, snapshot: dict, timeout: int = 1500, grace: int = 180, margin: int = 5) -> str | None:
    import glob

    total_limit = timeout + grace
    log.info("Esperando descarga completa (timeout %ds + gracia %ds = %ds max)...", timeout, grace, total_limit)

    tiempo_inicio = time.time()
    tamanos_previos = {}
    gracia_mostrada = False
    ultimo_reporte = tiempo_inicio
    arranco_viendo_temporal_o_nuevo = False
    aviso_inicio_pendiente = True

    while time.time() - tiempo_inicio < total_limit:
        transcurrido = time.time() - tiempo_inicio

        temporales = (
            glob.glob(os.path.join(carpeta, "*.crdownload"))
            + glob.glob(os.path.join(carpeta, "*.tmp"))
            + glob.glob(os.path.join(carpeta, "*.part"))
        )

        hay_registro = bool(temporales)
        if not hay_registro:
            for f in _glob_hojas(carpeta):
                if os.path.basename(f).startswith("~$"):
                    continue
                try:
                    if f not in snapshot or os.path.getmtime(f) > snapshot[f] + 5:
                        hay_registro = True
                        break
                except OSError:
                    continue

        if hay_registro:
            arranco_viendo_temporal_o_nuevo = True
            aviso_inicio_pendiente = False

        if (
            aviso_inicio_pendiente
            and transcurrido >= 90
        ):
            log.warning(
                "Aviso: tras %ds no se detecto inicio de descarga en '%s'. "
                "El clic 'Exportar a fichero' posiblemente no lanzo la descarga "
                "(sin .crdownload/.tmp ni .xlsx nuevo). Revisar si el export quedo "
                "pendiente en el ERP o cayo a otra carpeta.",
                int(transcurrido),
                carpeta,
            )
            aviso_inicio_pendiente = False

        if transcurrido >= timeout and not gracia_mostrada:
            log.warning("Timeout principal (%ds) alcanzado, periodo de gracia de %ds...", timeout, grace)
            gracia_mostrada = True

        if arranco_viendo_temporal_o_nuevo and transcurrido - (ultimo_reporte - tiempo_inicio) >= 30:
            log.info("Esperando descarga... %ds transcurridos (max %ds)", int(transcurrido), total_limit)
            ultimo_reporte = time.time()

        temporales = glob.glob(os.path.join(carpeta, "*.crdownload")) + glob.glob(os.path.join(carpeta, "*.tmp"))

        if not temporales:
            candidatos = []
            for f in _glob_hojas(carpeta):
                if os.path.basename(f).startswith("~$"):
                    continue
                try:
                    mtime = os.path.getmtime(f)
                except OSError:
                    continue
                if f not in snapshot:
                    candidatos.append((f, mtime))
                else:
                    prev = snapshot.get(f)
                    if prev is not None and mtime > prev + margin:
                        candidatos.append((f, mtime))

            if candidatos:
                arch = max(candidatos, key=lambda x: x[1])[0]
                try:
                    size = os.path.getsize(arch)
                except OSError:
                    size = None
                if size is not None:
                    prev_size = tamanos_previos.get(arch)
                    if prev_size == size and size > 0:
                        log.info("Descarga completada en %.0fs: %s (%d KB)", transcurrido, os.path.basename(arch), size // 1024)
                        return arch
                    tamanos_previos[arch] = size

        time.sleep(1)

    log.warning("Timeout esperando descarga (total %ds)", total_limit)
    return None
