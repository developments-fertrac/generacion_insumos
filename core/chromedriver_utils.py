"""
Utilidades para obtener un ChromeDriver compatible con el Chrome instalado.

Problema que resuelve
---------------------
Chrome se actualiza automaticamente con frecuencia (y tambien el ChromeDriver
queda desactualizado). Cuando la version mayor del ChromeDriver no coincide con
la del Chrome instalado, Selenium lanza:

    session not created: This version of ChromeDriver only supports Chrome version X
    Current browser version is Y [...]

Eso derrumbaba el envio de informes por WhatsApp Web (envio_informe_ventas).

Recuperacion automatica
-----------------------
1. Detecta la version mayor de Chrome instalado (registro de Windows / binario).
2. Valida el ChromeDriver local (raiz del proyecto) contra esa version.
3. Si el local no sirve, busca un driver compatible en las caches de
   webdriver-manager (.wdm / .wdm_cache), ya descargado en corridas previas.
4. Si no hay cache compatible, descarga el ChromeDriver correcto con
   ChromeDriverManager (webdriver-manager).
5. Reemplaza el ChromeDriver local obsoleto por el compatible (con backup
   .bak), para que las proximas corridas arranquen directo y sean rapidas.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path

from core.logger import get_logger

log = get_logger("chromedriver_utils")

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _state_dir() -> Path:
    """Directorio de estado operativo, fuera del repositorio.

    Phase 0: el repo no guarda cache de drivers ni perfiles de navegador.
    Configurable con la variable de entorno STATE_DIR.
    """
    configurado = os.environ.get("STATE_DIR", "").strip()
    if configurado:
        return Path(configurado)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "GeneracionInsumos"


# Rutas alternativas donde se buscan drivers cacheados por webdriver-manager
_RAICES_CACHE = (
    _state_dir() / ".wdm",
    _state_dir() / ".wdm_cache",
    Path.home() / ".wdm",
)

# Rutas del binario de Chrome (para detectar version como respaldo)
_RUTAS_CHROME = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
)

_CLAVES_REGISTRO = (
    r"HKLM\SOFTWARE\Google\Chrome\BLBeacon",
    r"HKCU\SOFTWARE\Google\Chrome\BLBeacon",
    r"HKLM\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
)


def _mayor(version: str | None) -> str | None:
    """Extrae la version mayor ('153.0.8010.53' -> '153')."""
    if not version:
        return None
    m = re.search(r"(\d+)", str(version))
    return m.group(1) if m else None


def _sortable(version: str | None) -> int:
    try:
        return int(str(version).split(".")[0])
    except Exception:
        return 0


def chrome_version_instalada() -> str | None:
    """Version completa de Chrome instalado o None si no se puede leer."""
    # 1) Registro de Windows (fuente principal)
    for clave in _CLAVES_REGISTRO:
        try:
            res = subprocess.run(
                ["reg", "query", clave, "/v", "version"],
                capture_output=True, text=True, timeout=5,
            )
            if res.returncode == 0 and res.stdout.strip():
                partes = res.stdout.strip().split()
                if partes:
                    return partes[-1]
        except Exception:
            continue

    # 2) Version del binario (respaldo)
    for ruta in _RUTAS_CHROME:
        try:
            if os.path.exists(ruta):
                res = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     f"(Get-Item '{ruta}').VersionInfo.ProductVersion"],
                    capture_output=True, text=True, timeout=10,
                )
                if res.returncode == 0 and res.stdout.strip():
                    return res.stdout.strip().splitlines()[-1].strip()
        except Exception:
            continue
    return None


def chromedriver_version(ruta) -> str | None:
    """Version completa del ChromeDriver en la ruta indicada (ej. '153.0.8010.52')."""
    try:
        if not ruta or not os.path.exists(str(ruta)):
            return None
        res = subprocess.run(
            [str(ruta), "--version"], capture_output=True, text=True, timeout=10,
        )
        if res.returncode == 0:
            m = re.search(r"ChromeDriver\s+(\d[\w.]*)", res.stdout)
            return m.group(1) if m else res.stdout.strip()[:60] or None
    except Exception as e:
        log.debug("No se pudo leer version de ChromeDriver %s: %s", ruta, e)
    return None


def chromedriver_mayor(ruta) -> str | None:
    """Version mayor del ChromeDriver en la ruta indicada o None."""
    return _mayor(chromedriver_version(ruta))


def es_compatible(ruta, chrome_mayor: str | None) -> bool:
    """True si el ChromeDriver de la ruta es compatible con la mayor de Chrome.

    Si no se pudo determinar la version de Chrome, se asume compatible (no hay
    forma de validarlo), manteniendo el comportamiento historico.
    """
    if not ruta or not os.path.exists(str(ruta)):
        return False
    if not chrome_mayor:
        return True
    return chromedriver_mayor(ruta) == chrome_mayor


def _drivers_en_cache() -> list[tuple[Path, str | None]]:
    """Lista (ruta, version_mayor) de todos los chromedriver.exe cacheados."""
    encontrados: list[tuple[Path, str | None]] = []
    for raiz in _RAICES_CACHE:
        try:
            if not raiz.exists():
                continue
            for archivo in raiz.rglob("chromedriver.exe"):
                try:
                    encontrados.append((archivo, chromedriver_mayor(archivo)))
                except Exception:
                    continue
        except Exception:
            continue
    return encontrados


def _buscar_en_cache(chrome_mayor: str | None) -> Path | None:
    """Busca en caches un ChromeDriver compatible. Devuelve la ruta o None."""
    drivers = _drivers_en_cache()
    if not drivers:
        return None

    if chrome_mayor:
        exactos = [d for d in drivers if d[1] == chrome_mayor]
        if exactos:
            mejor = max(exactos, key=lambda d: _sortable(d[1]))
            log.info("ChromeDriver %s encontrado en cache: %s", chrome_mayor, mejor[0])
            return mejor[0]

        # Un driver de otra version mayor SIEMPRE falla ("This version of
        # ChromeDriver only supports Chrome version X"): no se usa; se descarga.
        log.info("Sin ChromeDriver para mayor %s en cache; se intentara descargar", chrome_mayor)
        return None
    else:
        con_version = [d for d in drivers if d[1]]
        if con_version:
            return max(con_version, key=lambda d: _sortable(d[1]))[0]
    return None


def _descargar(chrome_mayor: str | None) -> str | None:
    """Descarga (si hace falta) un ChromeDriver compatible con webdriver-manager."""
    try:
        from webdriver_manager.chrome import ChromeDriverManager
    except ImportError:
        log.warning("webdriver-manager no instalado; no se puede descargar ChromeDriver")
        return None

    # Guardar descargas en la cache de estado (fuera del repo) para reutilizarlas
    try:
        os.environ.setdefault("WDM_CACHE_PATH", str(_state_dir() / ".wdm"))
        os.environ.setdefault("WDM_LOCAL", "1")
    except Exception:
        pass

    intentos = []
    if chrome_mayor:
        intentos.append(("por mayor", lambda: ChromeDriverManager(driver_version=chrome_mayor).install()))
    intentos.append(("auto-detect", lambda: ChromeDriverManager().install()))

    ultimo_error = None
    for nombre, fn in intentos:
        try:
            log.info("Descargando ChromeDriver (%s)...", nombre)
            path = fn()
            if path and os.path.exists(path):
                if es_compatible(path, chrome_mayor):
                    log.info("ChromeDriver descargado: %s (%s)", path, chromedriver_mayor(path))
                    return str(path)
                log.warning("ChromeDriver descargado %s no compatible con Chrome %s",
                            chromedriver_mayor(path), chrome_mayor)
        except Exception as e:
            ultimo_error = e
            log.warning("ChromeDriverManager (%s) fallo: %s", nombre, e)

    log.warning("No se pudo descargar un ChromeDriver compatible (ultimo error: %s)", ultimo_error)
    return None


def _reemplazar_local(local: Path, nuevo: Path) -> bool:
    """Reemplaza el ChromeDriver local por 'nuevo', respaldando el anterior.

    Devuelve True si el local queda compatible (o no existia y se creo).
    """
    try:
        if local.exists():
            version_anterior = chromedriver_mayor(local) or "desconocida"
            backup = local.with_suffix(".exe.bak")
            if backup.exists():
                backup.unlink()
            shutil.copy2(local, backup)
            log.info("ChromeDriver local obsoleto (%s) respaldado en %s",
                     version_anterior, backup.name)
        shutil.copy2(nuevo, local)
        log.info("ChromeDriver local actualizado a: %s", chromedriver_mayor(local))
        return True
    except Exception as e:
        log.warning("No se pudo actualizar el ChromeDriver local (probablemente "
                    "esta bloqueado por un proceso activo): %s", e)
        return False


def obtener_chromedriver_path(reemplazar_local: bool = True) -> str | None:
    """Devuelve la ruta de un ChromeDriver compatible con el Chrome instalado.

    Si el ChromeDriver local (raiz del proyecto) ya es compatible, lo usa.
    Si no, busca en cache / descarga y lo copia encima del local (con backup),
    dejando la recuperacion hecha para las proximas corridas.

    Si el archivo local esta bloqueado, devuelve la ruta del driver bueno
    encontrado en cache/descarga para no fallar igualmente.
    """
    chrome_full = chrome_version_instalada()
    chrome_mayor = _mayor(chrome_full)
    log.info("Chrome instalado: %s (mayor %s)", chrome_full or "?", chrome_mayor or "?")

    local = PROJECT_ROOT / "chromedriver.exe"

    if es_compatible(local, chrome_mayor):
        log.info("ChromeDriver local OK (mayor %s)", chromedriver_mayor(local))
        return str(local)

    if local.exists() and chrome_mayor:
        log.warning(
            "ChromeDriver local %s NO es compatible con Chrome %s. "
            "Recuperando automaticamente...",
            chromedriver_mayor(local) or "?", chrome_mayor,
        )

    # 1) cache local de webdriver-manager (sin internet)
    nuevo = _buscar_en_cache(chrome_mayor)
    # 2) descarga
    if nuevo is None:
        descargado = _descargar(chrome_mayor)
        nuevo = Path(descargado) if descargado else None

    if nuevo is not None and str(nuevo) != str(local):
        if reemplazar_local:
            if _reemplazar_local(local, nuevo):
                # Verificar que la copia quedo bien; si no, usar la fuente
                if es_compatible(local, chrome_mayor):
                    return str(local)
                return str(nuevo)
            return str(nuevo)
        return str(nuevo)
    if nuevo is not None:
        return str(nuevo)

    if chrome_mayor:
        # Con la version de Chrome conocida, un driver sin validar no sirve.
        # None => quien llama usa Selenium Manager (descarga automatica).
        log.warning("Sin ChromeDriver %s en cache ni por descarga; se delega en Selenium Manager",
                    chrome_mayor)
        return None

    # Sin version de Chrome detectada: ultimo recurso, el local tal cual
    for cand in (local, Path(r"C:\chromedriver\chromedriver.exe"), Path("chromedriver.exe")):
        if cand.exists():
            log.warning("Usando ChromeDriver %s sin validar (no se pudo "
                        "detectar/recuperar la version correcta)", cand)
            return str(cand)

    log.error("No se encontro ningun ChromeDriver compatible")
    return None

