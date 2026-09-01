"""Utilidades compartidas de navegacion en ERP Fertrac via Selenium.

Funciones reutilizables para interactuar con menus, dropdowns, modales,
paginacion y exportacion en la plataforma Odoo/Fertrac.
"""

from __future__ import annotations

import re
import time

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from config.erp_selectors import (
    BotonAccion,
    Carga,
    Checkbox,
    Exportar,
    Modal,
    Paginador,
    VistaLista,
)
from core.logger import get_logger

log = get_logger("erp_navigation")


def descartar_alert_js(driver) -> bool:
    """Acepta una alerta JavaScript abierta (si la hay) para evitar que
    bloquee el bucle de comandos de Selenium (causa comun de 'congelamiento')."""
    try:
        alert = driver.switch_to.alert
        texto = alert.text or ""
        alert.accept()
        if texto:
            log.info("Alerta JS aceptada: %s", texto)
        else:
            log.info("Alerta JS aceptada (vacia)")
        return True
    except Exception:
        return False


def click_with_fallback(driver, element, description: str = "elemento") -> bool:
    try:
        element.click()
        return True
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            log.warning("No se pudo hacer click en %s", description)
            return False


def find_visible_element(driver, selectors: list[str], timeout: int = 15):
    wait = WebDriverWait(driver, timeout)
    for selector in selectors:
        try:
            elem = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            if elem.is_displayed():
                return elem
        except Exception:
            continue
    return None


def find_all_visible(driver, selector: str) -> list:
    elementos = driver.find_elements(By.XPATH, selector)
    return [e for e in elementos if e.is_displayed()]


def open_menu_by_selectors(
    driver, selectors: list[str], menu_name: str, timeout: int = 15
) -> bool:
    wait = WebDriverWait(driver, timeout)
    for selector in selectors:
        try:
            elem = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
            if elem.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                time.sleep(0.3)
                click_with_fallback(driver, elem, menu_name)
                log.info("Menu '%s' abierto", menu_name)
                return True
        except Exception:
            continue
    log.error("No se encontro el menu '%s'", menu_name)
    return False


def select_option_from_autocomplete(
    driver, input_selectors: list[str], option_selectors: list[str],
    search_text: str, match_text: str, description: str = "opcion",
    timeout: int = 10,
) -> bool:
    campo = find_visible_element(driver, input_selectors, timeout)
    if not campo:
        log.error("No se encontro el campo de busqueda para %s", description)
        return False

    driver.execute_script("arguments[0].scrollIntoView(true);", campo)
    time.sleep(0.5)
    campo.clear()
    time.sleep(0.3)
    campo.send_keys(search_text)
    time.sleep(3)

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, " | ".join(option_selectors)))
        )
    except Exception:
        log.warning("Timeout esperando opciones del dropdown")

    time.sleep(1)

    for selector in option_selectors:
        try:
            opciones = driver.find_elements(By.XPATH, selector)
            opciones = [op for op in opciones if op.is_displayed()]
            for opcion in opciones:
                texto = opcion.text.strip().lower()
                if match_text.lower() in texto:
                    click_with_fallback(driver, opcion, description)
                    time.sleep(1)
                    log.info("%s seleccionada: '%s'", description, opcion.text.strip())
                    return True
        except Exception:
            continue

    log.error("No se encontro la %s con texto '%s'", description, match_text)
    return False


def select_option_from_dropdown(
    driver, dropdown_selectors: list[str], option_text: str,
    description: str = "opcion",
) -> bool:
    for selector in dropdown_selectors:
        try:
            elementos = driver.find_elements(By.XPATH, selector)
            for elem in elementos:
                if elem.is_displayed():
                    select = Select(elem)
                    for opcion in select.options:
                        if option_text.lower() in opcion.text.lower():
                            select.select_by_visible_text(opcion.text)
                            time.sleep(1)
                            log.info("%s seleccionada: '%s'", description, opcion.text)
                            return True
        except Exception:
            continue
    log.error("No se pudo seleccionar %s", description)
    return False


def get_pager_info(driver) -> tuple:
    """Devuelve (registro_final, total) del paginador Odoo.

    Soporta el formato moderno (spans separados: ``.o_pager_value`` con
    '1-80' y ``.o_pager_limit`` con '15941') y el clasico ('1-80 / 15941'),
    con o sin separadores de miles.
    Retorna (None, None) si no puede detectar los valores.
    """
    def _final(texto) -> int | None:
        texto = (texto or "").strip()
        m = re.search(r"(\d+)\s*[-–—]\s*(\d+)", texto)
        if m:
            return int(m.group(2))
        m = re.search(r"(\d+)", texto)
        return int(m.group(1)) if m else None

    def _total(texto) -> int | None:
        texto = (texto or "").strip()
        grupos = re.findall(r"\d[\d.,]*", texto)
        if not grupos:
            return None
        limpio = re.sub(r"[^0-9]", "", grupos[-1])
        return int(limpio) if limpio else None

    try:
        valores = find_all_visible(driver, Paginador.RANGO_VALUE)
        limites = find_all_visible(driver, Paginador.RANGO_LIMIT)

        final = _final(valores[0].text) if valores else None
        total = _total(limites[0].text) if limites else None

        if final is None or total is None:
            for elem in find_all_visible(driver, Paginador.CONTENEDOR):
                texto = elem.text.strip() or elem.get_attribute("textContent").strip()
                m = re.search(r"(\d+)\s*[-–—]\s*(\d+)\s*/\s*([\d.,]+)", texto)
                if m:
                    if final is None:
                        final = int(m.group(2))
                    if total is None:
                        total = _total(m.group(3))
                    if final is not None and total is not None:
                        break
        return final, total
    except Exception:
        return None, None


def wait_for_loading_complete(driver, timeout: int = 600, log_func=None) -> bool:
    if log_func is None:
        log_func = log.info
    start = time.time()
    last_report = start

    while time.time() - start < timeout:
        now = time.time()

        if now - last_report >= 60:
            minutes = int((now - start) / 60)
            log_func("Esperando carga... %d min(s) transcurrido(s)", minutes)
            last_report = now

        try:
            mensajes = find_all_visible(driver, Carga.MENSAJES)
            if mensajes:
                time.sleep(10)
                continue
        except Exception:
            pass

        try:
            spinners = find_all_visible(driver, Carga.SPINNERS)
            if spinners:
                time.sleep(5)
                continue
        except Exception:
            pass

        try:
            overlay = driver.execute_script(Carga.OVERLAY_JS)
            if overlay:
                time.sleep(10)
                continue
        except Exception:
            pass

        try:
            final, total = get_pager_info(driver)
            if final is not None and total and total > 0 and final >= total * 0.98:
                elapsed = int((time.time() - start) / 60)
                log_func("Carga completada: %d/%d (%d minutos)", final, total, elapsed)
                return True
        except Exception:
            pass

        time.sleep(5)

    log.warning("Timeout esperando carga completa (%ds)", timeout)
    return False


def open_action_menu(driver) -> bool:
    for selector in BotonAccion.SELECTORES:
        try:
            elementos = driver.find_elements(By.XPATH, selector)
            for boton in elementos:
                if boton.is_displayed():
                    texto = boton.text.strip()
                    if "acci" in texto.lower() or "action" in texto.lower():
                        driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                        time.sleep(0.3)
                        click_with_fallback(driver, boton, "menu Accion")
                        time.sleep(1)
                        log.info("Menu 'Accion' abierto")
                        return True
        except Exception:
            continue
    log.error("No se encontro el boton 'Accion'")
    return False


def select_export_option(driver) -> bool:
    for selector in Exportar.MENU_OPCIONES:
        try:
            opcion = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            opcion.click()
            time.sleep(2)
            log.info("'Exportar' seleccionado")
            return True
        except Exception:
            continue
    log.error("No se encontro la opcion 'Exportar'")
    return False


def select_export_file_option(driver) -> bool:
    for selector in Exportar.A_FICHERO:
        try:
            boton = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            boton.click()
            time.sleep(2)
            log.info("Exportacion iniciada")
            return True
        except Exception:
            continue
    log.error("No se encontro el boton 'Exportar a fichero'")
    return False


def select_from_modal_dropdown(
    driver, option_text: str, description: str = "opcion"
) -> bool:
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, Modal.CONTENEDOR))
        )
    except Exception:
        log.warning("No se detecto modal, continuando...")

    dropdown = None
    for selector in Modal.SELECT_EXPORT:
        try:
            elementos = driver.find_elements(By.XPATH, selector)
            for elem in elementos:
                if elem.is_displayed():
                    dropdown = elem
                    break
            if dropdown:
                break
        except Exception:
            continue

    if not dropdown:
        dropdown = driver.execute_script(Modal.SELECT_JS)

    if not dropdown:
        log.error("No se encontro el dropdown en el modal")
        return False

    driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
    time.sleep(0.3)

    select = Select(dropdown)
    for opc in select.options:
        texto = opc.text.strip()
        if option_text.upper() in texto.upper():
            try:
                select.select_by_visible_text(texto)
            except Exception:
                select.select_by_index(select.options.index(opc))
            time.sleep(0.5)
            log.info("%s seleccionada: '%s'", description, texto)
            return True

    for opc in select.options:
        valor = opc.get_attribute("value") or ""
        if option_text.lower() in valor.lower():
            select.select_by_value(valor)
            time.sleep(0.5)
            log.info("%s seleccionada por valor: '%s'", description, opc.text)
            return True

    log.error("No se encontro '%s' en las opciones del dropdown", option_text)
    return False


def close_modal(driver) -> bool:
    for selector in Modal.CERRAR:
        try:
            elementos = driver.find_elements(By.XPATH, selector)
            for elem in elementos:
                if elem.is_displayed():
                    elem.click()
                    time.sleep(1)
                    log.info("Modal cerrado")
                    return True
        except Exception:
            continue
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)
    log.info("Modal cerrado con ESC")
    return True


def detect_and_accept_empty_alert(driver, timeout: int = 5) -> bool:
    from config.erp_selectors import AlertaERP
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, AlertaERP.SIN_REGISTROS))
        )
        log.warning("ERP indica: 'No hay registros que coincidan'")
        try:
            boton = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, AlertaERP.BOTON_ACEPTAR))
            )
            boton.click()
            time.sleep(1)
            log.info("Alerta aceptada")
        except Exception:
            pass
        return True
    except Exception:
        return False


def wait_for_erp_section(driver, url: str, section_check: str, timeout: int = 30) -> bool:
    driver.get(url)
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, section_check))
        )
        log.info("Navegacion exitosa")
        return True
    except Exception:
        log.warning("Timeout esperando seccion, URL actual: %s", driver.current_url)
        return False
