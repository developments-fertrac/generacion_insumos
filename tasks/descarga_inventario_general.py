"""Tarea: descargar informe de Productos (Inventario General) desde ERP Fertrac.

Navega a Inventario > Productos, cambia vista a lista, detecta total de
registros, modifica el rango para mostrar todos, espera la carga completa,
selecciona todos los registros y exporta el fichero INVENTARIO GENERAL.
"""

from __future__ import annotations

import os
import re
import subprocess
import time
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from config.erp_selectors import (
    Checkbox,
    Menu,
    Modal,
    Navegacion,
    Paginador,
    VistaLista,
)
from config.settings import Settings
from core.browser import create_driver, do_login, take_xlsx_snapshot, wait_for_download_complete
from core.erp_navigation import (
    close_modal,
    click_with_fallback,
    find_all_visible,
    find_visible_element,
    get_pager_info,
    open_action_menu,
    open_menu_by_selectors,
    select_export_file_option,
    select_export_option,
    select_from_modal_dropdown,
    wait_for_erp_section,
    wait_for_loading_complete,
)
from core.email_notifier import EmailNotifier
from tasks.base_task import BaseTask


class DescargaInventarioGeneral(BaseTask):
    name = "descarga_inv_general"

    def setup(self) -> None:
        self.notifier = EmailNotifier(self.settings.smtp_inv_general, self.name)
        self.paths = self.settings.paths
        self.erp = self.settings.erp
        self.browser_cfg = self.settings.browser

    def execute(self) -> None:
        self._driver = None
        self.paths.ensure_dirs()
        carpeta_descarga = str(self.paths.inventario_general_mes)
        self.log.info("Carpeta de descarga: %s", carpeta_descarga)

        self._verificar_versiones_chrome()

        self._driver = create_driver(carpeta_descarga, self.browser_cfg)

        if not do_login(self._driver, self.erp, self.browser_cfg):
            raise RuntimeError("Fallo en el login")

        self._navegar_a_inventario()

        if not self._seleccionar_productos():
            raise RuntimeError("Fallo al seleccionar Productos")

        if not self._cambiar_a_vista_lista():
            raise RuntimeError("Fallo al cambiar a vista lista")

        total, _ = self._detectar_total_registros()
        if total is None:
            raise RuntimeError("Fallo al detectar total de registros")

        if not self._modificar_rango_registros(total):
            raise RuntimeError("Fallo al modificar rango de registros")

        if not wait_for_loading_complete(
            self._driver,
            timeout=self.browser_cfg.timeout_carga_maxima,
            log_func=self.log.info,
        ):
            raise RuntimeError("Fallo al esperar carga completa")

        self.log.info("=" * 70)
        self.log.info("INICIANDO EXPORTACION")

        snapshot = take_xlsx_snapshot(carpeta_descarga)

        if not self._seleccionar_todos_registros():
            raise RuntimeError("Fallo al seleccionar todos los registros")

        if not open_action_menu(self._driver):
            raise RuntimeError("Fallo al abrir menu Accion")

        if not select_export_option(self._driver):
            raise RuntimeError("Fallo al seleccionar Exportar")

        if not select_from_modal_dropdown(self._driver, "INVENTARIO GENERAL", "INVENTARIO GENERAL"):
            raise RuntimeError("Fallo al seleccionar INVENTARIO GENERAL")

        if not select_export_file_option(self._driver):
            raise RuntimeError("Fallo al exportar fichero")

        self.log.info("=" * 70)
        self.log.info("ESPERANDO PROCESAMIENTO DE EXPORTACION")
        self._esperar_procesamiento_post_exportacion()

        self.log.info("=" * 70)
        self.log.info("ESPERANDO DESCARGA DEL ARCHIVO")
        archivo_descargado = wait_for_download_complete(
            carpeta_descarga,
            snapshot,
            timeout=self.browser_cfg.timeout_descarga_completa,
            grace=self.browser_cfg.periodo_gracia,
        )

        if not archivo_descargado:
            raise RuntimeError("No se pudo verificar la descarga del archivo")

        archivo_final = self._renombrar_archivo_con_fecha(archivo_descargado)
        self.log.info("=" * 70)
        self.log.info("PROCESO COMPLETADO: %s", os.path.basename(archivo_final))

    def teardown(self) -> None:
        if hasattr(self, "_driver") and self._driver:
            try:
                self._driver.quit()
            except Exception:
                pass
            self.log.info("Navegador cerrado")

    # ------------------------------------------------------------------ helpers

    @staticmethod
    def _verificar_versiones_chrome() -> None:
        from core.logger import get_logger

        log = get_logger("chrome_versions")
        log.info("Verificando versiones de Chrome y ChromeDriver...")
        try:
            result = subprocess.run(
                ["reg", "query",
                 r"HKLM\SOFTWARE\Google\Chrome\BLBeacon", "/v", "version"],
                capture_output=True, text=True, timeout=5,
            )
            if result.returncode == 0:
                version_chrome = result.stdout.strip().split()[-1]
            else:
                result2 = subprocess.run(
                    ["reg", "query",
                     r"HKCU\SOFTWARE\Google\Chrome\BLBeacon", "/v", "version"],
                    capture_output=True, text=True, timeout=5,
                )
                version_chrome = result2.stdout.strip().split()[-1] if result2.returncode == 0 else "no encontrada"
            log.info("Chrome instalado : %s", version_chrome)
        except Exception as e:
            log.warning("No se pudo leer version de Chrome: %s", e)

        try:
            result = subprocess.run(
                ["chromedriver", "--version"],
                capture_output=True, text=True, timeout=5,
            )
            version_driver = result.stdout.strip() if result.returncode == 0 else "no encontrado en PATH"
            log.info("ChromeDriver     : %s", version_driver)
        except Exception as e:
            log.warning("No se pudo leer version de ChromeDriver: %s", e)

    def _navegar_a_inventario(self) -> None:
        self.log.info("Navegando a Inventario...")
        self._driver.get(self.erp.url_inventario)
        wait_for_erp_section(self._driver, self.erp.url_inventario, Navegacion.INVENTARIO, timeout=30)

    def _seleccionar_productos(self) -> bool:
        self.log.info("Seleccionando 'Datos principales' > 'Productos'...")

        if not open_menu_by_selectors(self._driver, Menu.DATOS_PRINCIPALES, "Datos principales"):
            return False

        wait = WebDriverWait(self._driver, 15)
        for selector in Menu.PRODUCTOS:
            try:
                productos = wait.until(EC.element_to_be_clickable(("xpath", selector)))
                productos.click()
                self.log.info("'Productos' seleccionado")

                WebDriverWait(self._driver, 15).until(
                    lambda d: "product.template" in d.current_url
                    or "action=278" in d.current_url
                    or d.current_url != self.erp.url_inventario
                )

                url_actual = self._driver.current_url
                if "product.template" not in url_actual and "action=278" not in url_actual:
                    self.log.warning("URL inesperada, navegando directamente a Productos...")
                    self._driver.get(self.erp.url_productos)
                    WebDriverWait(self._driver, 20).until(
                        EC.presence_of_element_located(("xpath", "//div[contains(@class,'o_content')]"))
                    )

                self.log.info("En Productos. URL: %s", self._driver.current_url)
                return True
            except Exception:
                continue

        self.log.error("No se encontro la opcion 'Productos'")
        return False

    def _cambiar_a_vista_lista(self) -> bool:
        self.log.info("Verificando/cambiando a vista Lista...")
        try:
            time.sleep(2)

            for selector in VistaLista.BOTONES:
                try:
                    botones = self._driver.find_elements(By.XPATH, selector)
                    for boton in botones:
                        if boton.is_displayed():
                            clases = boton.get_attribute("class") or ""
                            aria_pressed = boton.get_attribute("aria-pressed") or ""
                            if "active" in clases or "btn-primary" in clases or aria_pressed == "true":
                                self.log.info("Ya estamos en vista Lista")
                                return True
                            self._driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                            time.sleep(0.3)
                            click_with_fallback(self._driver, boton, "vista Lista")
                            time.sleep(2)
                            self.log.info("Vista Lista activada")
                            return True
                except Exception:
                    continue

            url_actual = self._driver.current_url
            if "view_type=kanban" in url_actual:
                self._driver.get(url_actual.replace("view_type=kanban", "view_type=list"))
                WebDriverWait(self._driver, 20).until(
                    EC.presence_of_element_located(("xpath", VistaLista.CONTENEDOR))
                )
                self.log.info("Vista lista forzada por URL")
                return True

            self.log.info("Asumiendo que ya estamos en vista adecuada")
            return True

        except Exception as e:
            self.log.error("Error cambiando a vista lista: %s", e)
            return False

    def _detectar_total_registros(self) -> tuple:
        self.log.info("Detectando total de registros...")
        try:
            WebDriverWait(self._driver, 20).until(
                EC.presence_of_element_located(("xpath", Paginador.CONTENEDOR))
            )

            total = None
            final = None
            deadline = time.time() + self.browser_cfg.timeout_detectar_total
            while time.time() < deadline:
                final, total = get_pager_info(self._driver)
                if total is not None:
                    break
                time.sleep(2)

            if total is not None:
                self.log.info("Total de registros detectado: %s", total)

                if total < 10000:
                    self.log.warning("Solo %d registros, reintentando...", total)
                    self._driver.get(self.erp.url_productos)
                    time.sleep(5)
                    _, retry_total = get_pager_info(self._driver)
                    if retry_total is not None and retry_total >= 10000:
                        self.log.info("Registros correctos: %d", retry_total)
                        return retry_total, None
                    raise RuntimeError(f"Vista incorrecta: solo {total} registros. URL: {self._driver.current_url}")

                if final is not None:
                    self.log.info("Rango actual: %d/%d", final, total)
                return total, None

            raise RuntimeError("No se pudo detectar el total de registros")

        except Exception as e:
            self.log.error("Error detectando total: %s", e)
            return None, None

    def _modificar_rango_registros(self, total: int) -> bool:
        self.log.info("Modificando rango a 1-%d...", total)
        try:
            time.sleep(1)

            nuevo_rango = f"1-{total}"
            self.log.info("Escribiendo: %s", nuevo_rango)

            for intento in range(1, 3):
                campo_rango = None
                for selector in [Paginador.RANGO_VALUE, Paginador.RANGO_DISPLAY]:
                    try:
                        elementos = self._driver.find_elements(By.XPATH, selector)
                        if selector == Paginador.RANGO_VALUE:
                            self.log.info(
                                "Elementos o_pager_value encontrados: %d", len(elementos)
                            )
                        for elem in elementos:
                            try:
                                if not elem.is_displayed():
                                    continue
                                texto = (elem.text or "").strip()
                                if re.match(r"^\d+-\d+$", texto) or (
                                    selector == Paginador.RANGO_VALUE
                                    and re.match(r"^\d+$", texto)
                                ):
                                    self.log.info("Campo de rango encontrado: '%s'", texto)
                                    campo_rango = elem
                                    break
                            except Exception:
                                continue
                        if campo_rango:
                            break
                    except Exception:
                        continue

                if not campo_rango:
                    raise RuntimeError("No se encontro el campo de rango")

                self._driver.execute_script(
                    "arguments[0].scrollIntoView(true);", campo_rango
                )
                time.sleep(0.3)

                click_with_fallback(self._driver, campo_rango, "campo de rango")
                time.sleep(0.5)

                # Tras el clic, Odoo sustituye el span por un <input>: siempre hay
                # que obtener una referencia fresca (la anterior queda stale).
                input_rango = None
                plazo_input = time.time() + 5
                while time.time() < plazo_input and input_rango is None:
                    try:
                        input_rango = find_visible_element(
                            self._driver,
                            [
                                "//*[contains(@class, 'o_pager')]//input",
                                "//div[contains(@class, 'o_cp_pager')]//input",
                                "//input[contains(@class, 'o_pager_value')]",
                            ],
                            timeout=1,
                        )
                    except Exception:
                        input_rango = None
                    if input_rango is None:
                        time.sleep(0.3)

                if input_rango is not None:
                    self.log.info("Input editable detectado (intento %d)", intento)
                    # Evitar clear()/send_keys() sobre el input (se vuelve stale en
                    # el re-render): fijar el valor por JS y confirmar con Enter.
                    self._driver.execute_script(
                        """
                        var el = arguments[0];
                        el.focus();
                        el.value = arguments[1];
                        el.dispatchEvent(new Event('input', {bubbles: true}));
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                        """,
                        input_rango,
                        nuevo_rango,
                    )
                    time.sleep(0.3)
                    ActionChains(self._driver).send_keys(Keys.RETURN).perform()
                else:
                    self.log.warning(
                        "No aparecio input editable (intento %d), usando teclado...",
                        intento,
                    )
                    nuevo_valor = find_visible_element(
                        self._driver, [Paginador.RANGO_VALUE], timeout=3
                    )
                    if nuevo_valor is not None:
                        try:
                            click_with_fallback(
                                self._driver, nuevo_valor, "campo de rango"
                            )
                        except Exception:
                            pass
                    time.sleep(0.3)
                    try:
                        ActionChains(self._driver).send_keys(nuevo_rango).perform()
                        time.sleep(0.5)
                        ActionChains(self._driver).send_keys(Keys.RETURN).perform()
                    except Exception as e:
                        self.log.warning("Fallback teclado fallo: %s", e)

                time.sleep(2)

                _, total_verificado = get_pager_info(self._driver)
                if total_verificado is not None:
                    self.log.info(
                        "Total tras modificar rango: %d (intento %d)",
                        total_verificado,
                        intento,
                    )
                if (
                    total_verificado is not None
                    and total_verificado >= total * 0.9
                ):
                    self.log.info("Rango modificado a %s", nuevo_rango)
                    return True

                self.log.warning("El rango no se aplico, reintentando...")
                time.sleep(1)

            raise RuntimeError("No se pudo modificar el rango a %s" % nuevo_rango)

        except Exception as e:
            self.log.error("Error modificando rango: %s", e)
            return False

    def _seleccionar_todos_registros(self) -> bool:
        self.log.info("Seleccionando todos los registros...")
        try:
            checkbox_header = find_visible_element(self._driver, Checkbox.HEADER, timeout=10)

            if not checkbox_header:
                checkbox_header = self._driver.execute_script(Checkbox.HEADER_JS)
                if checkbox_header:
                    self.log.info("Checkbox encontrado con JavaScript")

            if not checkbox_header:
                raise RuntimeError("No se pudo encontrar el checkbox del header")

            self._driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_header)
            time.sleep(0.3)

            try:
                checkbox_header.click()
                self.log.info("Click en checkbox realizado")
            except Exception as e:
                self.log.warning("Click normal fallo: %s, intentando con JS...", e)
                try:
                    self._driver.execute_script("arguments[0].click();", checkbox_header)
                    self.log.info("Click con JavaScript realizado")
                except Exception as js_error:
                    if "timeout" in str(js_error).lower():
                        self.log.warning("Timeout en JS (normal con muchos registros), continuando...")
                    else:
                        raise

            self.log.info("Esperando procesamiento de seleccion (hasta 5 minutos)...")
            for minuto in range(5):
                time.sleep(60)
                self.log.info("%d/5 minutos esperados...", minuto + 1)

            selectores_accion = [
                "//button[contains(., 'Accion')]",
                "//button[contains(., 'Acción')]",
                "//button[contains(text(), 'Acción')]",
            ]

            for selector in selectores_accion:
                try:
                    botones = self._driver.find_elements(By.XPATH, selector)
                    for boton in botones:
                        if boton.is_displayed():
                            self.log.info("Boton 'Accion' visible: '%s'", boton.text.strip())
                            return True
                except Exception:
                    continue

            raise RuntimeError("El boton 'Accion' NO aparece - registros no seleccionados")

        except Exception as e:
            self.log.error("Error seleccionando registros: %s", e)
            return False

    def _esperar_procesamiento_post_exportacion(self) -> None:
        tiempo_export = time.time()
        while time.time() - tiempo_export < 600:
            try:
                spinners = find_all_visible(self._driver, "//span[contains(@class, 'fa-spinner')] | //div[contains(@class, 'o_loading')]")
                if spinners:
                    minutos = int((time.time() - tiempo_export) / 60)
                    if minutos > 0:
                        self.log.info("Procesando exportacion... %d minuto(s)", minutos)
                    time.sleep(5)
                else:
                    self.log.info("Procesamiento completado")
                    break
            except Exception:
                time.sleep(2)

    def _renombrar_archivo_con_fecha(self, archivo_original: str) -> str:
        self.log.info("Renombrando archivo con fecha...")
        try:
            directorio = os.path.dirname(archivo_original)
            extension = os.path.splitext(archivo_original)[1]

            fecha_actual = datetime.now()
            meses = {
                1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
                5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
                9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
            }

            nuevo_nombre = (
                f"INVENTARIO GENERAL ACTUALIZADO "
                f"{fecha_actual.day:02d} DE {meses[fecha_actual.month]} DE {fecha_actual.year}"
                f"{extension}"
            )
            nueva_ruta = os.path.join(directorio, nuevo_nombre)

            # Reintentar el rename: puede fallar si el archivo aun esta en uso
            # (Excel/antivirus/temporal). El nombre original queda disponible hasta
            # que el proceso de exportacion deja de escribirlo.
            for _ in range(60):
                try:
                    if os.path.exists(nueva_ruta):
                        os.remove(nueva_ruta)
                    os.rename(archivo_original, nueva_ruta)
                    break
                except Exception as e:
                    time.sleep(2)
                    ultimo_error = e
            else:
                self.log.error("No se pudo renombrar a '%s' tras varias esperas: %s", nuevo_nombre, ultimo_error)
                return archivo_original

            self.log.info("Archivo renombrado a: %s", nuevo_nombre)
            return nueva_ruta

        except Exception as e:
            self.log.error("Error renombrando archivo: %s", e)
            return archivo_original
