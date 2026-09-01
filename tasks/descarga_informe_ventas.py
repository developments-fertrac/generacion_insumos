from __future__ import annotations

import glob
import os
import time
from datetime import datetime
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from config.erp_selectors import (
    InformeVentas,
    Menu,
    Navegacion,
    Valorizado,
)
from core.browser import create_driver, do_login, take_xlsx_snapshot, wait_for_download_complete
from core.email_notifier import EmailNotifier
from core.erp_navigation import (
    click_with_fallback,
    find_all_visible,
    find_visible_element,
    open_menu_by_selectors,
    wait_for_erp_section,
)
from tasks.base_task import BaseTask

ERROR_IMAGES_DIR = Path(__file__).resolve().parent.parent / "errorImages"


class DescargaInformeVentas(BaseTask):
    name = "descarga_ventas"

    def setup(self) -> None:
        self.notifier = EmailNotifier(self.settings.smtp_ventas, self.name)
        self.paths = self.settings.paths
        self.erp = self.settings.erp
        self.browser_cfg = self.settings.browser
        self.paths.ensure_dirs()
        ERROR_IMAGES_DIR.mkdir(parents=True, exist_ok=True)

    def execute(self) -> None:
        carpeta_descarga = str(self.paths.ventas_mes)
        self.log.info("Carpeta de descarga: %s", carpeta_descarga)

        self._driver = create_driver(carpeta_descarga, self.browser_cfg)
        try:
            if not do_login(self._driver, self.erp, self.browser_cfg):
                raise RuntimeError("Fallo en el login")

            self._navegar_a_crm()

            if not self._abrir_menu_informes():
                raise RuntimeError("Fallo al abrir menu de informes")

            if not self._seleccionar_ventas_y_remisiones():
                raise RuntimeError("Fallo al seleccionar Ventas y remisiones")

            if not self._seleccionar_informe_facturas():
                raise RuntimeError("Fallo al seleccionar informe de facturas")

            if not self._configurar_fechas_ano_completo():
                self.log.warning("No se pudieron configurar las fechas del ano completo, continuando...")

            if not self._marcar_mostrar_costo():
                self.log.warning("No se pudo marcar 'Mostrar costo', continuando...")

            if not self._configurar_tipo_detallado():
                raise RuntimeError("Fallo al configurar tipo detallado")

            snapshot_previo = take_xlsx_snapshot(carpeta_descarga)
            self._limpiar_crdownload_huerfanos(carpeta_descarga)

            if not self._generar_xlsx():
                raise RuntimeError("Fallo al generar XLSX")

            ruta_archivo = wait_for_download_complete(
                carpeta_descarga,
                snapshot_previo,
                timeout=self.browser_cfg.timeout_descarga_completa,
                grace=self.browser_cfg.periodo_gracia,
                margin=self.browser_cfg.margen_snapshot,
            )

            if not ruta_archivo:
                raise RuntimeError("No se pudo verificar la descarga del archivo")

            self.log.info("Archivo descargado: %s", os.path.basename(ruta_archivo))
            self.log.info("Ubicacion: %s", carpeta_descarga)
        finally:
            if self._driver:
                self._driver.quit()
                self._driver = None
                self.log.info("Navegador cerrado")

    def teardown(self) -> None:
        pass

    # ── helpers ──────────────────────────────────────────────────────

    def _save_error_screenshot(self, name: str) -> None:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = ERROR_IMAGES_DIR / f"{name}_{timestamp}.png"
            self._driver.save_screenshot(str(screenshot_path))
            self.log.info("Screenshot guardado en: %s", screenshot_path)
        except Exception as e:
            self.log.warning("No se pudo guardar screenshot: %s", e)

    def _limpiar_crdownload_huerfanos(self, carpeta: str) -> None:
        huerfanos = glob.glob(os.path.join(carpeta, "*.crdownload"))
        huerfanos += glob.glob(os.path.join(carpeta, "*.tmp"))
        if not huerfanos:
            self.log.info("Sin temporales huerfanos que limpiar")
            return
        for h in huerfanos:
            try:
                os.remove(h)
                self.log.info("Limpiado temporal huerfano: %s", os.path.basename(h))
            except OSError as e:
                self.log.warning("No se pudo eliminar %s: %s", os.path.basename(h), e)

    # ── navegacion CRM ──────────────────────────────────────────────

    def _navegar_a_crm(self) -> None:
        self.log.info("Navegando a CRM/Ventas...")
        url_crm = self.erp.url_crm

        self._driver.get(url_crm)
        time.sleep(20)

        self.log.info("URL actual: %s", self._driver.current_url)

        if "crm.lead" in self._driver.current_url or "menu_id=371" in self._driver.current_url:
            self.log.info("En la seccion CRM/Ventas")
            return

        self.log.warning("Navegacion directa no funciono, intentando con JavaScript...")
        self._driver.execute_script(f"window.location.href = '{url_crm}';")
        time.sleep(20)

        self.log.info("URL actual despues de JavaScript: %s", self._driver.current_url)

        if "crm.lead" in self._driver.current_url or "menu_id=371" in self._driver.current_url:
            self.log.info("En la seccion CRM/Ventas")
            return

        self.log.warning("Intentando navegar por el menu...")
        try:
            wait = WebDriverWait(self._driver, 20)
            menu_crm = wait.until(EC.element_to_be_clickable(
                (By.XPATH, Navegacion.CRM)
            ))
            menu_crm.click()
            time.sleep(20)
            self.log.info("En la seccion CRM/Ventas")
        except Exception:
            self.log.warning("No se pudo navegar por el menu, continuando de todas formas...")
            self.log.info("URL actual: %s", self._driver.current_url)

    def _abrir_menu_informes(self) -> bool:
        self.log.info("Abriendo menu de Informes...")
        return open_menu_by_selectors(self._driver, Menu.INFORMES, "Informes", timeout=5)

    def _seleccionar_ventas_y_remisiones(self) -> bool:
        self.log.info("Seleccionando 'Ventas y remisiones'...")
        wait = WebDriverWait(self._driver, 5)

        try:
            time.sleep(2)

            for selector in Menu.VENTAS_Y_REMISIONES:
                try:
                    self.log.info("Intentando selector: %s...", selector[:50])
                    opcion = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    opcion.click()
                    time.sleep(2)
                    self.log.info("Opcion 'Ventas y remisiones' seleccionada")
                    return True
                except Exception:
                    continue

            self.log.warning("No se encontro con selectores normales, buscando todas las opciones del menu...")
            try:
                opciones = self._driver.find_elements(By.XPATH, Menu.VENTAS_DROPDOWN_ITEMS)
                self.log.info("Se encontraron %d opciones en el menu", len(opciones))

                for i, opc in enumerate(opciones[:10]):
                    texto = opc.text.strip()
                    if texto:
                        self.log.info("Opcion %d: %s", i + 1, texto)
                        if "ventas" in texto.lower() or "remision" in texto.lower():
                            self.log.info("Intentando hacer click en: %s", texto)
                            opc.click()
                            time.sleep(2)
                            self.log.info("Opcion seleccionada")
                            return True
            except Exception as e_list:
                self.log.error("Error listando opciones: %s", e_list)

            raise RuntimeError("No se encontro la opcion 'Ventas y remisiones' en el menu")

        except Exception as e:
            self.log.error("Error seleccionando 'Ventas y remisiones': %s", e)
            self._save_error_screenshot("error_menu_ventas")
            return False

    def _seleccionar_informe_facturas(self, max_intentos: int = 3) -> bool:
        wait = WebDriverWait(self._driver, 10)

        for intento in range(1, max_intentos + 1):
            try:
                self.log.info("Intento %d/%d: Seleccionando 'Informes de ventas (Facturas)'...", intento, max_intentos)

                if intento == 1:
                    self.log.info("Esperando modal SINFO...")
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog")))
                    time.sleep(2)
                    self.log.info("Modal SINFO encontrado")

                self.log.info("Esperando contenido del modal cargar...")
                time.sleep(3)

                self.log.info("Buscando campo 'Seleccione un informe' con selector principal...")
                campo_informe = find_visible_element(self._driver, [InformeVentas.CAMPO_INFORME[0]], timeout=15)

                if campo_informe:
                    self.log.info("Campo encontrado con selector principal")
                else:
                    self.log.info("Selector principal no funciono, intentando selectores alternativos...")
                    campo_informe = find_visible_element(self._driver, InformeVentas.CAMPO_INFORME[1:], timeout=10)

                if not campo_informe:
                    self.log.info("Selectores de lista no funcionaron, buscando inputs visibles en modal...")
                    all_inputs = self._driver.find_elements(By.XPATH, "//div[contains(@class, 'modal')]//input")
                    self.log.info("Inputs encontrados en modal: %d", len(all_inputs))
                    for i, inp in enumerate(all_inputs):
                        try:
                            input_type = inp.get_attribute("type") or "text"
                            input_id = inp.get_attribute("id") or ""
                            input_name = inp.get_attribute("name") or ""
                            input_placeholder = inp.get_attribute("placeholder") or ""
                            is_displayed = inp.is_displayed()
                            self.log.info("  Input %d: type=%s, id=%s, name=%s, placeholder=%s, visible=%s",
                                          i, input_type, input_id, input_name, input_placeholder, is_displayed)
                            if is_displayed and input_type not in ("hidden", "checkbox", "radio"):
                                campo_informe = inp
                                self.log.info("  -> Seleccionado como campo de informe")
                                break
                        except Exception as e:
                            self.log.warning("  Input %d: error leyendo atributos: %s", i, e)

                if not campo_informe:
                    self._save_error_screenshot("error_campo_informe_no_encontrado")
                    raise RuntimeError("No se encontro el campo de busqueda")

                self.log.info("Haciendo scroll al campo...")
                self._driver.execute_script("arguments[0].scrollIntoView(true);", campo_informe)
                time.sleep(1)

                self.log.info("Limpiando campo...")
                campo_informe.clear()
                time.sleep(0.5)

                self.log.info("Escribiendo 'Informes de ventas' en el campo...")
                campo_informe.send_keys("Informes de ventas")
                time.sleep(3)

                valor_campo = campo_informe.get_attribute("value") or ""
                self.log.info("Valor actual del campo: '%s'", valor_campo)

                self.log.info("Texto escrito, verificando dropdown...")

                self.log.info("Esperando a que aparezcan las opciones del dropdown...")
                try:
                    wait_short = WebDriverWait(self._driver, 8)
                    wait_short.until(EC.presence_of_element_located(
                        (By.XPATH, InformeVentas.AUTOCOMPLETE_WAIT)
                    ))
                    self.log.info("Opciones del dropdown detectadas")
                except Exception:
                    self.log.warning("Timeout esperando opciones, verificando DOM...")
                    dropdown_html = self._driver.execute_script("""
                        var lists = document.querySelectorAll('ul.ui-autocomplete, ul[role="listbox"], ul.dropdown-menu');
                        var result = [];
                        lists.forEach(function(ul) {
                            result.push({
                                class: ul.className,
                                visible: ul.offsetParent !== null,
                                items: ul.querySelectorAll('li').length
                            });
                        });
                        return result;
                    """)
                    self.log.info("Listas encontradas en DOM: %s", dropdown_html)

                time.sleep(1)

                self.log.info("Buscando opciones del dropdown...")
                opciones = find_all_visible(self._driver, InformeVentas.AUTOCOMPLETE_OPCIONES[0])

                if not opciones:
                    for selector in InformeVentas.AUTOCOMPLETE_OPCIONES[1:]:
                        opciones = find_all_visible(self._driver, selector)
                        if opciones:
                            self.log.info("Opciones encontradas con selector: %s", selector[:60])
                            break

                if not opciones:
                    self.log.info("Buscando opciones con JavaScript...")
                    opciones_js = self._driver.execute_script("""
                        var items = [];
                        var lists = document.querySelectorAll('ul.ui-autocomplete li, ul[role="listbox"] li, ul.dropdown-menu li');
                        lists.forEach(function(li) {
                            if (li.offsetParent !== null && li.textContent.trim()) {
                                items.push({text: li.textContent.trim(), tag: li.tagName});
                            }
                        });
                        return items;
                    """)
                    self.log.info("Opciones via JavaScript: %s", opciones_js[:10] if opciones_js else [])

                if opciones:
                    self.log.info("Opciones disponibles (%d):", len(opciones))
                    for i, opcion in enumerate(opciones[:15]):
                        texto = opcion.text.strip()
                        self.log.info("    %d. '%s'", i + 1, texto)

                    self.log.info("Buscando 'Informes de ventas (Facturas)'...")
                    for opcion in opciones:
                        texto = opcion.text.strip()
                        texto_lower = texto.lower()

                        if "informes de ventas (facturas)" in texto_lower or (
                            ("factura" in texto_lower or "facturas" in texto_lower)
                            and "ventas" in texto_lower
                            and "informe" in texto_lower
                        ):
                            self.log.info("Encontrada opcion: '%s'", texto)
                            click_with_fallback(self._driver, opcion, "informe de facturas")
                            time.sleep(2)
                            return True

                    self.log.warning("No se encontro coincidencia exacta, buscando por 'factura' o 'ventas'...")
                    for opcion in opciones:
                        texto = opcion.text.strip()
                        texto_lower = texto.lower()
                        if "factura" in texto_lower or "ventas" in texto_lower:
                            self.log.info("Intentando con: '%s'", texto)
                            click_with_fallback(self._driver, opcion, "informe de facturas")
                            time.sleep(2)
                            return True

                if not opciones:
                    self.log.warning("No se encontraron opciones visibles en el dropdown")
                    self._save_error_screenshot("error_dropdown_vacio")

                if intento < max_intentos:
                    self.log.warning(
                        "No se encontraron opciones en intento %d, reintentando en 3 segundos...", intento
                    )
                    time.sleep(3)
                    continue
                else:
                    raise RuntimeError(
                        "No se encontraron opciones en el dropdown despues de multiples intentos"
                    )

            except Exception as e:
                if intento < max_intentos:
                    self.log.warning("Error en intento %d: %s", intento, e)
                    self.log.info("Reintentando en 5 segundos...")
                    time.sleep(5)
                else:
                    self.log.error(
                        "Error seleccionando informe de facturas despues de %d intentos: %s",
                        max_intentos,
                        e,
                    )
                    self._save_error_screenshot("error_informe_facturas")
                    return False

        self.log.error(
            "No se pudo seleccionar 'Informes de ventas (Facturas)' despues de %d intentos", max_intentos
        )
        return False

    def _configurar_fechas_ano_completo(self) -> bool:
        anio_actual = datetime.now().year
        fecha_desde = f"01/01/{anio_actual}"
        fecha_hasta = f"31/12/{anio_actual}"

        self.log.info("Configurando rango de fechas: %s - %s (ano completo)...", fecha_desde, fecha_hasta)

        try:
            campo_desde = find_visible_element(self._driver, InformeVentas.FECHA_DESDE, timeout=10)
            if not campo_desde:
                inputs_fecha = find_all_visible(self._driver, InformeVentas.FECHA_INPUT_MODAL)
                if len(inputs_fecha) >= 2:
                    campo_desde = inputs_fecha[0]
                    self.log.info("Campo 'Fecha desde' encontrado por fallback (primer input de fecha)")
                else:
                    inputs_modal = find_all_visible(self._driver, InformeVentas.FECHA_INPUT_GENERICO)
                    if len(inputs_modal) >= 2:
                        campo_desde = inputs_modal[0]
                        self.log.info("Campo 'Fecha desde' encontrado por fallback generico")

            if campo_desde:
                self._driver.execute_script("arguments[0].scrollIntoView(true);", campo_desde)
                time.sleep(0.3)
                campo_desde.click()
                time.sleep(0.3)
                campo_desde.clear()
                time.sleep(0.2)
                self._driver.execute_script(
                    """
                    arguments[0].value = '';
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    """,
                    campo_desde,
                )
                time.sleep(0.2)
                campo_desde.send_keys(fecha_desde)
                time.sleep(0.5)
                campo_desde.send_keys(Keys.TAB)
                time.sleep(1)
                self.log.info("'Fecha desde' configurada: %s", fecha_desde)
            else:
                self.log.warning("No se encontro campo 'Fecha desde'")
                return False

            campo_hasta = find_visible_element(self._driver, InformeVentas.FECHA_HASTA, timeout=10)
            if not campo_hasta:
                inputs_fecha = find_all_visible(self._driver, InformeVentas.FECHA_INPUT_MODAL)
                if len(inputs_fecha) >= 2:
                    campo_hasta = inputs_fecha[1]
                    self.log.info("Campo 'Fecha hasta' encontrado por fallback (segundo input de fecha)")
                else:
                    inputs_modal = find_all_visible(self._driver, InformeVentas.FECHA_INPUT_GENERICO)
                    if len(inputs_modal) >= 2:
                        campo_hasta = inputs_modal[1]
                        self.log.info("Campo 'Fecha hasta' encontrado por fallback generico")

            if campo_hasta:
                self._driver.execute_script("arguments[0].scrollIntoView(true);", campo_hasta)
                time.sleep(0.3)
                campo_hasta.click()
                time.sleep(0.3)
                campo_hasta.clear()
                time.sleep(0.2)
                self._driver.execute_script(
                    """
                    arguments[0].value = '';
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    """,
                    campo_hasta,
                )
                time.sleep(0.2)
                campo_hasta.send_keys(fecha_hasta)
                time.sleep(0.5)
                campo_hasta.send_keys(Keys.TAB)
                time.sleep(1)
                self.log.info("'Fecha hasta' configurada: %s", fecha_hasta)
            else:
                self.log.warning("No se encontro campo 'Fecha hasta'")
                return False

            self.log.info("Rango de fechas configurado: %s - %s", fecha_desde, fecha_hasta)
            return True

        except Exception as e:
            self.log.error("Error configurando fechas: %s", e)
            self._save_error_screenshot("error_fechas")
            return False

    def _marcar_mostrar_costo(self) -> bool:
        self.log.info("Marcando casilla 'Mostrar costo'...")

        try:
            time.sleep(2)

            self.log.info("Buscando label 'Mostrar costo'...")
            try:
                label = find_visible_element(self._driver, InformeVentas.CHECKBOX_MOSTRAR_COSTO, timeout=5)
                if label:
                    self.log.info("Label 'Mostrar costo' encontrado")
                    self._driver.execute_script("arguments[0].scrollIntoView(true);", label)
                    time.sleep(0.5)

                    self.log.info("Haciendo click en el label...")
                    if click_with_fallback(self._driver, label, "Mostrar costo"):
                        self.log.info("Click en label ejecutado")
                        time.sleep(1)
                        return True
            except Exception as e_label:
                self.log.warning("No se pudo hacer click en el label: %s", e_label)

            self.log.info("Buscando texto 'Mostrar costo' para hacer click...")
            try:
                elementos = find_all_visible(self._driver, "//*[contains(text(), 'Mostrar costo')]")
                self.log.info("Se encontraron %d elementos con texto 'Mostrar costo'", len(elementos))

                for elem in elementos:
                    if elem.is_displayed():
                        self.log.info("Elemento visible encontrado, haciendo click...")
                        self._driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                        time.sleep(0.5)
                        if click_with_fallback(self._driver, elem, "Mostrar costo"):
                            time.sleep(1)
                            return True
            except Exception as e_text:
                self.log.warning("Error con estrategia de texto: %s", e_text)

            self.log.info("Intentando con JavaScript puro...")
            try:
                script = """
                var checkboxes = document.querySelectorAll('input[type="checkbox"]');
                for (var i = 0; i < checkboxes.length; i++) {
                    var cb = checkboxes[i];
                    var parent = cb.parentElement;
                    var text = parent.textContent || parent.innerText;
                    if (text.includes('Mostrar costo') || text.includes('mostrar costo')) {
                        if (!cb.checked) {
                            cb.checked = true;
                            cb.dispatchEvent(new Event('change', { bubbles: true }));
                            cb.dispatchEvent(new Event('click', { bubbles: true }));
                        }
                        return 'MARCADO';
                    }
                }
                return 'NO_ENCONTRADO';
                """

                resultado = self._driver.execute_script(script)
                self.log.info("Resultado de JavaScript: %s", resultado)

                if resultado == "MARCADO":
                    self.log.info("Checkbox marcado con JavaScript puro")
                    time.sleep(1)
                    return True
                else:
                    self.log.warning("JavaScript no pudo encontrar el checkbox")
            except Exception as e_js:
                self.log.warning("Error con JavaScript: %s", e_js)

            self.log.info("Buscando input por atributos...")
            try:
                for selector in InformeVentas.CHECKBOX_MOSTRAR_COSTO_INPUTS:
                    try:
                        cb = self._driver.find_element(By.XPATH, selector)
                        if cb.is_displayed():
                            self.log.info("Checkbox encontrado con selector: %s...", selector[:50])
                            self._driver.execute_script("arguments[0].scrollIntoView(true);", cb)
                            time.sleep(0.5)
                            self._driver.execute_script(
                                """
                                arguments[0].checked = true;
                                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                                arguments[0].dispatchEvent(new Event('click', { bubbles: true }));
                                """,
                                cb,
                            )
                            self.log.info("Checkbox marcado")
                            time.sleep(1)
                            return True
                    except Exception:
                        continue
            except Exception as e_attr:
                self.log.warning("Error con busqueda por atributos: %s", e_attr)

            self.log.warning("No se pudo marcar 'Mostrar costo' con ninguna estrategia")
            return False

        except Exception as e:
            self.log.error("Error general marcando 'Mostrar costo': %s", e)
            self._save_error_screenshot("error_mostrar_costo")
            return False

    def _configurar_tipo_detallado(self) -> bool:
        self.log.info("Configurando tipo de informe a 'Detallado'...")

        try:
            dropdown_tipo = find_visible_element(self._driver, InformeVentas.SELECT_TIPO_INFORME, timeout=10)

            if dropdown_tipo is None:
                raise RuntimeError("No se encontro el dropdown de Tipo de informe")

            select_tipo = Select(dropdown_tipo)

            try:
                select_tipo.select_by_visible_text("Detallado")
            except Exception:
                try:
                    select_tipo.select_by_value("detailed")
                except Exception:
                    select_tipo.select_by_index(1)

            time.sleep(1)
            self.log.info("Tipo configurado a 'Detallado'")
            return True

        except Exception as e:
            self.log.error("Error configurando tipo detallado: %s", e)
            return False

    def _generar_xlsx(self) -> bool:
        self.log.info("Generando archivo XLSX...")
        wait = WebDriverWait(self._driver, 10)

        try:
            boton_xlsx = wait.until(
                EC.element_to_be_clickable((By.XPATH, InformeVentas.BOTON_GENERAR_XLSX))
            )
            boton_xlsx.click()

            self.log.info("Esperando descarga...")
            time.sleep(self.browser_cfg.timeout_descarga)
            self.log.info("Descarga iniciada")
            return True

        except Exception as e:
            self.log.error("Error generando XLSX: %s", e)
            return False
