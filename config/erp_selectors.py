"""Selectores XPath centralizados para la navegacion en ERP Fertrac.

Organizados por tipo de accion: menus, dropdowns, botones, modales, etc.
Cada grupo contiene una lista de selectores en orden de prioridad (el primero
que funcione se usa).
"""


class Navegacion:
    INVENTARIO = (
        "//nav | //div[contains(@class,'o_content')] | "
        "//div[contains(@class,'o_kanban')]"
    )
    CRM = (
        "//a[contains(@data-menu, '371')] | //a[contains(., 'CRM')] | "
        "//span[contains(., 'CRM')]"
    )
    POST_LOGIN = (
        "//nav | //div[contains(@class,'o_main_navbar')] | "
        "//div[contains(@class,'o_menu')]"
    )


class Menu:
    DATOS_PRINCIPALES = [
        "//a[contains(text(), 'Datos principales')]",
        "//span[contains(text(), 'Datos principales')]",
        "//div[contains(text(), 'Datos principales')]",
        "//button[contains(., 'Datos principales')]",
    ]
    PRODUCTOS = [
        "//a[normalize-space(text())='Productos']",
        "//a[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
        "//span[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
    ]
    INFORMES = [
        "//a[normalize-space(text())='Informes']",
        "//span[normalize-space(text())='Informes']",
        "//li[normalize-space(.)='Informes']/a",
        "//div[normalize-space(text())='Informes']",
        "//button[normalize-space(.)='Informes']",
        "//*[normalize-space(text())='Informes']",
    ]
    VENTAS_Y_REMISIONES = [
        "//a[contains(@data-menu-xmlid, 'co_reports.menu_reports_sales')]/span[normalize-space(.)='Ventas y remisiones']",
        "//span[contains(text(), 'Ventas y remisiones')]",
    ]
    VENTAS_DROPDOWN_ITEMS = (
        "//div[contains(@class, 'dropdown')]//a | "
        "//ul[contains(@class, 'dropdown')]//li"
    )


class VistaLista:
    BOTONES = [
        "//button[contains(@class, 'o_cp_switch_list')]",
        "//button[@data-view-type='list']",
        "//button[contains(@title, 'List')]",
        "//button[contains(@title, 'Lista')]",
        "//i[contains(@class, 'fa-list-ul')]/parent::button",
        "//i[contains(@class, 'oi-view-list')]/parent::button",
        "//button[contains(@class, 'o_list')]",
    ]
    CONTENEDOR = "//table | //div[contains(@class,'o_list')]"



class Paginador:
    CONTENEDOR = (
        "//span[contains(@class, 'o_pager')] | "
        "//div[contains(@class, 'o_cp_pager')]"
    )
    TEXTO = "//*[contains(text(), '/')]"
    RANGO_VALUE = "//span[contains(@class, 'o_pager_value')]"
    RANGO_LIMIT = "//span[contains(@class, 'o_pager_limit')]"
    RANGO_DISPLAY = "//*[contains(text(), '1-')]"
    DETECT_CARGA = (
        "//span[contains(@class, 'o_pager')] | "
        "//span[contains(@class, 'o_pager_value')]"
    )


class Carga:
    MENSAJES = (
        "//*[contains(text(), 'Cargando') or contains(text(), 'Loading') or "
        "contains(text(), 'Procesando') or contains(text(), 'Espere')]"
    )
    SPINNERS = (
        "//span[contains(@class, 'fa-spinner')] | "
        "//div[contains(@class, 'o_loading')]"
    )
    OVERLAY_JS = """
        var overlays = document.querySelectorAll(
            '[class*="blockUI"], [class*="o_loading"], [class*="modal-backdrop"], .o_blockUI, .blockUI'
        );
        for (var i = 0; i < overlays.length; i++) {
            var style = window.getComputedStyle(overlays[i]);
            if (style.display !== 'none' && style.visibility !== 'hidden') {
                var rect = overlays[i].getBoundingClientRect();
                if (rect.width > 500 && rect.height > 500) return true;
            }
        }
        return false;
    """


class Checkbox:
    HEADER = [
        "//thead//th[1]//input[@type='checkbox']",
        "//th[@class='o_list_record_selector']//input[@type='checkbox']",
        "//table//thead//th//input[@type='checkbox']",
    ]
    HEADER_JS = """
        var c = document.querySelector('thead th input[type="checkbox"]');
        if (c) return c;
        return document.querySelector('th.o_list_record_selector input[type="checkbox"]');
    """


class BotonAccion:
    SELECTORES = [
        "//button[contains(., 'Acción')]",
        "//button[contains(text(), 'Acción')]",
        "//a[contains(text(), 'Acción')]",
        "//button[contains(., 'Accion')]",
        "//button[contains(text(), 'Accion')]",
        "//button[contains(@class, 'dropdown') and contains(., 'Acc')]",
        "//*[contains(text(), 'Action')]",
    ]


class Exportar:
    MENU_OPCIONES = [
        "//a[contains(text(), 'Exportar')]",
        "//span[contains(text(), 'Exportar')]",
        "//*[contains(text(), 'Exportar') and not(contains(text(), 'fichero'))]",
    ]
    A_FICHERO = [
        "//button[contains(text(), 'Exportar a fichero')]",
        "//button[contains(., 'Exportar') and contains(., 'fichero')]",
        "//button[contains(@class, 'btn-primary') and contains(., 'Exportar')]",
    ]


class Modal:
    CONTENEDOR = (
        "//div[contains(@class,'modal') and contains(@class,'show')] | "
        "//div[@role='dialog']"
    )
    DIALOG = "modal-dialog"
    SELECT_EXPORT = [
        "//select[contains(@name, 'export')]",
        "//div[contains(@class, 'modal')]//select",
        "//select",
    ]
    SELECT_JS = """
        var selects = document.querySelectorAll('select');
        for (var i = 0; i < selects.length; i++) {
            if (selects[i].offsetParent !== null) return selects[i];
        }
        return null;
    """
    CERRAR = [
        "//button[contains(@class, 'close')]",
        "//button[@class='btn-close']",
        "//span[contains(@class, 'close')]",
        "//button[contains(@aria-label, 'Close')]",
    ]


class Valorizado:
    OPCION = [
        "//a[normalize-space(text())='Valorizado']",
        "//span[normalize-space(text())='Valorizado']",
        "//*[normalize-space(text())='Valorizado']",
    ]
    TIPO_CONSULTA_DROPDOWN = [
        "//div[contains(@class,'modal')]//select",
        "//div[contains(@class, 'modal')]//select",
        "//select",
    ]
    UBICACION_INPUT = (
        "//div[contains(@class, 'modal')]//input[not(@type='hidden') and "
        "not(@type='checkbox') and not(@type='radio')]"
    )
    AUTOCOMPLETE_OPCIONES = [
        "//li[contains(@id, 'ui-id')]",
        "//li[contains(@id, 'ui-id')]/a",
        "//ul[contains(@class, 'ui-autocomplete')]//li",
        "//ul[@role='listbox']//li",
        "//div[contains(@class, 'ui-menu')]//li",
        "//ul[contains(@class, 'dropdown-menu')]//li",
        "//ul[contains(@class, 'o_m2o')]//li",
    ]
    # Item "Buscar mas..." del dropdown de autocompletado de ubicacion
    BUSCAR_MAS = [
        "//li[contains(@id, 'ui-id')]/a[normalize-space(.)='Buscar más...']",
        "//li[contains(@id, 'ui-id')]/a[contains(., 'Buscar') and contains(., 'más')]",
        "//div[contains(@class, 'ui-menu')]//a[contains(., 'Buscar')]",
        "//ul[contains(@class, 'ui-autocomplete')]//a[contains(., 'Buscar')]",
    ]
    UBICACION_INPUT_NOMBRE = (
        "//div[@name='location_id']//input[contains(@id, 'o_field_input')]"
    )
    # Filas del modal de "Buscar mas..." que contienen el valor de la ubicacion
    MODAL_FILA_UBICACION = [
        "//td[contains(@class, 'o_data_cell') and normalize-space(.)='{VB}']",
        "//td[normalize-space(.)='{VB}']",
        "//td[contains(normalize-space(.), '{VB}')]",
        "//tr[.//td[normalize-space(.)='{VB}']]",
    ]
    MODAL_ABRIR = [
        "//div[contains(@class, 'modal') and contains(@class, 'show')]//table[contains(@class, 'o_list_view')]",
        "//div[contains(@class, 'modal')]//table[contains(@class, 'o_list_view')]",
        "//div[contains(@role, 'dialog')]//table",
    ]
    GENERAR_XLSX = [
        "//button[contains(text(), 'Generar XLSX')]",
        "//button[contains(., 'Generar XLSX')]",
        "//button[contains(@class, 'btn-primary') and contains(., 'Generar')]",
    ]


class AlertaERP:
    SIN_REGISTROS = "//*[contains(text(),'No hay registros que coincidan')]"
    BOTON_ACEPTAR = (
        "//button[normalize-space(text())='Aceptar'] | "
        "//button[normalize-space(text())='OK'] | "
        "//button[contains(@class,'btn-primary') and "
        "(contains(.,'Aceptar') or contains(.,'OK'))]"
    )


class InformeVentas:
    CAMPO_INFORME = [
        "//div[@name='report_id']//input[contains(@id, 'o_field_input')]",
        "//input[@placeholder='Seleccione un informe']",
        "//div[contains(@class, 'o_field_many2one')]//input",
        "//div[contains(@class, 'o_input_dropdown')]//input",
        "//input[contains(@class, 'o_input')]",
        "//div[contains(@class, 'modal')]//input[@type='text']",
        "//div[contains(@class, 'modal-body')]//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]",
        "//div[contains(@class, 'modal')]//input[not(@type='hidden')]",
        "//div[contains(@role, 'dialog')]//input[not(@type='hidden')]",
    ]
    AUTOCOMPLETE_OPCIONES = [
        "//ul[contains(@class, 'ui-autocomplete')]//li[contains(@class, 'ui-menu-item')]",
        "//ul[contains(@class, 'ui-menu')]//li",
        "//div[contains(@class, 'ui-menu')]//li",
        "//li[contains(@class, 'ui-menu-item')]",
        "//ul[contains(@class, 'dropdown-menu')]//li",
    ]
    AUTOCOMPLETE_WAIT = (
        "//ul[contains(@class, 'ui-autocomplete')]//li | "
        "//ul[contains(@class, 'ui-menu')]//li"
    )
    FECHA_DESDE = [
        "//label[contains(text(), 'Fecha desde')]/following::input[1]",
        "//input[@name='date_from']",
        "//div[contains(@class, 'o_field_widget')]//input[contains(@data-field, 'date_from')]",
    ]
    FECHA_HASTA = [
        "//label[contains(text(), 'Fecha hasta')]/following::input[1]",
        "//input[@name='date_to']",
        "//div[contains(@class, 'o_field_widget')]//input[contains(@data-field, 'date_to')]",
    ]
    FECHA_INPUT_MODAL = (
        "//div[contains(@class, 'modal')]//input[contains(@class, 'o_datepicker_input') "
        "or contains(@class, 'datetimepicker-input') or @data-toggle='datetimepicker']"
    )
    FECHA_INPUT_GENERICO = "//div[contains(@class, 'modal')]//input[@type='text']"
    CHECKBOX_MOSTRAR_COSTO = [
        "//label[contains(., 'Mostrar costo')]",
        "//*[contains(text(), 'Mostrar costo')]",
    ]
    CHECKBOX_MOSTRAR_COSTO_INPUTS = [
        "//input[@type='checkbox' and ancestor::*[contains(text(), 'Mostrar costo')]]",
        "//div[contains(text(), 'Mostrar costo')]/preceding::input[@type='checkbox'][1]",
        "//div[contains(text(), 'Mostrar costo')]/following::input[@type='checkbox'][1]",
    ]
    SELECT_TIPO_INFORME = [
        "//select[@name='report_type']",
        "//select[contains(@id, 'report_type')]",
        "//label[contains(text(), 'Tipo de informe')]/following::select[1]",
        "//div[contains(@class, 'o_field_widget')]//select",
    ]
    BOTON_GENERAR_XLSX = (
        "//button[contains(text(), 'Generar XLSX')] | "
        "//button[contains(@class, 'btn-primary') and contains(., 'XLSX')]"
    )


class ChromeVersion:
    REG_HKLM = r"HKLM\SOFTWARE\Google\Chrome\BLBeacon"
    REG_HKCU = r"HKCU\SOFTWARE\Google\Chrome\BLBeacon"
