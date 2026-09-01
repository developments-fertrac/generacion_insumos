"""Utilidades compartidas de procesamiento Excel con COM (Windows).

Funciones reutilizables para abrir, desencriptar, convertir y gestionar
archivos Excel via COM (win32com) y msoffcrypto.
"""

from __future__ import annotations

import gc
import io
import os
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from pathlib import Path

import msoffcrypto

from core.logger import get_logger

log = get_logger("excel_processing")

try:
    import win32com.client as win32

    HAS_COM = True
except Exception:
    HAS_COM = False


def safe_close_workbook(wb, max_attempts: int = 3, delay: float = 0.5):
    if wb is None:
        return
    for attempt in range(max_attempts):
        try:
            wb.Close(SaveChanges=False)
            time.sleep(0.1)
            return
        except Exception as e:
            if attempt < max_attempts - 1:
                time.sleep(delay)
            else:
                log.warning("No se pudo cerrar workbook: %s", e)
    try:
        del wb
    except Exception:
        pass


def safe_quit_excel(excel, max_attempts: int = 3, delay: float = 0.5):
    if excel is None:
        return
    for attempt in range(max_attempts):
        try:
            excel.Quit()
            time.sleep(0.2)
            return
        except Exception as e:
            if attempt < max_attempts - 1:
                time.sleep(delay)
            else:
                log.warning("No se pudo cerrar Excel: %s", e)
    try:
        del excel
    except Exception:
        pass
    gc.collect()


def verificar_archivo_disponible(archivo: Path) -> bool:
    try:
        with open(archivo, "rb") as f:
            f.read(1)
        return True
    except (PermissionError, IOError, OSError):
        return False
    except Exception as e:
        log.warning("Error al verificar archivo: %s", e)
        return False


def obtener_archivo_trabajo(
    archivo: Path, crear_copia_si_bloqueado: bool = True
) -> tuple[Path, bool]:
    if verificar_archivo_disponible(archivo):
        log.info("Archivo disponible para lectura directa")
        return archivo, False

    log.info("Archivo en uso: %s", archivo.name)
    if not crear_copia_si_bloqueado:
        raise PermissionError(f"ERROR: El archivo '{archivo.name}' esta abierto.")

    log.info("Creando copia temporal...")
    try:
        temp_dir = Path(tempfile.gettempdir()) / "fertrac_inventario_temp"
        temp_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_copia = f"TEMP_{timestamp}_{archivo.name}"
        ruta_copia = temp_dir / nombre_copia
        shutil.copy2(archivo, ruta_copia)
        log.info("Copia temporal creada: %s", nombre_copia)
        return ruta_copia, True
    except Exception as e:
        log.error("Error al crear copia temporal: %s", e)
        raise PermissionError(
            f"ERROR: No se pudo crear copia de '{archivo.name}'. Detalle: {e}"
        )


def limpiar_archivo_temporal(archivo: Path, es_temporal: bool):
    if not es_temporal or archivo is None:
        return
    try:
        if archivo.exists():
            archivo.unlink()
            log.info("Copia temporal eliminada: %s", archivo.name)
    except Exception as e:
        log.warning("No se pudo eliminar copia temporal: %s", e)


def limpiar_copias_temporales_antiguas(max_horas: int = 24):
    temp_dir = Path(tempfile.gettempdir()) / "fertrac_inventario_temp"
    if not temp_dir.exists():
        return
    try:
        ahora = time.time()
        eliminados = 0
        for archivo in temp_dir.glob("TEMP_*"):
            if not archivo.is_file():
                continue
            edad_horas = (ahora - archivo.stat().st_mtime) / 3600
            if edad_horas > max_horas:
                try:
                    archivo.unlink()
                    eliminados += 1
                except Exception:
                    pass
        if eliminados > 0:
            log.info("Limpieza: %d archivos temporales antiguos eliminados", eliminados)
    except Exception as e:
        log.warning("Error en limpieza automatica: %s", e)


def decrypt_to_stream_local(xlsx_path: Path, password: str) -> io.BytesIO:
    bio = io.BytesIO()
    with open(xlsx_path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password=password)
        office.decrypt(bio)
    bio.seek(0)
    return bio


def is_encrypted_xlsx(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            of = msoffcrypto.OfficeFile(f)
            return bool(getattr(of, "is_encrypted", True))
    except Exception:
        return False


def save_bytesio_to_temp(bio: io.BytesIO, stem: str) -> Path:
    tmp = Path(tempfile.gettempdir()) / f"~dec_{stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    with open(tmp, "wb") as out:
        out.write(bio.getvalue())
    return tmp


def com_convert_to_xlsx(path: Path, passwords=None) -> Path:
    if not HAS_COM:
        raise RuntimeError("win32com no disponible en este sistema")
    if passwords is None:
        passwords = []

    encrypted = False
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            with open(path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                encrypted = bool(getattr(of, "is_encrypted", False))
        except Exception:
            encrypted = False

    wb = None
    last_err = None
    pw_attempts = passwords if encrypted else [None] + list(passwords)

    # Si el archivo origen esta bloqueado (abierto por otra instancia de Excel),
    # trabajamos sobre una copia temporal para poder convertirlo.
    ruta_origen = path
    if not verificar_archivo_disponible(Path(path)):
        log.info("Archivo origen en uso, creando copia temporal para conversion")
        try:
            temp_dir = Path(tempfile.gettempdir()) / "fertrac_conv_temp"
            temp_dir.mkdir(exist_ok=True)
            copia = temp_dir / f"conv_{Path(path).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{Path(path).suffix}"
            shutil.copy2(path, copia)
            ruta_origen = copia
        except Exception as e:
            log.warning("No se pudo crear copia del origen bloqueado: %s", e)

    for intento in range(3):
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = True
        excel.Interactive = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        wb = None
        open_ok = False
        for pw in pw_attempts:
            try:
                if pw:
                    wb = excel.Workbooks.Open(
                        str(ruta_origen), UpdateLinks=0, ReadOnly=True,
                        IgnoreReadOnlyRecommended=True, Password=pw,
                    )
                else:
                    wb = excel.Workbooks.Open(
                        str(ruta_origen), UpdateLinks=0, ReadOnly=True,
                        IgnoreReadOnlyRecommended=True,
                    )
                open_ok = True
                break
            except Exception as e:
                last_err = e
                continue

        if not open_ok:
            safe_quit_excel(excel)
            if intento < 2:
                log.warning("Intento %d: no pude abrir '%s' (posible bloqueo), reintentando...", intento + 1, Path(path).name)
                time.sleep(2)
                continue
            msg = "archivo cifrado sin contrasena valida" if encrypted else "no pude abrir el archivo"
            raise RuntimeError(f"COM no pudo abrir '{Path(path).name}': {msg}. Detalle: {last_err}")

        tmp = Path(tempfile.gettempdir()) / f"~conv_{Path(path).stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
        try:
            wb.SaveAs(str(tmp), FileFormat=51)
            wb.Close(SaveChanges=False)
            excel.Quit()
            return tmp
        except Exception as e:
            last_err = e
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            safe_quit_excel(excel)
            if intento < 2:
                log.warning("Intento %d: fallo al guardar conversion (posible bloqueo): %s", intento + 1, e)
                time.sleep(2)
                continue
            break

    # Ultimo recurso: limpiar procesos Excel huerfanos que bloquean archivos
    log.warning("Conversion fallo por bloqueo persistente; limpiando procesos EXCEL.EXE huerfanos...")
    try:
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True, timeout=30)
    except Exception as e:
        log.warning("No se pudo limpiar EXCEL.EXE: %s", e)
    time.sleep(2)
    raise RuntimeError(
        f"COM no pudo convertir '{Path(path).name}' incluso tras reintentos: {last_err}"
    )


def excel_serial_from_date(dt) -> int | None:
    if dt is None:
        return None
    from datetime import datetime as dt_type
    if hasattr(dt, "year") and hasattr(dt, "month") and hasattr(dt, "day"):
        return (dt_type(dt.year, dt.month, dt.day) - dt_type(1899, 12, 30)).days
    return None


def convert_xls_via_xlrd(path: Path, sheet: str | None = None) -> Path:
    """Convierte un .xls legacy (BIFF) a un .xlsx temporal usando SOLO xlrd (puro Python).

    Evita depender de Excel COM, que falla cuando la tarea corre como servicio de
    Windows (Session 0) con ``Microsoft Excel no puede obtener acceso al archivo
    'Temp\\\\<numeros>'``. Retorna la ruta del .xlsx temporal.
    """
    import openpyxl
    import xlrd

    if xlrd.__version__.startswith("1."):
        # xlrd 1.x usa formattuing_info / cell_value_tuples: no soportado aqui
        raise RuntimeError("Se requiere xlrd>=2.0 para leer .xls legacy")

    tmp = (
        Path(tempfile.gettempdir())
        / f"~xlrd_{Path(path).stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    )

    wb = xlrd.open_workbook(str(path), formatting_info=False)
    chosen = wb.sheet_names()[0]
    if sheet:
        for name in wb.sheet_names():
            if name.strip().lower() == sheet.strip().lower():
                chosen = name
                break
    sh = wb.sheet_by_name(chosen)

    out = openpyxl.Workbook()
    ws = out.active
    ws.title = (chosen[:31] or "Hoja1")
    rows = []
    for r in range(sh.nrows):
        rows.append([_xlrd_cell_value(sh, r, c) for c in range(sh.ncols)])
    # quitar columnas totalmente vacias al final
    if rows:
        ncols = max(len(r) for r in rows)
        while ncols > 0 and all(
            (row[ncols - 1] in (None, "") if len(row) >= ncols else True)
            for row in rows
        ):
            ncols -= 1
        rows = [r[:ncols] for r in rows]
    for row in rows:
        ws.append(row)
    out.save(tmp)
    log.info("Convertido '%s' via xlrd (sin COM) -> %s", Path(path).name, tmp.name)
    return tmp


def _xlrd_cell_value(sh, r: int, c: int):
    try:
        cell = sh.cell(r, c)
        v = cell.value
        if cell.ctype == 3:  # fecha
            try:
                import xlrd

                dt = xlrd.xldate_as_datetime(v, sh.book.datemode)
                return dt
            except Exception:
                return v
        return v
    except Exception:
        return None


def convert_xls_legacy(path: Path, sheet: str | None = None) -> Path:
    """Convierte un .xls a .xlsx temporal SIN depender de Excel COM.

    Primero intenta leerlo como BIFF legacy con xlrd; si el archivo resulta ser
    una tabla HTML con extension .xls (comun en exportaciones de Odoo), usa
    pd.read_html. Levanta RuntimeError si ninguna via funciona.
    """
    try:
        return convert_xls_via_xlrd(path, sheet=sheet)
    except Exception as e1:
        log.warning("Era .xls BIFF? No con xlrd (%s); probando como tabla HTML...", e1)
        try:
            import pandas as pd

            frames = pd.read_html(str(path))
            if not frames:
                raise RuntimeError("read_html no encontro tablas")
            df = frames[0]
            tmp = (
                Path(tempfile.gettempdir())
                / f"~xlrd_{Path(path).stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
            )
            df.to_excel(tmp, index=False, engine="openpyxl")
            log.info("Convertido '%s' via read_html (sin COM) -> %s", Path(path).name, tmp.name)
            return tmp
        except Exception as e2:
            raise RuntimeError(
                f"No se pudo convertir '{Path(path).name}' sin COM (xlrd/HTML fallaron): {e2}"
            ) from e2
