"""Escritor de inventario sobre la plantilla real, via Excel COM (solo Windows).

Es el unico lugar que toca la plantilla de produccion. Recibe el DataFrame ya
calculado y validado por el dominio y lo proyecta sobre la hoja INVENTARIO
conservando lo que solo Excel sabe hacer bien: formatos, formulas Dif, tablas
dinamicas de RESUMEN LINEA, vinculos externos y el respaldo INVENTARIO COPIA.

Pasos (``escribir``):

 1. Descifra la plantilla a un temporal y la abre con Excel.
 2. Mapea los encabezados reales de la fila 2 a las columnas canonicas.
 3. Retitula la columna de existencia con la fecha (``EXISTENCIA OCT 02``).
 4. Borra el contenido de las filas de datos anteriores (conserva formatos).
 5. Escribe todas las filas por bloques: valores + formulas Dif.
 6. Copia el formato de la primera fila de datos a las demas y deja la
    columna REFERENCIA como texto.
 7. Fila de subtotales al final y SUBTOTAL de la fila 1 al nuevo rango.
 8. Apunta las tablas dinamicas al nuevo rango.
 9. Clona INVENTARIO en INVENTARIO COPIA (valores + formatos).
10. Guarda como ``$2026 INVENTARIO GENERAL ACTUALIZADO <fecha>.xlsx`` y cifra.

Las funciones sin COM (``mapear_columnas``, ``construir_matriz``...) son puras
y tienen pruebas en ``tests/unit/adapters``.
"""

from __future__ import annotations

import logging
import re
import shutil
import tempfile
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import date, datetime
from functools import partial
from pathlib import Path
from typing import Any

import pandas as pd

from insumos.adapters.excel.escritor_simple import titulo_existencia
from insumos.adapters.excel.lector import ExcelReader
from insumos.domain.rules.inventario import columnas as C

log = logging.getLogger("actualizacion_inventario")

_COM_REINTENTABLES = (-2147418111, -2147417848)  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
_XL_PASTE_VALUES = -4163
_XL_PASTE_FORMATS = -4122
_XL_PASTE_COLUMN_WIDTHS = 8
_XL_OPENXML = 51
_XL_DATABASE = 1
_FILAS_POR_BLOQUE = 4000

# Formulas Dif: columna Dif -> (columna sistema, columna copia). "=+N3=E3".
_FORMULAS_DIF: dict[str, tuple[str, str]] = {
    C.DIF_MARCA: (C.MARCA_SISTEMA, C.MARCA_COPIA),
    C.DIF_LINEA: (C.LINEA_SISTEMA, C.LINEA_COPIA),
    C.DIF_SUBLINEA: (C.SUBLINEA_SISTEMA, C.SUBLINEA_COPIA),
}
_REQUERIDAS = (C.REFERENCIA, C.EXISTENCIA, C.COSTO_PROMEDIO, C.TOTAL_INV)


# ----------------------------------------------------------------------
# Funciones puras (probadas sin Excel)
# ----------------------------------------------------------------------

def letra_columna(indice: int) -> str:
    """1 -> A, 27 -> AA."""
    letras = ""
    while indice > 0:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def mapear_columnas(encabezados: Sequence[Any]) -> dict[str, int]:
    """Columna canonica -> indice 1-based, a partir de la fila de encabezados."""
    mapa: dict[str, int] = {}
    for i, valor in enumerate(encabezados, start=1):
        if valor is None or not str(valor).strip():
            continue
        n = C.normalizar_encabezado(valor)
        if C.es_columna_existencia(n):
            mapa.setdefault(C.EXISTENCIA, i)
        elif n in C.COLUMNAS_PLANTILLA:
            mapa.setdefault(n, i)
    faltan = [c for c in _REQUERIDAS if c not in mapa]
    if faltan:
        raise ValueError(f"La hoja no tiene las columnas {faltan}. Encabezados: {list(encabezados)}")
    return mapa


def valor_celda(v: Any) -> Any:
    """Convierte un valor de pandas/numpy a algo que COM acepte."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        return v
    if hasattr(v, "item"):  # numpy escalar
        v = v.item()
    if isinstance(v, (bool, int, float)):
        return v
    return str(v)


def construir_matriz(
    datos: pd.DataFrame,
    columnas: dict[str, int],
    ancho: int,
    primera_fila: int,
    *,
    formulas: bool = True,
    formulas_dif: bool = False,
) -> list[tuple[Any, ...]]:
    """Filas listas para escribir en la hoja.

    - ``formulas=True`` (hoja INVENTARIO, como la salida de produccion):
      TOTAL INV = ``=G3*H3``. Con ``False`` (hoja COPIA) va el valor.
    - Dif marca/linea/sub-linea van como valor booleano (sistema = copia), que
      es como los deja la salida de produccion. ``formulas_dif=True`` escribe
      ``=+N3=E3`` en su lugar.
    - REFERENCIA siempre como texto (``1060`` -> ``"1060"``).
    """
    vacia = [None] * ancho
    por_columna: dict[int, list[Any]] = {}
    n = len(datos)
    filas_excel = range(primera_fila, primera_fila + n)
    for nombre, indice in columnas.items():
        if nombre in _FORMULAS_DIF:
            sistema, copia = _FORMULAS_DIF[nombre]
            if sistema not in columnas or copia not in columnas:
                continue
            if formulas and formulas_dif:
                ls, lc = letra_columna(columnas[sistema]), letra_columna(columnas[copia])
                por_columna[indice] = [f"=+{ls}{r}={lc}{r}" for r in filas_excel]
            elif sistema in datos.columns and copia in datos.columns:
                por_columna[indice] = [
                    _iguales(a, b) for a, b in zip(datos[sistema], datos[copia], strict=True)
                ]
            continue
        if nombre == C.TOTAL_INV and formulas and C.EXISTENCIA in columnas and C.COSTO_PROMEDIO in columnas:
            le, lc = letra_columna(columnas[C.EXISTENCIA]), letra_columna(columnas[C.COSTO_PROMEDIO])
            por_columna[indice] = [f"={le}{r}*{lc}{r}" for r in filas_excel]
            continue
        if nombre not in datos.columns:
            continue
        if nombre == C.REFERENCIA:
            por_columna[indice] = [C.clave_referencia(v) and _ref_texto(v) for v in datos[nombre]]
        else:
            por_columna[indice] = [valor_celda(v) for v in datos[nombre]]

    filas: list[tuple[Any, ...]] = []
    for i in range(n):
        fila = list(vacia)
        for indice, valores in por_columna.items():
            fila[indice - 1] = valores[i]
        filas.append(tuple(fila))
    return filas


def _iguales(a: Any, b: Any) -> bool:
    """Lo que daria ``=N3=E3`` en Excel: comparacion de texto sin distinguir mayusculas."""
    va, vb = valor_celda(a), valor_celda(b)
    if isinstance(va, str) and isinstance(vb, str):
        return va.strip().upper() == vb.strip().upper()
    return bool(va == vb)


def _ref_texto(v: Any) -> str:
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


_RANGO_SUBTOTAL = re.compile(r"SUBTOTAL\((\d+),\s*\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)", re.IGNORECASE)


def ajustar_formula_subtotal(formula: str, primera: int, ultima: int) -> str:
    """``=SUBTOTAL(109,G3:G15557)`` -> ``=SUBTOTAL(109,G3:G<ultima>)``."""

    def _sub(m: re.Match[str]) -> str:
        return f"SUBTOTAL({m.group(1)},{m.group(2)}{primera}:{m.group(4)}{ultima})"

    return _RANGO_SUBTOTAL.sub(_sub, formula)


def nombre_salida(prefijo: str, ahora: datetime) -> str:
    return f"{prefijo} {ahora:%Y%m%d_%H%M}.xlsx"


# ----------------------------------------------------------------------
# COM
# ----------------------------------------------------------------------

def _com[T](funcion: Callable[[], T], intentos: int = 5, espera: float = 2.0) -> T:
    """Reintenta las llamadas que Excel rechaza por estar ocupado."""
    for intento in range(1, intentos + 1):
        try:
            return funcion()
        except Exception as e:  # noqa: BLE001 - errores COM no tienen tipo propio util
            codigo = e.args[0] if e.args and isinstance(e.args[0], int) else None
            if codigo in _COM_REINTENTABLES and intento < intentos:
                time.sleep(espera * intento)
                continue
            raise
    raise RuntimeError("inalcanzable")


@dataclass
class EscritorInventarioCom:
    """Implementa ``EscritorInventario`` sobre la plantilla real con Excel."""

    plantilla: Path
    carpeta_salida: Path
    password: str | None
    cifrar_salida: bool = True
    prefijo_salida: str = "$2026 INVENTARIO GENERAL ACTUALIZADO"
    hoja: str = "INVENTARIO"
    hoja_copia: str = "INVENTARIO COPIA"
    fila_encabezado: int = 2

    def escribir(self, datos: pd.DataFrame, *, hoy: date) -> Path:
        import pythoncom  # type: ignore[import-untyped,import-not-found,unused-ignore]
        import win32com.client as win32  # type: ignore[import-untyped,import-not-found,unused-ignore]

        self.carpeta_salida.mkdir(parents=True, exist_ok=True)
        destino = self.carpeta_salida / nombre_salida(self.prefijo_salida, datetime.now())
        temporal = Path(tempfile.mkdtemp(prefix="insumos_inv_")) / "plantilla.xlsx"
        temporal.write_bytes(ExcelReader(ruta=self.plantilla, password=self.password)._bytes_planos())

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        wb = None
        try:
            for atributo, valor in (
                ("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False),
                ("EnableEvents", False), ("AskToUpdateLinks", False),
            ):
                try:
                    setattr(excel, atributo, valor)
                except Exception:  # noqa: BLE001
                    pass
            wb = _com(lambda: excel.Workbooks.Open(str(temporal), UpdateLinks=0, ReadOnly=False))
            try:
                excel.Calculation = -4135  # manual durante la escritura
            except Exception:  # noqa: BLE001
                pass
            ws = _com(lambda: wb.Worksheets(self.hoja))
            ultima = self._proyectar(ws, datos, hoy)
            # Recalcular ANTES de clonar: la COPIA guarda valores, no formulas.
            try:
                excel.Calculation = -4105  # automatico
                _com(lambda: excel.CalculateFull())
            except Exception as e:  # noqa: BLE001
                log.warning("No se pudo recalcular el libro: %s", e)
            self._actualizar_tablas_dinamicas(wb, ultima)
            self._clonar_copia(wb, ws)
            _com(lambda: wb.SaveAs(str(destino), FileFormat=_XL_OPENXML))
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass
            pythoncom.CoUninitialize()
            shutil.rmtree(temporal.parent, ignore_errors=True)

        if self.cifrar_salida and self.password:
            _cifrar(destino, self.password)
        log.info("Inventario escrito: %s (%d referencias)", destino, len(datos))
        return destino

    # ------------------------------------------------------------------

    def _proyectar(self, ws: Any, datos: pd.DataFrame, hoy: date) -> int:
        h = self.fila_encabezado
        usado = _com(lambda: ws.UsedRange)
        ancho = int(_com(lambda: usado.Column)) + int(_com(lambda: usado.Columns.Count)) - 1
        ultima_anterior = int(_com(lambda: usado.Row)) + int(_com(lambda: usado.Rows.Count)) - 1

        fila_enc = _com(lambda: ws.Range(ws.Cells(h, 1), ws.Cells(h, ancho)).Value)
        encabezados = list(fila_enc[0]) if isinstance(fila_enc, tuple) else [fila_enc]
        columnas = mapear_columnas(encabezados)

        # 3. Encabezado de existencia con la fecha
        _com(lambda: setattr(ws.Cells(h, columnas[C.EXISTENCIA]), "Value", titulo_existencia(hoy)))

        primera = h + 1
        n = len(datos)
        ultima = h + n

        # 4. Borrar datos anteriores (incluye la fila de subtotales de ayer)
        if ultima_anterior >= primera:
            _com(lambda: ws.Range(ws.Cells(primera, 1), ws.Cells(ultima_anterior, ancho)).ClearContents())

        # REFERENCIA como texto antes de escribir, para que "1060" no se vuelva numero
        col_ref = columnas[C.REFERENCIA]
        _com(lambda: setattr(ws.Range(ws.Cells(primera, col_ref), ws.Cells(max(ultima, primera), col_ref)),
                             "NumberFormat", "@"))

        # 5. Escribir por bloques
        matriz = construir_matriz(datos, columnas, ancho, primera)
        for inicio in range(0, n, _FILAS_POR_BLOQUE):
            bloque = matriz[inicio:inicio + _FILAS_POR_BLOQUE]
            f1 = primera + inicio
            f2 = f1 + len(bloque) - 1
            rango = ws.Range(ws.Cells(f1, 1), ws.Cells(f2, ancho))
            _com(partial(setattr, rango, "Value", bloque))

        # 6. Formato de la primera fila de datos a todas
        if n > 1:
            try:
                _com(lambda: ws.Range(ws.Cells(primera, 1), ws.Cells(primera, ancho)).Copy())
                _com(lambda: ws.Range(ws.Cells(primera + 1, 1), ws.Cells(ultima, ancho)).PasteSpecial(
                    Paste=_XL_PASTE_FORMATS))
            except Exception as e:  # noqa: BLE001
                log.warning("No se pudo copiar el formato de la fila %d: %s", primera, e)
            finally:
                try:
                    ws.Application.CutCopyMode = False
                except Exception:  # noqa: BLE001
                    pass
        _com(lambda: setattr(ws.Range(ws.Cells(primera, col_ref), ws.Cells(ultima, col_ref)), "NumberFormat", "@"))

        # Formatos sobrantes de filas que ya no existen
        fila_subtotal = ultima + 1
        if ultima_anterior > fila_subtotal:
            _com(lambda: ws.Range(ws.Cells(fila_subtotal + 1, 1), ws.Cells(ultima_anterior, ancho)).ClearFormats())

        # 7. Subtotales al final y en la fila 1
        for nombre in (C.EXISTENCIA, C.TOTAL_INV):
            letra = letra_columna(columnas[nombre])
            celda = ws.Cells(fila_subtotal, columnas[nombre])
            _com(partial(setattr, celda, "Formula", f"=SUBTOTAL(109,{letra}{primera}:{letra}{ultima})"))
        for fila in range(1, h):
            for col in range(1, ancho + 1):
                celda = ws.Cells(fila, col)
                formula = _com(partial(getattr, celda, "Formula"))
                if isinstance(formula, str) and "SUBTOTAL(" in formula.upper():
                    nueva = ajustar_formula_subtotal(formula, primera, ultima)
                    if nueva != formula:
                        _com(partial(setattr, celda, "Formula", nueva))

        self._rango = (h, ultima, ancho)
        return ultima

    def _actualizar_tablas_dinamicas(self, wb: Any, ultima: int) -> None:
        """Apunta las tablas dinamicas que leen INVENTARIO al rango nuevo y las refresca."""
        h, _, ancho = self._rango
        fuente = f"'{self.hoja}'!R{h}C1:R{ultima}C{ancho}"
        for i in range(1, int(wb.Worksheets.Count) + 1):
            hoja = wb.Worksheets(i)
            try:
                tablas = hoja.PivotTables()
                total = int(tablas.Count)
            except Exception:  # noqa: BLE001
                continue
            for j in range(1, total + 1):
                pt = tablas.Item(j)
                try:
                    origen = str(pt.SourceData)
                except Exception:  # noqa: BLE001
                    continue
                if self.hoja.upper() not in origen.upper() or self.hoja_copia.upper() in origen.upper():
                    continue
                try:
                    cache = wb.PivotCaches().Create(SourceType=_XL_DATABASE, SourceData=fuente)
                    pt.ChangePivotCache(cache)
                    pt.RefreshTable()
                    log.info("Tabla dinamica '%s' (%s) -> %s", pt.Name, hoja.Name, fuente)
                except Exception as e:  # noqa: BLE001
                    log.warning("No se pudo actualizar la tabla dinamica '%s' en %s: %s", pt.Name, hoja.Name, e)

    def _clonar_copia(self, wb: Any, ws: Any) -> None:
        """INVENTARIO COPIA = respaldo identico (valores, formatos y anchos)."""
        try:
            copia = wb.Worksheets(self.hoja_copia)
        except Exception:  # noqa: BLE001
            copia = wb.Worksheets.Add(Before=ws)
            copia.Name = self.hoja_copia
        h, ultima, ancho = self._rango
        origen = ws.Range(ws.Cells(1, 1), ws.Cells(ultima + 1, ancho))
        try:
            copia.Cells.ClearContents()
            copia.Cells.ClearFormats()
            origen.Copy()
            destino = copia.Cells(1, 1)
            for pegado in (_XL_PASTE_VALUES, _XL_PASTE_FORMATS, _XL_PASTE_COLUMN_WIDTHS):
                destino.PasteSpecial(Paste=pegado)
        except Exception as e:  # noqa: BLE001
            log.warning("Copia por portapapeles fallo (%s); se copian solo valores", e)
            copia.Range(copia.Cells(1, 1), copia.Cells(ultima + 1, ancho)).Value = origen.Value
        finally:
            try:
                ws.Application.CutCopyMode = False
            except Exception:  # noqa: BLE001
                pass
        ws.Activate()


def _cifrar(ruta: Path, password: str) -> None:
    from msoffcrypto.format.ooxml import OOXMLFile

    plano = ruta.with_name(f"~plano_{ruta.name}")
    shutil.move(str(ruta), plano)
    try:
        with plano.open("rb") as entrada, ruta.open("wb") as salida:
            OOXMLFile(entrada).encrypt(password, salida)
    except Exception:
        shutil.move(str(plano), ruta)
        raise
    finally:
        plano.unlink(missing_ok=True)
