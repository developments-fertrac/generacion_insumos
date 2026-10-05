"""Lectura de las dos entradas del inventario general.

- ``leer_inventario_bd``: exportacion de base de datos (``Inventario.xlsx``,
  sin cifrar, encabezado en la fila 1).
- ``leer_plantilla``: hoja INVENTARIO de la plantilla (cifrada, encabezado en
  la fila 2). Devuelve las columnas con nombre canonico; la de existencia
  ("EXISTENCIA SEP 10") se entrega como ``EXISTENCIA``.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

from insumos.adapters.excel.lector import ExcelReader
from insumos.domain.ports import PlantillaInventario
from insumos.domain.rules.inventario import columnas as C

HOJA_INVENTARIO = "INVENTARIO"
FILA_ENCABEZADO = 2  # visible, 1-based


def leer_inventario_bd(ruta: Path) -> pd.DataFrame:
    """La exportacion tal cual, con encabezados normalizados."""
    df = pd.read_excel(ruta, sheet_name=0, dtype={C.BD_REFERENCIA: "object"})
    df.columns = [C.normalizar_encabezado(c) for c in df.columns]
    df["_FILA"] = df.index + 2  # fila visible en Excel (encabezado en la 1)
    return df


@dataclass(frozen=True)
class PlantillaLeida:
    datos: pd.DataFrame
    encabezado_existencia: str
    columnas_desconocidas: tuple[str, ...]


def leer_plantilla(ruta: Path, password: str | None) -> PlantillaLeida:
    contenido = BytesIO(ExcelReader(ruta=ruta, password=password)._bytes_planos())
    with pd.ExcelFile(contenido) as xls:
        if HOJA_INVENTARIO not in xls.sheet_names:
            raise ValueError(f"{ruta.name} no tiene hoja {HOJA_INVENTARIO}. Tiene: {xls.sheet_names}")
        df = pd.read_excel(xls, sheet_name=HOJA_INVENTARIO, header=FILA_ENCABEZADO - 1, dtype=object)

    renombres: dict[str, str] = {}
    desconocidas: list[str] = []
    existencia_original = ""
    for col in df.columns:
        if str(col).startswith("Unnamed"):
            continue
        n = C.normalizar_encabezado(col)
        if C.es_columna_existencia(n):
            renombres[col] = C.EXISTENCIA
            existencia_original = str(col).strip()
        elif n in C.COLUMNAS_PLANTILLA:
            renombres[col] = n
        else:
            desconocidas.append(str(col))

    faltan = [c for c in (C.REFERENCIA, C.EXISTENCIA, C.COSTO_PROMEDIO, C.TOTAL_INV) if c not in renombres.values()]
    if faltan:
        raise ValueError(
            f"{ruta.name}/{HOJA_INVENTARIO}: faltan columnas {faltan} en la fila {FILA_ENCABEZADO}. "
            f"Encabezados: {list(df.columns)}"
        )

    datos = df[list(renombres)].rename(columns=renombres)
    datos["_FILA"] = datos.index + FILA_ENCABEZADO + 1  # fila visible en Excel
    # La fila de subtotales del final no tiene referencia.
    ref = datos[C.REFERENCIA]
    datos = datos[ref.notna() & ref.astype(str).str.strip().ne("")].reset_index(drop=True)
    return PlantillaLeida(
        datos=datos,
        encabezado_existencia=existencia_original,
        columnas_desconocidas=tuple(desconocidas),
    )


# ----------------------------------------------------------------------
# Adaptador del puerto FuenteInventario sobre carpetas de archivos
# ----------------------------------------------------------------------

PREFIJO_SALIDA = "2026 INVENTARIO GENERAL ACTUALIZADO"
PREFIJO_MAESTRO = "2026 INVENTARIO GENERAL"
PREFIJO_MATRIZ = "2026 MATRIZ USD"
PREFIJO_DISTRIBUCION = "DISTRIBUCION DE MATRICES"


@dataclass
class FuenteInventarioArchivos:
    """Implementa ``FuenteInventario`` leyendo los archivos de las carpetas de Fertrac.

    - Base de datos: ``archivo_bd`` (``Inventario.xlsx``).
    - Plantilla: la salida mas reciente en ``carpeta_salida``; si no hay, el
      maestro ``$2026 INVENTARIO GENERAL`` de ``carpeta_insumos``.
    - Maestros: Matriz USD y Distribucion en ``carpeta_insumos``. Si faltan, se
      devuelve vacio y la regla correspondiente lo reporta como advertencia.
    """

    archivo_bd: Path
    carpeta_salida: Path
    carpeta_insumos: Path
    password: str | None
    hoja_matriz: str = "2026"
    advertencias: list[str] | None = None

    def __post_init__(self) -> None:
        if self.advertencias is None:
            self.advertencias = []
        self._fuentes: list[dict[str, str]] = []

    def fuentes_usadas(self) -> list[dict[str, str]]:
        """Archivos leidos, para la hoja FUENTES DE DATOS del reporte (como el legacy)."""
        return list(self._fuentes)

    def _registrar(self, dato: str, ruta: Path | None) -> None:
        from insumos.adapters.system.reloj import desde_timestamp, hoy

        if ruta is None:
            self._fuentes.append({"DATO": dato, "ARCHIVO": "(no encontrado)", "FECHA MODIFICACION": "",
                                  "ADVERTENCIA": "ARCHIVO NO ENCONTRADO"})
            return
        modificado = desde_timestamp(ruta.stat().st_mtime)
        dias = (hoy() - modificado.date()).days
        advertencia = (
            f"ARCHIVO NO ES DE HOY (tiene {dias} dia(s) de antiguedad)" if dias > 0 else ""
        )
        self._fuentes.append({
            "DATO": dato,
            "ARCHIVO": ruta.name,
            "FECHA MODIFICACION": modificado.strftime("%Y-%m-%d %H:%M:%S"),
            "ADVERTENCIA": advertencia,
        })

    def inventario_bd(self) -> pd.DataFrame:
        if not self.archivo_bd.is_file():
            raise FileNotFoundError(f"No existe la exportacion de base de datos: {self.archivo_bd}")
        self._registrar("Inventario (base de datos)", self.archivo_bd)
        return leer_inventario_bd(self.archivo_bd)

    def ruta_plantilla(self) -> Path:
        from insumos.adapters.system.archivos import mas_reciente

        previa = mas_reciente(self.carpeta_salida, PREFIJO_SALIDA)
        if previa is not None:
            return previa
        maestro = mas_reciente(self.carpeta_insumos, PREFIJO_MAESTRO)
        if maestro is None:
            raise FileNotFoundError(
                f"No hay plantilla: ni salida previa '{PREFIJO_SALIDA}*' en {self.carpeta_salida} "
                f"ni maestro '{PREFIJO_MAESTRO}*' en {self.carpeta_insumos}"
            )
        return maestro

    def plantilla(self) -> PlantillaInventario:
        ruta = self.ruta_plantilla()
        self._registrar("Plantilla de inventario", ruta)
        leida = leer_plantilla(ruta, self.password)
        return PlantillaInventario(
            datos=leida.datos, origen=str(ruta), encabezado_existencia=leida.encabezado_existencia
        )

    def matriz_usd(self) -> pd.DataFrame:
        from insumos.adapters.excel.maestros import leer_matriz_usd
        from insumos.adapters.system.archivos import mas_reciente

        ruta = mas_reciente(self.carpeta_insumos, PREFIJO_MATRIZ)
        self._registrar("Matriz USD", ruta)
        if ruta is None:
            self._advertir(f"Matriz USD no encontrada en {self.carpeta_insumos}")
            return pd.DataFrame(columns=["CLAVE", "DESCRIPCION"])
        return leer_matriz_usd(ruta, self.password, self.hoja_matriz)

    def distribucion(self) -> dict[str, Any]:
        from insumos.adapters.excel.maestros import leer_distribucion
        from insumos.adapters.system.archivos import mas_reciente

        ruta = mas_reciente(self.carpeta_insumos, PREFIJO_DISTRIBUCION)
        self._registrar("Distribucion de matrices", ruta)
        if ruta is None:
            self._advertir(f"Distribucion de matrices no encontrada en {self.carpeta_insumos}")
            return {"gestor": {}, "clasificacion": {}}
        return leer_distribucion(ruta, self.password)

    def _advertir(self, mensaje: str) -> None:
        assert self.advertencias is not None
        self.advertencias.append(mensaje)
