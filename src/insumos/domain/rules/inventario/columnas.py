"""Nombres canonicos de columnas y clave de referencia del inventario.

El dominio trabaja con nombres estables. Los adaptadores traducen:

- La plantilla de Excel trae encabezados con espacios sobrantes ("NOMBRE MYR ",
  "Marca sistema "), mayusculas mezcladas ("MARCA copia") y una columna de
  existencia cuyo nombre cambia cada dia ("EXISTENCIA SEP 10"). El lector la
  entrega como ``EXISTENCIA``; el escritor la vuelve a titular con la fecha.
- La exportacion de base de datos (``Inventario.xlsx``) trae sus propios
  nombres (``EXISTENCIA NETA``, ``COSTO``, ``SUB-LINEA``...).

Las reglas solo conocen las constantes de este modulo.
"""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

# --- Plantilla de inventario general (hoja INVENTARIO) ---
REFERENCIA = "REFERENCIA"
NOMBRE_LISTA = "NOMBRE LISTA"
NOMBRE_ODOO = "NOMBRE ODOO"
NOMBRE_MYR = "NOMBRE MYR"
MARCA_COPIA = "MARCA COPIA"
INV_BODEGA_GERENCIA = "INV BODEGA GERENCIA"
EXISTENCIA = "EXISTENCIA"
COSTO_PROMEDIO = "COSTO PROMEDIO"
TOTAL_INV = "TOTAL INV"
LINEA_COPIA = "LINEA COPIA"
SUBLINEA_COPIA = "SUB-LINEA COPIA"
LIDER_LINEA = "LIDER LINEA"
CLASIFICACION = "CLASIFICACION"
MARCA_SISTEMA = "MARCA SISTEMA"
DIF_MARCA = "DIF MARCA"
LINEA_SISTEMA = "LINEA SISTEMA"
DIF_LINEA = "DIF LINEA"
SUBLINEA_SISTEMA = "SUB-LINEA SISTEMA"
DIF_SUBLINEA = "DIF SUB-LINEA"

COLUMNAS_PLANTILLA: tuple[str, ...] = (
    REFERENCIA,
    NOMBRE_LISTA,
    NOMBRE_ODOO,
    NOMBRE_MYR,
    MARCA_COPIA,
    INV_BODEGA_GERENCIA,
    EXISTENCIA,
    COSTO_PROMEDIO,
    TOTAL_INV,
    LINEA_COPIA,
    SUBLINEA_COPIA,
    LIDER_LINEA,
    CLASIFICACION,
    MARCA_SISTEMA,
    DIF_MARCA,
    LINEA_SISTEMA,
    DIF_LINEA,
    SUBLINEA_SISTEMA,
    DIF_SUBLINEA,
)

# Columnas que en la hoja son formulas (=+N3=E3). El escritor las genera; el
# dominio no las calcula ni las compara.
COLUMNAS_FORMULA: tuple[str, ...] = (DIF_MARCA, DIF_LINEA, DIF_SUBLINEA)

# --- Exportacion de base de datos (Inventario.xlsx) ---
BD_REFERENCIA = "REFERENCIA"
BD_NOMBRE_ODOO = "NOMBRE ODOO"
BD_MARCA = "MARCA"
BD_LINEA = "LINEA"
BD_SUBLINEA = "SUB-LINEA"
BD_EXISTENCIA_NETA = "EXISTENCIA NETA"
BD_COSTO = "COSTO"
BD_TOTAL_INV = "TOTAL INV"
BD_MOTIVO = "MOTIVO"

COLUMNAS_BD_REQUERIDAS: tuple[str, ...] = (
    BD_REFERENCIA,
    BD_NOMBRE_ODOO,
    BD_MARCA,
    BD_LINEA,
    BD_SUBLINEA,
    BD_EXISTENCIA_NETA,
    BD_COSTO,
    BD_TOTAL_INV,
    BD_MOTIVO,
)

# Columna auxiliar con la clave de cruce. Empieza por "_" para que el escritor
# la ignore.
CLAVE = "_CLAVE"


def normalizar_encabezado(texto: Any) -> str:
    """Forma comparable de un encabezado de Excel.

    ``"Sub- linea sistema "`` -> ``"SUB-LINEA SISTEMA"``;
    ``"MARCA copia"`` -> ``"MARCA COPIA"``.
    """
    s = str(texto).strip().upper()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s+", " ", s)
    for origen, destino in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")):
        s = s.replace(origen, destino)
    return s


def es_columna_existencia(encabezado_normalizado: str) -> bool:
    """``EXISTENCIA``, ``EXISTENCIA SEP 10``, ``EXISTENCIA OCT 02``..."""
    return encabezado_normalizado == EXISTENCIA or encabezado_normalizado.startswith(
        EXISTENCIA + " "
    )


def clave_referencia(valor: Any) -> str:
    """Clave de cruce de una referencia entre la plantilla y la base de datos.

    Insensible a mayusculas y espacios. Los numeros que pandas lee como float
    (``1060.0``) se llevan a su forma entera (``"1060"``), que es como los
    exporta la base de datos.
    """
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, bool):
        return str(valor).upper()
    if isinstance(valor, int):
        return str(valor)
    if isinstance(valor, float):
        return str(int(valor)) if valor.is_integer() else str(valor)
    s = str(valor).strip().upper()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def claves(serie: pd.Series) -> pd.Series:
    """``clave_referencia`` aplicada a una columna completa."""
    return serie.map(clave_referencia).astype("string")
