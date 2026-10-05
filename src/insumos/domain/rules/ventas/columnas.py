"""Nombres de columna y utilidades puras del dominio de ventas.

Todo lo que antes vivia como propiedades de ``ActualizacionVentas``
(``COLS_INFORME``, ``COLS_FORMULAS_PRE/POST``, el mapeo semantico) queda aqui
para que las reglas y los adaptadores hablen el mismo idioma.
"""

from __future__ import annotations

import numbers
import re
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd
from unidecode import unidecode

# --- Informe de ventas (InformesDeVentas(Facturas)_268), nombres originales ----
INF_NIT = "NIT"
INF_NUMERO = "Número"
INF_NRO_DOC = "Nro. documento"
INF_FECHA_ORIGEN = "Fecha documento origen"
INF_FECHA = "Fecha"
INF_PREFIJO = "Prefijo"
INF_DOC_ORIGEN = "Documento origen"
INF_REFERENCIA = "Referencia"
INF_MARCA = "Marca"
INF_CANTIDAD = "Cantidad facturada"
INF_VALOR_BRUTO = "Valor bruto"
INF_COSTO_UNITARIO = "Costo unitario"
INF_VALOR_BASE_PV = "Valor base precio de venta"
INF_VALOR_UNITARIO = "Valor unitario"

COLUMNAS_INFORME = (
    INF_NIT, INF_NUMERO, INF_NRO_DOC, INF_FECHA_ORIGEN, INF_FECHA, INF_PREFIJO,
    INF_DOC_ORIGEN, INF_REFERENCIA, INF_MARCA, INF_CANTIDAD, INF_VALOR_BRUTO,
    INF_COSTO_UNITARIO, INF_VALOR_BASE_PV, "VR DESCUENTO",
)

# --- Hoja VENTAS 2026 (encabezados normalizados: sin tildes, mayusculas) ------
REFERENCIA = "REFERENCIA"
NIT_CLIENTE = "NIT CLIENTE"
FECHA = "FECHA"
ANO = "ANO"
MES = "MES"
MES_NO = "MES NO."
LINEA = "LINEA"
SUBLINEA = "SUBLINEA"
LIDER_LINEA = "LIDER LINEA"
COSTO_FACTOR_HOY = "COSTO FACTOR HOY"
DCTO_CONDICIONADO = "DCTO CONDICIONADO"
PORC_PIE_FACTURA = "PORCENTAJE DCTO A PIE DE FACTURA"
VR_DESCUENTO = "VR DESCUENTO"
VTA_ACORDADA = "VTA ACORDADA X UNIDAD LICITADO"
FECHAS_ACTUALIZACION = ("FECHA DE ACTUALIZACION", "FECHA ACTUALIZACION")

PREFIJOS_NOTA_CREDITO = ("NC", "NCDTO", "NDCTO")

MAPEO_SEMANTICO: dict[str, str] = {
    "NRO. DOCUMENTO CLIENTE": NIT_CLIENTE,
    "CIUDAD/SUCURSAL": "CIUDAD",
    "DESCRIPCION": "DESCRPCION",  # asi se llama en la plantilla
    "VALOR UNITARIO": "VR UNITARIO",
    "CANTIDAD FACTURADA": "CANTIDAD",
    "VALOR BRUTO": "VR TOTAL",
    "COSTO UNITARIO": "COSTO PROMEDIO",
    "VENDEDOR": "VEND",
    "PREFIJO": "DV",
}

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# Columnas que la plantilla calcula con formula: el proceso nunca las escribe.
COLUMNAS_FORMULA_PRE = (
    "MES", "MES NO.", "ANO", "SEMANA NUMERO",
    PORC_PIE_FACTURA,
    "DIF EN MG FACTURADO vs MG LICITADO",
)
COLUMNAS_FORMULA_POST = (
    "VENTA A COSTO TOTAL", "VENTA TOTAL NETA",
    "VTA NETA X UNIDAD (DCTO PIE FACT)",
    "VTA NETA X UNIDAD (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
    "MARGEN NETO (DCTO PIE FACT)",
    "MARGEN NETO (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
    "MARGEN NETO FACTOR HOY (DCTO PIE FACT)",
    "MARGEN NETO FACTOR HOY (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
    "PART", "SUMA MP",
    "DIF EN PRECIO FACTURADO vs LICITADO",
)
COLUMNAS_FORMULA = tuple(dict.fromkeys(COLUMNAS_FORMULA_PRE + COLUMNAS_FORMULA_POST))


def normalizar_encabezado(nombre: object) -> str:
    """``" Descripción  producto"`` -> ``"DESCRIPCION PRODUCTO"`` (como el legacy)."""
    s = unidecode(str(nombre).strip()).upper()
    return re.sub(r"\s+", " ", s).strip()


def clave_comparacion(nombre: object) -> str:
    """Forma laxa para comparar encabezados: minusculas, solo letras y numeros."""
    t = unidecode(str(nombre)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def a_numero_si_todo_convierte(serie: pd.Series) -> pd.Series:
    """``pd.to_numeric(errors="ignore")``: convierte solo si todos los valores son numero."""
    try:
        return pd.to_numeric(serie, errors="raise")
    except (ValueError, TypeError):
        return serie


def normalizar_referencia(valor: object) -> str | None:
    """La referencia es TEXTO: ``"0123"`` sigue ``"0123"``; vacio -> None.

    Solo una celda numerica entera pierde el ``.0`` que le agrega pandas
    (4591.0 -> ``"4591"``): en Excel una celda numerica no tiene ceros a la
    izquierda, asi que no hay nada que conservar.
    """
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, bool):
        return str(valor)
    if isinstance(valor, numbers.Integral):
        return str(int(valor))
    if isinstance(valor, numbers.Real) and float(valor).is_integer():
        return str(int(valor))
    s = str(valor).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    return s


def texto_licitado(valor: object) -> str:
    """Referencia o NIT de la hoja PRECIO UNIT LICITADOS.

    Celda numerica: se redondea al entero (10.05 -> ``"10"``, 900111222.0 ->
    ``"900111222"``). Celda de texto: se respeta tal cual (``"0123"``).
    """
    if isinstance(valor, numbers.Real) and not isinstance(valor, bool):
        return str(int(Decimal(str(valor)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)))
    return str(valor).strip()


def referencia_valida(valor: object) -> bool:
    """Una referencia real: no vacia, maximo 30 caracteres y 3 palabras."""
    try:
        if pd.isna(valor):
            return False
    except (TypeError, ValueError):
        pass
    t = str(valor).strip()
    return bool(t) and len(t) <= 30 and len(t.split()) <= 3


def limpiar_nit(serie: pd.Series) -> pd.Series:
    """Texto sin espacios ni ``.0`` final (el NIT llega como numero desde Excel)."""
    s = serie.astype(str).str.strip()
    return s.str.replace(r"\.0$", "", regex=True)
