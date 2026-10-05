"""Lectura de las entradas de ventas desde libros ya descifrados (``BytesIO``).

Cada funcion devuelve un DataFrame con encabezados normalizados (sin tildes,
mayusculas), listo para ``ctx.tablas``. El descifrado y la busqueda del
archivo mas reciente siguen en ``tasks/actualizacion_ventas.py`` hasta que se
migren a un adaptador ``FuenteVentas`` (fase 4b).
"""

from __future__ import annotations

import re
from io import BytesIO

import pandas as pd
from unidecode import unidecode

from insumos.domain.rules.ventas import columnas as C

HOJAS_VENTAS = ("VENTAS 2026", "VENTAS 2025")
HOJA_MYR = "COSTOS INV FINAL"
HOJA_MATRIZ = "CLIENTES GENERAL"
HOJA_LICITADOS = "PRECIO UNIT LICITADOS"


def _clave_hoja(nombre: object) -> str:
    s = re.sub(r"\s+", " ", unidecode(str(nombre)).lower().strip())
    return re.sub(r"[^a-z0-9 ]", "", s)


def buscar_hoja(stream: BytesIO, objetivos: tuple[str, ...]) -> str:
    """Igual que ``core.excel_utils.find_sheet_name``: exacta, luego contenida, luego la primera."""
    stream.seek(0)
    nombres = [str(n) for n in pd.ExcelFile(stream, engine="openpyxl").sheet_names]
    claves = {_clave_hoja(n): n for n in nombres}
    buscadas = [_clave_hoja(t) for t in objetivos]
    for t in buscadas:
        if t in claves:
            return claves[t]
    for t in buscadas:
        for clave, real in claves.items():
            if t in clave:
                return real
    if nombres:
        return nombres[0]
    raise ValueError("El libro no tiene hojas.")


def _limpiar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalizar(df: pd.DataFrame) -> pd.DataFrame:
    salida = df.copy()
    salida.columns = [C.normalizar_encabezado(c) for c in df.columns]
    return salida


def leer_plantilla_ventas(stream: BytesIO, hojas: tuple[str, ...] = HOJAS_VENTAS) -> tuple[str, pd.DataFrame]:
    """Hoja de ventas de la plantilla (encabezado en la fila 2)."""
    for hoja in hojas:
        stream.seek(0)
        try:
            df = pd.read_excel(stream, sheet_name=hoja, engine="openpyxl", header=1)
        except ValueError:
            continue
        return hoja, normalizar(_limpiar(df))
    raise ValueError(f"No se encontro ninguna hoja de ventas valida. Intentadas: {list(hojas)}")


def leer_inventario_lineas(stream: BytesIO) -> pd.DataFrame:
    """Hoja INVENTARIO del inventario general (encabezado en la fila 2, o la 1)."""
    hoja = buscar_hoja(stream, ("INVENTARIO", "INVENTARIO GENERAL"))

    def _leer(fila: int) -> pd.DataFrame:
        stream.seek(0)
        return _limpiar(pd.read_excel(stream, sheet_name=hoja, engine="openpyxl", header=fila))

    df = _leer(1)
    if C.REFERENCIA not in df.columns:
        df = _leer(0)
    return normalizar(df)


def leer_myr(stream: BytesIO, motor: str = "openpyxl") -> pd.DataFrame:
    """Hoja COSTOS INV FINAL del MYR (encabezado en la fila 3); ``motor="pyxlsb"`` para .xlsb."""
    stream.seek(0)
    df = normalizar(_limpiar(pd.read_excel(stream, sheet_name=HOJA_MYR, engine=motor, header=2)))
    if "REFERENCIA FERTRAC" in df.columns:
        df = df.rename(columns={"REFERENCIA FERTRAC": C.REFERENCIA})
    return df


def leer_matriz_clientes(stream: BytesIO) -> pd.DataFrame:
    stream.seek(0)
    return normalizar(pd.read_excel(stream, sheet_name=HOJA_MATRIZ, engine="openpyxl"))


def leer_precios_licitados(stream: BytesIO) -> tuple[pd.DataFrame, tuple[str, ...]]:
    """Hoja PRECIO UNIT LICITADOS en formato largo.

    - Fila 1, columnas C:H: NIT de cada cliente licitado.
    - Fila 4: encabezados; la columna de referencia es la que dice
      REFERENCIA FERTRAC (o la primera que dice REFERENCIA).
    - Devuelve (``REFERENCIA``, ``NIT``, ``PRECIO``) con precio > 0, y la
      tupla de NIT licitados.

    Lanza ``ValueError`` si la hoja no tiene la forma esperada.
    """
    stream.seek(0)
    precios = pd.read_excel(stream, sheet_name=HOJA_LICITADOS, engine="openpyxl", header=3)
    precios.columns = [str(c).strip() for c in precios.columns]
    col_ref = next(
        (c for c in precios.columns if "REFERENCIA" in c.upper() and "FERTRAC" in c.upper()),
        next((c for c in precios.columns if "REFERENCIA" in c.upper()), None),
    )
    if col_ref is None:
        raise ValueError("No se encontro columna REFERENCIA")

    stream.seek(0)
    primera = pd.read_excel(stream, sheet_name=HOJA_LICITADOS, engine="openpyxl", header=None, nrows=1)
    nit_a_columna: dict[str, str] = {}
    for idx in range(2, 8):  # columnas C:H
        if idx < len(primera.columns):
            valor = primera.iloc[0, idx]
            if pd.notna(valor):
                nit = C.texto_licitado(valor)
                if nit and nit not in ("", "nan", "None"):
                    nit_a_columna[nit] = precios.columns[idx]
    if not nit_a_columna:
        raise ValueError("No hay NITs en headers")

    tabla: dict[str, dict[str, float]] = {}
    for _, fila in precios.iterrows():
        ref = C.texto_licitado(fila[col_ref]) if pd.notna(fila[col_ref]) else ""
        if not ref or ref in ("nan", "None"):
            continue
        por_nit: dict[str, float] = {}
        for nit, col in nit_a_columna.items():
            precio = fila[col]
            if pd.notna(precio):
                try:
                    valor = float(precio)
                except (TypeError, ValueError):
                    continue
                if valor > 0:
                    por_nit[nit] = valor
        if por_nit:
            tabla[ref] = por_nit  # una referencia repetida: gana la ultima fila

    filas = [{"REFERENCIA": r, "NIT": n, "PRECIO": p} for r, d in tabla.items() for n, p in d.items()]
    largo = pd.DataFrame(filas, columns=["REFERENCIA", "NIT", "PRECIO"])
    return largo, tuple(nit_a_columna)
