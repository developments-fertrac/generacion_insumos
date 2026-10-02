"""Lectura de los maestros que enriquecen el inventario.

- ``$2026 MATRIZ USD``             -> NOMBRE LISTA por referencia
- ``DISTRIBUCION DE MATRICES``     -> LIDER LINEA (gestor) y CLASIFICACION por LINEA

Port fiel de ``cargar_matriz_usd`` y ``cargar_distribucion`` del legacy: mismas
heuristicas de encabezado y de columnas, porque los archivos son mantenidos a
mano y su forma cambia sin aviso.
"""

from __future__ import annotations

import re
import unicodedata
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

from insumos.adapters.excel.lector import ExcelReader


def _norm(texto: Any) -> str:
    s = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[\s\-_]+", " ", s.lower()).strip()
    return s


def _bytes(ruta: Path, password: str | None) -> BytesIO:
    """Descifra una sola vez (la matriz pesa ~20 MB cifrada)."""
    return BytesIO(ExcelReader(ruta=ruta, password=password)._bytes_planos())


def leer_matriz_usd(ruta: Path, password: str | None, hoja_contiene: str = "2026") -> pd.DataFrame:
    """Devuelve ``CLAVE`` (referencia Fertrac) y ``DESCRIPCION`` (lista de precios)."""
    contenido = _bytes(ruta, password)
    with pd.ExcelFile(contenido) as xls:
        hoja = next(
            (h for h in xls.sheet_names if re.sub(r"^\$+", "", h).strip() == hoja_contiene),
            None,
        ) or next((h for h in xls.sheet_names if hoja_contiene in h), xls.sheet_names[0])
        crudo = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=20)
        fila = None
        for i in range(len(crudo)):
            valores = [_norm(v) for v in crudo.iloc[i] if pd.notna(v)]
            if any("referencia" in v and "fertrac" in v for v in valores) or any(
                "descripcion" in v and "lista" in v for v in valores
            ):
                fila = i
                break
        if fila is None:
            raise ValueError(f"{ruta.name}/{hoja}: no se encontro la fila de encabezado de la matriz")
        df = pd.read_excel(xls, sheet_name=hoja, header=fila)

    columnas = {c: _norm(c) for c in df.columns}
    col_ref = next(
        (c for c, n in columnas.items() if "referencia" in n and ("fertrac" in n or "inventario" in n)),
        None,
    )
    col_desc = next(
        (c for c, n in columnas.items() if "descripcion" in n and "lista" in n and "precio" in n),
        None,
    )
    if col_ref is None or col_desc is None:
        raise ValueError(
            f"{ruta.name}/{hoja}: faltan columnas REFERENCIA INVENTARIO FERTRAC o "
            f"DESCRIPCION LISTA DE PRECIOS. Columnas: {list(df.columns)}"
        )
    salida = pd.DataFrame({"CLAVE": df[col_ref], "DESCRIPCION": df[col_desc]})
    salida = salida[salida["CLAVE"].notna() & salida["CLAVE"].astype(str).str.strip().ne("")]
    return salida.reset_index(drop=True)


def leer_distribucion(ruta: Path, password: str | None) -> dict[str, dict[str, str]]:
    """``{"gestor": {LINEA: gestor}, "clasificacion": {LINEA: categoria}}``.

    Igual que el legacy: la LINEA se limpia de sufijos entre parentesis
    ("CEI (EN DESARROLLO)" -> "CEI") y, si una LINEA se repite, gana la ultima
    fila.
    """
    contenido = _bytes(ruta, password)
    crudo = pd.read_excel(contenido, sheet_name=0, header=None, nrows=15)
    fila = 2
    for i in range(min(10, len(crudo))):
        texto = " ".join(str(v).upper() for v in crudo.iloc[i] if pd.notna(v))
        if "LINEA" in texto and "GESTOR" in texto:
            fila = i
            break
    contenido.seek(0)
    df = pd.read_excel(contenido, sheet_name=0, header=fila)
    df.columns = [str(c).strip() for c in df.columns]
    idx = {_norm(c): c for c in df.columns}

    col_linea = idx.get("linea") or next((r for n, r in idx.items() if "linea" in n), None)
    col_gestor = idx.get("gestor") or idx.get("lider") or next(
        (r for n, r in idx.items() if "gestor" in n or "lider" in n), None
    )
    col_clasif = idx.get("categoria") or idx.get("clasificacion") or next(
        (r for n, r in idx.items() if "categ" in n or "clasificac" in n), None
    )
    gestor: dict[str, str] = {}
    clasificacion: dict[str, str] = {}
    if col_linea is None:
        return {"gestor": gestor, "clasificacion": clasificacion}

    for _, fila_df in df.iterrows():
        linea = fila_df[col_linea]
        if pd.isna(linea) or not str(linea).strip():
            continue
        llave = re.sub(r"\s*\([^)]*\)\s*", "", str(linea).strip().upper()).strip()
        if not llave:
            continue
        if col_gestor is not None and pd.notna(fila_df[col_gestor]) and str(fila_df[col_gestor]).strip():
            gestor[llave] = str(fila_df[col_gestor]).strip()
        if col_clasif is not None and pd.notna(fila_df[col_clasif]) and str(fila_df[col_clasif]).strip():
            clasificacion[llave] = str(fila_df[col_clasif]).strip()
    return {"gestor": gestor, "clasificacion": clasificacion}
