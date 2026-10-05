"""Comparacion de salidas para el protocolo golden (seccion 7 del plan).

Un golden file responde a una pregunta concreta: *¿la salida nueva dice lo mismo
que la salida de la implementacion actual?* Para responderla no basta con
comparar bytes ni con ``assert_frame_equal``: los .xlsx que produce el legacy
traen marcas de tiempo, orden de escritura y formatos que cambian sin que
cambie un solo dato de negocio. Comparar asi produce cientos de diferencias
falsas y el equipo deja de mirar el reporte.

Por eso la comparacion es **por clave de negocio** y clasifica cada diferencia:

===========================  ==========================================================
Clase                        Que significa
===========================  ==========================================================
``solo_legacy``              La fila esta en la salida actual y no en la nueva.
``solo_nuevo``               La fila esta en la nueva y no en la actual.
``valor_distinto``           La fila esta en ambas y algun dato difiere.
``columna_faltante``         La columna existe en la actual y no en la nueva.
``columna_extra``            La columna existe en la nueva y no en la actual.
``filas_extra``              La nueva tiene filas que no corresponden a ninguna clave.
===========================  ==========================================================

El criterio de exito del plan es *cero diferencias inexplicadas*: este modulo
no decide cuales son explicables, las lista todas y deja esa lectura a quien
conoce el negocio. Preferible un reporte de 40 filas a un booleano que nadie
audita.
"""

from __future__ import annotations

import math
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

# Valores que el negocio considera "no informado". ``#N/A`` NO entra aqui: en el
# inventario significa "el ERP no encontro la linea", que es informacion.
_VACIOS = frozenset({"", " ", "-", "none", "nan", "null"})

# Tolerancia por defecto al comparar numeros. El legacy acumula operaciones en
# coma flotante a traves de Excel, asi que una diferencia de 1e-9 no es una
# diferencia de negocio.
_TOLERANCIA_POR_DEFECTO = 1e-6


def _es_vacio(valor: Any) -> bool:
    if valor is None:
        return True
    if isinstance(valor, float) and math.isnan(valor):
        return True
    if isinstance(valor, str):
        return valor.strip().lower() in _VACIOS
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return False


def _normalizar(valor: Any) -> Any:
    """Deja comparable lo que el negocio considera el mismo valor.

    No redondea: la tolerancia la aplica ``_iguales`` con ``math.isclose``.
    Redondear aqui seria una segunda politica de comparacion que ademas habria
    que mantener sincronizada con la otra.
    """
    if _es_vacio(valor):
        return None
    if isinstance(valor, str):
        return valor.strip()
    if isinstance(valor, bool):
        return valor
    try:
        numero = float(valor)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return valor
    if math.isnan(numero):
        return None
    return numero


def _iguales(a: Any, b: Any, tolerancia: float) -> bool:
    na, nb = _normalizar(a), _normalizar(b)
    if na is None and nb is None:
        return True
    if isinstance(na, float) and isinstance(nb, float):
        return math.isclose(na, nb, rel_tol=0.0, abs_tol=tolerancia)
    return bool(na == nb)


@dataclass(frozen=True)
class Diferencia:
    """Una diferencia puntual entre dos salidas."""

    clase: str
    clave: Any
    columna: str = ""
    valor_actual: Any = None
    valor_nuevo: Any = None

    def __str__(self) -> str:
        if self.clase in {"solo_legacy", "solo_nuevo", "filas_extra"}:
            return f"{self.clase}: clave={self.clave!r}"
        return (
            f"{self.clase}: clave={self.clave!r} columna={self.columna!r} "
            f"actual={self.valor_actual!r} nuevo={self.valor_nuevo!r}"
        )


@dataclass
class ResultadoComparacion:
    """Todo lo que hay que revisar antes de dar el switch-over por bueno."""

    diferencias: list[Diferencia] = field(default_factory=list)
    filas_actual: int = 0
    filas_nuevo: int = 0
    claves_comunes: int = 0
    columnas_comunes: list[str] = field(default_factory=list)

    @property
    def iguales(self) -> bool:
        return not self.diferencias

    @property
    def por_clase(self) -> dict[str, int]:
        conteo: dict[str, int] = {}
        for d in self.diferencias:
            conteo[d.clase] = conteo.get(d.clase, 0) + 1
        return dict(sorted(conteo.items()))

    def por_regla(self) -> pd.DataFrame:
        """Las diferencias como tabla, para pegarlas en el reporte de switch-over."""
        if not self.diferencias:
            return pd.DataFrame(columns=["clase", "clave", "columna", "actual", "nuevo"])
        return pd.DataFrame(
            [
                {
                    "clase": d.clase,
                    "clave": d.clave,
                    "columna": d.columna,
                    "actual": d.valor_actual,
                    "nuevo": d.valor_nuevo,
                }
                for d in self.diferencias
            ]
        )

    def resumen(self) -> str:
        if self.iguales:
            return (
                f"Sin diferencias: {self.filas_nuevo} filas, "
                f"{len(self.columnas_comunes)} columnas comparadas."
            )
        partes = [f"{n} {clase}" for clase, n in self.por_clase.items()]
        return (
            f"{len(self.diferencias)} diferencias ({', '.join(partes)}) "
            f"sobre {self.claves_comunes} claves comunes."
        )


def comparar(
    actual: pd.DataFrame,
    nuevo: pd.DataFrame,
    claves: Sequence[str],
    columnas: Sequence[str] | None = None,
    tolerancia: float = _TOLERANCIA_POR_DEFECTO,
) -> ResultadoComparacion:
    """Compara dos salidas de inventario por clave de negocio.

    ``claves`` son las columnas que identifican una referencia (por ejemplo
    ``["REFERENCIA"]``). Se comparan solo las filas cuya clave existe en ambas
    salidas; las que no, se reportan como ``solo_legacy`` / ``solo_nuevo``.

    ``columnas`` acota la comparacion. Si se omite, se comparan las columnas
    comunes. Omitirla a proposito es una decision: comparar tambien las columnas
    de formato (color, ancho) produce ruido que nadie revisa.
    """
    resultado = ResultadoComparacion(
        filas_actual=len(actual), filas_nuevo=len(nuevo)
    )

    faltantes = [c for c in claves if c not in actual.columns or c not in nuevo.columns]
    if faltantes:
        raise ValueError(
            f"La clave de negocio no existe en ambas salidas: {faltantes}. "
            "Sin clave no se puede comparar por referencia; comparar por indice "
            "produce diferencias falsas en cuanto el orden de filas cambia."
        )

    idx_actual = _indexar(actual, claves)
    idx_nuevo = _indexar(nuevo, claves)

    resultado.claves_comunes = len(idx_actual.keys() & idx_nuevo.keys())

    for clave in sorted(idx_actual.keys() - idx_nuevo.keys(), key=str):
        resultado.diferencias.append(Diferencia("solo_legacy", clave))
    for clave in sorted(idx_nuevo.keys() - idx_actual.keys(), key=str):
        resultado.diferencias.append(Diferencia("solo_nuevo", clave))

    if columnas is None:
        comunes = [c for c in actual.columns if c in nuevo.columns]
        for c in actual.columns:
            if c not in nuevo.columns and c not in claves:
                resultado.diferencias.append(Diferencia("columna_faltante", "", c))
        for c in nuevo.columns:
            if c not in actual.columns and c not in claves:
                resultado.diferencias.append(Diferencia("columna_extra", "", c))
    else:
        comunes = list(columnas)

    resultado.columnas_comunes = comunes

    for clave in sorted(idx_actual.keys() & idx_nuevo.keys(), key=str):
        fila_a = actual.loc[idx_actual[clave]]
        fila_n = nuevo.loc[idx_nuevo[clave]]
        for columna in comunes:
            if columna in claves:
                continue
            valor_a = fila_a[columna]
            valor_n = fila_n[columna]
            if not _iguales(valor_a, valor_n, tolerancia):
                resultado.diferencias.append(
                    Diferencia("valor_distinto", clave, columna, valor_a, valor_n)
                )

    return resultado


def _indexar(df: pd.DataFrame, claves: Sequence[str]) -> dict[Any, Any]:
    """Indexa por clave de negocio, detectando claves duplicadas.

    Una clave repetida hace que la comparacion sea ambigua: sin avisar, se
    compararia siempre contra la primera coincidencia y las diferencias de las
    demas quedarian ocultas. Se falla en vez de adivinar.
    """
    if df.empty:
        return {}

    valores: Any
    if len(claves) == 1:
        valores = df[claves[0]].map(_normalizar)
    else:
        valores = pd.MultiIndex.from_frame(df[list(claves)])

    if valores.duplicated().any():
        repetidas = valores[valores.duplicated()].unique()[:5]
        raise ValueError(
            f"Clave de negocio repetida en la salida ({list(claves)}): {list(repetidas)}. "
            "Con claves duplicadas la comparacion hide diferencias; corrige los datos "
            "de entrada antes de comparar."
        )

    return dict(zip(valores, df.index, strict=True))


# ----------------------------------------------------------------------
# Comparacion por posicion (ventas: no hay clave unica por linea)
# ----------------------------------------------------------------------


def comparar_por_posicion(
    actual: pd.DataFrame,
    nuevo: pd.DataFrame,
    *,
    tolerancia: float = 1e-9,
    max_ejemplos: int = 20,
) -> pd.DataFrame:
    """Diferencias fila a fila entre dos salidas que deben venir en el mismo orden.

    Devuelve un DataFrame (TIPO, COLUMNA, FILA, ACTUAL, NUEVO); vacio = paridad.
    Por columna se reportan hasta ``max_ejemplos`` filas y el total en una fila
    ``TOTAL``, para que una columna mal calculada no produzca 50.000 lineas.
    """
    filas: list[dict[str, Any]] = []
    for c in actual.columns.difference(nuevo.columns):
        filas.append({"TIPO": "columna_faltante", "COLUMNA": c, "FILA": None, "ACTUAL": None, "NUEVO": None})
    for c in nuevo.columns.difference(actual.columns):
        filas.append({"TIPO": "columna_extra", "COLUMNA": c, "FILA": None, "ACTUAL": None, "NUEVO": None})
    if len(actual) != len(nuevo):
        filas.append({"TIPO": "filas", "COLUMNA": "", "FILA": None, "ACTUAL": len(actual), "NUEVO": len(nuevo)})

    n = min(len(actual), len(nuevo))
    a = actual.reset_index(drop=True).iloc[:n]
    b = nuevo.reset_index(drop=True).iloc[:n]
    for c in a.columns.intersection(b.columns):
        distintas = [i for i, (x, y) in enumerate(zip(a[c], b[c], strict=True)) if not _iguales(x, y, tolerancia)]
        for i in distintas[:max_ejemplos]:
            filas.append({"TIPO": "valor_distinto", "COLUMNA": c, "FILA": i, "ACTUAL": a.at[i, c], "NUEVO": b.at[i, c]})
        if len(distintas) > max_ejemplos:
            filas.append({"TIPO": "TOTAL", "COLUMNA": c, "FILA": None, "ACTUAL": len(distintas), "NUEVO": None})
    return pd.DataFrame(filas, columns=["TIPO", "COLUMNA", "FILA", "ACTUAL", "NUEVO"])
