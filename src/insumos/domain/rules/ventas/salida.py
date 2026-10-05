"""Reglas de salida: forma final de la hoja VENTAS 2026 y validacion."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult
from insumos.domain.rules.registry import rule
from insumos.domain.rules.ventas import columnas as C


class SalidaVentasInvalida(ValueError):
    """El resultado no cumple lo minimo para escribirse en la plantilla."""


@rule("ven.alinear_con_plantilla")
@dataclass(frozen=True)
class AlinearConPlantilla(ReglaBase):
    """Columnas en el orden de la plantilla; las de formula las calcula Excel.

    ``ctx.params["columnas_plantilla"]``: encabezados normalizados de la hoja
    VENTAS 2026. Las columnas que la plantilla no tiene van al final (ordenadas
    por nombre; el escritor solo usa las de la plantilla).
    """

    description = "Ordena columnas como la plantilla y omite las columnas de formula"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        plantilla = list(ctx.params.get("columnas_plantilla") or ())
        if not plantilla:
            raise SalidaVentasInvalida("El contexto no trae 'columnas_plantilla'")
        formula = {C.clave_comparacion(c) for c in C.COLUMNAS_FORMULA}
        df = df.copy()
        for col in plantilla:
            if col not in df.columns and C.clave_comparacion(col) not in formula:
                df[col] = np.nan
        nuevas = sorted(
            c for c in set(df.columns) - set(plantilla) if C.clave_comparacion(c) not in formula
        )
        finales = [c for c in plantilla if c in df.columns] + nuevas
        return RuleResult(df=df[finales], metrics={"columnas": len(finales), "fuera_de_plantilla": nuevas})


@rule("ven.validar_salida")
@dataclass(frozen=True)
class ValidarSalida(ReglaBase):
    """Falla antes de escribir si el resultado no tiene sentido."""

    description = "Verifica columnas clave, filas y que todas las fechas sean del ano en curso"
    requeridas: tuple[str, ...] = (C.REFERENCIA, C.NIT_CLIENTE, C.FECHA)

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        faltan = [c for c in self.requeridas if c not in df.columns]
        if faltan:
            raise SalidaVentasInvalida(f"La salida no tiene {faltan}")
        if df.empty:
            raise SalidaVentasInvalida("La salida no tiene filas")
        if ctx.hoy is not None:
            anios = pd.to_datetime(df[C.FECHA], errors="coerce").dt.year
            otros = int((anios != ctx.hoy.year).sum())
            if otros:
                raise SalidaVentasInvalida(f"{otros} filas con FECHA fuera de {ctx.hoy.year}")
        resultado = RuleResult(df=df, metrics={"filas": len(df)})
        if "MES NO." in df.columns:
            resultado.metrics["por_mes"] = {
                int(m): int(n) for m, n in df["MES NO."].value_counts().sort_index().items()
            }
        return resultado
