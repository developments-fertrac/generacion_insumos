"""Caso de uso: informe de facturas -> datos de la hoja VENTAS 2026.

Aplica ``config/rules/ventas.yaml`` con las tablas de referencia como contexto.
No escribe: la proyeccion sobre la plantilla (formulas, Resum Mes, tablas
dinamicas, contrasena) la sigue haciendo el escritor COM del proceso actual.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

import insumos.domain.rules  # noqa: F401  (registra las reglas)
from insumos.domain.audit import AuditTrail
from insumos.domain.pipeline import RulePipeline
from insumos.domain.ports import EntradasVentas
from insumos.domain.rules.base import RuleContext


@dataclass(frozen=True)
class ResultadoVentas:
    datos: pd.DataFrame
    audit: AuditTrail

    def resumen(self) -> dict[str, Any]:
        eliminadas = self.audit.tabla_eliminaciones()
        por_regla = (
            eliminadas["_regla"].value_counts().to_dict() if "_regla" in eliminadas.columns else {}
        )
        return {
            "registros": len(self.datos),
            "eliminadas": len(eliminadas),
            "eliminadas_por_regla": por_regla,
        }


@dataclass(frozen=True)
class TransformarVentas:
    reglas: Path
    hoy: date

    def __call__(self, entradas: EntradasVentas) -> ResultadoVentas:
        tablas = {
            "inventario": entradas.inventario,
            "myr": entradas.myr,
            "matriz_clientes": entradas.matriz_clientes,
        }
        if entradas.precios_licitados is not None:
            tablas["precios_licitados"] = entradas.precios_licitados
        ctx = RuleContext(
            tablas=tablas,
            hoy=self.hoy,
            params={
                "columnas_plantilla": entradas.columnas_plantilla,
                "nits_licitados": entradas.nits_licitados,
            },
        )
        audit = AuditTrail()
        resultado = RulePipeline.from_yaml(self.reglas, audit).run(entradas.informe, ctx)
        return ResultadoVentas(datos=resultado.df, audit=audit)
