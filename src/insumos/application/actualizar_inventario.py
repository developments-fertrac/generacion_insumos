"""Caso de uso: actualizar el inventario general desde la base de datos.

1. Lee la exportacion de base de datos y la filtra (``inventario_bd.yaml``).
2. Lee la plantilla y le aplica las reglas (``inventario.yaml``) usando la base
   de datos filtrada, la Matriz USD y la Distribucion como contexto.
3. Si todo cuadra, escribe el archivo y el reporte de auditoria.

No sabe de Excel, COM ni rutas: recibe puertos ya construidos.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

import insumos.domain.rules  # noqa: F401  (registra las reglas)
from insumos.domain.audit import AuditTrail
from insumos.domain.pipeline import RulePipeline
from insumos.domain.ports import (
    EscritorInventario,
    EscritorReporte,
    EvidenciaEjecucion,
    FuenteInventario,
)
from insumos.domain.rules.base import RuleContext
from insumos.domain.rules.inventario import columnas as C


@dataclass(frozen=True)
class ResultadoActualizacion:
    datos: pd.DataFrame
    audit_bd: AuditTrail
    audit_inventario: AuditTrail
    archivo_salida: Path | None
    archivo_reporte: Path | None
    origen_plantilla: str

    def resumen(self) -> dict[str, Any]:
        nuevas = int(self.datos["_NUEVA"].sum()) if "_NUEVA" in self.datos.columns else 0
        total = pd.to_numeric(self.datos[C.TOTAL_INV], errors="coerce").fillna(0).sum()
        existencia = pd.to_numeric(self.datos[C.EXISTENCIA], errors="coerce").fillna(0).sum()
        return {
            "referencias": len(self.datos),
            "referencias_nuevas": nuevas,
            "existencia_total": float(existencia),
            "total_inv": float(total),
            "excluidas_por_bd": len(self.audit_bd.tabla_eliminaciones()),
            "eliminadas_de_plantilla": len(self.audit_inventario.tabla_eliminaciones()),
            "plantilla": self.origen_plantilla,
            "archivo_salida": str(self.archivo_salida) if self.archivo_salida else "",
            "archivo_reporte": str(self.archivo_reporte) if self.archivo_reporte else "",
        }


@dataclass
class ActualizarInventario:
    fuente: FuenteInventario
    escritor: EscritorInventario
    reporte: EscritorReporte
    reglas_bd: Path
    reglas_inventario: Path
    hoy: date = field(default_factory=date.today)

    def __call__(self) -> ResultadoActualizacion:
        # 1. Base de datos
        audit_bd = AuditTrail()
        bd = RulePipeline.from_yaml(self.reglas_bd, audit_bd).run(
            self.fuente.inventario_bd(), RuleContext(hoy=self.hoy)
        )
        excluidas = audit_bd.tabla_eliminaciones()
        motivos = {
            C.clave_referencia(ref): str(mot)
            for ref, mot in zip(
                excluidas.get(C.BD_REFERENCIA, pd.Series(dtype=object)),
                excluidas.get(C.BD_MOTIVO, pd.Series(dtype=object)),
                strict=True,
            )
            if C.clave_referencia(ref)
        }

        # 2. Plantilla
        plantilla = self.fuente.plantilla()
        ctx = RuleContext(
            inventario_bd=bd.df,
            matriz_usd=self.fuente.matriz_usd(),
            distribucion=self.fuente.distribucion(),
            hoy=self.hoy,
            params={"motivos_excluidos_bd": motivos},
        )
        audit_inv = AuditTrail()
        resultado = RulePipeline.from_yaml(self.reglas_inventario, audit_inv).run(plantilla.datos, ctx)

        # 3. Salida (solo si el pipeline completo cuadro: validar_salida lanza si no)
        salida = self.escritor.escribir(resultado.df, hoy=self.hoy)
        reporte = self.reporte.escribir(
            EvidenciaEjecucion(
                audit_bd=audit_bd,
                audit_inventario=audit_inv,
                datos=resultado.df,
                archivo_salida=salida,
                hoy=self.hoy,
                fuentes=tuple(self.fuente.fuentes_usadas()),
            )
        )
        return ResultadoActualizacion(
            datos=resultado.df,
            audit_bd=audit_bd,
            audit_inventario=audit_inv,
            archivo_salida=salida,
            archivo_reporte=reporte,
            origen_plantilla=plantilla.origen,
        )
