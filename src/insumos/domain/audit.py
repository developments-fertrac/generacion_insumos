"""Rastro de auditoria: que hizo cada regla, con evidencia.

Sustituye al ``EliminacionTracker`` de 400 lineas que vive hoy dentro de
``actualizacion_inventario.py``. Ahi el rastro se imprimia en consola y se
perdia; el unico registro durable era un log de texto.

Aqui el rastro es un objeto que se pasa al escritor de Excel (hoja
"FUENTES DE DATOS") y a los golden files, de modo que "por que desaparecio
esta referencia" tiene una respuesta consultable y no una busqueda en logs.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class PasoAuditoria:
    """Lo que hizo una regla en una pasada del pipeline."""

    id_regla: str
    description: str
    filas_antes: int
    filas_despues: int
    modificadas: int
    eliminadas: int
    warnings: tuple[str, ...] = ()
    metrics: dict[str, Any] = field(default_factory=dict)

    @property
    def elimino_filas(self) -> bool:
        return self.filas_antes > self.filas_despues


@dataclass
class AuditTrail:
    """Acumula los pasos de una ejecucion y las filas eliminadas con su motivo."""

    pasos: list[PasoAuditoria] = field(default_factory=list)
    eliminaciones: list[pd.DataFrame] = field(default_factory=list)
    pipeline: str = ""
    version: str = ""

    def registrar(self, paso: PasoAuditoria) -> None:
        self.pasos.append(paso)

    def acumular_eliminaciones(self, filas: pd.DataFrame, id_regla: str) -> None:
        """Guarda filas eliminadas, agregando el nombre de la regla que las quito."""
        if filas.empty:
            return
        copia = filas.copy()
        if "_regla" not in copia.columns:
            copia["_regla"] = id_regla
        self.eliminaciones.append(copia)

    def resumen(self) -> dict[str, Any]:
        """Diccionario de metricas, apto para log o para el resumen del correo."""
        return {
            "pipeline": self.pipeline,
            "version": self.version,
            "pasos": len(self.pasos),
            "filas_finales": self.pasos[-1].filas_despues if self.pasos else 0,
            "filas_iniciales": self.pasos[0].filas_antes if self.pasos else 0,
            "total_eliminadas": sum(p.eliminadas for p in self.pasos),
            "total_modificadas": sum(p.modificadas for p in self.pasos),
            "warnings": sum(len(p.warnings) for p in self.pasos),
            "por_regla": {
                p.id_regla: {
                    "eliminadas": p.eliminadas,
                    "modificadas": p.modificadas,
                }
                for p in self.pasos
            },
        }

    def tabla_eliminaciones(self) -> pd.DataFrame:
        """Todas las filas eliminadas en una sola tabla, con motivo y regla."""
        if not self.eliminaciones:
            return pd.DataFrame(columns=["_motivo", "_regla"])
        return pd.concat(self.eliminaciones, ignore_index=True)

    def tabla_pasos(self) -> pd.DataFrame:
        """Un paso por fila, para el reporte de ejecucion."""
        filas = [
            {
                "orden": i,
                "regla": p.id_regla,
                "descripcion": p.description,
                "filas_antes": p.filas_antes,
                "filas_despues": p.filas_despues,
                "eliminadas": p.eliminadas,
                "modificadas": p.modificadas,
                "warnings": "; ".join(p.warnings),
            }
            for i, p in enumerate(self.pasos, 1)
        ]
        return pd.DataFrame(filas)
