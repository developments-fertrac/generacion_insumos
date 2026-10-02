"""El motor de reglas: encadena reglas en el orden que declara el YAML.

Este es el objeto que reemplaza las fases 1-7 de ``execute()`` en
``actualizacion_inventario.py`` (700 lineas con sub-fases 4b, 5b, 5c y ramas
duplicadas). Aqui el orden vive en ``config/rules/inventario.yaml`` y cambiar
una validacion es editar una linea, no un metodo de 700 lineas.
"""

from __future__ import annotations

import time
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from insumos.domain.audit import AuditTrail, PasoAuditoria
from insumos.domain.reglas_declarativas import (
    DeclaracionPipeline,
    leer_declaracion,
)
from insumos.domain.rules import registry
from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult


class ErrorDePipeline(RuntimeError):
    """Una regla fallo dentro del pipeline."""


@dataclass(frozen=True)
class ResultadoPipeline:
    """Salida completa del pipeline: el df final y el rastro."""

    df: pd.DataFrame
    audit: AuditTrail

    @property
    def eliminadas(self) -> pd.DataFrame:
        return self.audit.tabla_eliminaciones()


class RulePipeline:
    """Ejecuta una lista de reglas en orden sobre un DataFrame."""

    def __init__(self, reglas: Sequence[ReglaBase], audit: AuditTrail) -> None:
        self.reglas = list(reglas)
        self.audit = audit

    def __len__(self) -> int:
        return len(self.reglas)

    @property
    def ids(self) -> list[str]:
        return [r.id for r in self.reglas]

    def run(self, df: pd.DataFrame, ctx: RuleContext) -> ResultadoPipeline:
        """Aplica cada regla en orden y devuelve el df final con su rastro.

        Una regla que lanza detiene el pipeline. Se elige propagar el error en
        vez de tragarselo porque un fallo silencioso en una regla de inventario
        produce un archivo de salida incorrecto que parece correcto, y el
        dano se descubre semanas despues al comparar totales.
        """
        actual = df
        for i, regla in enumerate(self.reglas, 1):
            antes = len(actual)
            inicio = time.perf_counter()

            try:
                resultado = regla.apply(actual, ctx)
            except Exception as e:
                raise ErrorDePipeline(
                    f"Regla {i}/{len(self.reglas)} '{regla.id}' fallo: {e}"
                ) from e

            if not isinstance(resultado, RuleResult):
                raise ErrorDePipeline(
                    f"Regla '{regla.id}' devolvio {type(resultado).__name__}, "
                    "se esperaba RuleResult"
                )

            duracion_ms = (time.perf_counter() - inicio) * 1000
            paso = PasoAuditoria(
                id_regla=regla.id,
                description=regla.description,
                filas_antes=antes,
                filas_despues=len(resultado.df),
                modificadas=resultado.modified,
                eliminadas=resultado.eliminadas,
                warnings=tuple(resultado.warnings),
                metrics={**resultado.metrics, "ms": round(duracion_ms, 2)},
            )
            self.audit.registrar(paso)
            self.audit.acumular_eliminaciones(resultado.removed, regla.id)
            actual = resultado.df

        return ResultadoPipeline(df=actual, audit=self.audit)

    # ------------------------------------------------------------------
    # Construccion desde declaracion
    # ------------------------------------------------------------------

    @classmethod
    def desde_declaracion(
        cls,
        declaracion: DeclaracionPipeline,
        audit: AuditTrail,
    ) -> RulePipeline:
        """Construye el pipeline resolviendo ids del YAML contra el registro."""
        reglas: list[ReglaBase] = []
        for entrada in declaracion.reglas_activas:
            reglas.append(registry.construir(entrada.id, dict(entrada.params)))
        audit.pipeline = declaracion.nombre
        audit.version = declaracion.version
        return cls(reglas, audit)

    @classmethod
    def from_yaml(cls, path: Path | str, audit: AuditTrail) -> RulePipeline:
        """Carga el pipeline desde ``config/rules/*.yaml``.

        Falla de forma ruidosa y temprano si el YAML pide una regla que no
        existe o si no declara reglas activas, en vez de dejar pasar los datos
        sin transformar.
        """
        declaracion = leer_declaracion(Path(path))
        return cls.desde_declaracion(declaracion, audit)

    @classmethod
    def desde_ids(
        cls,
        ids: Iterable[str],
        audit: AuditTrail,
    ) -> RulePipeline:
        """Construye desde una lista de ids. Util para pruebas."""
        reglas = [registry.construir(i) for i in ids]
        return cls(reglas, audit)
