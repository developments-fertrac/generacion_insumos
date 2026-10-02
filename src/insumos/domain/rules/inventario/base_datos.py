"""Reglas sobre la exportacion de base de datos (``Inventario.xlsx``).

Desde 2026-10 la base de datos entrega el inventario ya depurado: resta
faltantes, impo y aforo (``EXISTENCIA NETA``), clasifica cada referencia y
explica por que entra o no en ``MOTIVO``. Estas reglas no recalculan nada de
eso; solo aplican el criterio de negocio acordado sobre lo que llega:

1. Solo entran las filas cuyo ``MOTIVO`` inicia con ``INVENTARIO``.
2. Si la existencia es 0, el costo (y por tanto el total) es 0.

Tambien hay reglas genericas (duplicados, costo cero) que el pipeline de la
plantilla reutiliza con otros nombres de columna.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult, marcar_eliminadas
from insumos.domain.rules.inventario import columnas as C
from insumos.domain.rules.registry import rule


class ColumnasFaltantes(ValueError):
    """La exportacion no trae las columnas que el proceso necesita."""


@rule("inv.validar_columnas_bd")
@dataclass(frozen=True)
class ValidarColumnasBd(ReglaBase):
    """Falla temprano si la exportacion cambio de estructura.

    Sin esta regla, una columna renombrada en la base de datos (``COSTO`` ->
    ``COSTO UNITARIO``) produciria un inventario con costo vacio que parece
    correcto.
    """

    description = "Verifica que la exportacion de base de datos trae las columnas requeridas"
    requeridas: tuple[str, ...] = C.COLUMNAS_BD_REQUERIDAS

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        faltan = [c for c in self.requeridas if c not in df.columns]
        if faltan:
            raise ColumnasFaltantes(
                f"La exportacion de base de datos no trae {faltan}. "
                f"Columnas recibidas: {list(df.columns)}"
            )
        if df.empty:
            raise ColumnasFaltantes("La exportacion de base de datos no trae filas.")
        return RuleResult(df=df, metrics={"filas": len(df), "columnas": len(df.columns)})


@rule("inv.filtrar_motivo_inventario")
@dataclass(frozen=True)
class FiltrarMotivoInventario(ReglaBase):
    """Deja solo las filas cuyo MOTIVO inicia con INVENTARIO.

    El resto (``EXCLUIDO: marca Invalida - Varios``, ``EXCLUIDO: referencia -
    Valera``...) sale al reporte de eliminaciones con su motivo original, para
    que el negocio vea por que una referencia no esta en el informe.

    Nota: ``LLEGA AL INFORME`` no se usa como filtro. El criterio acordado es
    ``MOTIVO`` (las referencias con existencia 0 llegan con ``NO`` pero deben
    actualizarse a 0 en el informe).
    """

    description = "Conserva solo las referencias cuyo MOTIVO inicia con INVENTARIO"
    columna: str = C.BD_MOTIVO
    prefijo: str = "INVENTARIO"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        motivo = df[self.columna].astype("string").str.strip()
        entra = motivo.str.upper().str.startswith(self.prefijo.upper()).fillna(False).astype(bool)

        fuera = df[~entra]
        removidas = marcar_eliminadas(
            fuera.assign(_detalle=motivo[~entra].fillna("(MOTIVO vacio)")),
            f"MOTIVO no inicia con {self.prefijo}",
            self.id,
        )
        por_motivo = motivo[~entra].fillna("(vacio)").value_counts().to_dict()
        return RuleResult(
            df=df[entra].copy(),
            removed=removidas,
            metrics={
                "conservadas": int(entra.sum()),
                "excluidas": int((~entra).sum()),
                "excluidas_por_motivo": {str(k): int(v) for k, v in por_motivo.items()},
            },
        )


@rule("inv.eliminar_referencias_duplicadas")
@dataclass(frozen=True)
class EliminarReferenciasDuplicadas(ReglaBase):
    """Deja una sola fila por referencia y quita las filas sin referencia.

    Se conserva la primera aparicion (mismo criterio del paso 1 del legacy,
    ``PASO_1_REFERENCIA_DUPLICADA``). La comparacion usa la clave normalizada:
    ``"abc-1 "`` y ``"ABC-1"`` son la misma referencia.
    """

    description = "Elimina referencias vacias y duplicadas (conserva la primera)"
    columna: str = C.REFERENCIA

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        clave = C.claves(df[self.columna])
        vacia = clave.isna() | clave.eq("")
        duplicada = ~vacia & clave.duplicated(keep="first")

        removidas = pd.concat(
            [
                marcar_eliminadas(df[vacia], "Referencia vacia", self.id),
                marcar_eliminadas(df[duplicada], "Referencia duplicada", self.id),
            ]
        )
        quitar = vacia | duplicada
        return RuleResult(
            df=df[~quitar].copy(),
            removed=removidas,
            warnings=(
                [f"{int(duplicada.sum())} referencias duplicadas en {self.columna}"]
                if duplicada.any()
                else []
            ),
            metrics={"vacias": int(vacia.sum()), "duplicadas": int(duplicada.sum())},
        )


@rule("inv.costo_cero_sin_existencia")
@dataclass(frozen=True)
class CostoCeroSinExistencia(ReglaBase):
    """Si la existencia es 0, el costo y el total son 0.

    Regla de negocio explicita: una referencia sin unidades no aporta valor al
    inventario, aunque el sistema conserve su ultimo costo. Una existencia vacia
    se trata como 0.
    """

    description = "Pone costo y total en 0 cuando la existencia es 0"
    col_existencia: str = C.BD_EXISTENCIA_NETA
    col_costo: str = C.BD_COSTO
    col_total: str = C.BD_TOTAL_INV

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        salida = df.copy()
        existencia = pd.to_numeric(salida[self.col_existencia], errors="coerce").fillna(0)
        salida[self.col_existencia] = existencia
        costo = pd.to_numeric(salida[self.col_costo], errors="coerce").fillna(0)
        total = pd.to_numeric(salida[self.col_total], errors="coerce").fillna(0)

        sin_existencia = existencia.eq(0)
        cambia = sin_existencia & (costo.ne(0) | total.ne(0))
        costo = costo.mask(sin_existencia, 0.0)
        total = total.mask(sin_existencia, 0.0)
        salida[self.col_costo] = costo
        salida[self.col_total] = total

        negativas = int(existencia.lt(0).sum())
        return RuleResult(
            df=salida,
            modified=int(cambia.sum()),
            warnings=([f"{negativas} referencias con existencia negativa"] if negativas else []),
            metrics={
                "sin_existencia": int(sin_existencia.sum()),
                "costo_puesto_en_cero": int(cambia.sum()),
                "existencia_negativa": negativas,
            },
        )
