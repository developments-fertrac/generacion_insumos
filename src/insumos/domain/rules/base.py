"""Contrato base de una regla de negocio.

Reglas del contrato (seccion 4.1 del plan):

1. **Puras.** Sin filesystem, sin Selenium, sin Excel, sin red, sin reloj.
   Entrada y salida son datos.
2. **Vectorizadas.** Operan sobre un DataFrame de pandas, no sobre una hoja.
3. **Uniformes.** Todas devuelven ``RuleResult``, sin excepciones para la
   salida: una regla que falla lanza; una regla que elimina filas las reporta.
4. **Auditables.** Cada resultado dice que hizo, cuantas filas toco y por que.
   Eso alimenta el ``AuditTrail`` y la hoja "FUENTES DE DATOS".
5. **Componibles.** Una regla no sabe cual es la siguiente ni si es la ultima.

Por que ``RuleResult`` en vez de excepciones
-------------------------------------------
Un dia de inventario con 12.000 referencias tiene errores en todas partes. Si
la validacion de linea lanzara al primer fallo, el operador veria un error y
nada mas. Con ``RuleResult`` ve las 340 filas eliminadas, el motivo de cada una
y las 12 que quedaron con advertencia. El reporte de eliminaciones (hoja
"FUENTES DE DATOS") sale de ahi, no de un log.
"""

from __future__ import annotations

from collections.abc import Set
from dataclasses import dataclass, field
from datetime import date
from typing import Any, ClassVar

import pandas as pd


@dataclass(frozen=True)
class RuleContext:
    """Datos de referencia que las reglas necesitan, de solo lectura.

    Lo arma el caso de uso a partir de los puertos; las reglas nunca lo
    modifican. Cada tarea de inventario recibe el mismo contexto, asi que anadir
    una tabla de referencia no obliga a cambiar ninguna regla existente.
    """

    matriz_usd: pd.DataFrame = field(default_factory=pd.DataFrame)
    marcas_propias: Set[str] = frozenset()
    distribucion: dict[str, Any] = field(default_factory=dict)
    remisiones: pd.DataFrame = field(default_factory=pd.DataFrame)
    valorizados: dict[str, pd.DataFrame] = field(default_factory=dict)
    referencias_nuevas: Set[str] = frozenset()
    # Exportacion de base de datos ya filtrada (MOTIVO = INVENTARIO*). Es la
    # fuente de existencia, costo y total desde 2026-10 (ver ADR 0008).
    inventario_bd: pd.DataFrame = field(default_factory=pd.DataFrame)
    hoy: date | None = None
    params: dict[str, Any] = field(default_factory=dict)


@dataclass
class RuleResult:
    """Salida de una regla: el DataFrame nuevo y la evidencia de lo que hizo."""

    df: pd.DataFrame
    removed: pd.DataFrame = field(default_factory=pd.DataFrame)
    modified: int = 0
    warnings: list[str] = field(default_factory=list)
    metrics: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not isinstance(self.df, pd.DataFrame):
            raise TypeError(f"RuleResult.df debe ser DataFrame, hay {type(self.df).__name__}")

    @property
    def eliminadas(self) -> int:
        """Filas que la regla quito."""
        return len(self.removed)

    def con_advertencia(self, mensaje: str) -> RuleResult:
        """Devuelve self con una advertencia mas. Encadena sin copiar el df."""
        self.warnings.append(mensaje)
        return self


class ReglaBase:
    """El contrato, y la unica jerarquia, de toda regla registrada.

    Existe para que ``self.id`` y ``self.description`` esten declarados. Si el
    ``@rule(...)`` se limita a inyectar ``id`` sin que la clase lo declare, cada
    ``self.id`` dentro de una regla es un ``attr-defined`` que mypy marca, y la
    garantia de que toda regla tenga id se vuelve una convencion social en vez de
    algo que el comprobador de tipos garantiza.

    ``id`` y ``description`` son ``ClassVar``: los fija el decorador y la clase,
    no quien instancia la regla. Los parametros configurables (``lineas_validas``,
    ``almacenes``, ...) si son campos del dataclass, y por eso se anotan.

    Se sustituyo a un ``Protocol`` equivalente: dos nombres para lo mismo invita a
    que las reglas implementen uno y se anoten contra el otro. Ademas un Protocol
    con atributos mutables (``id: str``) no lo satisface una clase que los expone
    como ``ClassVar``, que es justo como deben ser.
    """

    id: ClassVar[str]
    description: ClassVar[str]

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        """Transforma ``df`` y devuelve el resultado auditable. No muta ``df``."""
        raise NotImplementedError


def marcar_eliminadas(
    filas: pd.DataFrame,
    motivo: str,
    id_regla: str,
) -> pd.DataFrame:
    """Anade ``_motivo`` y ``_regla`` a las filas que una regla elimina.

    Centralizado para que el reporte de eliminaciones tenga las mismas columnas
    sin importar que regla produjo la fila.
    """
    if filas.empty:
        return filas
    salida = filas.copy()
    salida["_motivo"] = motivo
    salida["_regla"] = id_regla
    return salida
