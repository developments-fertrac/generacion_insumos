"""Puertos: lo que el dominio necesita del mundo exterior, sin decir como.

Los adaptadores (``insumos.adapters``) los implementan; los casos de uso
(``insumos.application``) los reciben ya construidos. Cambiar la fuente del
inventario (archivo exportado hoy, consulta SQL directa manana) es escribir
otro adaptador, no tocar reglas.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Protocol

import pandas as pd

from insumos.domain.audit import AuditTrail


@dataclass(frozen=True)
class PlantillaInventario:
    """La hoja INVENTARIO de la plantilla, con columnas canonicas."""

    datos: pd.DataFrame
    origen: str
    encabezado_existencia: str = ""


class FuenteInventario(Protocol):
    """Entradas del inventario general."""

    def inventario_bd(self) -> pd.DataFrame:
        """Exportacion de base de datos, sin filtrar."""
        ...

    def plantilla(self) -> PlantillaInventario:
        """Plantilla del dia anterior (o el maestro si no hay salida previa)."""
        ...

    def matriz_usd(self) -> pd.DataFrame:
        """``CLAVE`` y ``DESCRIPCION``; vacio si el archivo no esta disponible."""
        ...

    def distribucion(self) -> dict[str, Any]:
        """``{"gestor": {...}, "clasificacion": {...}}``."""
        ...

    def fuentes_usadas(self) -> list[dict[str, str]]:
        """Archivos leidos (DATO, ARCHIVO, FECHA MODIFICACION, ADVERTENCIA)."""
        ...


class EscritorInventario(Protocol):
    """Proyecta el inventario calculado sobre el archivo de salida."""

    def escribir(self, datos: pd.DataFrame, *, hoy: date) -> Path:
        ...


@dataclass(frozen=True)
class EvidenciaEjecucion:
    """Lo que el reporte de auditoria necesita saber de una corrida."""

    audit_bd: AuditTrail
    audit_inventario: AuditTrail
    datos: pd.DataFrame
    archivo_salida: Path | None
    hoy: date
    fuentes: tuple[dict[str, str], ...] = ()


class EscritorReporte(Protocol):
    """Persiste la evidencia: pasos, eliminaciones y referencias nuevas."""

    def escribir(self, evidencia: EvidenciaEjecucion) -> Path:
        ...
