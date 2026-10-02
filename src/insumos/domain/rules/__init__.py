"""Reglas de negocio: unidades puras y componibles.

Contrato unico:

    Rule.apply(df: pd.DataFrame, ctx: RuleContext) -> RuleResult

Reglas obligatorias:
1. Puras: sin filesystem, sin Selenium, sin Excel, sin red.
2. Vectorizadas sobre pandas.
3. Con ``id`` estable (``<dominio>.<verbo>_<objeto>``) y ``description``
   escrita para el negocio.
4. Registradas con ``@rule("...")`` y declaradas en ``config/rules/*.yaml``.

Importar este paquete registra las reglas de inventario.
"""

from __future__ import annotations

# Importa los modulos de reglas para que @rule(...) se ejecute al importar.
from insumos.domain.rules import inventario  # noqa: F401
