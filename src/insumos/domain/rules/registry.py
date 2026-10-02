"""Registro de reglas: la pieza que hace composable el pipeline.

Cada regla se registra con ``@rule("inv.eliminar_linea_invalida")`` al
definirse. El pipeline se arma por *nombre* desde el YAML, no importando
clases, asi que:

- el YAML puede reordenar o desactivar reglas sin tocar Python;
- agregar una regla es escribir su clase mas una linea en el YAML;
- un typo en el YAML falla al arrancar, con la lista de ids disponibles, en vez
  de ejecutarse a medias del inventario.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, TypeVar, cast

from insumos.domain.rules.base import ReglaBase

_TRegla = TypeVar("_TRegla", bound=type[ReglaBase])

# id -> clase. Ordered por escritura, pero el YAML manda el orden real.
_REGISTRO: dict[str, type[ReglaBase]] = {}


class ReglaDesconocida(KeyError):
    """El YAML pide una regla que no existe en el registro."""

    def __init__(self, id_regla: str) -> None:
        self.id_regla = id_regla
        disponibles = ", ".join(sorted(_REGISTRO)) or "(ninguna registrada)"
        super().__init__(
            f"Regla desconocida: '{id_regla}'. Registradas: {disponibles}"
        )


class ReglaInvalida(TypeError):
    """La clase registrada no cumple el contrato Rule."""


def rule(id_regla: str) -> Callable[[_TRegla], _TRegla]:
    """Decorador que registra una clase de regla bajo su id.

    Inyecta ``id`` como atributo de clase y verifica el contrato en el momento
    del registro, no en el primer inventario: un error de tipeo se ve al
    importar el modulo, no a las 4 de la tarde un jueves.

        @rule("inv.eliminar_linea_invalida")
        @dataclass(frozen=True)
        class EliminarLineaInvalida:
            description = "..."
            def apply(self, df, ctx): ...

    ``description`` debe ser un atributo de clase (sin anotacion), no un campo
    del dataclass: es texto para el reporte, no estado por instancia.
    """

    def deco(clase: _TRegla) -> _TRegla:
        destino = cast("type[ReglaBase]", clase)
        destino.id = id_regla

        if not isinstance(getattr(destino, "description", None), str) or not destino.description:
            raise ReglaInvalida(
                f"{clase.__name__}: falta 'description'. Una regla sin descripcion "
                "no se puede reportar al negocio."
            )
        if not callable(getattr(destino, "apply", None)):
            raise ReglaInvalida(f"{clase.__name__}: falta 'apply(df, ctx) -> RuleResult'")

        if clase.id in _REGISTRO and _REGISTRO[clase.id] is not clase:
            raise ReglaInvalida(
                f"Id de regla duplicado: '{clase.id}' ya lo usa "
                f"{_REGISTRO[clase.id].__name__}. Los ids no se reciclan: un id "
                "que cambia de significado deja el historial de auditoria inservible."
            )
        _REGISTRO[clase.id] = clase
        return clase

    return deco


def obtener(id_regla: str) -> type[ReglaBase]:
    """Devuelve la clase registrada, o falla con la lista de disponibles."""
    try:
        return _REGISTRO[id_regla]
    except KeyError:
        raise ReglaDesconocida(id_regla) from None


def _normalizar(valor: Any) -> Any:
    """Convierte lo que viene del YAML al tipo que la regla declara.

    El YAML no distingue lista de conjunto, y escribir ``[A1, B2]`` a mano es
    mas natural que ``frozenset: [A1, B2]``. Se normaliza aqui para que cada
    regla no tenga que defensively hacer ``frozenset(params.get(...))``.
    """
    if isinstance(valor, list):
        try:
            return frozenset(valor)
        except TypeError:  # contiene listas/dicts: no es un conjunto
            return tuple(valor)
    return valor


def construir(id_regla: str, params: dict[str, Any] | None = None) -> ReglaBase:
    """Instancia la regla con los parametros del YAML.

    La clase recibe solo lo que declare en ``__init__``. Un parametro que la
    regla no espera es un error de configuracion, no algo que se ignore en
    silencio: asi un typo en el YAML no desactiva medio pipeline.
    """
    clase = obtener(id_regla)
    try:
        return clase(**{k: _normalizar(v) for k, v in (params or {}).items()})
    except TypeError as e:
        raise ReglaInvalida(
            f"Parametros invalidos para '{id_regla}': {e}. "
            f"Firma de {clase.__name__}: {list(_firma(clase))}"
        ) from None


def _firma(clase: type[Any]) -> list[str]:
    import inspect

    try:
        return list(inspect.signature(clase.__init__).parameters)
    except (ValueError, TypeError):
        return []


def registradas() -> dict[str, type[ReglaBase]]:
    """Copia del registro, para el catalogo y los diagnosticos."""
    return dict(_REGISTRO)


def catalogo() -> list[dict[str, Any]]:
    """Catalogo de reglas para generar Markdown (ver seccion 10 del plan)."""
    filas = []
    for id_regla, clase in sorted(_REGISTRO.items()):
        filas.append(
            {
                "id": id_regla,
                "description": clase.description,
                "clase": clase.__name__,
                "modulo": clase.__module__,
                "parametros": [p for p in _firma(clase) if p != "self"],
            }
        )
    return filas


def _cargar_modulos_de_reglas() -> None:
    """Importa los modulos que contienen ``@rule(...)``.

    Se llama desde el bootstrap y desde los tests. Importar aqui evita que un
    modulo nuevo quede sin registrar por olvido.
    """
    import importlib
    import pkgutil

    from insumos.domain import rules as _rules_pkg

    for info in pkgutil.walk_packages(_rules_pkg.__path__, f"{_rules_pkg.__name__}."):
        if ".base" in info.name or info.name.endswith(".base"):
            continue
        try:
            importlib.import_module(info.name)
        except ImportError:  # pragma: no cover - depende del entorno
            # Un modulo que no importa no debe tumbar el pipeline; se registra
            # el problema al montar el bootstrap.
            continue
