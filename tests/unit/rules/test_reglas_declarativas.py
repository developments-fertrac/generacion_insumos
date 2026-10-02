"""Pruebas del contrato de config/rules/*.yaml.

El caso central es ``test_pipeline_vacio_falla``: codifica el fallo silencioso
que se quiere evitar (inventario que pasa sin transformar y parece correcto).
"""

from __future__ import annotations

from pathlib import Path

import pytest

from insumos.domain.reglas_declarativas import (
    ConfiguracionInvalida,
    EntradaRegla,
    leer_declaracion,
)

_RAIZ = Path(__file__).resolve().parents[3]  # tests/unit/rules/ -> raiz del repo
_DIR_YAML = _RAIZ / "config" / "rules"
_YAMLS_REALES = sorted(_DIR_YAML.glob("*.yaml"))

pytestmark = pytest.mark.unit


def _escribir(tmp_path: Path, cuerpo: str) -> Path:
    ruta = tmp_path / "pipeline.yaml"
    ruta.write_text(cuerpo, encoding="utf-8")
    return ruta


# --- el caso que nos importa ---

def test_pipeline_vacio_falla(tmp_path: Path) -> None:
    """Un pipeline sin reglas debe ser un error visible, no un no-op."""
    ruta = _escribir(tmp_path, """
pipeline: inventario_general
version: 2026.10.0
rules: []
""")
    with pytest.raises(ConfiguracionInvalida, match="no tiene reglas activas"):
        leer_declaracion(ruta)


def test_pipeline_vacio_pasa_con_allow_empty(tmp_path: Path) -> None:
    ruta = _escribir(tmp_path, """
pipeline: inventario_general
version: 2026.10.0
allow_empty: true
rules: []
""")
    declaracion = leer_declaracion(ruta)
    assert declaracion.reglas_activas == ()
    assert declaracion.allow_empty is True


def test_todas_las_reglas_deshabilitadas_tambien_falla(tmp_path: Path) -> None:
    """No basta con que la lista tenga items: deben quedar activas."""
    ruta = _escribir(tmp_path, """
pipeline: inventario_general
version: 2026.10.0
rules:
  - id: inv.eliminar_linea_invalida
    enabled: false
""")
    with pytest.raises(ConfiguracionInvalida, match="no tiene reglas activas"):
        leer_declaracion(ruta)


# --- forma de las entradas ---

def test_acepta_forma_corta_y_larga(tmp_path: Path) -> None:
    ruta = _escribir(tmp_path, """
pipeline: inventario_general
version: "2026.10.1"
rules:
  - inv.eliminar_linea_invalida
  - id: inv.eliminar_criterios_negocio
    enabled: true
    params: { maximo_paso_1: 1000 }
""")
    declaracion = leer_declaracion(ruta)
    assert [r.id for r in declaracion.reglas_activas] == [
        "inv.eliminar_linea_invalida",
        "inv.eliminar_criterios_negocio",
    ]
    assert declaracion.reglas_activas[1].params == {"maximo_paso_1": 1000}


def test_respeta_el_orden_declarado(tmp_path: Path) -> None:
    """El orden del YAML es el orden de ejecucion: no se puede reordenar."""
    orden = [
        "inv.primera",
        "inv.segunda",
        "inv.tercera",
        "inv.cuarta",
    ]
    ruta = _escribir(tmp_path, "pipeline: inventario_general\nversion: 1\nrules:\n"
                     + "".join(f"  - {i}\n" for i in orden))
    declaracion = leer_declaracion(ruta)
    assert [r.id for r in declaracion.reglas_activas] == orden


@pytest.mark.parametrize(
    "cuerpo, expected",
    [
        ("pipeline: x\nrules: []\n", "version"),
        ("version: 1\nrules: []\n", "pipeline"),
        ("pipeline: x\nversion: 1\nrules: {}\n", "rules"),
        ("pipeline: x\nversion: 1\nrules: []\nallow_empty: si\n", "allow_empty"),
        ("pipeline: x\nversion: 1\nrules: [{enabled: true}]\n", "falta 'id'"),
        ("pipeline: x\nversion: 1\nrules: [{id: a.b, enabled: 'si'}]\n", "enabled"),
        ("pipeline: x\nversion: 1\nrules: [{id: a.b, params: 5}]\n", "params"),
        ("pipeline: x\nversion: 1\nrules: [42]\n", "se esperaba str o mapa"),
    ],
)
def test_rechaza_yaml_mal_formado(tmp_path: Path, cuerpo: str, expected: str) -> None:
    with pytest.raises(ConfiguracionInvalida, match=expected):
        leer_declaracion(_escribir(tmp_path, cuerpo))


def test_archivo_inexistente_falla_con_nombre(tmp_path: Path) -> None:
    with pytest.raises(ConfiguracionInvalida, match="No existe el archivo"):
        leer_declaracion(tmp_path / "no_existe.yaml")


def test_yaml_vacio_falla(tmp_path: Path) -> None:
    with pytest.raises(ConfiguracionInvalida, match="falta 'pipeline'"):
        leer_declaracion(_escribir(tmp_path, "\n"))


# --- los YAML reales del repo ---

def test_existe_la_carpeta_de_reglas() -> None:
    assert _DIR_YAML.is_dir(), "debe existir config/rules/"
    assert _YAMLS_REALES, "debe haber al menos un pipeline declarado"


@pytest.mark.parametrize("ruta", _YAMLS_REALES, ids=lambda p: p.name)
def test_yaml_real_declara_pipeline_y_version(ruta: Path) -> None:
    """Estructural: valida aunque este vacio a proposito en esta fase."""
    from insumos.domain.reglas_declarativas import _leer

    datos = _leer(ruta)
    assert isinstance(datos.get("pipeline"), str) and datos["pipeline"]
    assert datos.get("version") is not None
    assert isinstance(datos.get("rules"), list)


@pytest.mark.parametrize("ruta", _YAMLS_REALES, ids=lambda p: p.name)
def test_yaml_real_se_resuelve_contra_el_registro(ruta: Path) -> None:
    """Cada YAML de produccion declara reglas que existen (Fase 3 inventario).

    ventas.yaml sigue vacio a proposito hasta la Fase 4; se acepta solo si lo
    declara con ``allow_empty`` o sin reglas.
    """
    import insumos.domain.rules  # noqa: F401
    from insumos.domain.reglas_declarativas import _leer
    from insumos.domain.rules import registry

    datos = _leer(ruta)
    if not datos["rules"]:
        assert ruta.name == "ventas.yaml", f"{ruta.name} no deberia estar vacio"
        return
    declaracion = leer_declaracion(ruta)
    for entrada in declaracion.reglas_activas:
        registry.construir(entrada.id, dict(entrada.params))


def test_ids_de_regla_siguen_la_convencion() -> None:
    """<dominio>.<verbo>_<objeto>. Se verifica al registrar cada regla."""
    for entrada in (
        EntradaRegla(id="inv.eliminar_linea_invalida"),
        EntradaRegla(id="ven.integrar_costo_factor_hoy"),
    ):
        dominio, _, verbo = entrada.id.partition(".")
        assert dominio in {"inv", "ven"}, f"dominio desconocido: {entrada.id}"
        assert "_" in verbo, f"id sin verbo_objeto: {entrada.id}"
