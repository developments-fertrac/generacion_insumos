"""Reglas de inventario con fuente base de datos (ADR 0008).

Cada prueba fija una regla de negocio acordada:

- solo entra MOTIVO que inicia con INVENTARIO;
- existencia = EXISTENCIA NETA, costo promedio = COSTO, total = TOTAL INV;
- existencia 0 -> costo 0;
- la plantilla pierde lo que la base de datos no trae y gana lo nuevo.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

import insumos.domain.rules  # noqa: F401
from insumos.domain.audit import AuditTrail
from insumos.domain.pipeline import ErrorDePipeline, RulePipeline
from insumos.domain.rules import registry
from insumos.domain.rules.base import RuleContext
from insumos.domain.rules.inventario import columnas as C
from insumos.domain.rules.inventario.actualizacion import CuadreFallido
from insumos.domain.rules.inventario.base_datos import ColumnasFaltantes

_RAIZ = Path(__file__).resolve().parents[3]
pytestmark = pytest.mark.unit


def _bd(filas: list[tuple]) -> pd.DataFrame:
    cols = ["REFERENCIA", "NOMBRE ODOO", "MARCA", "LINEA", "SUB-LINEA",
            "EXISTENCIA NETA", "COSTO", "TOTAL INV", "MOTIVO"]
    return pd.DataFrame(filas, columns=cols)


BD = _bd([
    ("A1", "Disco A", "EATON", "FULLER", "EATON", 10, 100.0, 1000.0, "INVENTARIO"),
    ("B2", "Kit B", "CUMMINS", "CUMMINS", "CUMMINS", 0, 0.0, 0.0, "INVENTARIO: tiene existencia pero no tiene costo cargado"),
    ("C3", "Rodamiento", "TIMKEN", "TIMKEN", "TIMKEN", 5, 50.0, 250.0, "EXCLUIDO: referencia - Arme/Desarme"),
    ("D4", "Nuevo", "MERITOR", "MERITOR", "MERITOR", 2, 30.0, 60.0, "INVENTARIO: neta después de descontar faltantes"),
    ("E5", "Varios", "VARIOS", "VARIOS", "VARIOS", 1, 1.0, 1.0, "EXCLUIDO: marca Invalida - Varios"),
])


def _correr(ids: list[str] | str, df: pd.DataFrame, ctx: RuleContext | None = None) -> tuple[pd.DataFrame, AuditTrail]:
    audit = AuditTrail()
    if isinstance(ids, str):
        pipeline = RulePipeline.from_yaml(_RAIZ / "config" / "rules" / ids, audit)
    else:
        pipeline = RulePipeline.desde_ids(ids, audit)
    return pipeline.run(df, ctx or RuleContext()).df, audit


# --- columnas -------------------------------------------------------------

@pytest.mark.parametrize("entrada, esperado", [
    ("Sub- linea sistema ", "SUB-LINEA SISTEMA"),
    ("MARCA copia", "MARCA COPIA"),
    ("NOMBRE MYR ", "NOMBRE MYR"),
    ("CLASIFICACIÓN", "CLASIFICACION"),
])
def test_normalizar_encabezado(entrada: str, esperado: str) -> None:
    assert C.normalizar_encabezado(entrada) == esperado


@pytest.mark.parametrize("entrada, esperado", [
    (1060.0, "1060"), (1060, "1060"), (" abc-1 ", "ABC-1"), ("1060.0", "1060"),
    (None, ""), (float("nan"), ""), ("K3220/3002466", "K3220/3002466"),
])
def test_clave_referencia(entrada: object, esperado: str) -> None:
    assert C.clave_referencia(entrada) == esperado


def test_columna_existencia_con_fecha() -> None:
    assert C.es_columna_existencia("EXISTENCIA SEP 10")
    assert C.es_columna_existencia("EXISTENCIA")
    assert not C.es_columna_existencia("EXISTENCIAS")


# --- pipeline inventario_bd ----------------------------------------------

def test_validar_columnas_bd_falla_si_falta_una() -> None:
    with pytest.raises(ErrorDePipeline) as e:
        _correr(["inv.validar_columnas_bd"], BD.drop(columns="COSTO"))
    assert isinstance(e.value.__cause__, ColumnasFaltantes)


def test_filtro_por_motivo_conserva_solo_inventario() -> None:
    df, audit = _correr(["inv.filtrar_motivo_inventario"], BD)
    assert list(df["REFERENCIA"]) == ["A1", "B2", "D4"]
    eliminadas = audit.tabla_eliminaciones()
    assert set(eliminadas["REFERENCIA"]) == {"C3", "E5"}
    assert eliminadas.loc[eliminadas["REFERENCIA"] == "C3", "_detalle"].iloc[0].startswith("EXCLUIDO")


def test_filtro_por_motivo_es_insensible_a_mayusculas_y_espacios() -> None:
    bd = BD.assign(MOTIVO=["  inventario", "Inventario: x", "excluido", None, "INVENTARIOS"])
    df, _ = _correr(["inv.filtrar_motivo_inventario"], bd)
    assert list(df["REFERENCIA"]) == ["A1", "B2", "E5"]


def test_costo_cero_cuando_existencia_cero() -> None:
    bd = _bd([("X", "n", "m", "l", "s", 0, 999.0, 0.0, "INVENTARIO"),
              ("Y", "n", "m", "l", "s", 3, 10.0, 30.0, "INVENTARIO")])
    regla = registry.construir("inv.costo_cero_sin_existencia")
    r = regla.apply(bd, RuleContext())
    assert list(r.df["COSTO"]) == [0.0, 10.0]
    assert list(r.df["TOTAL INV"]) == [0.0, 30.0]
    assert r.modified == 1


def test_duplicados_conserva_la_primera_y_quita_vacias() -> None:
    df = pd.DataFrame({"REFERENCIA": ["A", "a ", None, "B", ""], "V": [1, 2, 3, 4, 5]})
    regla = registry.construir("inv.eliminar_referencias_duplicadas")
    r = regla.apply(df, RuleContext())
    assert list(r.df["V"]) == [1, 4]
    assert r.metrics == {"vacias": 2, "duplicadas": 1}


def test_yaml_bd_completo() -> None:
    df, audit = _correr("inventario_bd.yaml", BD)
    assert list(df["REFERENCIA"]) == ["A1", "B2", "D4"]
    assert audit.pipeline == "inventario_bd"


# --- pipeline inventario_general -------------------------------------------

def _plantilla() -> pd.DataFrame:
    return pd.DataFrame({
        C.REFERENCIA: ["A1", "B2", "C3", "ZZ", "A1"],
        C.NOMBRE_LISTA: ["Lista A", None, "x", "x", "dup"],
        C.NOMBRE_ODOO: ["viejo", "Kit B", "x", "x", "dup"],
        C.NOMBRE_MYR: ["Lista A", "Kit B", "x", "x", "dup"],
        C.MARCA_COPIA: ["EATON", "CUMMINS", "x", "x", "dup"],
        C.INV_BODEGA_GERENCIA: [0, 0, 0, 0, 0],
        C.EXISTENCIA: [7, 4, 1, 1, 1],
        C.COSTO_PROMEDIO: [90.0, 20.0, 1.0, 1.0, 1.0],
        C.TOTAL_INV: [630.0, 80.0, 1.0, 1.0, 1.0],
        C.LINEA_COPIA: ["FULLER", "CUMMINS", "x", "x", "dup"],
        C.SUBLINEA_COPIA: ["EATON", "CUMMINS", "x", "x", "dup"],
        C.LIDER_LINEA: ["JORDAN", "KAREN", "x", "x", "dup"],
        C.CLASIFICACION: ["TRANSMISION", "MOTOR", "x", "x", "dup"],
        C.MARCA_SISTEMA: ["EATON", "CUMMINS", "x", "x", "dup"],
        C.LINEA_SISTEMA: ["FULLER", "CUMMINS", "x", "x", "dup"],
        C.SUBLINEA_SISTEMA: ["EATON", "CUMMINS", "x", "x", "dup"],
    })


def _contexto(bd_filtrada: pd.DataFrame) -> RuleContext:
    return RuleContext(
        inventario_bd=bd_filtrada,
        matriz_usd=pd.DataFrame({"CLAVE": ["A1", "D4"], "DESCRIPCION": ["Lista A nueva", "Lista D"]}),
        distribucion={"gestor": {"MERITOR": "PEDRO"}, "clasificacion": {"MERITOR": "EJES"}},
        params={"motivos_excluidos_bd": {"C3": "EXCLUIDO: referencia - Arme/Desarme"}},
    )


@pytest.fixture
def resultado() -> tuple[pd.DataFrame, AuditTrail]:
    bd, _ = _correr("inventario_bd.yaml", BD)
    return _correr("inventario.yaml", _plantilla(), _contexto(bd))


def test_salida_tiene_exactamente_las_referencias_de_la_bd(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    assert sorted(df[C.REFERENCIA]) == ["A1", "B2", "D4"]


def test_existencia_costo_y_total_vienen_de_la_bd(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    a1 = df.set_index(C.REFERENCIA).loc["A1"]
    assert a1[C.EXISTENCIA] == 10
    assert a1[C.COSTO_PROMEDIO] == 100.0
    assert a1[C.TOTAL_INV] == 1000.0
    assert a1[C.NOMBRE_ODOO] == "Disco A"


def test_existencia_cero_deja_costo_cero(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    b2 = df.set_index(C.REFERENCIA).loc["B2"]
    assert b2[C.EXISTENCIA] == 0 and b2[C.COSTO_PROMEDIO] == 0 and b2[C.TOTAL_INV] == 0


def test_eliminadas_llevan_motivo_de_la_bd(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    _, audit = resultado
    elim = audit.tabla_eliminaciones().set_index(C.REFERENCIA)
    assert elim.loc["C3", "_detalle"] == "Base de datos: EXCLUIDO: referencia - Arme/Desarme"
    assert elim.loc["ZZ", "_detalle"] == "No existe en la base de datos"
    assert elim.loc["A1", "_motivo"] == "Referencia duplicada"


def test_referencia_nueva_se_agrega_y_enriquece(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    d4 = df.set_index(C.REFERENCIA).loc["D4"]
    assert bool(d4["_NUEVA"]) is True
    assert d4[C.NOMBRE_LISTA] == "Lista D"
    assert d4[C.NOMBRE_MYR] == "Lista D"
    assert d4[C.MARCA_COPIA] == "MERITOR" and d4[C.LINEA_COPIA] == "MERITOR"
    assert d4[C.LIDER_LINEA] == "PEDRO" and d4[C.CLASIFICACION] == "EJES"
    assert d4[C.INV_BODEGA_GERENCIA] == 0


def test_nombre_lista_cero_si_no_esta_en_matriz_y_myr_cae_a_odoo(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    b2 = df.set_index(C.REFERENCIA).loc["B2"]
    assert b2[C.NOMBRE_LISTA] == "0"
    assert b2[C.NOMBRE_MYR] == "Kit B"


def test_curaduria_manual_no_se_pisa(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    a1 = df.set_index(C.REFERENCIA).loc["A1"]
    assert a1[C.LIDER_LINEA] == "JORDAN" and a1[C.CLASIFICACION] == "TRANSMISION"


def test_orden_por_total_desc(resultado: tuple[pd.DataFrame, AuditTrail]) -> None:
    df, _ = resultado
    assert list(df[C.REFERENCIA]) == ["A1", "D4", "B2"]


def test_validar_salida_detecta_descuadre() -> None:
    bd, _ = _correr("inventario_bd.yaml", BD)
    df, _ = _correr("inventario.yaml", _plantilla(), _contexto(bd))
    roto = df.copy()
    roto.loc[0, C.TOTAL_INV] = 1.0
    regla = registry.construir("inv.validar_salida")
    with pytest.raises(CuadreFallido, match="no cuadra"):
        regla.apply(roto, _contexto(bd))


def test_sin_bd_en_contexto_falla_explicito() -> None:
    regla = registry.construir("inv.actualizar_desde_bd")
    with pytest.raises(ValueError, match="inventario_bd"):
        regla.apply(_plantilla(), RuleContext())
