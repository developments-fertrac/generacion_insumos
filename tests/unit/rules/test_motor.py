"""Pruebas del motor de reglas: registro, pipeline y auditoria.

Criterio de salida de la Fase 1 del plan: motor al 100%.
"""
from __future__ import annotations

from dataclasses import FrozenInstanceError, dataclass
from datetime import date
from pathlib import Path
from typing import cast

import pandas as pd
import pytest

from insumos.domain.audit import AuditTrail, PasoAuditoria
from insumos.domain.pipeline import ErrorDePipeline, ResultadoPipeline, RulePipeline
from insumos.domain.rules import registry
from insumos.domain.rules.base import (
    ReglaBase,
    RuleContext,
    RuleResult,
    marcar_eliminadas,
)
from insumos.domain.rules.inventario.base_datos import FiltrarMotivoInventario

_RAIZ = Path(__file__).resolve().parents[3]


@dataclass(frozen=True)
class _ReglaDePrueba(ReglaBase):
    id = 'test.regla'
    description = "regla de prueba"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        return RuleResult(df=df, metrics={"ok": True})


@dataclass(frozen=True)
class _ReglaQueElimina(ReglaBase):
    id = 'test.elimina'
    description = "elimina la primera fila"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if df.empty:
            return RuleResult(df=df)
        quitar = df.iloc[:1]
        return RuleResult(df=df.iloc[1:], removed=marcar_eliminadas(quitar, "prueba", "test"))


@dataclass(frozen=True)
class _ReglaQueExplota(ReglaBase):
    id = 'test.explota'
    description = "lanza excepcion"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        raise ValueError("boom")


@dataclass(frozen=True)
class _ReglaQueDevuelveNada(ReglaBase):
    id = 'test.mal'
    description = "no devuelve RuleResult"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        return "no soy un RuleResult"  # type: ignore[return-value]


# --- registro ---

def test_registro_inyecta_el_id() -> None:
    @registry.rule("test.inyecta_id")
    @dataclass(frozen=True)
    class InyectaId(ReglaBase):
        description = "d"

        def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
            return RuleResult(df=df)

    assert InyectaId.id == "test.inyecta_id"
    assert registry.obtener("test.inyecta_id") is InyectaId


def test_registro_rechaza_sin_descripcion() -> None:
    with pytest.raises(registry.ReglaInvalida, match="description"):

        @registry.rule("test.sin_descripcion")
        @dataclass(frozen=True)
        class SinDescripcion(ReglaBase):
            def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
                return RuleResult(df=df)


def test_registro_rechaza_sin_apply() -> None:
    """Una clase que no hereda ``ReglaBase`` y no define ``apply`` se rechaza.

    mypy ya lo prohibe por el ``bound`` del TypeVar; esta prueba cubre el caso
    runtime, que es lo que pasa si alguien registra una regla desde un modulo
    sin tipos o con ``# type: ignore``.
    """

    @dataclass(frozen=True)
    class SinApply:
        description = "d"

    with pytest.raises(registry.ReglaInvalida, match="apply"):
        registry.rule("test.sin_apply")(SinApply)  # type: ignore[type-var,arg-type]


def test_registro_rechaza_id_duplicado() -> None:
    @registry.rule("test.duplicado")
    @dataclass(frozen=True)
    class Uno(ReglaBase):
        description = "d"

        def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
            return RuleResult(df=df)

    with pytest.raises(registry.ReglaInvalida, match="duplicado"):

        @registry.rule("test.duplicado")
        @dataclass(frozen=True)
        class Otro(ReglaBase):
            description = "d2"

            def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
                return RuleResult(df=df)


def test_obtener_id_desconocido_lista_los_disponibles() -> None:
    with pytest.raises(registry.ReglaDesconocida, match="Registradas"):
        registry.obtener("test.no_existe_nunca")


def test_construir_pasa_params() -> None:
    """El YAML entrega listas; la regla las recibe ya normalizadas."""
    regla = cast(
        FiltrarMotivoInventario,
        registry.construir("inv.filtrar_motivo_inventario", {"prefijo": "INVENTARIO", "columna": "MOTIVO"}),
    )
    assert regla.prefijo == "INVENTARIO"
    assert regla.columna == "MOTIVO"


def test_construir_rechaza_param_inesperado() -> None:
    """Un typo en el YAML no puede desactivar medio pipeline en silencio."""
    with pytest.raises(registry.ReglaInvalida, match="Parametros invalidos"):
        registry.construir("inv.filtrar_motivo_inventario", {"prefijoo": "X"})


def test_catalogo_tiene_las_reglas_de_inventario() -> None:
    import insumos.domain.rules  # noqa: F401

    ids = {f["id"] for f in registry.catalogo()}
    for esperado in (
        "inv.filtrar_motivo_inventario",
        "inv.costo_cero_sin_existencia",
        "inv.actualizar_desde_bd",
        "inv.validar_salida",
    ):
        assert esperado in ids


def test_ninguna_regla_de_inventario_importa_infraestructura() -> None:
    """Guardarraíl de Phase 1: las reglas son puras."""
    import ast

    from insumos.domain.rules import inventario as pkg

    for archivo in pkg.__path__:
        for f in Path(archivo).rglob("*.py"):
            arbol = ast.parse(f.read_text(encoding="utf-8"))
            for nodo in ast.walk(arbol):
                nombres = []
                if isinstance(nodo, ast.Import):
                    nombres = [a.name for a in nodo.names]
                elif isinstance(nodo, ast.ImportFrom) and nodo.module:
                    nombres = [nodo.module]
                for nombre in nombres:
                    raiz = nombre.split(".")[0]
                    assert raiz not in {
                        "win32com", "selenium", "openpyxl", "msoffcrypto", "os", "io", "glob",
                    }, f"{f.name} importa {nombre}"


# --- pipeline ---

def test_pipeline_devuelve_df_igual_si_no_hay_reglas() -> None:
    df = pd.DataFrame({"A": [1, 2, 3]})
    resultado = RulePipeline([], AuditTrail()).run(df, RuleContext())
    assert resultado.df.equals(df)
    assert resultado.audit.pasos == []


def test_pipeline_encadena_en_orden() -> None:
    df = pd.DataFrame({"A": range(5)})
    pasos: list[str] = []

    @dataclass(frozen=True)
    class Primera(ReglaBase):
        id = "test.primera"
        description = "primera"

        def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
            pasos.append("primera")
            return RuleResult(df=df)

    @dataclass(frozen=True)
    class Segunda(ReglaBase):
        id = "test.segunda"
        description = "segunda"

        def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
            pasos.append("segunda")
            return RuleResult(df=df)

    RulePipeline([Primera(), Segunda()], AuditTrail()).run(df, RuleContext())
    assert pasos == ["primera", "segunda"]


def test_pipeline_propaga_el_df_de_una_regla_a_la_siguiente() -> None:
    df = pd.DataFrame({"A": range(3)})
    resultado = RulePipeline([_ReglaQueElimina()], AuditTrail()).run(df, RuleContext())
    assert len(resultado.df) == 2
    assert resultado.audit.pasos[0].filas_antes == 3
    assert resultado.audit.pasos[0].filas_despues == 2
    assert resultado.audit.pasos[0].eliminadas == 1


def test_pipeline_las_reglas_no_mutan_el_df_original() -> None:
    df = pd.DataFrame({"A": range(3)})
    original = df.copy()
    RulePipeline([_ReglaQueElimina()], AuditTrail()).run(df, RuleContext())
    pd.testing.assert_frame_equal(df, original)


def test_pipeline_registra_el_audit() -> None:
    df = pd.DataFrame({"A": range(3)})
    resultado = RulePipeline([_ReglaQueElimina()], AuditTrail()).run(df, RuleContext())
    paso = resultado.audit.pasos[0]
    assert paso.id_regla == "test.elimina"
    assert paso.elimino_filas
    assert resultado.eliminadas.shape[0] == 1


def test_pipeline_mide_el_tiempo() -> None:
    df = pd.DataFrame({"A": range(3)})
    resultado = RulePipeline([_ReglaDePrueba()], AuditTrail()).run(df, RuleContext())
    assert "ms" in resultado.audit.pasos[0].metrics


def test_pipeline_una_excepcion_no_se_traga() -> None:
    """Un fallo silencioso produce un archivo correcto en apariencia. No."""
    df = pd.DataFrame({"A": range(3)})
    with pytest.raises(ErrorDePipeline, match="boom"):
        RulePipeline([_ReglaQueExplota()], AuditTrail()).run(df, RuleContext())


def test_pipeline_rechaza_salida_que_no_es_ruleresult() -> None:
    df = pd.DataFrame({"A": range(3)})
    with pytest.raises(ErrorDePipeline, match="se esperaba RuleResult"):
        RulePipeline([_ReglaQueDevuelveNada()], AuditTrail()).run(df, RuleContext())


def test_pipeline_soporta_df_vacio() -> None:
    resultado = RulePipeline([_ReglaQueElimina()], AuditTrail()).run(
        pd.DataFrame(), RuleContext()
    )
    assert resultado.df.empty


def test_ids_y_len() -> None:
    p = RulePipeline([_ReglaDePrueba()], AuditTrail())
    assert len(p) == 1
    assert p.ids == ["test.regla"]


# --- carga desde YAML ---

def test_from_yaml_falla_con_id_inexistente(tmp_path: Path) -> None:
    ruta = tmp_path / "p.yaml"
    ruta.write_text(
        "pipeline: x\nversion: 1\nrules:\n  - inv.regla_que_no_existe\n", encoding="utf-8"
    )
    with pytest.raises(registry.ReglaDesconocida, match="Registradas"):
        RulePipeline.from_yaml(ruta, AuditTrail())


def test_from_yaml_rechaza_pipeline_vacio(tmp_path: Path) -> None:
    """El fallo silencioso que se quiere evitar, ahora en el motor."""
    ruta = tmp_path / "p.yaml"
    ruta.write_text("pipeline: x\nversion: 1\nrules: []\n", encoding="utf-8")
    with pytest.raises(ValueError, match="no tiene reglas activas"):
        RulePipeline.from_yaml(ruta, AuditTrail())


def test_desde_ids_construye(monkeypatch: pytest.MonkeyPatch) -> None:
    import insumos.domain.rules  # noqa: F401

    p = RulePipeline.desde_ids(["inv.filtrar_motivo_inventario"], AuditTrail())
    assert p.ids == ["inv.filtrar_motivo_inventario"]


# --- audit trail ---

def test_resumen_del_audit() -> None:
    audit = AuditTrail(pipeline="inventario", version="1")
    audit.registrar(PasoAuditoria("a", "d", 10, 8, 0, 2))
    audit.registrar(PasoAuditoria("b", "d", 8, 8, 3, 0))
    resumen = audit.resumen()
    assert resumen["total_eliminadas"] == 2
    assert resumen["total_modificadas"] == 3
    assert resumen["filas_iniciales"] == 10
    assert resumen["filas_finales"] == 8
    assert resumen["por_regla"]["a"]["eliminadas"] == 2


def test_audit_acumula_eliminaciones_de_varias_reglas() -> None:
    audit = AuditTrail()
    audit.acumular_eliminaciones(pd.DataFrame({"A": [1]}), "r1")
    audit.acumular_eliminaciones(pd.DataFrame({"A": [2, 3]}), "r2")
    tabla = audit.tabla_eliminaciones()
    assert len(tabla) == 3
    assert set(tabla["_regla"]) == {"r1", "r2"}


def test_audit_tabla_pasos_ordenada() -> None:
    audit = AuditTrail()
    for i in range(3):
        audit.registrar(PasoAuditoria(f"r{i}", f"d{i}", 5, 5, 0, 0))
    tabla = audit.tabla_pasos()
    assert list(tabla["orden"]) == [1, 2, 3]
    assert list(tabla["regla"]) == ["r0", "r1", "r2"]


def test_audit_vacio_no_rompe() -> None:
    audit = AuditTrail()
    assert audit.resumen()["pasos"] == 0
    assert audit.tabla_eliminaciones().empty
    assert audit.tabla_pasos().empty


# --- base ---

def test_rulecontext_es_de_solo_lectura() -> None:
    ctx = RuleContext()
    with pytest.raises(FrozenInstanceError):
        ctx.hoy = date.today()  # type: ignore[misc]


def test_rulecontext_por_defecto_no_comparte_estado() -> None:
    a, b = RuleContext(), RuleContext()
    assert a.valorizados is not b.valorizados


def test_marcar_eliminadas_agrega_motivo_y_regla() -> None:
    df = pd.DataFrame({"A": [1, 2]})
    out = marcar_eliminadas(df, "porque si", "regla.x")
    assert out["_motivo"].iloc[0] == "porque si"
    assert out["_regla"].iloc[0] == "regla.x"
    assert "_motivo" not in df.columns


def test_marcar_eliminadas_sobre_df_vacio() -> None:
    out = marcar_eliminadas(pd.DataFrame(), "m", "r")
    assert out.empty


def test_ruleresult_rechaza_df_que_no_es_dataframe() -> None:
    with pytest.raises(TypeError, match="debe ser DataFrame"):
        RuleResult(df=[1, 2, 3])  # type: ignore[arg-type]


def test_con_advertencia_encadena() -> None:
    r = RuleResult(df=pd.DataFrame({"A": [1]}))
    assert r.con_advertencia("ojo").con_advertencia("otro") is r
    assert r.warnings == ["ojo", "otro"]


def test_resultado_pipeline_es_eliminadas() -> None:
    audit = AuditTrail()
    audit.acumular_eliminaciones(pd.DataFrame({"A": [1, 2]}), "r")
    resultado = ResultadoPipeline(df=pd.DataFrame(), audit=audit)
    assert len(resultado.eliminadas) == 2
