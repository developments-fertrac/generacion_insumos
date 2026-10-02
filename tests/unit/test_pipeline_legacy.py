"""Pruebas de humo del pipeline legacy.

Verifican que el contrato externo (CLI -> Orchestrator -> BaseTask) sigue
funcionando mientras el interior se migra tarea por tarea. Un fallo aqui
significa que la migracion rompio produccion.
"""
from __future__ import annotations

import pytest

from orchestrator import PIPELINES, Orchestrator
from tasks import TASK_REGISTRY
from tasks.base_task import BaseTask

pytestmark = pytest.mark.unit


# Clave del registro -> atributo `name` de la clase. No siempre coinciden:
# `actualizacion_inv` es la clave de CLI, `actualizacion_inventario` es el
# nombre de la clase (y del archivo de log). Este mapeo es parte del contrato
# externo y no debe cambiar durante la migracion.
NOMBRES_DE_LOG = {
    # Las tres descargas por Selenium se retiraron en 2026-10 (ADR 0008).
    "actualizacion_inv": "actualizacion_inventario",
    "actualizacion_ventas": "actualizacion_ventas",
    "envio_informe_ventas": "envio_informe_ventas",
}


def test_registro_de_tareas_conserva_las_tareas_vigentes() -> None:
    assert set(TASK_REGISTRY) == set(NOMBRES_DE_LOG)


@pytest.mark.parametrize("nombre", sorted(TASK_REGISTRY))
def test_todas_las_tareas_heredan_de_basetask(nombre: str) -> None:
    clase = TASK_REGISTRY[nombre]
    assert issubclass(clase, BaseTask)
    assert clase.name == NOMBRES_DE_LOG[nombre]
    assert isinstance(clase.name, str) and clase.name


@pytest.mark.parametrize("workflow", sorted(PIPELINES))
def test_workflows_conservan_su_composicion(workflow: str) -> None:
    fases = PIPELINES[workflow]
    assert fases, "un workflow no puede tener fases vacias"
    for fase in fases:
        assert fase, "una fase no puede tener tareas vacias"
        for tarea in fase:
            assert tarea in TASK_REGISTRY, f"{workflow} referencia una tarea inexistente: {tarea}"


def test_orchestrator_rechaza_tarea_desconocida() -> None:
    orq = Orchestrator(settings=None)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="Task desconocida"):
        orq.run_task("no_existe")


def test_workflows_completos_cubren_las_tareas() -> None:
    tasks_pipeline = {t for fases in PIPELINES.values() for f in fases for t in f}
    assert tasks_pipeline == set(TASK_REGISTRY)
