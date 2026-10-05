from __future__ import annotations

import time
from dataclasses import dataclass, field

from config.settings import Settings
from core.logger import get_logger
from tasks import TASK_REGISTRY
from tasks.base_task import BaseTask


log = get_logger("orchestrator")


@dataclass
class PhaseResult:
    task_name: str
    success: bool
    elapsed_seconds: int


@dataclass
class PipelineResult:
    phases: list[list[PhaseResult]] = field(default_factory=list)
    total_elapsed: int = 0

    @property
    def all_success(self) -> bool:
        return all(r.success for phase in self.phases for r in phase)


# Desde 2026-10 los insumos llegan exportados por la base de datos: ya no hay
# fase de descarga por Selenium.
PIPELINE_COMPLETO = [
    ["actualizacion_inv", "actualizacion_ventas"],
]

PIPELINE_VENTAS = [
    ["actualizacion_ventas"],
    ["envio_informe_ventas"],
]

PIPELINE_PARCIAL_VENTAS = [
    ["actualizacion_ventas"],
]

PIPELINE_INVENTARIO = [
    ["actualizacion_inv"],
]

# Solo el envio del informe (recorte de imagenes + WhatsApp), sobre el
# $2026 VENTAS_Actualizacion ya generado.
PIPELINE_ENVIO = [
    ["envio_informe_ventas"],
]

PIPELINES = {
    "completo": PIPELINE_COMPLETO,
    "ventas": PIPELINE_VENTAS,
    "inventario": PIPELINE_INVENTARIO,
    "parcialVentas": PIPELINE_PARCIAL_VENTAS,
    "envio": PIPELINE_ENVIO,
}


class Orchestrator:
    def __init__(self, settings: Settings, pipeline: list[list[str]] | None = None):
        self.settings = settings
        self.pipeline = pipeline or PIPELINE_COMPLETO

    def run_all(self) -> PipelineResult:
        log.info("=" * 70)
        log.info("PIPELINE - Generacion de Insumos")
        log.info("=" * 70)
        result = PipelineResult()
        start = time.time()

        for phase_idx, phase_tasks in enumerate(self.pipeline, 1):
            log.info("")
            log.info("FASE %d/%d: %s", phase_idx, len(self.pipeline), " → ".join(phase_tasks))
            log.info("-" * 70)

            phase_results = self._run_phase(phase_idx, phase_tasks)
            result.phases.append(phase_results)

            failed = [r for r in phase_results if not r.success]
            if failed:
                log.error("Fase %d tuvo fallos: %s", phase_idx, [r.task_name for r in failed])
                if phase_idx < len(self.pipeline):
                    log.warning("Continuando con fase siguiente a pesar de fallos...")

        result.total_elapsed = int(time.time() - start)
        log.info("")
        log.info("=" * 70)
        log.info("RESUMEN PIPELINE (%ds total)", result.total_elapsed)
        log.info("=" * 70)
        for phase_idx, phase_results in enumerate(result.phases, 1):
            for r in phase_results:
                status = "OK" if r.success else "FALLO"
                log.info("  Fase %d | %-30s | %s (%ds)", phase_idx, r.task_name, status, r.elapsed_seconds)
        log.info("=" * 70)
        return result

    def run_task(self, task_name: str) -> PhaseResult:
        task_class = TASK_REGISTRY.get(task_name)
        if not task_class:
            raise ValueError(f"Task desconocida: '{task_name}'. Disponibles: {list(TASK_REGISTRY.keys())}")
        return self._execute_task(task_class)

    def _run_phase(self, phase_idx: int, task_names: list[str]) -> list[PhaseResult]:
        results = []
        for task_name in task_names:
            task_class = TASK_REGISTRY.get(task_name)
            if not task_class:
                log.error("Task desconocida: '%s' - saltando", task_name)
                results.append(PhaseResult(task_name=task_name, success=False, elapsed_seconds=0))
                continue
            results.append(self._execute_task(task_class))
        return results

    def _execute_task(self, task_class: type[BaseTask]) -> PhaseResult:
        start = time.time()
        task = task_class(self.settings)
        success = task.run()
        elapsed = int(time.time() - start)
        return PhaseResult(task_name=task.name, success=success, elapsed_seconds=elapsed)
