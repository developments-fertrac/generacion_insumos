from __future__ import annotations

import argparse
import sys
from pathlib import Path

_root = Path(__file__).resolve().parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from config.settings import Settings
from core.logger import get_logger
from orchestrator import Orchestrator, PIPELINES
from tasks import TASK_REGISTRY

AVAILABLE_TASKS = list(TASK_REGISTRY.keys())


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generacion de Insumos - Pipeline de automatizacion",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python run.py                          # Ejecutar pipeline completo
  python run.py --workflow ventas        # Solo ventas: descarga → actualizacion → envio informe
  python run.py --workflow inventario    # Solo inventario: descarga inv → descarga valorizados → actualizacion
  python run.py --task descarga_ventas   # Ejecutar solo una task
  python run.py --task envio_informe_ventas  # Recortar imagenes y enviarlas por WhatsApp Web
  python run.py --phase 1                # Ejecutar solo fase 1
  python run.py --phase 2                # Ejecutar solo fase 2
  python run.py --list                   # Listar workflows y tasks disponibles
        """,
    )
    parser.add_argument("--workflow", "-w", choices=list(PIPELINES.keys()),
                        help="Ejecutar un workflow completo (ventas, inventario, completo)")
    parser.add_argument("--task", "-t", type=str, help="Ejecutar solo una task especifica")
    parser.add_argument("--phase", "-p", type=int, choices=[1, 2], help="Ejecutar solo una fase especifica")
    parser.add_argument("--list", "-l", action="store_true", help="Listar workflows y tasks disponibles")

    args = parser.parse_args()

    log = get_logger("run")
    settings = Settings()

    if args.list:
        print("Workflows disponibles:")
        print("  completo    descarga_inv_general → descarga_ventas → descarga_valorizados → actualizacion_inv + actualizacion_ventas")
        print("  ventas      descarga_ventas → actualizacion_ventas → envio_informe_ventas")
        print("  inventario  descarga_inv_general → descarga_valorizados → actualizacion_inv")
        print()
        print("Tasks individuales:")
        for name in AVAILABLE_TASKS:
            print(f"  - {name}")
        return 0

    if args.workflow:
        pipeline = PIPELINES[args.workflow]
        log.info("Workflow '%s': %s", args.workflow, " → ".join([" → ".join(phase) for phase in pipeline]))
        orchestrator = Orchestrator(settings, pipeline=pipeline)
        result = orchestrator.run_all()
        return 0 if result.all_success else 1

    if args.task:
        if args.task not in AVAILABLE_TASKS:
            log.error("Task '%s' no encontrada. Disponibles: %s", args.task, AVAILABLE_TASKS)
            return 1
        orchestrator = Orchestrator(settings)
        result = orchestrator.run_task(args.task)
        return 0 if result.success else 1

    if args.phase:
        orchestrator = Orchestrator(settings)
        phase_idx = args.phase - 1
        if phase_idx >= len(orchestrator.pipeline):
            log.error("Fase %d no existe en el pipeline actual", args.phase)
            return 1
        phase_tasks = orchestrator.pipeline[phase_idx]
        log.info("Ejecutando Fase %d: %s", args.phase, " → ".join(phase_tasks))
        results = []
        for task_name in phase_tasks:
            task_class = TASK_REGISTRY.get(task_name)
            if not task_class:
                log.error("Task '%s' no encontrada", task_name)
                continue
            task = task_class(settings)
            success = task.run()
            results.append(success)
        return 0 if all(results) else 1

    orchestrator = Orchestrator(settings)
    result = orchestrator.run_all()
    return 0 if result.all_success else 1


if __name__ == "__main__":
    sys.exit(main())
