"""Ratchet de errores de tipos sobre el codigo legacy.

Que hace
--------
`mypy` sobre `src/` y `tests/` exige cero errores. Para el legacy eso no es
factible hoy (117 errores, sobre todo `WebDriver | None` de Selenium), asi que
en vez de ignorarlo se lo mide: este script cuenta los errores y falla si el
numero **sube**.

Eso convierte "arreglalo cuando puedas" en "no lo empeores". El conteo solo
puede bajar, y baja solo cuando una tarea se migra.

Uso
---
    python scripts/ratchet_types.py            # verifica (usar en CI)
    python scripts/ratchet_types.py --show     # imprime el desglose
    python scripts/ratchet_types.py --update   # re-linea la base a mano

Notas de diseno
---------------
- La linea base esta en `quality-baseline.json`, versionada. Cambiarla es una
  decision consciente, no un efecto colateral.
- Se comparan **codigos de error**, no lineas de texto: mypy cambia sus
  mensajes entre versiones y un reformulado no debe romper el ratchet.
- El conteo depende de la version de mypy y de los stubs instalados. Si se
  actualizan las dependencias, el numero puede moverse sin que nadie toque el
  codigo. Por eso `--show` imprime el desglose: ante una suba inesperada, lo
  primero es comprobar si cambio el entorno.

Punto ciego conocido
--------------------
Con `check_untyped_defs = false`, mypy **no revisa el cuerpo de una funcion
sin anotaciones**. Una funcion nueva agregada al legacy sin tipos de retorno
(por ejemplo `def f(x): return x + 1`) pasa el ratchet sin que nada lo note.

Por eso el ratchet es una red secundaria, no la unica:
- codigo nuevo (`src/`, `tests/`): `mypy` estricto, cero errores, sin excusas.
- legacy: solo se controla que no empeore lo ya anotado.

Regla practica: si agregas funciones al legacy, anotalas. Las que lleguen sin
tipos pasan inadvertidas hasta que las migres.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CONFIG = RAIZ / "mypy-legacy.toml"
LINEA_BASE = RAIZ / "quality-baseline.json"

# mypy: file.py:12: error: mensaje  [codigo]
# Los mensajes contienen dos puntos y rutas, asi que se ancla al final.
_RE_ERROR = re.compile(r"^(?P<archivo>[^:]+):(?P<linea>\d+):\s+error:\s+(?P<mensaje>.*)$")


def _mypy() -> str:
    """Ejecuta mypy con la config del legacy y devuelve stdout."""
    proceso = subprocess.run(
        [sys.executable, "-m", "mypy", "--no-pretty", "--no-error-summary",
         "--config-file", str(CONFIG)],
        cwd=RAIZ,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    # mypy sale con 1 cuando encuentra errores: es lo esperado, no un fallo.
    if proceso.returncode not in (0, 1):
        raise SystemExit(
            f"mypy fallo inesperadamente (codigo {proceso.returncode}):\n{proceso.stderr}"
        )
    return proceso.stdout


def contar() -> Counter[str]:
    """Devuelve Counter de 'archivo|codigo-de-error'."""
    conteo: Counter[str] = Counter()
    for linea in _mypy().splitlines():
        m = _RE_ERROR.match(linea.strip())
        if not m:
            continue
        archivo = m.group("archivo").replace("\\", "/").lower()
        codigo = m.group("mensaje").rsplit("[", 1)[-1].rstrip("]").strip()
        if not codigo or "[" not in m.group("mensaje"):
            codigo = "sin-codigo"
        conteo[f"{archivo}|{codigo}"] += 1
    return conteo


def leer_base() -> dict[str, int]:
    if not LINEA_BASE.exists():
        raise SystemExit(f"No existe {LINEA_BASE.name}. Generalo con --update")
    datos = json.loads(LINEA_BASE.read_text(encoding="utf-8"))
    return dict(datos["errors"])


def guardar_base(conteo: Counter[str], nota: str) -> None:
    LINEA_BASE.write_text(
        json.dumps(
            {
                "_comentario": (
                    "Linea base del ratchet de tipos del legacy. "
                    "Ver scripts/ratchet_types.py. Solo puede BAJAR: "
                    "la actualiza quien migra una tarea."
                ),
                "nota": nota,
                "total": sum(conteo.values()),
                "errors": dict(sorted(conteo.items())),
            },
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    grupo = parser.add_mutually_exclusive_group()
    grupo.add_argument("--show", action="store_true", help="imprime el desglose")
    grupo.add_argument("--update", action="store_true", help="re-linea la base")
    args = parser.parse_args()

    conteo = contar()
    total = sum(conteo.values())

    if args.show:
        print(f"Errores de tipos en el legacy: {total}\n")
        for clave, n in conteo.most_common():
            print(f"  {n:>4}  {clave}")
        return 0

    if args.update:
        guardar_base(conteo, "re-lineado a mano")
        print(f"Linea base actualizada: {total} errores")
        return 0

    base = leer_base()
    total_base = sum(base.values())

    if total > total_base:
        nuevos = dict(conteo - Counter(base))
        print(f"RATCHET ROTO: {total} errores de tipos, la base es {total_base}.")
        print(f"Subieron {total - total_base}.\n")
        for clave, n in sorted(nuevos.items()):
            print(f"  +{n:<4} {clave}")
        print(
            "\nEl legacy no puede empeorar. Arreglalo, o si el cambio del "
            "conteo viene de una actualizacion de dependencias, revisa con "
            "--show y luego --update."
        )
        return 1

    if total < total_base:
        print(f"Mejoró: {total} errores (la base era {total_base}).")
        print("Actualiza la base con --update para aprovechar el trinquete.")
    else:
        print(f"Ratchet OK: {total} errores, igual que la base.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
