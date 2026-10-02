"""Muestra el catalogo de reglas registradas. Herramienta de diagnostico."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from insumos.domain.rules import registry  # noqa: E402
from insumos.domain.rules import inventario  # noqa: E402,F401  (dispara @rule)

for f in registry.catalogo():
    print(f"{f['id']:<50} {f['clase']}")

print()
print("total:", len(registry.registradas()))
