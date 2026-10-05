"""Paridad con datos reales: motor legacy vs pipeline de reglas de ventas.

Uso (desde la carpeta del proyecto, con el .env de produccion):
    uv run python scripts/comparar_ventas.py

Lee los MISMOS archivos que ``actualizacion_ventas`` (plantilla, informe _268,
inventario actualizado, MYR .xlsb, matriz de clientes), corre los pasos 2 a 9
con ambos motores y compara el resultado fila a fila.

NO escribe la plantilla, NO hace backup, NO envia correos. Solo deja en la
carpeta de pruebas ``COMPARACION_VENTAS_<fecha_hora>.xlsx`` con:

- DIFERENCIAS     vacia = paridad total
- ELIMINACIONES   lineas que el pipeline excluye y por que regla
- RESUMEN REGLAS  filas antes/despues de cada regla

Codigo de salida 0 = paridad; 1 = hay diferencias (revisar el archivo).
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

_RAIZ = Path(__file__).resolve().parent.parent
for ruta in (_RAIZ, _RAIZ / "src"):
    if str(ruta) not in sys.path:
        sys.path.insert(0, str(ruta))

import pandas as pd  # noqa: E402

from config.settings import Settings  # noqa: E402
from insumos.adapters.excel.comparador import comparar_por_posicion  # noqa: E402
from tasks.actualizacion_ventas import ZONA_COLOMBIA, ActualizacionVentas, _hoy_bogota  # noqa: E402


def main() -> int:
    tarea = ActualizacionVentas(Settings())
    tarea.setup()
    log = tarea.log
    log.info("== Comparacion de motores de ventas (no escribe la plantilla) ==")

    p_ventas, p_inv, p_myr, p_mat, p_inf = tarea.localizar_archivos()
    stream, hoja, plantilla = tarea.cargar_plantilla(p_ventas)
    entradas = tarea.preparar_entradas(p_inf, p_inv, p_myr, p_mat, stream, list(plantilla.columns))

    log.info("Motor legacy...")
    actual = tarea._transformar_legacy(entradas, stream, _hoy_bogota().year)
    log.info("Motor reglas...")
    nuevo, resultado = tarea.transformar_reglas(entradas)

    # Las columnas fuera de la plantilla el legacy las ordenaba al azar (set):
    # se comparan en el mismo orden; el escritor COM solo usa las de la plantilla.
    nuevo = nuevo[[c for c in actual.columns if c in nuevo.columns] +
                  [c for c in nuevo.columns if c not in actual.columns]]
    diferencias = comparar_por_posicion(actual, nuevo)

    salida = tarea.DIR_PRUEBAS / f"COMPARACION_VENTAS_{datetime.now(ZONA_COLOMBIA):%Y%m%d_%H%M%S}.xlsx"
    salida.parent.mkdir(parents=True, exist_ok=True)
    eliminadas = resultado.audit.tabla_eliminaciones()
    with pd.ExcelWriter(salida, engine="openpyxl") as xw:
        diferencias.to_excel(xw, sheet_name="DIFERENCIAS", index=False)
        eliminadas.to_excel(xw, sheet_name="ELIMINACIONES", index=False)
        resultado.audit.tabla_pasos().to_excel(xw, sheet_name="RESUMEN REGLAS", index=False)
        pd.DataFrame(
            [("Plantilla", p_ventas.name), ("Hoja", hoja), ("Informe", p_inf.name), ("Inventario", p_inv.name),
             ("MYR", p_myr.name), ("Matriz", p_mat.name),
             ("Registros legacy", len(actual)), ("Registros reglas", len(nuevo))],
            columns=["DATO", "VALOR"],
        ).to_excel(xw, sheet_name="FUENTES", index=False)

    log.info("Registros: legacy=%d reglas=%d", len(actual), len(nuevo))
    if diferencias.empty:
        log.info("PARIDAD TOTAL. Reporte: %s", salida)
        log.info("Para activar el pipeline: VENTAS_MOTOR=reglas en el .env")
        return 0
    log.warning("HAY %d DIFERENCIAS. Revisar %s", len(diferencias), salida)
    for tipo, n in diferencias["TIPO"].value_counts().items():
        log.warning("  %s: %d", tipo, n)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
