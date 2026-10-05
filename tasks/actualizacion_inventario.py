"""Tarea: Actualizacion de Inventario General (fuente: base de datos).

Composicion (raiz de composicion de la arquitectura hexagonal):

    FuenteInventarioArchivos  ->  ActualizarInventario  ->  EscritorInventarioCom
         (adaptador)               (caso de uso)              (adaptador, Windows)
                                         |
                              config/rules/inventario_bd.yaml
                              config/rules/inventario.yaml

Reemplaza la version de 5.300 lineas que descargaba el ERP por Selenium y
recalculaba existencias desde los valorizados. Desde 2026-10 la base de datos
exporta ``Inventario.xlsx`` con la existencia neta, el costo y el total ya
calculados (ver docs/adr/0008-fuente-base-de-datos.md).

Modo revision: ``python run.py --task actualizacion_inv --dry-run`` escribe en
``Pruebas Inv General/dry-run`` un archivo con la misma estructura que el de
produccion (openpyxl sobre una copia de la plantilla), sin abrir Excel ni tocar
la plantilla original.
"""

from __future__ import annotations

import os
import sys
from datetime import date, datetime
from pathlib import Path

_SRC = Path(__file__).resolve().parent.parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from config.settings import Settings  # noqa: E402
from core.email_notifier import EmailNotifier  # noqa: E402
from insumos.adapters.excel.escritor_openpyxl import EscritorInventarioOpenpyxl  # noqa: E402
from insumos.adapters.excel.fuentes_inventario import FuenteInventarioArchivos  # noqa: E402
from insumos.adapters.excel.reporte_inventario import EscritorReporteXlsx  # noqa: E402
from insumos.adapters.system import reloj  # noqa: E402  (hora Colombia)
from insumos.application.actualizar_inventario import (  # noqa: E402
    ActualizarInventario,
    ResultadoActualizacion,
)
from tasks.base_task import BaseTask  # noqa: E402

_RAIZ = Path(__file__).resolve().parent.parent
REGLAS_BD = _RAIZ / "config" / "rules" / "inventario_bd.yaml"
REGLAS_INVENTARIO = _RAIZ / "config" / "rules" / "inventario.yaml"


class ActualizacionInventario(BaseTask):
    name = "actualizacion_inventario"

    def __init__(self, settings: Settings):
        super().__init__(settings)
        self.resultado: ResultadoActualizacion | None = None
        self.dry_run = os.getenv("INSUMOS_DRY_RUN", "").strip() in ("1", "true", "yes")

    def setup(self) -> None:
        super().setup()
        if not self.dry_run:
            self.notifier = EmailNotifier(self.settings.smtp_inv_general, self.name)

    def construir(self) -> ActualizarInventario:
        paths = self.settings.paths
        password = self.settings.excel.password
        salida = paths.output_inv_general
        fuente = FuenteInventarioArchivos(
            archivo_bd=paths.inventario_bd,
            carpeta_salida=salida,
            carpeta_insumos=paths.base,
            password=password,
        )
        if self.dry_run:
            # Misma estructura que produccion (hojas, formatos, formulas, COPIA),
            # escrita con openpyxl sobre una copia de la plantilla: no abre Excel.
            carpeta = salida / "dry-run"
            escritor = EscritorInventarioOpenpyxl(
                plantilla=fuente.ruta_plantilla(),
                carpeta_salida=carpeta,
                password=password,
            )
            reporte = EscritorReporteXlsx(carpeta)
        else:
            from insumos.adapters.excel.com_inventario import EscritorInventarioCom

            escritor = EscritorInventarioCom(
                plantilla=fuente.ruta_plantilla(),
                carpeta_salida=salida,
                password=password,
                carpeta_temporal=paths.temporal_inventario,  # INVENTARIO_TMP_DIR en .env
            )
            reporte = EscritorReporteXlsx(salida)  # misma carpeta que el proceso anterior
        return ActualizarInventario(
            fuente=fuente,
            escritor=escritor,
            reporte=reporte,
            reglas_bd=REGLAS_BD,
            reglas_inventario=REGLAS_INVENTARIO,
            hoy=reloj.hoy(),
        )

    def execute(self) -> None:
        caso = self.construir()
        self.log.info("Base de datos: %s", self.settings.paths.inventario_bd)
        self.log.info("Modo: %s", "DRY-RUN (sin tocar la plantilla)" if self.dry_run else "produccion")

        self.resultado = caso()

        for audit in (self.resultado.audit_bd, self.resultado.audit_inventario):
            for paso in audit.pasos:
                self.log.info(
                    "  [%s] %-38s %6d -> %6d  elim=%-5d mod=%-5d %s",
                    audit.pipeline, paso.id_regla, paso.filas_antes, paso.filas_despues,
                    paso.eliminadas, paso.modificadas, "; ".join(paso.warnings),
                )
        advertencias = getattr(caso.fuente, "advertencias", None) or []
        for aviso in advertencias:
            self.log.warning(aviso)
        for clave, valor in self.resultado.resumen().items():
            self.log.info("  %s: %s", clave, valor)

    def _notify_success(self, detail: str) -> None:
        """Mismo correo del proceso anterior: resumen + archivo generado adjunto."""
        if not self.notifier or self.resultado is None:
            return
        r = self.resultado.resumen()
        salida = self.resultado.archivo_salida
        reporte = self.resultado.archivo_reporte
        fuente_bd = self.settings.paths.inventario_bd
        lineas = [
            "Proceso de Actualizacion de Inventario General completado.",
            "",
            f"Fecha de ejecucion: {reloj.ahora():%d/%m/%Y %H:%M:%S}",
            f"Carpeta de trabajo: {self.settings.paths.output_inv_general}",
            f"Archivo generado: {salida.name if salida else ''}",
            "",
            _linea_fuente(fuente_bd),
            "",
            "ESTADISTICAS:",
            f"  - Registros inventario original: {self.resultado.audit_inventario.pasos[0].filas_antes:,}",
            f"  - Referencias en el informe: {r['referencias']:,}",
            f"  - Referencias nuevas agregadas: {r['referencias_nuevas']:,}",
            f"  - Excluidas por la base de datos (MOTIVO): {r['excluidas_por_bd']:,}",
            f"  - Eliminaciones en la plantilla: {r['eliminadas_de_plantilla']:,}",
            f"  - Existencia total: {r['existencia_total']:,.0f}",
            f"  - Total inventario: ${r['total_inv']:,.2f}",
            "",
            "ARCHIVOS GENERADOS:",
            f"  - {salida.name if salida else ''}",
        ]
        if reporte:
            lineas.append(f"  - {reporte.name}")
        adjunto = salida if salida and salida.exists() else None
        self.notifier.notify_success("<br>".join(lineas), attachment=adjunto)


def _linea_fuente(ruta: Path) -> str:
    """Como el aviso de 'Fuente ERP' del proceso anterior, ahora para la base de datos."""
    if not ruta.exists():
        return f"Fuente base de datos: {ruta.name} (no disponible)"
    modificado = reloj.desde_timestamp(ruta.stat().st_mtime)
    linea = f"Fuente base de datos: {ruta.name} (modificado: {modificado:%d/%m/%Y %H:%M:%S})"
    dias = (reloj.hoy() - modificado.date()).days
    if dias > 0:
        linea += f"<br>*** ATENCION: este archivo NO es de hoy (tiene {dias} dia(s) de antiguedad). ***"
    return linea
