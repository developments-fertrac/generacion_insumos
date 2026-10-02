from tasks.actualizacion_inventario import ActualizacionInventario
from tasks.actualizacion_ventas import ActualizacionVentas
from tasks.base_task import BaseTask
from tasks.envio_informe_ventas import EnvioInformeVentas

# Desde 2026-10 las descargas del ERP por Selenium (descarga_inv_general,
# descarga_ventas, descarga_valorizados) no existen: la base de datos exporta
# los insumos directamente (ver docs/adr/0008).
TASK_REGISTRY: dict[str, type[BaseTask]] = {
    "actualizacion_inv": ActualizacionInventario,
    "actualizacion_ventas": ActualizacionVentas,
    "envio_informe_ventas": EnvioInformeVentas,
}

__all__ = ["TASK_REGISTRY"]
