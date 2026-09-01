from tasks.base_task import BaseTask
from tasks.descarga_valorizados import DescargaValorizados
from tasks.descarga_inventario_general import DescargaInventarioGeneral
from tasks.descarga_informe_ventas import DescargaInformeVentas
from tasks.actualizacion_inventario import ActualizacionInventario
from tasks.actualizacion_ventas import ActualizacionVentas
from tasks.envio_informe_ventas import EnvioInformeVentas

TASK_REGISTRY: dict[str, type[BaseTask]] = {
    "descarga_valorizados": DescargaValorizados,
    "descarga_inv_general": DescargaInventarioGeneral,
    "descarga_ventas": DescargaInformeVentas,
    "actualizacion_inv": ActualizacionInventario,
    "actualizacion_ventas": ActualizacionVentas,
    "envio_informe_ventas": EnvioInformeVentas,
}

__all__ = ["TASK_REGISTRY"]
