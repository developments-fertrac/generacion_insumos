# 0010. Cierre de hallazgos de la auditoria de documentacion (2026-10-05)

## Estado
Aceptada (2026-10-05). Modifica parcialmente la 0002 (pandera).

## Contexto

La documentacion exhaustiva del 2026-10-05
(`docs/Documentacion_Generacion_de_Insumos.docx`) dejo 11 hallazgos: correos
duplicados, hora del servidor en el envio, procesos de Excel/Chrome del
usuario terminados a la fuerza, codigo y dependencias sin uso, documentacion
desalineada, linea base del ratchet desactualizada, riesgo de hojas duplicadas
en el reporte y logs antiguos con contrasenas.

## Decision

1. **Notificaciones.** Cada tarea envia un solo correo por corrida, desde
   `BaseTask.run` (`_notify_success` / `_notify_failure` sobrescritos en
   ventas y envio con el log adjunto). El de exito sale solo cuando todo
   termino; el envio valida que algun destinatario haya recibido el informe
   antes de considerarse exitoso.
2. **Hora de Colombia en todo el proceso**: envio (mes, dia, carpeta del log),
   correos, carpeta mensual de `settings` e indicador `PROCESANDO.txt`.
3. **Solo se terminan procesos propios**: el envio cierra el Chrome que usa el
   perfil de WhatsApp (por su `user-data-dir`) y el Excel que el mismo abrio
   (por PID). Nunca `taskkill /IM chrome.exe` ni `/IM EXCEL.EXE`.
4. **Se retira codigo sin uso** (git conserva la historia): conversiones COM y
   xlrd y copias temporales de `core/excel_processing.py`; busquedas y
   resolucion de columnas sin uso de `core/excel_utils.py`; configuracion de
   REMISIONES; ayudas sin uso de `chromedriver_utils` y `archivos`;
   `EscritorInventarioSimple`; constantes y metodos sin uso de
   `ActualizacionVentas`.
5. **pandera y PyMuPDF salen de las dependencias**: ninguno se importa. La 0002
   se mantiene en lo demas (pandas es el portador de datos del dominio); si se
   escribe el primer esquema pandera, se vuelve a declarar la dependencia.
   `HEADLESS` sale de `.env.example` (ningun modulo la leia).
6. **Reporte de inventario**: los nombres de hoja por PASO son unicos aunque
   coincidan los primeros 31 caracteres.
7. **Logs antiguos**: las contrasenas encontradas en logs de
   `Archivos Validacion/` y `_obsoleto_2026-10-02/` se reemplazan por su huella
   (`sha256:...`) en el mismo archivo; no se borra nada.

## Consecuencias

Se gana: un correo por corrida y en el momento correcto, fechas consistentes
con el negocio, ninguna sesion del usuario cerrada a la fuerza, menos codigo
que mantener y logs sin secretos.

Se paga: el envio depende de poder leer el command line de los procesos
(psutil si esta instalado; si no, PowerShell/CIM). Si no lo logra, no cierra
Chrome en lugar de cerrar el del usuario.

Queda pendiente: re-linear el ratchet de tipos en Windows
(`uv run python scripts/ratchet_types.py --update`) y regenerar `uv.lock` con
red (`uv lock`) para que refleje el retiro de pandera y PyMuPDF.
