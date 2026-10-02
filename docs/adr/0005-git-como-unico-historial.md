# 0005. Git como unico historial; se borra el versionado por archivos

## Estado
Aceptada (Fase 0, 2026-10-01)

## Contexto

El versionado se hacia copiando archivos a mano:

```
tasks/actualizacion_inventario.BACKUP_20260907.py
tasks/actualizacion_ventas.BACKUP_20261001.py
tasks/descarga_informe_ventas.BACKUP_20260915.py
tasks/envio_informe_ventas.BACKUP_20261001.py
tasks/backup_descarga_informe_ventas.py
limpiar.BACKUP_20261001.py
Archivos Validacion/.../actualizar_inventario_general_viejo.py
backup.7z
```

Ese era el diagnostico P7 del plan: ruido, duda sobre cual es la fuente de
verdad, y riesgo de que un parche quede activo en el archivo equivocado.

## Decision

Git es el unico historial. Los archivos de respaldo se eliminan del arbol de
trabajo.

El orden importa y no es negociable:

1. `git init` + `.gitignore` + commit **baseline** con el estado completo,
   incluidos los respaldos.
2. `git rm` de los respaldos en un commit aparte.

 Asi el commit baseline conserva la historia y el borrado es reversible con
`git checkout f3f2ec5 -- <ruta>`.

## Consecuencias

Se gana: una sola fuente de verdad, diffs legibles, `git bisect` util sobre el
pipeline.

Se paga: 7.188 lineas borradas del arbol de trabajo (estan en el baseline).

Queda bloqueado: nada. `TASK_REGISTRY` no importaba ninguno de los archivos
eliminados, verificado antes de borrar.

## Nota

`chromedriver.exe.bak` tambien se elimino (duplicado binario de
`chromedriver.exe`). Este ultimo **no** se versiona: lo regenera
`core/chromedriver_utils.py`.
