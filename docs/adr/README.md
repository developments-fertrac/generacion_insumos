# Architecture Decision Records

Registro de las decisiones tecnicas de la migracion a arquitectura hexagonal.
Una decision = un archivo `NNNN-titulo-corto.md` inmutable: si se revierte, se
escribe un ADR nuevo que referencia al anterior.

Formato (MADR, resumido):

```markdown
# NNNN. Titulo

## Estado
Propuesta | Aceptada | Reemplazada por NNNN | Obsoleta

## Contexto
Que problema obliga a decidir.

## Decision
Que se decide.

## Consecuencias
Que se gana, que se paga, que queda bloqueado.
```

## Indice

| # | Decision | Estado |
|---|----------|--------|
| 0001 | Migracion incremental (strangler fig), no reescritura | Aceptada |
| 0002 | pandas + pandera como portador de datos del dominio | Aceptada |
| 0003 | Orden de reglas en YAML, no en Python | Aceptada |
| 0004 | COM se queda en fase 1; openpyxl se evalua en fase 6 | Aceptada |
| 0005 | Git como unico historial; se borra el versionado por archivos | Aceptada |
| 0006 | El estado operativo (logs, perfiles, cache) vive fuera del repo | Aceptada |
| 0007 | Ratchet de tipos sobre el legacy; mypy estricto en codigo nuevo | Aceptada |
| 0008 | La base de datos reemplaza la descarga del ERP por Selenium | Aceptada |

## Decisiones de la seccion 6 del plan aun sin ADR

Se redactan al cierre de la fase correspondiente: composicion root manual en
`bootstrap.py`, pandera `lazy=True`, eliminaciones por filtrado del DataFrame,
y la evaluacion de Prefect/Dagster para el orquestador.
