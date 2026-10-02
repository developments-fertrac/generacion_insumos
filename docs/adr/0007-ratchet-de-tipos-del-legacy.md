# 0007. El legacy se mide con un ratchet; el codigo nuevo no negocia

## Estado
Aceptada (Fase 0, revision 2026-10-01)

## Contexto

El codigo legacy (`tasks/`, `core/`, `config/`, `orchestrator.py`, `run.py`)
tiene **117 errores de tipos** de mypy, sobre todo `WebDriver | None` de
Selenium y asignaciones laxas. `src/` y `tests/` tienen cero.

Hay tres caminos:

1. Arreglar los 117 ahora. Es reescritura con un disfraz de lint: mezclar
   cambios mecanicos con cambios de comportamiento de negocio justo antes de
   migrar.
2. `ignore_errors = true` y seguir. Silencioso: el numero puede subir sin que
   nadie entere, y para entonces el codigo viejo ya estara peor que hoy.
3. Fijar el numero y fallar si sube.

## Decision

Opcion 3, con dos redes:

- **Codigo nuevo** (`src/`, `tests/`): `mypy` estricto, cero errores, sin
  excepciones. `disallow_untyped_defs`, `warn_return_any` en el dominio.
- **Legacy**: `scripts/ratchet_types.py` cuenta los errores y falla si el total
  **supera** la linea base de `quality-baseline.json` (117). Solo puede bajar,
  y baja solo cuando se migra una tarea.

El ratchet compara **codigos de error**, no lineas de texto, para que un
reformulado de mypy entre versiones no lo dispare.

## Consecuencias

Se gana: la deuda no crece mientras se migra. Cuando `tasks/actualizacion_
ventas.py` se migre en la Fase 4, su bloque de errores desaparece de la linea
base y el ratchet sigue funcionando sin tocarlo.

Se paga: dos configuraciones de mypy (`pyproject.toml` para lo nuevo,
`mypy-legacy.toml` para el conteo). Se evita que alguien las mezcle porque
`mypy-legacy.toml` dice en su encabezado que no se usa para dar el codigo por
bueno.

Queda bloqueado: nada. `mypy-legacy.toml` desaparece con el legacy en la Fase 5.

## Limitacion aceptada

Con `check_untyped_defs = false`, mypy no revisa el cuerpo de una funcion **sin
anotaciones**. Una funcion nueva agregada al legacy sin tipos pasa el ratchet
sin que nada lo note.

El ratchet es una red secundaria. La primaria es el mypy estricto sobre codigo
nuevo. Regla practica: al agregar funciones al legacy, anotalas.

Ademas, el conteo depende de la version de mypy y de los stubs instalados.
Actualizar dependencias puede mover el numero sin que nadie toque el codigo;
por eso `--show` imprime el desglose y `--update` exige una decision manual.

## Relacionado

El mismo motivo (la proteccion que depende de que alguien lea un comentario no
es una proteccion) se aplico a los pipelines de reglas vacios:
`src/insumos/domain/reglas_declarativas.py` lanza `ConfiguracionInvalida` cuando
`rules` queda vacio sin `allow_empty: true`.
