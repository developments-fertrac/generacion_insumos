# Generacion de Insumos — Field Manual

Fertrac · Data Area

## Qué es esto

Pipeline que toma las exportaciones de la base de datos de Fertrac y produce
los archivos consolidados de Inventario y Ventas con los que trabaja el negocio.

Desde 2026-10 ya no hay descargas del ERP por Selenium: la base de datos deja
`Inventario.xlsx` e `InformesDeVentas(Facturas)_268*.xlsx` en `DB_EXPORT_DIR`
(ver `docs/adr/0008-fuente-base-de-datos.md`).

```
run.py  (CLI)  →  orchestrator.py  (fases)  →  BaseTask.execute()  (el trabajo)
```

## Cómo se ejecuta

```powershell
python run.py                                   # inventario + ventas
python run.py --workflow ventas                 # actualizacion → envio por WhatsApp
python run.py --workflow inventario             # solo inventario
python run.py --task actualizacion_inv --dry-run  # inventario de revision, sin tocar la plantilla
python run.py --list
```

Códigos de salida: `0` todo OK, `1` algún fallo.

## Estado del proyecto

| Tarea | Estado | Dónde vive |
|-------|--------|------------|
| `actualizacion_inv` | **Migrada** (hexagonal, fuente BD) | `src/insumos/` + `config/rules/inventario*.yaml` |
| `actualizacion_ventas` | Legacy; solo cambió la carpeta de origen | `tasks/actualizacion_ventas.py` |
| `envio_informe_ventas` | Legacy | `tasks/envio_informe_ventas.py` |
| Descargas ERP (Selenium) | **Retiradas** | — |

Pendiente: certificar el escritor COM de inventario en Windows (primera
corrida real), migrar ventas (Fase 4) y envío (Fase 5).

Plan completo: `docs/Plan_Arquitectura_Hexagonal_Generacion_Insumos.docx`.

## Reglas de desarrollo

```powershell
uv sync                     # instala dependencias
uv run ruff check .         # lint
uv run mypy                 # tipos: src/ y tests/, cero errores
uv run pytest               # pruebas
uv run lint-imports         # regla de dependencia hexagonal
```

Instalar herramientas de desarrollo:

```powershell
uv sync --extra dev
```

### El legacy se mide, no se ignora

`tasks/`, `core/`, `config/`, `orchestrator.py` y `run.py` tienen **117 errores
de tipos** conocidos. No se corrigen ahora (sería reescribir en Fase 0), pero
tampoco se los deja crecer: hay un *ratchet* que falla si el número sube.

```powershell
uv run python scripts/ratchet_types.py           # falla si sube de 117
uv run python scripts/ratchet_types.py --show    # desglose por archivo y código
uv run python scripts/ratchet_types.py --update  # re-línea la base (a mano)
```

Cuando se migra una tarea, el conteo baja solo y la base se actualiza. Con el
retiro de las descargas por Selenium y la migración de inventario, el conteo
real ya está por debajo de la base: correr `--update` en Windows para fijarla.

Limitación conocida: con `check_untyped_defs = false`, una función nueva en el
legacy **sin anotaciones** no se revisa y pasa el ratchet. Si agregas funciones
al legacy, anótalas.

### La regla que no se negocia

Las dependencias apuntan **hacia adentro**: `adapters → application → domain`.

El dominio solo importa `pandas`, `pandera` y la biblioteca estándar. Prohibido
`win32com`, `selenium`, `openpyxl`, `msoffcrypto`. `lint-imports` y
`tests/contract/test_regla_dependencia.py` bloquean el merge si se rompe.

## Configuración

Credenciales y rutas salen de `.env` (copiar de `.env.example`).

| Variable | Para qué |
|----------|----------|
| `DB_EXPORT_DIR` | Carpeta de las exportaciones de la base de datos (**requerida** para inventario) |
| `INVENTARIO_BD_FILE` | Nombre del archivo de inventario (por defecto `Inventario.xlsx`) |
| `EXCEL_PASSWORD` / `EXCEL_PASSWORDS_TRY` | Desencriptar archivos |
| `SMTP_*` | Notificaciones por correo |
| `BASE_PATH` / `REMISIONES_BASE` | Carpetas de producción |
| `STATE_DIR` / `LOGS_DIR` | Estado operativo (fuera del repo) |
| `WHATSAPP_CHATS` | Destinatarios del informe |

### Dónde viven los archivos

El repositorio contiene **solo código y configuración**. El estado operativo
va a `%LOCALAPPDATA%\GeneracionInsumos\`:

| Qué | Dónde |
|-----|-------|
| Logs | `%LOCALAPPDATA%\GeneracionInsumos\logs\<fecha>\<tarea>.log` |
| Sesión de WhatsApp | `...\chrome_profile_whatsapp_session\` |
| Cache de chromedriver | `...\.wdm\`, `...\.wdm_cache\` |
| Imágenes del informe | `...\Img informe\` |
| Screenshots de error | `...\errorImages\` |

Para cambiar la ubicación, `STATE_DIR` / `LOGS_DIR` en `.env`.

## Estructura

```
├── run.py, orchestrator.py        # CLI y orquestación (se mantienen)
├── tasks/                         # BaseTask: un archivo por tarea
├── core/                          # infraestructura compartida
├── config/
│   ├── settings.py                # .env → dataclasses
│   └── rules/*.yaml               # orden y parámetros de reglas
├── src/insumos/
│   ├── domain/                    # pandas + pandera. Sin I/O.
│   ├── application/               # casos de uso
│   └── adapters/                  # Excel (lectura, COM, reporte), sistema
├── scripts/ratchet_types.py       # control de degradación del legacy
├── quality-baseline.json          # línea base del ratchet (117)
├── mypy-legacy.toml               # config exclusiva del ratchet
├── tests/
│   ├── unit/ contract/ fixtures/
└── docs/adr/                      # decisiones técnicas
```

## Pruebas

```powershell
uv run pytest -m unit              # reglas, rápido
uv run pytest -m contract          # adaptores cumplen su puerto
uv run pytest -m e2e               # pipeline completo (Windows + Excel)
```

## Inventario general: cómo funciona

```
Inventario.xlsx (BD) ──► inventario_bd.yaml ──┐
                                               ├─► inventario.yaml ──► validar_salida ──► Excel (COM) + reporte
Plantilla INVENTARIO ──────────────────────────┘        ▲
Matriz USD + Distribución ─────────────────────────────┘
```

Reglas de negocio (todas en `config/rules/`):

| Regla | Qué hace |
|-------|----------|
| `inv.filtrar_motivo_inventario` | Solo entra `MOTIVO` que inicia con `INVENTARIO` |
| `inv.costo_cero_sin_existencia` | Existencia 0 → costo y total 0 |
| `inv.eliminar_referencias_duplicadas` | Una fila por referencia |
| `inv.eliminar_ausentes_en_bd` | Lo que la BD no trae sale del informe (con su motivo) |
| `inv.actualizar_desde_bd` | EXISTENCIA = EXISTENCIA NETA, COSTO PROMEDIO = COSTO, TOTAL INV = TOTAL INV, columnas *sistema* |
| `inv.agregar_referencias_nuevas_bd` | Referencias nuevas de la BD se agregan |
| `inv.calcular_nombre_lista_y_myr` | NOMBRE LISTA (Matriz USD) y NOMBRE MYR |
| `inv.completar_campos_faltantes` | Columnas COPIA, bodega gerencia, líder y clasificación vacías |
| `inv.ordenar_por_total_inv` | Mayor a menor TOTAL INV |
| `inv.validar_salida` | Si no cuadra contra la BD, **no se escribe** el archivo |

Cada corrida deja `Pruebas Inv General\Reportes\REPORTE INVENTARIO <fecha>.xlsx`
con el resumen por regla, lo excluido por la BD, lo eliminado de la plantilla y
las referencias nuevas.

## Añadir una regla de negocio

1. Clase pura en `src/insumos/domain/rules/<dominio>/` con `@rule("id")`,
   `description` para el negocio y `apply(df, ctx) -> RuleResult`.
2. Una entrada en `config/rules/*.yaml`.
3. Pruebas unitarias: camino feliz y 2 casos extremos.

Cambiar validaciones es editar el YAML, no el código.

## Documentación

- Plan de arquitectura: `docs/Plan_Arquitectura_Hexagonal_Generacion_Insumos.docx`
- Decisiones técnicas: `docs/adr/` (0008: fuente base de datos)
- Documentación funcional: `docs/DOCUMENTACION.md`, `docs/DOCUMENTACION_TASKS.md`
- Manual operativo: `docs/MANUAL_USUARIO.docx`
