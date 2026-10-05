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

Siempre con `uv` (el python global no tiene las dependencias del proyecto):

```powershell
uv run python run.py                                   # completo: inventario + ventas (sin envio)
uv run python run.py --workflow inventario             # solo inventario (debe correr ANTES que ventas)
uv run python run.py --workflow ventas                 # actualizacion de ventas → envio por WhatsApp
uv run python run.py --workflow parcialVentas          # solo actualizacion de ventas
uv run python run.py --workflow envio                  # solo envio del informe por WhatsApp
uv run python run.py --task actualizacion_inv --dry-run  # inventario de revision, sin tocar la plantilla
uv run python run.py --list
```

Códigos de salida: `0` todo OK, `1` algún fallo.

Cada tarea envía **un solo correo** por corrida (`[OK]` o `[ERROR]`), desde
`BaseTask.run` y con el log del día adjunto. El de éxito sale solo cuando todo
terminó (en ventas, después de tablas dinámicas, reemplazo y copia `.xlsb`; en
el envío, solo si algún destinatario recibió el informe). Todas las fechas
(carpetas, logs, correos, filtros) son en hora de Colombia.

En producción lo dispara n8n (`schtasks /Run`) mediante los `.bat` de cada
tarea, que ejecutan `uv run --frozen` y dejan `resultado_<tarea>.txt` con el
código de salida. La tarea programada debe correr en **sesión interactiva**
(usuario con sesión abierta; bloqueada sirve, desconectada no): Excel COM y el
portapapeles no funcionan sin escritorio.

## Estado del proyecto

| Tarea | Estado | Dónde vive |
|-------|--------|------------|
| `actualizacion_inv` | **Migrada** (hexagonal, fuente BD). Corrida real exitosa 2026-10-02 | `src/insumos/` + `config/rules/inventario*.yaml` |
| `actualizacion_ventas` | **Motor dual** (ADR 0009): pasos 2–9 como 21 reglas `ven.*`; paridad probada con datos sintéticos. Por defecto `VENTAS_MOTOR=legacy` | `tasks/actualizacion_ventas.py` + `src/insumos/` + `config/rules/ventas.yaml` |
| `envio_informe_ventas` | Legacy (Fase 5 pendiente) | `tasks/envio_informe_ventas.py` |
| Descargas ERP (Selenium) | **Retiradas** (ADR 0008) | — |

Pendiente:

1. Ventas: `scripts/comparar_ventas.py` con datos reales; con paridad total,
   `VENTAS_MOTOR=reglas`. Dos semanas sin diferencias → borrar
   `_transformar_legacy` (Fase 4b: lectura/descifrado al adaptador `FuenteVentas`).
2. Fase 5: migrar `envio_informe_ventas`.
3. Operativos: `VENTAS_ACTUALIZACION_PASSWORD` en el `.env` de producción,
   regenerar `uv.lock` con red (`uv lock`: retira pandera y PyMuPDF, ADR 0010)
   y borrar `_obsoleto_2026-10-02/` (sus logs ya no tienen contraseñas).

Plan completo: `docs/Old/Plan_Arquitectura_Hexagonal_Generacion_Insumos.docx`.

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

`tasks/`, `core/`, `config/`, `orchestrator.py` y `run.py` tienen **12 errores
de tipos** conocidos (línea base del 2026-10-05, `quality-baseline.json`). No se corrigen ahora (sería reescribir en Fase 0), pero
tampoco se los deja crecer: hay un *ratchet* que falla si el número sube.

```powershell
uv run python scripts/ratchet_types.py           # falla si sube de la base (12)
uv run python scripts/ratchet_types.py --show    # desglose por archivo y código
uv run python scripts/ratchet_types.py --update  # re-línea la base (a mano)
```

Cuando se migra una tarea, el conteo baja solo y la base se actualiza.
`mypy-legacy.toml` analiza siempre como Windows (`platform = "win32"`) y con
`src` en el path, así que el conteo es el mismo en cualquier máquina; los
errores del núcleo nuevo no cuentan aquí (los revisa el mypy estricto).

Limitación conocida: con `check_untyped_defs = false`, una función nueva en el
legacy **sin anotaciones** no se revisa y pasa el ratchet. Si agregas funciones
al legacy, anótalas.

### La regla que no se negocia

Las dependencias apuntan **hacia adentro**: `adapters → application → domain`.

El dominio solo importa `pandas`, `PyYAML`, `unidecode` y la biblioteca estándar. Prohibido
`win32com`, `selenium`, `openpyxl`, `msoffcrypto`. `lint-imports` y
`tests/contract/test_regla_dependencia.py` bloquean el merge si se rompe.

## Configuración

Credenciales y rutas salen de `.env` (copiar de `.env.example`).

| Variable | Para qué |
|----------|----------|
| `DB_EXPORT_DIR` | Carpeta de las exportaciones de la base de datos (**requerida** para inventario; ventas toma de aquí el `_268`) |
| `INVENTARIO_BD_FILE` | Nombre del archivo de inventario (por defecto `Inventario.xlsx`) |
| `INVENTARIO_TMP_DIR` | Carpeta de la copia temporal de la plantilla que abre Excel (vacío = carpeta de salida) |
| `EXCEL_PASSWORD` / `EXCEL_PASSWORDS_TRY` | Desencriptar archivos (la principal es obligatoria) |
| `VENTAS_ACTUALIZACION_PASSWORD` | Clave exclusiva de `$2026 VENTAS_Actualizacion.xlsx` (área autorizada); si falta se usa `EXCEL_PASSWORD` |
| `VENTAS_MOTOR` | Motor de ventas: `legacy` (defecto) o `reglas` (ADR 0009) |
| `SMTP_*` | Notificaciones por correo (`SMTP_SENDER_EMAIL` y `SMTP_PASSWORD` obligatorias) |
| `BASE_PATH` | Carpeta raíz de producción (`ARCHIVOS DIARIOS 2026`) |
| `STATE_DIR` / `LOGS_DIR` | Estado operativo (fuera del repo) |
| `WHATSAPP_CHATS` | Destinatarios del informe, separados por coma |

Ninguna contraseña de estas variables llega a los logs: `core/seguridad.py`
las reemplaza por su huella (`sha256:...`).

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
│   ├── domain/                    # pandas. Sin I/O.
│   ├── application/               # casos de uso
│   └── adapters/                  # Excel (lectura, COM, reporte), sistema
├── scripts/
│   ├── ratchet_types.py           # control de degradación del legacy
│   ├── comparar_ventas.py         # paridad legacy vs reglas de ventas (datos reales, no escribe)
│   ├── comparar_estructura.py     # estructura de dos archivos de inventario
│   ├── ver_reglas.py              # catálogo de reglas registradas
│   └── publicar_a_produccion.ps1  # copia código a producción + uv sync
├── quality-baseline.json          # línea base del ratchet (12)
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

## Ventas: cómo funciona

```
InformesDeVentas(Facturas)_268 (BD) ─┐
Inventario actualizado (salida inv) ─┤
MYR EXISTENCIA (.xlsb) ──────────────┼─► VENTAS_MOTOR ─► legacy | ventas.yaml ─► Excel (COM) ─► tablas dinámicas
Matriz de clientes ──────────────────┤                                                        ─► reducir tamaño
PRECIO UNIT LICITADOS (plantilla) ───┘                                                        ─► copia .xlsb
```

1. Respaldo de `Pruebas\$2026 VENTAS_Actualizacion.xlsx` en `Backups_Ventas`
   (si falla, se detiene).
2. Transformación (pasos 2–9): excluye NC/NCDTO/NDCTO, fletes, publicidad y
   otros años; cruza LINEA/SUBLINEA/LIDER, COSTO FACTOR HOY, DCTO CONDICIONADO
   y VTA ACORDADA LICITADO. La referencia es **texto** (`"0123"` sigue `"0123"`).
3. Escritura COM (paso 10, sin cambios): fórmulas, `Resum Mes` (L20, L21,
   festivos L23:L27), metas, subtotales, contraseña del área.
4. Post-proceso: un solo PivotCache compartido para todas las tablas
   dinámicas, eliminación de dibujos sobrantes (`core/xlsx_cleaner.py`),
   reemplazo atómico del archivo y copia `$2026 VENTAS_<fecha>.xlsb`.

Paridad: `uv run pytest tests/unit/rules/test_paridad_ventas.py` (sintético) y
`uv run python scripts/comparar_ventas.py` (real; código 0 = paridad, deja
`COMPARACION_VENTAS_<fecha>.xlsx` en `Pruebas`). Mientras exista el motor dual,
todo cambio de negocio en ventas se aplica en **ambos** motores.

## Añadir una regla de negocio

1. Clase pura en `src/insumos/domain/rules/<dominio>/` con `@rule("id")`,
   `description` para el negocio y `apply(df, ctx) -> RuleResult`.
2. Una entrada en `config/rules/*.yaml`.
3. Pruebas unitarias: camino feliz y 2 casos extremos.

Cambiar validaciones es editar el YAML, no el código.

## Documentación

- **Documentación vigente (usuario + técnica, función por función):**
  `docs/Documentacion_Generacion_de_Insumos.docx` (2026-10-05)
- Decisiones técnicas: `docs/adr/` (0008: fuente base de datos; 0009: ventas con motor dual;
  0010: cierre de hallazgos de la auditoría)
- Histórico (anterior a la migración, puede estar desactualizado): `docs/Old/`
  — plan de arquitectura, `DOCUMENTACION*.md` y `MANUAL_USUARIO.docx`
