# 03 · Flujos de datos

> Fuente: hiperaristas `he_composicion_inventario`, `he_motor_dual_paridad`, `he_paso10_postproceso` y comunidades «Composición de inventario», «Ejecución y post-proceso de ventas», «Paso 10 COM de ventas», «Motor dual (VENTAS_MOTOR)», «WhatsApp Web (Selenium)» y «Capturas del informe» de `graphify-out/graph.json`.
> Rutas expresadas relativas a `BASE_PATH` / `DB_EXPORT_DIR` (nunca rutas absolutas de usuario).

## 1. Inventario General (`actualizacion_inv`) — migrado

> **Diagrama interactivo** (Archify, verificado contra `088f6f0`): [abrir en el índice](diagramas/main.html#04_inventario) · [abrir aparte](diagramas/03_flujo-inventario.html). El bloque Mermaid de abajo es la versión resumida para leer en GitHub/VS Code.

```mermaid
sequenceDiagram
    autonumber
    participant N as n8n / operador
    participant R as run.py + Orchestrator
    participant T as ActualizacionInventario (raíz de composición)
    participant F as FuenteInventarioArchivos
    participant U as ActualizarInventario
    participant P as RulePipeline
    participant W as EscritorInventarioCom | EscritorInventarioOpenpyxl
    participant Rep as EscritorReporteXlsx
    participant M as EmailNotifier
    N->>R: run.py --workflow inventario [--dry-run]
    R->>T: BaseTask.run() → setup() → execute()
    T->>T: construir(): adaptadores según INSUMOS_DRY_RUN
    T->>U: caso()
    U->>F: inventario_bd()  (DB_EXPORT_DIR/Inventario.xlsx)
    U->>P: from_yaml(inventario_bd.yaml).run(bd)
    Note over P: inv.validar_columnas_bd → inv.filtrar_motivo_inventario(INVENTARIO)<br/>→ inv.eliminar_referencias_duplicadas → inv.costo_cero_sin_existencia
    P-->>U: BD filtrada + AuditTrail (excluidas con MOTIVO)
    U->>F: plantilla() (última "ACTUALIZADO" o maestro), matriz_usd(), distribucion()
    U->>P: from_yaml(inventario.yaml).run(plantilla, ctx)
    Note over P: 9 reglas: duplicados → ausentes en BD → actualizar desde BD → nuevas →<br/>NOMBRE LISTA/MYR → campos faltantes → costo 0 → ordenar → validar_salida
    alt inv.validar_salida falla
        P-->>U: CuadreFallido → ErrorDePipeline
        U-->>T: excepción (no se escribe nada)
        T->>M: notify_failure (correo [ERROR])
    else cuadra
        U->>W: escribir(df, hoy)
        W-->>U: $2026 INVENTARIO GENERAL ACTUALIZADO «fecha».xlsx (cifrado)
        U->>Rep: escribir(EvidenciaEjecucion)
        Rep-->>U: REPORTE_ELIMINACIONES_«fecha».xlsx
        U-->>T: ResultadoActualizacion
        T->>M: notify_success(resumen + adjunto) — omitido en dry-run
    end
```

| Etapa | Entrada | Transformación / validación | Salida | Evidencia |
|---|---|---|---|---|
| Composición | `Settings`, `INSUMOS_DRY_RUN` | Selecciona escritor COM (producción) u openpyxl (`dry-run/`) | `ActualizarInventario` | `tasks/actualizacion_inventario.py:L63` · `tasks_actualizacion_inventario_actualizacioninventario_construir` |
| Lectura BD | `Inventario.xlsx` | Encabezados normalizados | DataFrame BD | `src/insumos/adapters/excel/fuentes_inventario.py` · `src_insumos_adapters_excel_fuentes_inventario_leer_inventario_bd` |
| Pipeline BD | DataFrame BD | 4 reglas (`config/rules/inventario_bd.yaml`) | BD filtrada + motivos de exclusión | `config_rules_inventario_bd` |
| Plantilla | Última salida `ACTUALIZADO` o maestro | `ruta_plantilla()` vía `archivos.mas_reciente` | `PlantillaInventario` | `src/insumos/adapters/excel/fuentes_inventario.py:L94` · `src_insumos_adapters_excel_fuentes_inventario_fuenteinventarioarchivos` |
| Pipeline plantilla | Plantilla + ctx | 9 reglas (`config/rules/inventario.yaml`); `inv.validar_salida` bloquea | df final | `config_rules_inventario`; `src_insumos_domain_rules_inventario_actualizacion_validarsalida` |
| Escritura | df final | 10 pasos COM: descifrar, mapear encabezados, retitular `EXISTENCIA <MES DD>`, escribir por bloques, fórmulas Dif, subtotales, tablas dinámicas, `INVENTARIO COPIA`, guardar y cifrar | Libro cifrado | `src/insumos/adapters/excel/com_inventario.py:L231` · `src_insumos_adapters_excel_com_inventario_escritorinventariocom_escribir` |
| Evidencia | Auditorías + fuentes | Hojas por paso con nombres únicos ≤ 31 car. | Reporte de eliminaciones | `src/insumos/adapters/excel/reporte_inventario.py:L128` · `src_insumos_adapters_excel_reporte_inventario_escritorreportexlsx`; `src_insumos_adapters_excel_reporte_inventario_nombres_de_hoja` |
| Notificación | Resultado | Resumen + aviso si la BD no es de hoy | Correo con adjunto | `tasks/actualizacion_inventario.py:L122` · `tasks_actualizacion_inventario_actualizacioninventario_notify_success` |

## 2. Ventas (`actualizacion_ventas`) — motor dual

> **Diagrama interactivo** (Archify, verificado contra `088f6f0`): [abrir en el índice](diagramas/main.html#05_ventas) · [abrir aparte](diagramas/03_flujo-ventas.html). El bloque Mermaid de abajo es la versión resumida para leer en GitHub/VS Code.

```mermaid
sequenceDiagram
    autonumber
    participant R as Orchestrator
    participant V as ActualizacionVentas
    participant FS as Sistema de archivos
    participant FV as fuentes_ventas (adaptador)
    participant TV as TransformarVentas
    participant L as _transformar_legacy (oráculo)
    participant X as Excel COM
    participant XC as core.xlsx_cleaner
    R->>V: BaseTask.run() → execute()
    V->>FS: crea PROCESANDO.txt
    V->>FS: localizar_archivos(): plantilla Pruebas/$2026 VENTAS_Actualizacion.xlsx,<br/>último INVENTARIO GENERAL ACTUALIZADO, MYR, matriz, _268 (hoy o último modificado)
    V->>FV: cargar_plantilla() con VENTAS_ACTUALIZACION_PASSWORD (fallback EXCEL_PASSWORD)
    V->>FS: backup atómico en Backups_Ventas (si falla → se detiene)
    V->>FV: preparar_entradas() → EntradasVentas
    alt VENTAS_MOTOR = reglas
        V->>TV: TransformarVentas(ventas.yaml, hoy)(entradas)
        TV-->>V: df + AuditTrail (21 pasos)
    else legacy (defecto)
        V->>L: pasos 2–9
        L-->>V: df
    end
    alt HAS_COM
        V->>X: Paso 10 com_write_df_into_template → _tmp_actualizacion_ventas.xlsx<br/>(fórmulas fill-down, Resum Mes L20/L21/L23:L27, subtotales, clave del área)
        V->>X: actualizar_tabla_dinamica_resumen_dia (PivotCache compartido)
        V->>XC: reducir_tamano_archivo (dibujos y partes huérfanas)
        V->>FS: os.replace(tmp → $2026 VENTAS_Actualizacion.xlsx)
        V->>X: guardar_copia_xlsb → $2026 VENTAS_«AAAA-MM-DD».xlsb (EXCEL_PASSWORD)
    else sin COM
        Note over V: Paso 10 y post-proceso omitidos — la tarea termina en éxito (riesgo)
    end
    V->>FS: elimina PROCESANDO.txt (finally)
    V-->>R: True/False → un solo correo (BaseTask.run)
```

| Etapa | Entrada | Transformación / validación | Salida | Evidencia |
|---|---|---|---|---|
| Indicador | — | `PROCESANDO.txt` en `BASE_PATH` mientras corre | archivo indicador | `tasks/actualizacion_ventas.py:L322` · `tasks_actualizacion_ventas_actualizacionventas_crear_indicador_progreso` |
| Localización | Carpetas `BASE_PATH`, `DB_EXPORT_DIR` | `_268` de hoy o último modificado; inventario actualizado más reciente | 5 rutas | `tasks/actualizacion_ventas.py:L2176` · `tasks_actualizacion_ventas_actualizacionventas_localizar_archivos`; L399 `..._find_informe_facturas_by_prefix`; L500 `..._find_inventario_actualizado` |
| Plantilla | Libro del área | Descifra con clave del área; transición desde la clave anterior | stream + columnas | `tasks/actualizacion_ventas.py:L202` · `tasks_actualizacion_ventas_actualizacionventas_abrir_actualizacion` |
| Respaldo | Plantilla | Copia a temporal + `os.replace`; falla → `RuntimeError` | `Backups_Ventas/$2026 VENTAS_Actualizacion.xlsx` | `tasks/actualizacion_ventas.py:L2553` · `tasks_actualizacion_ventas_actualizacionventas_ejecutar` |
| Entradas | Informe, inventario, MYR (.xlsb con `pyxlsb`), matriz, licitados | Normalización de encabezados; licitados en formato largo con precio > 0 | `EntradasVentas` | `tasks/actualizacion_ventas.py:L2218` · `tasks_actualizacion_ventas_actualizacionventas_preparar_entradas`; `src_insumos_adapters_excel_fuentes_ventas_leer_precios_licitados` |
| Transformación | `EntradasVentas` | Motor según `VENTAS_MOTOR` | df alineado a la plantilla | `tasks/actualizacion_ventas.py:L2260` · `tasks_actualizacion_ventas_actualizacionventas_transformar` |
| Paso 10 | df + plantilla | Escritura COM vectorizada, fórmulas, periodo del mes y festivos, subtotales, contraseña del área | `_tmp_actualizacion_ventas.xlsx` | `tasks/actualizacion_ventas.py:L1286` · `tasks_actualizacion_ventas_actualizacionventas_com_write_df_into_template`; L1164 `..._actualizar_periodo_mes`; L1204 `..._calcular_y_escribir_subtotales` |
| Post-proceso | Temporal | Tablas dinámicas → reducción de tamaño → reemplazo atómico → `.xlsb` | Libro final + copia | Hiperarista `he_paso10_postproceso` |

## 3. Envío del informe (`envio_informe_ventas`) — legacy

> **Diagrama interactivo** (Archify, verificado contra `088f6f0`): [abrir en el índice](diagramas/main.html#06_envio) · [abrir aparte](diagramas/03_flujo-envio.html). El bloque Mermaid de abajo es la versión resumida para leer en GitHub/VS Code.

```mermaid
sequenceDiagram
    autonumber
    participant R as Orchestrator
    participant E as EnvioInformeVentas
    participant X as Excel COM (visible)
    participant CB as Portapapeles (win32clipboard/PIL)
    participant CD as chromedriver_utils
    participant W as WhatsAppWeb (Selenium)
    participant M as EmailNotifier
    R->>E: BaseTask.run() → setup() (destinatarios WHATSAPP_CHATS)
    E->>X: abre $2026 VENTAS_Actualizacion.xlsx (clave del área)
    loop 4 capturas (CAPTURAS)
        E->>X: filtra pivote al mes / oculta días previos / resuelve hoja
        X->>CB: CopyPicture del rango
        CB-->>E: PNG en Img informe/
    end
    E->>X: extraer_factor_mg_neto
    E->>X: cierra solo su instancia (PID)
    E->>CD: obtener_chromedriver_path (compatible con Chrome)
    E->>W: iniciar_sesion() (perfil persistente, hasta 3 intentos; si fallan, revalida chromedriver y 3 más)
    loop por destinatario
        W->>W: buscar_chat → enviar_texto("Cordial saludo... VTAS «MES» MG NETO PONDERADO «factor»")
        W->>W: enviar_imagen × 4 (Ctrl+V)
    end
    alt exitosos == 0
        E-->>R: RuntimeError → correo [ERROR]
    else ≥ 1
        E->>M: notify_success (resumen + log)
    end
    E->>W: teardown → cerrar() (solo Chrome del perfil)
```

| Etapa | Entrada | Validación | Salida | Evidencia |
|---|---|---|---|---|
| Capturas | Libro de ventas | Omite captura si no existe la hoja; ≥ 1 imagen válida | `foto1..foto4 *.png` | `tasks/envio_informe_ventas.py:L1427` · `tasks_envio_informe_ventas_capturar_multiples_rangos` |
| Factor | Hoja de resumen | — | "MG NETO PONDERADO" | `tasks/envio_informe_ventas.py:L1234` · `tasks_envio_informe_ventas_extraer_factor_mg_neto` |
| Sesión | Perfil Chrome en estado operativo | Reintento tras revalidar ChromeDriver | sesión WhatsApp | `tasks/envio_informe_ventas.py:L442` · `tasks_envio_informe_ventas_whatsappweb_iniciar_sesion` |
| Envío | Chats | `exitosos == 0` → fallo | mensajes | `tasks/envio_informe_ventas.py:L957` · `tasks_envio_informe_ventas_whatsappweb_enviar_reporte_completo` |
| Cierre | PIDs | Solo Chrome del `user-data-dir` y Excel propio | — | `tasks/envio_informe_ventas.py:L277` · `tasks_envio_informe_ventas_cerrar_chrome_del_perfil`; L1416 `tasks_envio_informe_ventas_pid_de_excel` |

Capturas configuradas (`tasks/envio_informe_ventas.py:L118`): `Resumen Dia` (VTA DIA, dinámica), `Resum Mes` (RESUMEN MES), `Resumen Meta 2026` (A1:R14), `Resumen XDiaVend INF` (RESUMEN DIA POR VENDEDOR, día actual).
