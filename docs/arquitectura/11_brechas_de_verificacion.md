# 11 · Brechas: lo que el grafo no pudo confirmar

> Lista para validar con el equipo. Cada ítem indica por qué el grafo no lo confirma y quién debería validarlo.

| # | Brecha | Por qué no se confirma | Dónde aparece | Validar con |
|---|---|---|---|---|
| B1 | Contenido real de los `.bat` de producción (`PATH_UV`, `--frozen`, `nopause`, bitácora, `resultado_<tarea>.txt`) | No están en el repositorio; solo `test.bat` (diagnóstico) | 01, 07, 10 | Operación / responsable de n8n |
| B2 | Flujo de n8n y nombres de tareas programadas (p. ej. `FertracActualizacionInventario`) | Solo descritos en el docx | 01, 07 | Operación |
| B3 | Contexto, fecha y alcance de ADR 0001–0004 | Nunca tuvieron archivo; el "porqué" se reconstruyó de docx y docstrings | 06 | Arquitectura / Laura |
| B4 | Vigencia de ADR 0005–0010 tras su eliminación en `86f7093` (¿se eliminaron por obsoletos o por error?) | El commit dice "documentación obsoleta" pero README y código los citan como vigentes | 06, 09 R8 | Laura |
| B5 | Plan completo de fases (`docs/Old/Plan_Arquitectura_Hexagonal_*.docx`) | Eliminado del árbol | 10 | Laura |
| B6 | Cobertura de líneas real | No hay reporte de `pytest-cov`; no se pudieron ejecutar las pruebas en esta sesión (sin red para dependencias) | 08 | Ejecutar `uv run pytest --cov` en Windows |
| B7 | Que las 220 pruebas pasen en la rama actual (`088f6f0`) | Dato del docx/ADR 0010 (2026-10-05), previo a los dos últimos commits | 08 | CI / ejecución local |
| B8 | Existencia de CI y protección de rama (gobierno de YAML con aprobación de negocio) | No hay workflows ni CODEOWNERS en el repo | 02, 05, 10 | Plataforma / Laura |
| B9 | Configuración actual del `.env` de producción (`VENTAS_ACTUALIZACION_PASSWORD`, `DB_EXPORT_DIR`, `VENTAS_MOTOR`) | El grafo solo ve `.env.example` y el código | 07, 10 H2 | Operación |
| B10 | Presencia de `uv.lock` en la carpeta de producción | El lock se borró en desarrollo; la publicación no lo elimina en destino | 07, 09 R3 | Operación |
| B11 | Resultado de `scripts/comparar_ventas.py` con datos reales | No hay evidencia de ejecución en el repositorio | 04, 10 §1 | Área de Datos |
| B12 | Firma formal de las decisiones de negocio N1–N14 | Registradas en ADR/estado del proyecto, sin acta | 05 | Control de Ventas / Inventarios |
| B13 | Si `FERTRAC_PASS` sigue definido en algún `.env` | Ningún módulo la lee; solo la lista de secretos | 07 | Operación |
| B14 | Comportamiento de `com_write_df_into_template` frente a cambios de columnas de la plantilla | ~490 líneas COM sin prueba (solo simulación del escritor de inventario) | 03, 08 | Pruebas en Windows con plantilla real |
| B15 | Rutas exactas de producción (`BASE_PATH`, carpeta de publicación) | Se omitieron a propósito (sin rutas absolutas de usuario) | 07 | Operación |
| B16 | Método de generación del grafo y los diagramas | `graphifyy` no se pudo instalar (PyPI/GitHub bloqueados); el grafo salió de un extractor compatible con el mismo esquema. **Diagramas: resuelto 2026-10-05** — los 7 diagramas se regeneraron con Archify v3.0.1 en `diagramas/`, con evidencia `SRC` anclada al commit `088f6f0` y finalize (validate · deliver · check · browser-check) en verde | todos | Pendiente solo el grafo: regenerar con graphify oficial cuando haya red y comparar `GRAPH_REPORT.md` |
