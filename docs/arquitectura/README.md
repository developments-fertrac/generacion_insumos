# Documentación de arquitectura — Generación de Insumos

Generada el 2026-10-05 sobre la rama `main` (commit `088f6f0`) a partir de `graphify-out/graph.json` y `graphify-out/GRAPH_REPORT.md` (1.389 nodos, 3.263 aristas, 69 comunidades, 91 % EXTRACTED; graphify 0.9.77 oficial).

| Documento | Contenido |
|---|---|
| [00_resumen_ejecutivo.md](00_resumen_ejecutivo.md) | Propósito, usuarios, valor operativo, estado de la migración |
| [01_contexto_C4.md](01_contexto_C4.md) | C4 nivel 1 y 2 |
| [02_arquitectura_hexagonal.md](02_arquitectura_hexagonal.md) | Dominio, casos de uso, puertos→adaptadores, violaciones y desviaciones |
| [03_flujos_de_datos.md](03_flujos_de_datos.md) | Secuencias de inventario, ventas y envío |
| [04_motor_de_reglas_ventas.md](04_motor_de_reglas_ventas.md) | Catálogo de las 21 entradas `ven.*`, paridad y criterio de corte |
| [05_reglas_de_negocio.md](05_reglas_de_negocio.md) | Decisiones aprobadas y trazabilidad al código |
| [06_decisiones_ADR.md](06_decisiones_ADR.md) | Índice y detalle de ADR 0001–0010 |
| [07_operacion_y_despliegue.md](07_operacion_y_despliegue.md) | `.env`, comandos `uv`, publicación, n8n, logs, troubleshooting |
| [08_calidad_y_pruebas.md](08_calidad_y_pruebas.md) | Estrategia, cobertura estructural, paridad |
| [09_riesgos_y_deuda_tecnica.md](09_riesgos_y_deuda_tecnica.md) | Matriz de riesgos y deuda por acoplamiento |
| [10_roadmap.md](10_roadmap.md) | Higiene, Fases 4, 4b, 5 y 6 con criterios de salida |
| [11_brechas_de_verificacion.md](11_brechas_de_verificacion.md) | Lo que el grafo no pudo confirmar |
| [diagramas/main.html](diagramas/main.html) | Índice navegable de los 7 diagramas generados con Archify (HTML + fuente `.archify.json`), verificados contra el código |

Convenciones: `archivo:Lnn` · nodo `id` del grafo; **[NO VERIFICADO]** = no respaldado por el grafo; diagramas en Mermaid. Para explorar el grafo: abrir `graphify-out/graph.html` (requiere acceso a `unpkg.com` para la librería vis-network).
