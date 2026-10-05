# ADR 0009 — Ventas: pipeline de reglas con motor dual hasta probar paridad

Estado: Aceptada (2026-10-02)

## Contexto

`tasks/actualizacion_ventas.py` (2.700 lineas) mezcla en `_ejecutar` la
lectura de 5 archivos, 10 pasos de transformacion y la escritura COM sobre una
plantilla que el area de control protege con contrasena propia. Un error en la
transformacion llega al informe que ven las areas comerciales.

## Decision

1. Los pasos 2 a 9 (informe -> datos de VENTAS 2026) se reescriben como 21
   reglas `ven.*` declaradas en `config/rules/ventas.yaml`, con el caso de uso
   `TransformarVentas` y el puerto `EntradasVentas`.
2. La lectura de inventario, MYR, matriz y precios licitados pasa al adaptador
   `insumos.adapters.excel.fuentes_ventas`. Ambos motores reciben las mismas
   entradas.
3. El codigo anterior se conserva intacto como `_transformar_legacy` y es el
   **oraculo de paridad**: `tests/unit/rules/test_paridad_ventas.py` (datos
   sinteticos con todas las ramas) y `scripts/comparar_ventas.py` (datos
   reales, sin escribir nada).
4. `VENTAS_MOTOR` en `.env` elige el motor: `legacy` (defecto) o `reglas`. Se
   cambia a `reglas` solo cuando `comparar_ventas.py` da paridad total con datos
   reales; tras dos semanas sin diferencias se borra `_transformar_legacy`.
5. La escritura (paso 10: formulas, Resum Mes L20/L21/L23+, tablas dinamicas,
   contrasena de VENTAS_Actualizacion) no cambia en esta fase.

## Ajustes de negocio aprobados por el area (2026-10-02)

Aplicados en ambos motores (legacy y reglas), que usan las mismas funciones
de `insumos.domain.rules.ventas`:

- **La referencia es texto.** `"0123"` sigue `"0123"` en el informe, el cruce
  con el inventario, el MYR y los licitados; ya no se convierte a numero. Solo
  una celda numerica entera pierde el `.0` de pandas (4591.0 -> `"4591"`).
- **PRECIO UNIT LICITADOS:** referencia o NIT que llega como numero se
  redondea al entero (10.05 -> `"10"`, 10.5 -> `"11"`); si llega como texto se
  respeta. Antes se borraban todos los `.0` (`"10.05"` -> `"105"`).
- **Matriz de clientes:** `"12,5%"` se lee 0.125 (antes 0.05).

Se conserva: el informe trae una columna `NIT` que el cruce con la matriz
renombra a `NIT_mat` y descarta; no llega a la plantilla.

## Consecuencias

- Las eliminaciones (notas credito, fletes, publicidad, otros anos, sin
  referencia, referencias invalidas) quedan auditadas por regla.
- Mientras exista el motor dual hay dos implementaciones: el defecto `legacy`
  evita que un error del pipeline nuevo llegue a produccion.
