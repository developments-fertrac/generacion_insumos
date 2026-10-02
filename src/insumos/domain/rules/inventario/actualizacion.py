"""Reglas que actualizan la plantilla de inventario general con la base de datos.

Flujo (``config/rules/inventario.yaml``):

1. ``inv.eliminar_referencias_duplicadas``   la plantilla trae duplicados historicos
2. ``inv.eliminar_ausentes_en_bd``           la base de datos decide que entra
3. ``inv.actualizar_desde_bd``               existencia, costo, total y columnas "sistema"
4. ``inv.agregar_referencias_nuevas_bd``     referencias INVENTARIO que la plantilla no tiene
5. ``inv.calcular_nombre_lista_y_myr``       NOMBRE LISTA (Matriz USD) y NOMBRE MYR
6. ``inv.completar_campos_faltantes``        copia, bodega gerencia, lider y clasificacion
7. ``inv.costo_cero_sin_existencia``         existencia 0 -> costo 0
8. ``inv.ordenar_por_total_inv``             mayor a menor TOTAL INV
9. ``inv.validar_salida``                    controles de cuadre antes de escribir

La base de datos llega a las reglas por ``ctx.inventario_bd``, ya filtrada por
el pipeline ``inventario_bd``. Ninguna regla abre archivos.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any

import pandas as pd

from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult, marcar_eliminadas
from insumos.domain.rules.inventario import columnas as C
from insumos.domain.rules.registry import rule

# Columnas de la plantilla que se reemplazan con el valor de la base de datos.
_DESDE_BD: tuple[tuple[str, str], ...] = (
    (C.EXISTENCIA, C.BD_EXISTENCIA_NETA),
    (C.COSTO_PROMEDIO, C.BD_COSTO),
    (C.TOTAL_INV, C.BD_TOTAL_INV),
    (C.NOMBRE_ODOO, C.BD_NOMBRE_ODOO),
    (C.MARCA_SISTEMA, C.BD_MARCA),
    (C.LINEA_SISTEMA, C.BD_LINEA),
    (C.SUBLINEA_SISTEMA, C.BD_SUBLINEA),
)
_NUMERICAS = frozenset({C.EXISTENCIA, C.COSTO_PROMEDIO, C.TOTAL_INV})


def _bd_por_clave(ctx: RuleContext) -> pd.DataFrame:
    bd = ctx.inventario_bd
    if bd is None or bd.empty:
        raise ValueError(
            "No hay inventario de base de datos en el contexto. El pipeline "
            "'inventario_bd' debe ejecutarse antes que 'inventario_general'."
        )
    indexada = bd.assign(**{C.CLAVE: C.claves(bd[C.BD_REFERENCIA])})
    indexada = indexada[indexada[C.CLAVE].ne("")]
    indexada = indexada.drop_duplicates(subset=C.CLAVE, keep="first")
    return indexada.set_index(C.CLAVE)


def _asegurar_columnas(df: pd.DataFrame, cols: tuple[str, ...]) -> pd.DataFrame:
    faltan = [c for c in cols if c not in df.columns]
    if not faltan:
        return df
    return df.assign(**dict.fromkeys(faltan, pd.NA))


def _texto_util(valor: Any, *, cero_es_vacio: bool) -> str:
    """Texto limpio, o "" si el valor cuenta como vacio."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(valor).strip()
    if s in ("", "None", "nan", "NaN", "<NA>"):
        return ""
    if cero_es_vacio and s in ("0", "0.0"):
        return ""
    return s


@rule("inv.eliminar_ausentes_en_bd")
@dataclass(frozen=True)
class EliminarAusentesEnBd(ReglaBase):
    """Quita de la plantilla las referencias que la base de datos no deja entrar.

    La base de datos es la fuente de verdad sobre que referencias forman el
    inventario. Si una referencia de la plantilla no esta entre las de
    ``MOTIVO = INVENTARIO*``, sale del informe. El motivo del reporte dice si
    la base de datos la excluyo explicitamente (con su MOTIVO) o si ya no
    existe en la exportacion.
    """

    description = "Elimina las referencias de la plantilla que no llegan como INVENTARIO en la base de datos"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        bd = _bd_por_clave(ctx)
        clave = C.claves(df[C.REFERENCIA])
        presente = clave.isin(bd.index)

        excluidas: Mapping[str, str] = ctx.params.get("motivos_excluidos_bd", {}) or {}
        fuera = df[~presente].copy()
        detalle = clave[~presente].map(
            lambda k: f"Base de datos: {excluidas[k]}" if k in excluidas else "No existe en la base de datos"
        )
        fuera["_detalle"] = detalle
        return RuleResult(
            df=df[presente].copy(),
            removed=marcar_eliminadas(fuera, "No llega como INVENTARIO en la base de datos", self.id),
            metrics={
                "eliminadas": int((~presente).sum()),
                "excluidas_por_bd": int(detalle.str.startswith("Base de datos").sum()),
                "inexistentes_en_bd": int(detalle.eq("No existe en la base de datos").sum()),
            },
        )


@rule("inv.actualizar_desde_bd")
@dataclass(frozen=True)
class ActualizarDesdeBd(ReglaBase):
    """Reemplaza existencia, costo, total y columnas "sistema" con la base de datos.

    - EXISTENCIA      = EXISTENCIA NETA
    - COSTO PROMEDIO  = COSTO
    - TOTAL INV       = TOTAL INV
    - NOMBRE ODOO, Marca/Linea/Sub-linea sistema = NOMBRE ODOO, MARCA, LINEA, SUB-LINEA

    Las columnas "COPIA", LIDER LINEA y CLASIFICACION son de curaduria manual
    y no se tocan aqui.
    """

    description = "Actualiza existencia, costo promedio, total y columnas sistema desde la base de datos"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        bd = _bd_por_clave(ctx)
        salida = _asegurar_columnas(df.copy(), tuple(destino for destino, _ in _DESDE_BD))
        clave = C.claves(salida[C.REFERENCIA])
        coincide = clave.isin(bd.index)

        cambios: dict[str, int] = {}
        for destino, origen in _DESDE_BD:
            nuevos = clave.map(bd[origen])
            if destino in _NUMERICAS:
                nuevos = pd.to_numeric(nuevos, errors="coerce")
                anteriores = pd.to_numeric(salida[destino], errors="coerce")
                distinto = coincide & ~((anteriores - nuevos).abs().fillna(1) < 1e-6)
                salida[destino] = anteriores.astype("float64")
            else:
                anteriores = salida[destino].map(lambda v: _texto_util(v, cero_es_vacio=False))
                nuevos_txt = nuevos.map(lambda v: _texto_util(v, cero_es_vacio=False))
                distinto = coincide & anteriores.ne(nuevos_txt)
                salida[destino] = salida[destino].astype("object")
                nuevos = nuevos.astype("object")
            salida.loc[coincide, destino] = nuevos[coincide]
            cambios[destino] = int(distinto.sum())

        return RuleResult(
            df=salida,
            modified=int(coincide.sum()),
            metrics={"actualizadas": int(coincide.sum()), "valores_cambiados": cambios},
        )


@rule("inv.agregar_referencias_nuevas_bd")
@dataclass(frozen=True)
class AgregarReferenciasNuevasBd(ReglaBase):
    """Agrega las referencias INVENTARIO de la base de datos que la plantilla no tiene.

    Las filas nuevas llevan los datos de la base de datos; NOMBRE LISTA,
    NOMBRE MYR, columnas COPIA, LIDER LINEA y CLASIFICACION los completan las
    reglas siguientes. La columna auxiliar ``_NUEVA`` las identifica en el
    reporte.
    """

    description = "Agrega las referencias de la base de datos que no estan en la plantilla"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        bd = _bd_por_clave(ctx)
        salida = df.copy()
        salida["_NUEVA"] = False
        existentes = set(C.claves(salida[C.REFERENCIA]))
        nuevas = bd[~bd.index.isin(existentes)]
        if nuevas.empty:
            return RuleResult(df=salida, metrics={"agregadas": 0})

        filas = pd.DataFrame(index=range(len(nuevas)), columns=list(salida.columns), dtype="object")
        filas[C.REFERENCIA] = nuevas[C.BD_REFERENCIA].astype("object").to_numpy()
        for destino, origen in _DESDE_BD:
            valores = nuevas[origen].to_numpy()
            filas[destino] = pd.to_numeric(valores, errors="coerce") if destino in _NUMERICAS else valores
        filas["_NUEVA"] = True

        combinado = pd.concat([salida, filas], ignore_index=True)
        return RuleResult(
            df=combinado,
            modified=len(nuevas),
            metrics={"agregadas": len(nuevas)},
        )


@rule("inv.calcular_nombre_lista_y_myr")
@dataclass(frozen=True)
class CalcularNombreListaYMyr(ReglaBase):
    """NOMBRE LISTA desde la Matriz USD y NOMBRE MYR con prioridad LISTA -> ODOO.

    Port de ``calcular_nombre_lista_y_myr`` (legacy), aplicado a todas las
    filas como hacia el legacy:

    - Con Matriz USD disponible: NOMBRE LISTA = descripcion de la matriz, o
      ``"0"`` si la referencia no esta. Sin matriz: se conserva lo que haya.
    - NOMBRE MYR = NOMBRE LISTA si es un nombre real (no vacio ni ``"0"``),
      si no NOMBRE ODOO. Si ambos estan vacios, se conserva el valor previo.

    ``ctx.matriz_usd`` debe traer las columnas ``CLAVE`` y ``DESCRIPCION``.
    """

    description = "Calcula NOMBRE LISTA desde la Matriz USD y NOMBRE MYR (LISTA, si no ODOO)"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        salida = _asegurar_columnas(df.copy(), (C.NOMBRE_LISTA, C.NOMBRE_ODOO, C.NOMBRE_MYR))
        salida[C.NOMBRE_LISTA] = salida[C.NOMBRE_LISTA].astype("object")
        salida[C.NOMBRE_MYR] = salida[C.NOMBRE_MYR].astype("object")
        clave = C.claves(salida[C.REFERENCIA])

        matriz = ctx.matriz_usd
        con_matriz = matriz is not None and not matriz.empty
        encontradas = 0
        if con_matriz:
            mapa = (
                matriz.assign(CLAVE=C.claves(matriz["CLAVE"]))
                .drop_duplicates(subset="CLAVE", keep="first")
                .set_index("CLAVE")["DESCRIPCION"]
                .map(lambda v: _texto_util(v, cero_es_vacio=False))
            )
            descripcion = clave.map(mapa)
            encontradas = int(descripcion.fillna("").ne("").sum())
            salida[C.NOMBRE_LISTA] = descripcion.where(descripcion.fillna("").ne(""), "0").astype("object")

        lista = salida[C.NOMBRE_LISTA].map(lambda v: _texto_util(v, cero_es_vacio=True))
        odoo = salida[C.NOMBRE_ODOO].map(lambda v: _texto_util(v, cero_es_vacio=True))
        myr = lista.where(lista.ne(""), odoo)
        hay_valor = myr.ne("")
        salida.loc[hay_valor, C.NOMBRE_MYR] = myr[hay_valor]

        return RuleResult(
            df=salida,
            modified=int(hay_valor.sum()),
            warnings=([] if con_matriz else ["Matriz USD no disponible: NOMBRE LISTA se conserva"]),
            metrics={
                "con_descripcion_en_matriz": encontradas,
                "nombre_myr_desde_lista": int(lista.ne("").sum()),
                "nombre_myr_desde_odoo": int((lista.eq("") & odoo.ne("")).sum()),
            },
        )


@rule("inv.completar_campos_faltantes")
@dataclass(frozen=True)
class CompletarCamposFaltantes(ReglaBase):
    """Completa los campos vacios sin pisar lo que ya tiene valor.

    Port de ``calcular_completado_inventario_copia`` (legacy):

    - NOMBRE MYR vacio o "0": el mas largo entre NOMBRE ODOO y NOMBRE LISTA.
    - MARCA / LINEA / SUB-LINEA COPIA vacias: Marca / Linea / Sub-linea sistema.
    - INV BODEGA GERENCIA vacio: 0.
    - LIDER LINEA / CLASIFICACION vacios: cruce de LINEA COPIA contra la
      Distribucion de matrices (``ctx.distribucion["gestor"]`` y
      ``["clasificacion"]``).
    """

    description = "Completa NOMBRE MYR, columnas COPIA, bodega gerencia, lider y clasificacion vacios"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        cols = (
            C.NOMBRE_MYR, C.NOMBRE_ODOO, C.NOMBRE_LISTA, C.MARCA_COPIA, C.LINEA_COPIA,
            C.SUBLINEA_COPIA, C.MARCA_SISTEMA, C.LINEA_SISTEMA, C.SUBLINEA_SISTEMA,
            C.INV_BODEGA_GERENCIA, C.LIDER_LINEA, C.CLASIFICACION,
        )
        salida = _asegurar_columnas(df.copy(), cols)
        completadas: dict[str, int] = {}

        def _txt(col: str, *, cero_es_vacio: bool = False) -> pd.Series:
            return salida[col].map(lambda v: _texto_util(v, cero_es_vacio=cero_es_vacio))

        def _rellenar(col: str, vacio: pd.Series, valores: pd.Series) -> None:
            aplicar = vacio & valores.ne("")
            if aplicar.any():
                salida[col] = salida[col].astype("object")
                salida.loc[aplicar, col] = valores[aplicar]
            completadas[col] = int(aplicar.sum())

        # NOMBRE MYR
        odoo = _txt(C.NOMBRE_ODOO, cero_es_vacio=True)
        lista = _txt(C.NOMBRE_LISTA, cero_es_vacio=True)
        mas_largo = pd.Series(
            [o if len(o) >= len(lst) else lst for o, lst in zip(odoo, lista, strict=True)],
            index=salida.index,
        )
        _rellenar(C.NOMBRE_MYR, _txt(C.NOMBRE_MYR, cero_es_vacio=True).eq(""), mas_largo)

        # Columnas COPIA desde sistema
        for copia, sistema in (
            (C.MARCA_COPIA, C.MARCA_SISTEMA),
            (C.LINEA_COPIA, C.LINEA_SISTEMA),
            (C.SUBLINEA_COPIA, C.SUBLINEA_SISTEMA),
        ):
            _rellenar(copia, _txt(copia).eq(""), _txt(sistema))

        # INV BODEGA GERENCIA
        vacio_bodega = _txt(C.INV_BODEGA_GERENCIA).eq("")
        if vacio_bodega.any():
            salida[C.INV_BODEGA_GERENCIA] = salida[C.INV_BODEGA_GERENCIA].astype("object")
            salida.loc[vacio_bodega, C.INV_BODEGA_GERENCIA] = 0
        completadas[C.INV_BODEGA_GERENCIA] = int(vacio_bodega.sum())

        # LIDER LINEA / CLASIFICACION por LINEA COPIA
        linea = _txt(C.LINEA_COPIA).str.upper()
        dist = ctx.distribucion or {}
        for col, llave in ((C.LIDER_LINEA, "gestor"), (C.CLASIFICACION, "clasificacion")):
            mapa: Mapping[str, str] = dist.get(llave, {}) or {}
            _rellenar(col, _txt(col).eq(""), linea.map(lambda k, m=mapa: m.get(k, "")))

        sin_lider = int(_txt(C.LIDER_LINEA).eq("").sum())
        return RuleResult(
            df=salida,
            modified=sum(completadas.values()),
            warnings=([f"{sin_lider} referencias sin LIDER LINEA (LINEA COPIA sin distribucion)"] if sin_lider else []),
            metrics={"completadas": completadas},
        )


@rule("inv.ordenar_por_total_inv")
@dataclass(frozen=True)
class OrdenarPorTotalInv(ReglaBase):
    """Ordena de mayor a menor TOTAL INV (instruccion de negocio del 2026-09-10).

    Se hace aqui y no en Excel: ordenar 15.000 filas por COM es lento y deja
    las formulas Dif apuntando a filas que ya no son las suyas.
    """

    description = "Ordena el inventario por TOTAL INV de mayor a menor"
    columna: str = C.TOTAL_INV

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        orden = pd.to_numeric(df[self.columna], errors="coerce").fillna(0)
        salida = df.assign(_orden=orden).sort_values(
            "_orden", ascending=False, kind="mergesort"
        ).drop(columns="_orden").reset_index(drop=True)
        return RuleResult(df=salida, metrics={"filas": len(salida)})


class CuadreFallido(ValueError):
    """La salida no cuadra con la base de datos: no se debe escribir."""


@rule("inv.validar_salida")
@dataclass(frozen=True)
class ValidarSalida(ReglaBase):
    """Controles de cuadre antes de escribir el archivo.

    Si alguno falla, el pipeline se detiene y no se genera un informe
    incorrecto que parezca correcto:

    - referencias unicas y no vacias;
    - mismas referencias que la base de datos filtrada;
    - suma de TOTAL INV igual a la de la base de datos (tolerancia en pesos);
    - TOTAL INV = EXISTENCIA x COSTO PROMEDIO fila a fila;
    - costo 0 donde la existencia es 0.
    """

    description = "Verifica unicidad, cobertura y cuadre de totales contra la base de datos"
    tolerancia_total: float = 1.0
    tolerancia_fila: float = 1.0

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        bd = _bd_por_clave(ctx)
        clave = C.claves(df[C.REFERENCIA])
        errores: list[str] = []

        if clave.eq("").any():
            errores.append(f"{int(clave.eq('').sum())} filas sin referencia")
        if clave.duplicated().any():
            errores.append(f"{int(clave.duplicated().sum())} referencias duplicadas")

        sobran = set(clave) - set(bd.index)
        faltan = set(bd.index) - set(clave)
        if sobran:
            errores.append(f"{len(sobran)} referencias que no estan en la base de datos")
        if faltan:
            errores.append(f"{len(faltan)} referencias de la base de datos que faltan")

        existencia = pd.to_numeric(df[C.EXISTENCIA], errors="coerce").fillna(0)
        costo = pd.to_numeric(df[C.COSTO_PROMEDIO], errors="coerce").fillna(0)
        total = pd.to_numeric(df[C.TOTAL_INV], errors="coerce").fillna(0)

        suma_salida = float(total.sum())
        suma_bd = float(
            pd.to_numeric(bd[C.BD_TOTAL_INV], errors="coerce").fillna(0)
            .where(pd.to_numeric(bd[C.BD_EXISTENCIA_NETA], errors="coerce").fillna(0).ne(0), 0)
            .sum()
        )
        if abs(suma_salida - suma_bd) > self.tolerancia_total:
            errores.append(f"TOTAL INV {suma_salida:,.2f} no cuadra con la base de datos {suma_bd:,.2f}")

        descuadre = (total - existencia * costo).abs() > self.tolerancia_fila
        if descuadre.any():
            errores.append(f"{int(descuadre.sum())} filas con TOTAL INV distinto de EXISTENCIA x COSTO")

        costo_sin_existencia = existencia.eq(0) & costo.ne(0)
        if costo_sin_existencia.any():
            errores.append(f"{int(costo_sin_existencia.sum())} filas con costo y existencia 0")

        if errores:
            raise CuadreFallido("La salida no cuadra; no se escribe el archivo: " + "; ".join(errores))

        return RuleResult(
            df=df,
            metrics={
                "referencias": len(df),
                "existencia_total": float(existencia.sum()),
                "total_inv": suma_salida,
                "total_inv_bd": suma_bd,
            },
        )
