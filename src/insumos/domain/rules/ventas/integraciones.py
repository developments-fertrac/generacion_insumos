"""Reglas que cruzan las ventas con las tablas de referencia.

Las tablas llegan ya leidas en ``ctx.tablas`` (las arma el adaptador):

- ``inventario``         ``$2026 INVENTARIO GENERAL ACTUALIZADO`` hoja INVENTARIO
- ``myr``                ``$2026 INVENTARIO MYR EXISTENCIA`` hoja COSTOS INV FINAL
- ``matriz_clientes``    ``MATRIZ COMPLETA DE CLIENTES`` hoja CLIENTES GENERAL
- ``precios_licitados``  hoja PRECIO UNIT LICITADOS en formato largo
                         (``REFERENCIA``, ``NIT``, ``PRECIO``); los NIT
                         licitados van en ``ctx.params["nits_licitados"]``.

Port fiel de los pasos 4 a 7.5 de ``ActualizacionVentas._ejecutar``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd
from unidecode import unidecode

from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult, marcar_eliminadas
from insumos.domain.rules.registry import rule
from insumos.domain.rules.ventas import columnas as C


def _tabla(ctx: RuleContext, nombre: str) -> pd.DataFrame | None:
    tabla = ctx.tablas.get(nombre)
    if tabla is None or len(tabla.columns) == 0:
        return None
    return tabla


@rule("ven.integrar_linea_sublinea")
@dataclass(frozen=True)
class IntegrarLineaSublinea(ReglaBase):
    """LINEA, SUBLINEA y LIDER LINEA desde el inventario general actualizado.

    Ambas referencias se normalizan (``"4591.0"`` -> ``"4591"``) antes del cruce.
    """

    description = "Trae LINEA / SUBLINEA / LIDER LINEA del inventario general por REFERENCIA"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        inv = _tabla(ctx, "inventario")
        if inv is None:
            raise ValueError("Falta la tabla 'inventario' en el contexto")
        sinonimos = {"LINEA COPIA": ("LINEA COPIA", "LINEA"), "SUB-LINEA COPIA": ("SUB-LINEA COPIA", "SUBLINEA", "SUB-LINEA")}
        inv = inv.copy()
        for destino, candidatos in sinonimos.items():
            if destino not in inv.columns:
                for cand in candidatos:
                    if cand in inv.columns:
                        inv = inv.rename(columns={cand: destino})
                        break
        inv = inv.rename(columns={"LINEA COPIA": C.LINEA, "SUB-LINEA COPIA": C.SUBLINEA})
        columnas = [C.REFERENCIA, C.LINEA, C.SUBLINEA] + ([C.LIDER_LINEA] if C.LIDER_LINEA in inv.columns else [])
        inv = inv[columnas].drop_duplicates(C.REFERENCIA, keep="last").copy()
        inv[C.REFERENCIA] = inv[C.REFERENCIA].apply(C.normalizar_referencia)

        df = df.copy()
        if C.REFERENCIA in df.columns:
            df[C.REFERENCIA] = df[C.REFERENCIA].apply(C.normalizar_referencia)
        df = df.merge(inv, on=C.REFERENCIA, how="left", suffixes=("", "_inv"))
        for c in (C.LINEA, C.SUBLINEA, C.LIDER_LINEA):
            if c + "_inv" in df.columns:
                df[c] = df[c + "_inv"] if c not in df.columns else df[c].fillna(df[c + "_inv"])
                df = df.drop(columns=[c + "_inv"])

        resultado = RuleResult(df=df, metrics={
            c: int(df[c].notna().sum()) for c in (C.LINEA, C.SUBLINEA, C.LIDER_LINEA) if c in df.columns
        })
        sin_linea = len(df) - resultado.metrics.get(C.LINEA, 0)
        if sin_linea:
            resultado.con_advertencia(f"{sin_linea} ventas sin LINEA en el inventario")
        return resultado


@rule("ven.integrar_costo_factor_hoy")
@dataclass(frozen=True)
class IntegrarCostoFactorHoy(ReglaBase):
    """COSTO FACTOR HOY desde el MYR de existencia mas reciente."""

    description = "Trae COSTO FACTOR HOY del MYR por REFERENCIA"
    columnas_referencia: tuple[str, ...] = ("REFERENCIA FERTRAC", "REFERENCIA", "REF", "CODIGO")
    columnas_costo: tuple[str, ...] = ("COSTO FACTOR HOY", "COSTO FACTOR DOLAR HOY", "COSTO FACTOR USD HOY")

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        myr = _tabla(ctx, "myr")
        if myr is None:
            return RuleResult(df=df).con_advertencia("MYR no disponible: COSTO FACTOR HOY queda vacio")
        ref = next((c for c in self.columnas_referencia if c in myr.columns), None)
        if ref is None:
            return RuleResult(df=df).con_advertencia("No se encontro columna de REFERENCIA en MYR")
        costo = next((c for c in self.columnas_costo if c in myr.columns), None)
        if costo is None:
            return RuleResult(df=df).con_advertencia("No se encontro columna de COSTO FACTOR HOY en MYR")

        tabla = myr[[ref, costo]].rename(columns={ref: C.REFERENCIA, costo: C.COSTO_FACTOR_HOY})
        tabla[C.REFERENCIA] = tabla[C.REFERENCIA].apply(C.normalizar_referencia)
        tabla = tabla.drop_duplicates(C.REFERENCIA, keep="last")
        tabla = tabla[tabla[C.REFERENCIA].notna()]

        df = df.copy()
        df[C.REFERENCIA] = df[C.REFERENCIA].apply(C.normalizar_referencia)
        df = df.merge(tabla, on=C.REFERENCIA, how="left", suffixes=("", "_myr"))
        sufijo = C.COSTO_FACTOR_HOY + "_myr"
        if sufijo in df.columns:
            df[C.COSTO_FACTOR_HOY] = df[C.COSTO_FACTOR_HOY].fillna(df[sufijo])
            df = df.drop(columns=[sufijo])
        con_costo = int(df[C.COSTO_FACTOR_HOY].notna().sum()) if C.COSTO_FACTOR_HOY in df.columns else 0
        return RuleResult(df=df, metrics={"con_costo_factor_hoy": con_costo})


def descuento_a_decimal(valor: object) -> float:
    """Texto libre de la matriz de clientes -> fraccion (``"8%"`` -> 0.08).

    ``"SIN DTO"`` es 0; ``"5% maximo 60 dias"`` es 0.05; numeros mayores a 1 se
    leen como porcentaje.
    """
    if valor is None:
        return 0
    try:
        if pd.isna(valor):
            return 0
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (int, float)):
        if 0 <= valor <= 1:
            return valor
        return valor / 100 if valor > 1 else 0
    texto = str(valor).strip()
    if texto == "" or texto.upper() == "NAN":
        return 0
    norm = unidecode(texto).upper().strip()
    if "SIN DTO" in norm or "SIN DESCUENTO" in norm:
        return 0
    if "5%" in norm and ("MAXIMO" in norm or "MAX" in norm) and "60" in norm:
        return 0.05
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", texto)  # "12,5%" y "12.5%" -> 0.125
    if m:
        return float(m.group(1).replace(",", ".")) / 100
    try:
        num = float(texto.replace("%", "").replace(",", ".").strip())
    except ValueError:
        return 0
    if 0 <= num <= 1:
        return num
    return num / 100 if num > 1 else 0


@rule("ven.integrar_dcto_condicionado")
@dataclass(frozen=True)
class IntegrarDctoCondicionado(ReglaBase):
    """DCTO CONDICIONADO del cliente (solo TIPO DESCUENTO = CONDICIONADO); 0 si no tiene."""

    description = "Trae DCTO CONDICIONADO de la matriz de clientes por NIT"
    tipo: str = "CONDICIONADO"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        mat = _tabla(ctx, "matriz_clientes")
        if mat is None:
            return RuleResult(df=df).con_advertencia("Matriz de clientes no disponible: se salta DCTO CONDICIONADO")
        mat = mat.copy()
        if "TIPO DESCUENTO" in mat.columns:
            mat = mat[mat["TIPO DESCUENTO"].astype(str).str.strip().str.upper() == self.tipo.upper()].copy()
        if "DESCUENTO" in mat.columns:
            mat["DESCUENTO"] = mat["DESCUENTO"].apply(descuento_a_decimal)

        if not (C.NIT_CLIENTE in df.columns and "NIT" in mat.columns and "DESCUENTO" in mat.columns):
            return RuleResult(df=df).con_advertencia("Faltan NIT CLIENTE / NIT / DESCUENTO: se salta DCTO CONDICIONADO")

        tabla = mat[["NIT", "DESCUENTO"]].drop_duplicates("NIT", keep="last")
        tabla = tabla[tabla["NIT"].notna()].copy()
        tabla["NIT"] = tabla["NIT"].astype(str).str.strip().replace({"nan": None, "None": None, "": None})
        tabla["NIT"] = tabla["NIT"].str.replace(r"\.0$", "", regex=True)

        df = df.copy()
        nit = df[C.NIT_CLIENTE].astype(str).str.strip().replace({"nan": None, "None": None, "": None})
        df[C.NIT_CLIENTE] = nit.str.replace(r"\.0$", "", regex=True)
        df = df.merge(tabla, left_on=C.NIT_CLIENTE, right_on="NIT", how="left", suffixes=("", "_mat"))
        df = df.rename(columns={"DESCUENTO": C.DCTO_CONDICIONADO})
        con_dcto = int(df[C.DCTO_CONDICIONADO].notna().sum())
        df[C.DCTO_CONDICIONADO] = df[C.DCTO_CONDICIONADO].fillna(0)
        if "NIT" in df.columns:
            df = df.drop(columns=["NIT"])
        return RuleResult(df=df, metrics={"clientes_condicionados": len(tabla), "ventas_con_dcto": con_dcto})


@rule("ven.eliminar_referencias_invalidas")
@dataclass(frozen=True)
class EliminarReferenciasInvalidas(ReglaBase):
    """Sin LINEA y con referencia que no parece codigo (texto largo): no es producto."""

    description = "Elimina ventas sin LINEA cuya referencia no es un codigo valido"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if C.LINEA not in df.columns or C.REFERENCIA not in df.columns:
            return RuleResult(df=df)
        refs = df[C.REFERENCIA].astype(str).str.strip()
        fuera = (df[C.LINEA].fillna(0) == 0) & (~refs.apply(C.referencia_valida))
        removidas = marcar_eliminadas(df[fuera], "Sin LINEA y referencia invalida", self.id)
        return RuleResult(df=df[~fuera].copy(), removed=removidas)


def _pct_a_numero(valor: object) -> float:
    try:
        if pd.isna(valor):
            return 0.0
    except (TypeError, ValueError):
        pass
    try:
        return float(str(valor).replace("%", "").replace(",", ".").strip())
    except ValueError:
        return 0.0


@rule("ven.normalizar_dctos")
@dataclass(frozen=True)
class NormalizarDctos(ReglaBase):
    """Si la factura ya trae descuento comercial, el condicionado no aplica (queda 0%)."""

    description = "DCTO CONDICIONADO = 0% cuando la venta ya tiene VR DESCUENTO"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        df = df.copy()
        if C.PORC_PIE_FACTURA not in df.columns:
            df[C.PORC_PIE_FACTURA] = ""
        if C.DCTO_CONDICIONADO not in df.columns:
            return RuleResult(df=df)
        dcto = df[C.DCTO_CONDICIONADO].apply(_pct_a_numero)
        if C.VR_DESCUENTO in df.columns:
            otro = pd.to_numeric(df[C.VR_DESCUENTO], errors="coerce").fillna(0)
        else:
            otro = df[C.PORC_PIE_FACTURA].apply(_pct_a_numero)
        mascara = (dcto != 0) & (otro != 0)
        if mascara.any():
            # Mezcla numeros y "0%": columna object (pandas 2 lo hacia solo; pandas 3 lo exige).
            df[C.DCTO_CONDICIONADO] = df[C.DCTO_CONDICIONADO].astype(object)
        df.loc[mascara, C.DCTO_CONDICIONADO] = "0%"
        return RuleResult(df=df, modified=int(mascara.sum()))


@rule("ven.vta_acordada_licitado")
@dataclass(frozen=True)
class VtaAcordadaLicitado(ReglaBase):
    """Precio licitado por (REFERENCIA, NIT) de la hoja PRECIO UNIT LICITADOS; 0 si no aplica."""

    description = "Calcula VTA ACORDADA X UNIDAD LICITADO por referencia y NIT licitado"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        precios = ctx.tablas.get("precios_licitados")
        nits = set(ctx.params.get("nits_licitados") or ())
        if precios is None or not nits:
            # El proceso anterior registraba el error y seguia sin la columna.
            return RuleResult(df=df).con_advertencia(
                "Hoja PRECIO UNIT LICITADOS no disponible: no se calcula VTA ACORDADA"
            )
        faltan = [c for c in (C.NIT_CLIENTE, C.REFERENCIA) if c not in df.columns]
        if faltan:
            return RuleResult(df=df).con_advertencia(f"Faltan {faltan}: no se calcula VTA ACORDADA")

        tabla = precios.rename(columns={"REFERENCIA": "_REF", "NIT": "_NIT", "PRECIO": "_PRECIO"})
        df = df.copy()
        df["_NIT_CLEAN"] = C.limpiar_nit(df[C.NIT_CLIENTE])
        df["_REF_CLEAN"] = df[C.REFERENCIA].apply(C.normalizar_referencia)
        df = df.merge(tabla, left_on=["_REF_CLEAN", "_NIT_CLEAN"], right_on=["_REF", "_NIT"], how="left")
        licitado = df["_NIT_CLEAN"].isin(nits)
        df[C.VTA_ACORDADA] = 0
        df.loc[licitado, C.VTA_ACORDADA] = df.loc[licitado, "_PRECIO"].fillna(0)
        encontrados = int((df["_PRECIO"].notna() & licitado).sum())
        df = df.drop(columns=["_NIT_CLEAN", "_REF_CLEAN", "_REF", "_NIT", "_PRECIO"], errors="ignore")
        return RuleResult(df=df, metrics={"con_precio_licitado": encontrados, "nits_licitados": len(nits)})
