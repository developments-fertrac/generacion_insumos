"""Reglas sobre el informe de ventas (``InformesDeVentas(Facturas)_268``).

Port fiel de ``ActualizacionVentas.transformar_informe_ventas`` y de los pasos
2 y 3 de ``_ejecutar`` (filtros FLETE/PUBLICIDAD, ano en curso, MES, mapeo
semantico). El orden en ``config/rules/ventas.yaml`` es el del proceso
anterior; la paridad la fija ``tests/unit/rules/test_paridad_ventas.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from insumos.domain.rules.base import ReglaBase, RuleContext, RuleResult, marcar_eliminadas
from insumos.domain.rules.registry import rule
from insumos.domain.rules.ventas import columnas as C


class InformeSinDatos(ValueError):
    """El informe no trae ventas del ano en curso o le faltan columnas clave."""


def _mover_despues(df: pd.DataFrame, columna: str, ancla: str) -> pd.DataFrame:
    cols = df.columns.tolist()
    cols.remove(columna)
    cols.insert(cols.index(ancla) + 1, columna)
    return df[cols]


@rule("ven.detectar_encabezado_informe")
@dataclass(frozen=True)
class DetectarEncabezadoInforme(ReglaBase):
    """Si el informe trae filas de titulo antes del encabezado, las salta."""

    description = "Ubica la fila de encabezados del informe (hasta 15 filas de titulo)"
    max_filas: int = 15

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        objetivo = set(C.COLUMNAS_INFORME)
        minimo = max(3, int(0.6 * len(objetivo)))
        presentes = objetivo & {str(c).strip() for c in df.columns}
        if C.INF_NIT in df.columns or len(presentes) >= minimo:
            return RuleResult(df=df)  # el encabezado ya esta en la fila 1
        for i in range(min(self.max_filas, len(df))):
            fila = {str(x).strip() for x in df.iloc[i].tolist()}
            if len(objetivo & fila) >= minimo:
                salida = df.iloc[i + 1:].reset_index(drop=True)
                salida.columns = df.iloc[i].tolist()
                return RuleResult(df=salida, metrics={"fila_encabezado": i + 2})
        return RuleResult(df=df).con_advertencia("No se encontro la fila de encabezados del informe")


@rule("ven.excluir_prefijos")
@dataclass(frozen=True)
class ExcluirPrefijos(ReglaBase):
    """Las notas credito (NC, NCDTO, NDCTO) no son venta."""

    description = "Excluye notas credito y notas debito por descuento (Prefijo NC/NCDTO/NDCTO)"
    prefijos: tuple[str, ...] = C.PREFIJOS_NOTA_CREDITO
    columna: str = C.INF_PREFIJO

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if self.columna not in df.columns:
            return RuleResult(df=df).con_advertencia(f"Columna '{self.columna}' no encontrada: no se filtran NC/NCDTO")
        fuera = df[self.columna].astype(str).str.strip().str.upper().isin([p.upper() for p in self.prefijos])
        removidas = marcar_eliminadas(df[fuera], f"Prefijo {'/'.join(sorted(self.prefijos))}", self.id)
        return RuleResult(df=df[~fuera].copy(), removed=removidas)


@rule("ven.estructurar_informe")
@dataclass(frozen=True)
class EstructurarInforme(ReglaBase):
    """Renombra el descuento comercial y reserva la columna de valor unitario."""

    description = "Prepara columnas del informe: VR DESCUENTO y Valor unitario"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        for col in df.columns:
            bajo = str(col).lower()
            if "valor" in bajo and "descuento" in bajo and "comercial" in bajo:
                df = df.rename(columns={col: C.VR_DESCUENTO})
                break
        for cand in ("Número documento", "Numero documento", "Nº documento"):
            if cand in df.columns:
                df = df.drop(columns=[cand])
        df = df.copy()
        if C.INF_VALOR_BRUTO in df.columns:
            pos = list(df.columns).index(C.INF_VALOR_BRUTO)
            df.insert(pos, "COL_TMP_1", np.nan)
            df.insert(pos + 1, C.INF_VALOR_UNITARIO, np.nan)
        else:
            for c in ("COL_TMP_1", C.INF_VALOR_UNITARIO):
                if c not in df.columns:
                    df[c] = np.nan
        return RuleResult(df=df)


@rule("ven.fecha_de_venta")
@dataclass(frozen=True)
class FechaDeVenta(ReglaBase):
    """FE originada en un pedido (PV) toma la fecha del documento origen; el resto, la fecha.

    Asi la venta cae en el mes del pedido, no en el de la factura electronica.
    """

    description = "Fecha de venta: 'Fecha documento origen' si es FE de un PV, si no 'Fecha'"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        df = df.copy()
        for col in (C.INF_FECHA_ORIGEN, C.INF_FECHA):
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce", dayfirst=True).dt.normalize()

        necesarias = (C.INF_PREFIJO, C.INF_DOC_ORIGEN, C.INF_FECHA_ORIGEN, C.INF_FECHA)
        resultado = RuleResult(df=df)
        if all(c in df.columns for c in necesarias):
            es_fe = df[C.INF_PREFIJO].astype(str).str.strip().str.upper() == "FE"
            es_pv = df[C.INF_DOC_ORIGEN].astype(str).str.strip().str.upper().str.startswith("PV")
            usar_origen = es_fe & es_pv
            temporal = "_FECHA_TEMPORAL_PROCESAMIENTO_"
            df[temporal] = df[C.INF_FECHA_ORIGEN].where(usar_origen, df[C.INF_FECHA])
            df = df.drop(columns=[C.INF_FECHA_ORIGEN, C.INF_FECHA, C.INF_DOC_ORIGEN])
            df = df.rename(columns={temporal: C.INF_FECHA_ORIGEN})
            resultado = RuleResult(df=df, metrics={
                "fecha_documento_origen": int(usar_origen.sum()),
                "fecha_factura": int((~usar_origen).sum()),
                "fecha_nula": int(df[C.INF_FECHA_ORIGEN].isna().sum()),
            })
            if resultado.metrics["fecha_nula"]:
                resultado.con_advertencia(f"{resultado.metrics['fecha_nula']} registros sin fecha")
        else:
            resultado.con_advertencia("Faltan columnas para elegir la fecha de venta; se deja 'Fecha documento origen'")

        if C.INF_FECHA_ORIGEN in df.columns and C.INF_NRO_DOC in df.columns:
            df = _mover_despues(df, C.INF_FECHA_ORIGEN, C.INF_NRO_DOC)
        resultado.df = df
        return resultado


@rule("ven.calcular_valor_unitario")
@dataclass(frozen=True)
class CalcularValorUnitario(ReglaBase):
    """Valor unitario = Valor bruto / Cantidad facturada; ordena columnas como la plantilla."""

    description = "Calcula Valor unitario y reubica Referencia, Cantidad y Costo unitario"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if C.INF_REFERENCIA in df.columns and C.INF_NUMERO in df.columns:
            df = _mover_despues(df, C.INF_REFERENCIA, C.INF_NUMERO)
        if C.INF_CANTIDAD in df.columns and C.INF_MARCA in df.columns:
            df = _mover_despues(df, C.INF_CANTIDAD, C.INF_MARCA)
        df = df.copy()
        if C.INF_VALOR_BRUTO in df.columns and C.INF_CANTIDAD in df.columns:
            with np.errstate(divide="ignore", invalid="ignore"):
                df[C.INF_VALOR_UNITARIO] = (
                    pd.to_numeric(df[C.INF_VALOR_BRUTO], errors="coerce")
                    / pd.to_numeric(df[C.INF_CANTIDAD], errors="coerce")
                )
        if C.INF_COSTO_UNITARIO in df.columns and C.INF_VALOR_BASE_PV in df.columns:
            cols = df.columns.tolist()
            cols.remove(C.INF_COSTO_UNITARIO)
            cols.insert(cols.index(C.INF_VALOR_BASE_PV), C.INF_COSTO_UNITARIO)
            df = df[cols]
        return RuleResult(df=df)


@rule("ven.eliminar_sin_referencia")
@dataclass(frozen=True)
class EliminarSinReferencia(ReglaBase):
    """Las lineas sin referencia (subtotales, comentarios) no son venta."""

    description = "Elimina lineas del informe sin Referencia"
    columna: str = C.INF_REFERENCIA

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if self.columna not in df.columns:
            return RuleResult(df=df)
        ref = df[self.columna]
        queda = ~ref.isna() & (ref.astype(str).str.strip() != "")
        removidas = marcar_eliminadas(df[~queda], "Sin referencia", self.id)
        return RuleResult(df=df[queda].copy(), removed=removidas)


@rule("ven.tipificar_informe")
@dataclass(frozen=True)
class TipificarInforme(ReglaBase):
    """Tipos finales del informe: numeros, costo positivo, FECHA y encabezados normalizados."""

    description = "Convierte tipos, costo unitario positivo, FECHA y encabezados en mayusculas"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        df = df.copy()
        # La referencia NO se convierte a numero: es texto ("0123" sigue "0123").
        for c in (C.INF_NUMERO, "Documento"):
            if c in df.columns:
                df[c] = C.a_numero_si_todo_convierte(df[c])
        if C.INF_COSTO_UNITARIO in df.columns:
            df[C.INF_COSTO_UNITARIO] = pd.to_numeric(df[C.INF_COSTO_UNITARIO], errors="coerce").abs()
        if "COL_TMP_1" in df.columns:
            df = df.drop(columns=["COL_TMP_1"])
        if C.INF_REFERENCIA in df.columns:
            df[C.INF_REFERENCIA] = df[C.INF_REFERENCIA].apply(C.normalizar_referencia)
        if C.INF_FECHA_ORIGEN in df.columns:
            df = df.rename(columns={C.INF_FECHA_ORIGEN: C.FECHA})
        df = df.reset_index(drop=True)
        df.columns = [C.normalizar_encabezado(c) for c in df.columns]
        return RuleResult(df=df)


@rule("ven.excluir_referencias_no_comerciales")
@dataclass(frozen=True)
class ExcluirReferenciasNoComerciales(ReglaBase):
    """Fletes y publicidad se facturan pero no son venta de producto."""

    description = "Excluye REFERENCIA = FLETE VENTAS y las que contienen PUBLICIDAD"
    exactas: tuple[str, ...] = ("FLETE VENTAS",)
    contienen: tuple[str, ...] = ("PUBLICIDAD",)

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if C.REFERENCIA not in df.columns:
            return RuleResult(df=df)
        removidas = []
        for exacta in self.exactas:
            fuera = df[C.REFERENCIA].astype(str).str.strip().str.upper() == exacta.upper()
            removidas.append(marcar_eliminadas(df[fuera], f"REFERENCIA = {exacta}", self.id))
            df = df[~fuera].copy()
        for texto in self.contienen:
            fuera = df[C.REFERENCIA].astype(str).str.upper().str.contains(texto.upper(), na=False, regex=False)
            removidas.append(marcar_eliminadas(df[fuera], f"REFERENCIA contiene {texto}", self.id))
            df = df[~fuera].copy()
        return RuleResult(df=df, removed=pd.concat(removidas) if removidas else pd.DataFrame())


@rule("ven.filtrar_anio_en_curso")
@dataclass(frozen=True)
class FiltrarAnioEnCurso(ReglaBase):
    """El archivo de ventas es del ano completo en curso (hora Colombia)."""

    description = "Conserva solo las ventas del ano en curso"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        if C.FECHA not in df.columns:
            raise InformeSinDatos("Columna FECHA no encontrada en informe")
        if ctx.hoy is None:
            raise InformeSinDatos("El contexto no trae la fecha de hoy (hora Colombia)")
        df = df.copy()
        df[C.FECHA] = pd.to_datetime(df[C.FECHA], dayfirst=True, errors="coerce").dt.normalize()
        entra = df[C.FECHA].dt.year == ctx.hoy.year
        fuera = df[~entra]
        removidas = marcar_eliminadas(
            fuera, f"FECHA fuera de {ctx.hoy.year} o vacia", self.id
        )
        salida = df[entra].copy()
        if salida.empty:
            raise InformeSinDatos("No hay datos del ano actual en el informe")
        return RuleResult(df=salida, removed=removidas, metrics={"anio": ctx.hoy.year})


@rule("ven.calcular_periodo")
@dataclass(frozen=True)
class CalcularPeriodo(ReglaBase):
    """ANO, MES NO. y MES (nombre en espanol) desde FECHA."""

    description = "Calcula ANO, MES NO. y MES desde FECHA"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        df = df.copy()
        df[C.FECHA] = pd.to_datetime(df[C.FECHA], dayfirst=True, errors="coerce").dt.normalize()
        df[C.ANO] = df[C.FECHA].dt.year
        df[C.MES_NO] = df[C.FECHA].dt.month
        df[C.MES] = df[C.MES_NO].map(C.MESES)
        meses = sorted(int(m) for m in df[C.MES_NO].dropna().unique())
        return RuleResult(df=df, metrics={"meses": meses})


@rule("ven.mapeo_semantico")
@dataclass(frozen=True)
class MapeoSemantico(ReglaBase):
    """Nombres del informe -> nombres de la plantilla VENTAS 2026."""

    description = "Renombra columnas del informe a los nombres de la plantilla"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        mapeo = {k: v for k, v in C.MAPEO_SEMANTICO.items() if k in df.columns}
        return RuleResult(df=df.rename(columns=mapeo), metrics={"renombradas": len(mapeo)})


@rule("ven.ordenar_por_fecha")
@dataclass(frozen=True)
class OrdenarPorFecha(ReglaBase):
    """Mas reciente primero; marca FECHA DE ACTUALIZACION con hoy (hora Colombia)."""

    description = "Ordena por FECHA descendente y fija FECHA DE ACTUALIZACION"

    def apply(self, df: pd.DataFrame, ctx: RuleContext) -> RuleResult:
        df = df.copy()
        resultado = RuleResult(df=df)
        if C.FECHA in df.columns:
            df[C.FECHA] = pd.to_datetime(df[C.FECHA], dayfirst=True, errors="coerce").dt.normalize()
            df = df.sort_values(C.FECHA, ascending=False, na_position="last")
        else:
            resultado.con_advertencia("No se puede ordenar (falta columna FECHA)")
        for cand in C.FECHAS_ACTUALIZACION:
            if cand in df.columns:
                df[cand] = ctx.hoy
                break
        resultado.df = df.reset_index(drop=True)
        return resultado
