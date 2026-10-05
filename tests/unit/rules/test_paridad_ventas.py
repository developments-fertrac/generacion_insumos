"""Paridad: el pipeline ``config/rules/ventas.yaml`` produce lo mismo que los pasos 2-9 del proceso anterior.

El oraculo es ``ActualizacionVentas._transformar_legacy`` (el codigo de antes,
extraido sin cambios de logica). Los datos son sinteticos pero cubren cada
rama: notas credito, FE de pedido, fletes, publicidad, otros anos, referencias
invalidas, descuentos condicionados y precios licitados.

Con datos reales la misma comparacion la hace ``scripts/comparar_ventas.py``.
"""

from __future__ import annotations

import logging
from datetime import date
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

pytestmark = pytest.mark.unit

openpyxl = pytest.importorskip("openpyxl")

from insumos.adapters.excel import fuentes_ventas as fv  # noqa: E402
from insumos.application.transformar_ventas import TransformarVentas  # noqa: E402
from insumos.domain.ports import EntradasVentas  # noqa: E402
from insumos.domain.rules.ventas.integraciones import descuento_a_decimal  # noqa: E402

_RAIZ = Path(__file__).resolve().parents[3]
HOY = date(2026, 10, 2)

COLUMNAS_INFORME = [
    "Prefijo", "Número", "Nro. documento", "Fecha", "Fecha documento origen", "Documento origen",
    "Nro. documento cliente", "Cliente", "Ciudad/Sucursal", "Vendedor", "Referencia", "Descripción",
    "Marca", "Cantidad facturada", "Valor bruto", "Valor descuento comercial", "Costo unitario",
    "Valor base precio de venta", "NIT",
]

FILAS_INFORME = [
    # pref, num, nro, fecha, fecha origen, doc origen, nit cli, cliente, ciudad, vend, ref, desc, marca, cant, bruto, vdesc, costo, base, nit
    ("FE", 1001, "FE1001", "01/10/2026", "28/09/2026", "PV-77", 900111222, "ACME", "BOGOTA", "V1", 4591, "Disco", "EATON", 2, 200.0, 0, -50.0, 100.0, "x"),
    ("FE", 1002, "FE1002", "02/10/2026", "15/09/2026", "OV-12", "900333444", "BETA", "CALI", "V2", "A-77", "Kit", "CUMMINS", 1, 90.0, 5.0, 30.0, 90.0, "x"),
    ("NC", 1003, "NC1003", "02/10/2026", "", "", 900111222, "ACME", "BOGOTA", "V1", 4591, "Disco", "EATON", -1, -100.0, 0, 50.0, 100.0, "x"),
    ("NCDTO", 1004, "NCD1004", "02/10/2026", "", "", 900111222, "ACME", "BOGOTA", "V1", 4591, "Disco", "EATON", 0, -10.0, 0, 0.0, 0.0, "x"),
    ("FV", 1005, "FV1005", "10/03/2026", "", "", 800555666, "GAMA", "MEDELLIN", "V3", "FLETE VENTAS", "Flete", "", 1, 20.0, 0, 0.0, 20.0, "x"),
    ("FV", 1006, "FV1006", "11/03/2026", "", "", 800555666, "GAMA", "MEDELLIN", "V3", "PUBLICIDAD CATALOGO", "Pub", "", 1, 5.0, 0, 0.0, 5.0, "x"),
    ("FV", 1007, "FV1007", "20/12/2025", "", "", 800555666, "GAMA", "MEDELLIN", "V3", "B-1", "Viejo", "TIMKEN", 3, 30.0, 0, 7.0, 10.0, "x"),
    ("FV", 1008, "FV1008", "21/05/2026", "", "", 800555666, "GAMA", "MEDELLIN", "V3", None, "Sin ref", "", 1, 1.0, 0, 0.0, 1.0, "x"),
    ("FV", 1009, "FV1009", "22/05/2026", "", "", 700777888, "DELTA", "PASTO", "V4",
     "SERVICIO DE REPARACION GENERAL DE CAJA", "Servicio", "", 1, 500.0, 0, 0.0, 500.0, "x"),
    ("FE", 1010, "FE1010", "05/08/2026", "01/08/2026", "PV-90", 900333444, "BETA", "CALI", "V2", "A-77", "Kit", "CUMMINS", 4, 360.0, 0, 30.0, 90.0, "x"),
    ("FV", 1011, "FV1011", "05/08/2026", "", "", 700777888, "DELTA", "PASTO", "V4", 9999, "Sin inv", "VARIOS", 2, 10.0, 0, 3.0, 5.0, "x"),
    ("FV", 1012, "FV1012", "06/08/2026", "", "", 900111222, "ACME", "BOGOTA", "V1", 4591, "Disco", "EATON", 0, 0.0, 0, 50.0, 0.0, "x"),
    ("FV", 1013, "FV1013", "07/08/2026", "", "", 600999000, "EPSILON", "TUNJA", "V5", "0123", "Reten", "SKF", 2, 40.0, 0, 9.0, 20.0, "x"),
    ("FV", 1014, "FV1014", "08/08/2026", "", "", 900111222, "ACME", "BOGOTA", "V1", 10, "Buje", "SKF", 1, 15.0, 0, 4.0, 15.0, "x"),
]


def _informe() -> pd.DataFrame:
    return pd.DataFrame(FILAS_INFORME, columns=COLUMNAS_INFORME)


def _inventario() -> pd.DataFrame:
    return pd.DataFrame({
        "REFERENCIA": [4591, "A-77", "A-77", "C-3", 9999, "0123", 123, 10],
        "LINEA COPIA": ["FULLER", "CUMMINS-OLD", "CUMMINS", "TIMKEN", 0, "SKF-RETEN", "OTRA-123", "SKF-BUJE"],
        "SUB-LINEA COPIA": ["EATON", "X", "CUMMINS", "TIMKEN", 0, "SKF", "X", "SKF"],
        "LIDER LINEA": ["ANA", "LUIS", "LUIS", "EVA", None, "EVA", "EVA", "EVA"],
        "TOTAL INV": [1, 2, 3, 4, 5, 6, 7, 8],
    })


def _myr() -> pd.DataFrame:
    return pd.DataFrame({"REFERENCIA": [4591.0, "A-77", "A-77", "0123", 123], "COSTO FACTOR HOY": [55.5, 31.0, 32.0, 7.0, 99.0]})


def _matriz() -> pd.DataFrame:
    return pd.DataFrame({
        "NIT": [900111222.0, 900333444, 800555666, 700777888, 600999000],
        "CLIENTE": ["ACME", "BETA", "GAMA", "DELTA", "EPSILON"],
        "TIPO DESCUENTO": ["CONDICIONADO", "condicionado ", "FINANCIERO", "CONDICIONADO", "CONDICIONADO"],
        "DESCUENTO": ["8%", "5% MAXIMO 60 DIAS", "3%", "SIN DTO", "12,5%"],
    })


def _libro_licitados() -> BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VENTAS 2026"
    lic = wb.create_sheet("PRECIO UNIT LICITADOS")
    lic["C1"], lic["D1"], lic["E1"] = 900111222, 900333444.0, None
    for i, h in enumerate(["#", "REFERENCIA FERTRAC", "ACME", "BETA", "OTRO"], start=1):
        lic.cell(4, i, h)
    lic.append([1, 4591, 120.0, None, None])
    lic.append([2, "A-77", None, 88.0, 10])
    lic.append([3, "Z-1", 0, -5, None])
    lic.append([4, 10.05, 14.0, None, None])   # numero con decimales -> "10"
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


COLUMNAS_PLANTILLA = (
    "FECHA DE ACTUALIZACION", "DV", "NUMERO", "NRO. DOCUMENTO", "FECHA", "NIT CLIENTE", "CLIENTE",
    "CIUDAD", "VEND", "REFERENCIA", "DESCRPCION", "MARCA", "CANTIDAD", "VR UNITARIO", "VR TOTAL",
    "VR DESCUENTO", "COSTO PROMEDIO", "LINEA", "SUBLINEA", "LIDER LINEA", "COSTO FACTOR HOY",
    "DCTO CONDICIONADO", "PORCENTAJE DCTO A PIE DE FACTURA", "VTA ACORDADA X UNIDAD LICITADO",
    "MES", "MES NO.", "ANO", "SEMANA NUMERO", "VENTA TOTAL NETA", "COLUMNA MANUAL",
)


def _entradas(libro: BytesIO, informe: pd.DataFrame | None = None) -> EntradasVentas:
    precios, nits = fv.leer_precios_licitados(libro)
    return EntradasVentas(
        informe=_informe() if informe is None else informe,
        inventario=_inventario(),
        myr=_myr(),
        matriz_clientes=_matriz(),
        columnas_plantilla=COLUMNAS_PLANTILLA,
        precios_licitados=precios,
        nits_licitados=nits,
    )


@pytest.fixture
def legacy(monkeypatch):
    modulo = pytest.importorskip("tasks.actualizacion_ventas")
    monkeypatch.setattr(modulo, "_hoy_bogota", lambda: HOY)
    tarea = modulo.ActualizacionVentas.__new__(modulo.ActualizacionVentas)
    tarea.log = logging.getLogger("paridad_ventas")
    return tarea


def _correr_ambos(legacy, informe: pd.DataFrame | None = None):
    libro = _libro_licitados()
    entradas = _entradas(libro, informe)
    viejo = legacy._transformar_legacy(entradas, libro, HOY.year)
    nuevo = TransformarVentas(reglas=_RAIZ / "config" / "rules" / "ventas.yaml", hoy=HOY)(entradas)
    return viejo, nuevo


def _assert_paridad(viejo: pd.DataFrame, nuevo: pd.DataFrame) -> None:
    assert set(viejo.columns) == set(nuevo.columns)
    # Las columnas de la plantilla, en el mismo orden; las extra el legacy las
    # ordenaba como salieran de un set (no determinista), por eso solo conjunto.
    comunes = [c for c in COLUMNAS_PLANTILLA if c in viejo.columns]
    assert list(viejo.columns[: len(comunes)]) == comunes == list(nuevo.columns[: len(comunes)])
    pd.testing.assert_frame_equal(
        nuevo[list(viejo.columns)].reset_index(drop=True),
        viejo.reset_index(drop=True),
        check_dtype=False,
    )


def test_paridad_con_proceso_anterior(legacy):
    viejo, nuevo = _correr_ambos(legacy)
    _assert_paridad(viejo, nuevo.datos)
    assert len(nuevo.datos) == 7  # 12 lineas - NC - NCDTO - flete - publicidad - 2025 - sin ref - servicio


def test_paridad_con_filas_de_titulo(legacy):
    titulo = pd.DataFrame([["INFORME DE VENTAS"] + [None] * (len(COLUMNAS_INFORME) - 1),
                           [None] * len(COLUMNAS_INFORME),
                           COLUMNAS_INFORME], columns=[f"Unnamed: {i}" for i in range(len(COLUMNAS_INFORME))])
    cuerpo = pd.DataFrame(FILAS_INFORME, columns=titulo.columns)
    informe = pd.concat([titulo, cuerpo], ignore_index=True)
    viejo, nuevo = _correr_ambos(legacy, informe)
    _assert_paridad(viejo, nuevo.datos)


def test_reglas_de_negocio_en_el_resultado(legacy):
    _, nuevo = _correr_ambos(legacy)
    df = nuevo.datos.set_index("NUMERO")

    # FE de un pedido (PV) toma la fecha del pedido; FE de otra cosa, la de factura.
    assert df.loc[1001, "FECHA"] == pd.Timestamp("2026-09-28")
    assert df.loc[1002, "FECHA"] == pd.Timestamp("2026-10-02")
    assert df.loc[1001, "MES"] == "SEPTIEMBRE" and df.loc[1001, "MES NO."] == 9
    # Orden: mas reciente primero.
    assert list(nuevo.datos["FECHA"]) == sorted(nuevo.datos["FECHA"], reverse=True)
    # Cruces.
    assert df.loc[1001, "LINEA"] == "FULLER" and df.loc[1002, "LINEA"] == "CUMMINS"  # ultima del inventario
    assert df.loc[1002, "COSTO FACTOR HOY"] == 32.0
    assert df.loc[1001, "COSTO PROMEDIO"] == 50.0  # costo unitario en positivo
    assert df.loc[1001, "VR UNITARIO"] == 100.0
    # Condicionado: 8% para ACME; BETA tiene VR DESCUENTO en 1002 -> 0%.
    assert df.loc[1001, "DCTO CONDICIONADO"] == pytest.approx(0.08)
    assert df.loc[1002, "DCTO CONDICIONADO"] == "0%"
    assert df.loc[1010, "DCTO CONDICIONADO"] == pytest.approx(0.05)
    # Licitados.
    assert df.loc[1001, "VTA ACORDADA X UNIDAD LICITADO"] == 120.0
    assert df.loc[1010, "VTA ACORDADA X UNIDAD LICITADO"] == 88.0
    assert df.loc[1011, "VTA ACORDADA X UNIDAD LICITADO"] == 0
    # Formulas no se escriben; columnas manuales de la plantilla quedan vacias.
    assert "VENTA TOTAL NETA" not in nuevo.datos.columns
    assert nuevo.datos["COLUMNA MANUAL"].isna().all()
    assert (nuevo.datos["FECHA DE ACTUALIZACION"] == HOY).all()


def test_referencia_es_texto_y_ajustes_de_negocio(legacy):
    """Ajustes pedidos por el area (2026-10-02)."""
    _, nuevo = _correr_ambos(legacy)
    df = nuevo.datos.set_index("NUMERO")
    # 1. "0123" sigue "0123" y cruza con "0123" del inventario y del MYR, no con 123.
    assert df.loc[1013, "REFERENCIA"] == "0123"
    assert df.loc[1013, "LINEA"] == "SKF-RETEN"
    assert df.loc[1013, "COSTO FACTOR HOY"] == 7.0
    assert df.loc[1001, "REFERENCIA"] == "4591"  # celda numerica: sin ".0"
    assert df.loc[1001, "COSTO FACTOR HOY"] == 55.5
    # 2. Licitados: 10.05 numerico se redondea a "10".
    assert df.loc[1014, "VTA ACORDADA X UNIDAD LICITADO"] == 14.0
    # 3. "12,5%" -> 0.125.
    assert df.loc[1013, "DCTO CONDICIONADO"] == pytest.approx(0.125)


def test_eliminaciones_quedan_auditadas(legacy):
    _, nuevo = _correr_ambos(legacy)
    por_regla = nuevo.resumen()["eliminadas_por_regla"]
    assert por_regla == {
        "ven.excluir_prefijos": 2,
        "ven.eliminar_sin_referencia": 1,
        "ven.excluir_referencias_no_comerciales": 2,
        "ven.filtrar_anio_en_curso": 1,
        "ven.eliminar_referencias_invalidas": 1,
    }


def test_sin_hoja_licitados_sigue_sin_la_columna(legacy):
    entradas = _entradas(_libro_licitados())
    sin = EntradasVentas(**{**entradas.__dict__, "precios_licitados": None, "nits_licitados": ()})
    nuevo = TransformarVentas(reglas=_RAIZ / "config" / "rules" / "ventas.yaml", hoy=HOY)(sin)
    # La plantilla tiene la columna: queda vacia (como antes, cuando el paso 7.5 fallaba).
    assert nuevo.datos["VTA ACORDADA X UNIDAD LICITADO"].isna().all()
    assert any("PRECIO UNIT LICITADOS" in w for p in nuevo.audit.pasos for w in p.warnings)


@pytest.mark.parametrize("valor, esperado", [
    ("8%", 0.08), ("12,5 %", 0.125), ("12,5%", 0.125), ("12.5%", 0.125), ("SIN DTO", 0), ("5% maximo 60 dias", 0.05),
    (0.07, 0.07), (7, 0.07), (-1, 0), ("abc", 0), (None, 0), (np.nan, 0), ("0,1", 0.1),
])
def test_descuento_a_decimal(valor, esperado):
    assert descuento_a_decimal(valor) == pytest.approx(esperado)


def test_comparador_por_posicion_detecta_y_acepta(legacy):
    from insumos.adapters.excel.comparador import comparar_por_posicion

    viejo, nuevo = _correr_ambos(legacy)
    nuevo_df = nuevo.datos[list(viejo.columns)]
    assert comparar_por_posicion(viejo, nuevo_df).empty

    alterado = nuevo_df.copy()
    alterado.loc[0, "LINEA"] = "OTRA"
    dif = comparar_por_posicion(viejo, alterado.drop(columns=["CIUDAD"]))
    assert set(dif["TIPO"]) == {"valor_distinto", "columna_faltante"}
    assert dif.loc[dif["TIPO"] == "valor_distinto", "COLUMNA"].tolist() == ["LINEA"]


@pytest.mark.parametrize("valor, esperado", [
    ("0123", "0123"), (" A-77 ", "A-77"), (4591, "4591"), (4591.0, "4591"), (np.int64(10), "10"),
    ("10.05", "10.05"), (10.05, "10.05"), (None, None), (np.nan, None), ("", None), ("nan", None),
])
def test_normalizar_referencia_es_texto(valor, esperado):
    from insumos.domain.rules.ventas.columnas import normalizar_referencia

    assert normalizar_referencia(valor) == esperado


@pytest.mark.parametrize("valor, esperado", [
    (10.05, "10"), (10.5, "11"), (900111222.0, "900111222"), (4591, "4591"), ("0123", "0123"), (" X-1 ", "X-1"),
])
def test_texto_licitado(valor, esperado):
    from insumos.domain.rules.ventas.columnas import texto_licitado

    assert texto_licitado(valor) == esperado


def test_informe_sin_columna_nit_no_genera_advertencia_falsa():
    """El _268 real no trae columna NIT: el encabezado en la fila 1 basta."""
    from insumos.domain.rules.base import RuleContext
    from insumos.domain.rules.ventas.informe import DetectarEncabezadoInforme

    informe = _informe().drop(columns=["NIT"])
    resultado = DetectarEncabezadoInforme().apply(informe, RuleContext())
    assert resultado.warnings == []
    assert list(resultado.df.columns) == list(informe.columns)
