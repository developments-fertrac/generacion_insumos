"""Pruebas del protocolo golden: el comparador y el lector.

Se prueban con archivos sinteticos, no con los de ``Archivos Validacion``. Una
prueba que depende de un archivo de 380 MB que el equipo reemplaza cada semana
deja de ser una prueba y pasa a ser un temporizador con pasar/fallar.

Los archivos reales se usan solo en ``tests/golden/test_escenarios.py``, que
salta con un motivo explicito cuando no estan.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from insumos.adapters.excel.comparador import comparar
from insumos.adapters.excel.lector import (
    ExcelReader,
    HojaInexistente,
    PasswordIncorrecta,
)

pytestmark = pytest.mark.unit


def _salida(**overrides: object) -> pd.DataFrame:
    base: dict[str, object] = {
        "REFERENCIA": ["R1", "R2", "R3"],
        "LINEA": ["A1", "B2", "C3"],
        "EXISTENCIA": [10.0, 20.0, 30.0],
        "COSTO PROMEDIO": [100.0, 200.0, 300.0],
    }
    base.update(overrides)
    return pd.DataFrame(base)


# ==================================================================
# comparador: camino feliz
# ==================================================================

class TestComparadorCaminoFeliz:
    def test_salidas_identicas_no_reportan_diferencias(self) -> None:
        resultado = comparar(_salida(), _salida(), claves=["REFERENCIA"])
        assert resultado.iguales
        assert resultado.diferencias == []
        assert resultado.claves_comunes == 3

    def test_el_orden_de_filas_no_es_una_diferencia(self) -> None:
        """El legacy y el motor nuevo ordenan distinto; no es un cambio de dato."""
        actual = _salida()
        nuevo = actual.iloc[::-1].reset_index(drop=True)
        resultado = comparar(actual, nuevo, claves=["REFERENCIA"])
        assert resultado.iguales, resultado.resumen()

    def test_espacios_sin_significado_no_cuentan(self) -> None:
        nuevo = _salida(LINEA=[" A1 ", "B2", "C3 "])
        assert comparar(_salida(), nuevo, claves=["REFERENCIA"]).iguales

    def test_vacios_equivalentes_no_cuentan(self) -> None:
        """None, NaN, "", "  " y el texto "NAN" son la misma ausencia de dato."""
        actual = _salida(LINEA=[None, "  ", float("nan")])
        nuevo = _salida(LINEA=["", " ", "NAN"])
        assert comparar(actual, nuevo, claves=["REFERENCIA"]).iguales

    def test_un_vacio_contra_un_valor_si_es_una_diferencia(self) -> None:
        """Lo contrario del anterior: vacio contra dato es una diferencia real."""
        nuevo = _salida(LINEA=[None, "B2", "C3"])
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert resultado.por_clase == {"valor_distinto": 1}

    def test_diferencia_de_punto_flotante_no_cuenta(self) -> None:
        nuevo = _salida(EXISTENCIA=[10.0000001, 20.0, 30.0])
        assert comparar(_salida(), nuevo, claves=["REFERENCIA"]).iguales

    def test_compara_solo_las_columnas_pedidas(self) -> None:
        """Comparar tambien columnas de formato llena el reporte de ruido."""
        nuevo = _salida()
        nuevo["COLOR"] = "rojo"
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"], columnas=["LINEA"])
        assert resultado.iguales


# ==================================================================
# comparador: casos extremos
# ==================================================================

class TestComparadorCasosExtremos:
    def test_detecta_valor_distinto(self) -> None:
        nuevo = _salida(EXISTENCIA=[10.0, 999.0, 30.0])
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert not resultado.iguales
        assert resultado.por_clase == {"valor_distinto": 1}
        d = resultado.diferencias[0]
        assert d.clave == "R2"
        assert d.columna == "EXISTENCIA"
        assert d.valor_actual == 20.0
        assert d.valor_nuevo == 999.0

    def test_detecta_fila_que_solo_esta_en_la_salida_actual(self) -> None:
        nuevo = _salida().iloc[:2]
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert resultado.por_clase == {"solo_legacy": 1}
        assert resultado.diferencias[0].clave == "R3"

    def test_detecta_fila_nueva(self) -> None:
        nuevo = pd.concat(
            [_salida(), pd.DataFrame({"REFERENCIA": ["R9"], "LINEA": ["Z9"]})],
            ignore_index=True,
        )
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert resultado.por_clase == {"solo_nuevo": 1}
        assert resultado.diferencias[0].clave == "R9"

    def test_detecta_columna_faltante_y_extra(self) -> None:
        actual = _salida()
        actual["ES MARCA PROPIA"] = [True, False, False]
        nuevo = _salida()
        nuevo["OBSERVACION"] = ["a", "b", "c"]
        resultado = comparar(actual, nuevo, claves=["REFERENCIA"])
        clases = resultado.por_clase
        assert clases["columna_faltante"] == 1
        assert clases["columna_extra"] == 1

    def test_clave_repetida_falla_en_vez_de_ocultar_diferencias(self) -> None:
        """Con clave duplicada, comparar por indice esconde diferencias."""
        duplicado = pd.concat([_salida(), _salida().iloc[[0]]], ignore_index=True)
        with pytest.raises(ValueError, match="repetida"):
            comparar(duplicado, _salida(), claves=["REFERENCIA"])

    def test_clave_inexistente_falla_con_mensaje_util(self) -> None:
        with pytest.raises(ValueError, match="no existe en ambas salidas"):
            comparar(_salida(), _salida(), claves=["NO_EXISTE"])

    def test_salida_vacia_a_un_lado(self) -> None:
        resultado = comparar(_salida(), _salida().iloc[0:0], claves=["REFERENCIA"])
        assert resultado.por_clase == {"solo_legacy": 3}
        assert resultado.filas_nuevo == 0

    def test_todas_vacias(self) -> None:
        vacio = pd.DataFrame(columns=["REFERENCIA", "LINEA"])
        resultado = comparar(vacio, vacio, claves=["REFERENCIA"])
        assert resultado.iguales
        assert resultado.filas_actual == 0

    def test_clave_compuesta(self) -> None:
        """El valorizado se identifica por referencia Y almacen."""
        actual = pd.DataFrame(
            {"REFERENCIA": ["R1", "R1"], "ALMACEN": ["GENERAL", "TOBERIN"], "EXISTENCIA": [1.0, 2.0]}
        )
        nuevo = pd.DataFrame(
            {"REFERENCIA": ["R1", "R1"], "ALMACEN": ["GENERAL", "TOBERIN"], "EXISTENCIA": [1.0, 5.0]}
        )
        resultado = comparar(actual, nuevo, claves=["REFERENCIA", "ALMACEN"])
        assert resultado.por_clase == {"valor_distinto": 1}

    def test_texto_que_parece_numero_no_se_confunde(self) -> None:
        """'10' y 10 no son el mismo dato de negocio."""
        nuevo = _salida(EXISTENCIA=["10", 20.0, 30.0])
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert resultado.por_clase == {"valor_distinto": 1}

    def test_reporte_tabla_y_resumen(self) -> None:
        nuevo = _salida(EXISTENCIA=[10.0, 999.0, 30.0])
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        assert "valor_distinto" in resultado.resumen()
        tabla = resultado.por_regla()
        assert len(tabla) == 1
        assert tabla.iloc[0]["columna"] == "EXISTENCIA"

    def test_diferencia_es_legible(self) -> None:
        nuevo = _salida(EXISTENCIA=[10.0, 999.0, 30.0])
        resultado = comparar(_salida(), nuevo, claves=["REFERENCIA"])
        texto = str(resultado.diferencias[0])
        assert "valor_distinto" in texto
        assert "EXISTENCIA" in texto


# ==================================================================
# lector
# ==================================================================

class TestLector:
    def test_archivo_inexistente_falla_temprano(self, tmp_path: Path) -> None:
        with pytest.raises(FileNotFoundError, match="No existe"):
            ExcelReader(ruta=tmp_path / "no.xlsx")

    def test_lee_archivo_plano(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        pd.DataFrame({"A": [1, 2], "B": ["x", "y"]}).to_excel(ruta, index=False)
        lector = ExcelReader(ruta=ruta)
        assert lector.esta_cifrado is False
        assert lector.hojas() == ["Sheet1"]
        assert lector.leer().iloc[0]["B"] == "x"

    def test_lee_por_nombre_de_hoja(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        with pd.ExcelWriter(ruta) as w:
            pd.DataFrame({"A": [1]}).to_excel(w, sheet_name="GENERAL", index=False)
            pd.DataFrame({"A": [2]}).to_excel(w, sheet_name="TOBERIN", index=False)
        lector = ExcelReader(ruta=ruta)
        assert lector.leer("TOBERIN").iloc[0]["A"] == 2

    def test_lee_por_indice_de_hoja(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        with pd.ExcelWriter(ruta) as w:
            pd.DataFrame({"A": [1]}).to_excel(w, sheet_name="GENERAL", index=False)
            pd.DataFrame({"A": [2]}).to_excel(w, sheet_name="TOBERIN", index=False)
        assert ExcelReader(ruta=ruta).leer(1).iloc[0]["A"] == 2

    def test_hoja_inexistente_lista_las_disponibles(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        pd.DataFrame({"A": [1]}).to_excel(ruta, index=False, sheet_name="GENERAL")
        with pytest.raises(HojaInexistente, match="GENERAL"):
            ExcelReader(ruta=ruta).leer("NO_EXISTE")

    def test_indice_de_hoja_fuera_de_rango(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        pd.DataFrame({"A": [1]}).to_excel(ruta, index=False)
        with pytest.raises(HojaInexistente, match="no existe la"):
            ExcelReader(ruta=ruta).leer(7)

    def test_archivo_cifrado_sin_password_falla_explicito(self, tmp_path: Path) -> None:
        ruta = _archivo_cifrado(tmp_path)
        with pytest.raises(PasswordIncorrecta, match="EXCEL_PASSWORD"):
            ExcelReader(ruta=ruta).leer()

    def test_archivo_cifrado_con_password_incorrecta(self, tmp_path: Path) -> None:
        ruta = _archivo_cifrado(tmp_path)
        with pytest.raises(PasswordIncorrecta, match="no abre"):
            ExcelReader(ruta=ruta, password="incorrecta").leer()

    def test_archivo_cifrado_con_password_correcta(self, tmp_path: Path) -> None:
        ruta = _archivo_cifrado(tmp_path, password="secreta")
        lector = ExcelReader(ruta=ruta, password="secreta")
        assert lector.esta_cifrado is True
        assert lector.leer().iloc[0]["A"] == 1

    def test_con_password_devuelve_otro_lector(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        pd.DataFrame({"A": [1]}).to_excel(ruta, index=False)
        otro = ExcelReader(ruta=ruta).con_password("otra")
        assert otro.ruta == ruta
        assert otro.password == "otra"

    def test_iterar_hojas(self, tmp_path: Path) -> None:
        ruta = tmp_path / "p.xlsx"
        with pd.ExcelWriter(ruta) as w:
            pd.DataFrame({"A": [1]}).to_excel(w, sheet_name="GENERAL", index=False)
            pd.DataFrame({"A": [2]}).to_excel(w, sheet_name="TOBERIN", index=False)
        pares = dict(ExcelReader(ruta=ruta).iterar())
        assert set(pares) == {"GENERAL", "TOBERIN"}

    def test_el_descifrado_no_escribe_nada_en_disco(self, tmp_path: Path) -> None:
        """Un temporal plaintext de inventario es un riesgo, no una comodidad."""
        ruta = _archivo_cifrado(tmp_path, password="secreta")
        antes = {p.name for p in tmp_path.iterdir()}
        ExcelReader(ruta=ruta, password="secreta").leer()
        assert {p.name for p in tmp_path.iterdir()} == antes


def _archivo_cifrado(tmp_path: Path, password: str | None = None) -> Path:
    """Crea un .xlsx cifrado con msoffcrypto, para probar el lector de verdad."""
    from msoffcrypto.format.ooxml import OOXMLFile

    origen = tmp_path / "origen.xlsx"
    pd.DataFrame({"A": [1, 2]}).to_excel(origen, index=False)

    cifrado = tmp_path / "cifrado.xlsx"
    with origen.open("rb") as fh, cifrado.open("wb") as salida:
        oficina = OOXMLFile(fh)
        oficina.encrypt(password or "clave-por-defecto", salida)
    return cifrado


def test_la_ayuda_de_msoffcrypto_esta_disponible() -> None:
    """Si msoffcrypto cambia la API, este test avisa antes que los de arriba."""
    import msoffcrypto.format.ooxml

    assert hasattr(msoffcrypto.format.ooxml.OOXMLFile, "encrypt")
