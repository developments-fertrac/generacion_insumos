"""Lectura de archivos Excel a DataFrames.

Es el unico lugar del sistema nuevo que sabe que existen `.xlsx` cifrados con
contrasena. Todo lo que entra al dominio llega como ``pd.DataFrame``: si una
regla necesita el nombre de una hoja, es senal de que la regla esta haciendo
trabajo de lectura, que es del adaptador.

Por que existe ``ExcelReader`` y no funciones sueltas: la fase 2 necesita leer
los mismos archivos miles de veces (goldens) y no puede reinventar el
descifrado en cada script. La clase concentra el comportamiento real, incluido
el PasswordError explicito cuando la contrasena no sirve.
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Protocol


def _nombres(xls: Any) -> list[str]:
    """Nombres de hoja como texto.

    pandas no declara el tipo de ``sheet_names``, asi que mypy lo ve como
    ``list[int | str]``. Los nombres de hoja de Excel son texto siempre; el
    ``str()`` deja eso dicho en un solo lugar en vez de repetir un ``cast`` en
    cada llamada.
    """
    return [str(n) for n in xls.sheet_names]


class PasswordIncorrecta(RuntimeError):
    """El archivo esta cifrado y la contrasena no lo abre.

    Se distingue de "el archivo no existe" porque son fallos muy distintos: uno
    es de configuracion, el otro es de entrada. Confundirlos mando a revisar
    permisos cuando el problema es una contrasena que cambio.
    """


class HojaInexistente(KeyError):
    """Se pidio una hoja que el archivo no tiene."""


class LectorExcel(Protocol):
    """El puerto de lectura que consume el caso de uso."""

    def hojas(self) -> list[str]:
        """Nombres de las hojas, en el orden del archivo."""
        ...

    def leer(self, hoja: str, **kwargs: Any) -> Any:
        """Devuelve la hoja como DataFrame."""
        ...


@dataclass(frozen=True)
class OpcionesLectura:
    """Ajustes de lectura, con los valores que usa el legacy."""

    encabezados: int | None = 0
    dtype: Any = None


@dataclass
class ExcelReader:
    """Lee un `.xlsx` (cifrado o no) con pandas.

    El descifrado va con ``msoffcrypto``, que es lo que ya usa el legacy. Se
    descifra a memoria, no a disco: escribir un temporal plaintext de un archivo
    de inventario de 380 MB dejaria datos sensibles tirados en la carpeta del
    sistema, y el beneficio de un temporal es nulo.
    """

    ruta: Path
    password: str | None = None

    def __post_init__(self) -> None:
        if not self.ruta.exists():
            raise FileNotFoundError(f"No existe el archivo: {self.ruta}")

    # ------------------------------------------------------------------

    @property
    def esta_cifrado(self) -> bool:
        """Si el archivo tiene cifrado de Office.

        Se lee la cabecera OOXML en vez de intentar descifrar: la excepcion de
        msoffcrypto al abrir un archivo plano no es un error de contrasena, y
        confundirlas lleva a pedirle al usuario una contrasena que no es el
        problema.
        """
        with self.ruta.open("rb") as fh:
            magic = fh.read(8)
        # Un OOXML cifrado es un contenedor OLE (D0 CF 11 E0), no un ZIP (PK).
        return magic[:4] == b"\xd0\xcf\x11\xe0"

    def _bytes_planos(self) -> bytes:
        """Devuelve el contenido descifrado en memoria."""
        if not self.esta_cifrado:
            return self.ruta.read_bytes()

        import msoffcrypto

        buf = BytesIO()
        with self.ruta.open("rb") as fh:
            archivo = msoffcrypto.OfficeFile(fh)
            if not self.password:
                raise PasswordIncorrecta(
                    f"{self.ruta.name} esta cifrado y no hay contrasena configurada "
                    "(EXCEL_PASSWORD). Sin ella no se puede leer."
                )
            try:
                archivo.load_key(password=self.password)
                archivo.decrypt(buf)
            except Exception as e:
                raise PasswordIncorrecta(
                    f"La contrasena no abre {self.ruta.name}: {e}"
                ) from e
        buf.seek(0)
        return buf.getvalue()

    # ------------------------------------------------------------------

    def hojas(self) -> list[str]:
        """Nombres de las hojas del archivo."""
        import pandas as pd

        with pd.ExcelFile(BytesIO(self._bytes_planos())) as xls:
            return _nombres(xls)

    def leer(self, hoja: str | int = 0, opciones: OpcionesLectura | None = None) -> Any:
        """Devuelve una hoja como DataFrame.

        ``hoja`` acepta nombre o indice. Aceptar ambos importa porque el legacy
        llama a las hojas por posicion y los datos por nombre, y cada uno de los
        dos lados se equivoca alguna vez.
        """
        import pandas as pd

        op = opciones or OpcionesLectura()
        contenido = BytesIO(self._bytes_planos())

        with pd.ExcelFile(contenido) as xls:
            disponibles = _nombres(xls)
            if isinstance(hoja, int):
                if not (-len(disponibles) <= hoja < len(disponibles)):
                    raise HojaInexistente(
                        f"El archivo tiene {len(disponibles)} hojas; no existe la {hoja}."
                    )
                nombre = disponibles[hoja]
            else:
                if hoja not in disponibles:
                    raise HojaInexistente(
                        f"{self.ruta.name} no tiene la hoja '{hoja}'. Tiene: {disponibles}"
                    )
                nombre = hoja
            return pd.read_excel(
                xls,
                sheet_name=nombre,
                header=op.encabezados,
                dtype=op.dtype,
            )

    def iterar(self, hojas: Sequence[str] | None = None) -> Iterator[tuple[str, Any]]:
        """Itera ``(nombre_hoja, DataFrame)``.

        Recorre de a uno: el archivo de ventas pesa 380 MB y leer todas las hojas
        a la vez se queda sin memoria en la maquina de produccion.
        """
        for nombre in hojas if hojas is not None else self.hojas():
            yield nombre, self.leer(nombre)

    def con_password(self, password: str) -> ExcelReader:
        """Otro lector del mismo archivo con otra contrasena.

        Los insumosarios llegan con distintas contrasenas segun quien los
        exporto, y probarlas en bucle era el patron del legacy. Aqui queda
        explicito: se intenta una, y si falla se cambia el argumento.
        """
        return ExcelReader(ruta=self.ruta, password=password)
