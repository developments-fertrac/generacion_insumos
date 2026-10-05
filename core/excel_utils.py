from __future__ import annotations

import io
import os
import re
from pathlib import Path

import msoffcrypto
import pandas as pd
from unidecode import unidecode

from config.settings import ExcelConfig, MESES_ES_NOMBRE
from core.seguridad import huella
from core.logger import get_logger

log = get_logger("excel_utils")


_NORM_CACHE: dict[str, str] = {}


def norm(s: str) -> str:
    if s in _NORM_CACHE:
        return _NORM_CACHE[s]
    t = unidecode(str(s)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    _NORM_CACHE[s] = t
    return t


def norm_simple(s: str) -> str:
    s = unidecode(str(s)).lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_colname(s: str) -> str:
    t = unidecode(str(s)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def norm_label(s: str | None) -> str:
    if s is None:
        return ""
    t = unidecode(str(s)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def norm_base_filename(name: str) -> str:
    p = Path(name)
    base = p.stem
    base = base.replace("~$", "")
    base = re.sub(r"^\$+", "", base)
    base = unidecode(base).lower()
    base = re.sub(r"\s+", " ", base).strip()
    return base


def strip_dolares_temporales(name: str) -> str:
    base = Path(name).stem
    base = base.replace("~$", "")
    base = re.sub(r"^\$+", "", base)
    return base


def extract_fecha_es(filename: str):
    filename_low = filename.lower()
    patron = r"(\d{1,2})\s+([a-z\u00e1\u00e9\u00ed\u00f3\u00fa]+)"
    matches = re.findall(patron, filename_low)
    if not matches:
        return None
    from datetime import date
    for dia_str, mes_name in matches:
        mes = MESES_ES_NOMBRE.get(mes_name)
        if mes:
            try:
                dia = int(dia_str)
                year_match = re.search(r"(20\d{2})", filename)
                if year_match:
                    year = int(year_match.group(1))
                    return date(year, mes, dia)
            except (ValueError, TypeError):
                continue
    return None


def decrypt_to_stream(xlsx_path: Path, password: str | None = None, config: ExcelConfig | None = None) -> io.BytesIO:
    if config is None:
        config = ExcelConfig()

    passwords_to_try: list[str] = []
    if password:
        passwords_to_try.append(password)
    for pwd in config.passwords_try:
        if pwd not in passwords_to_try:
            passwords_to_try.append(pwd)

    return _decrypt_with_retry(xlsx_path, passwords_to_try)


def _decrypt_with_retry(xlsx_path: Path, passwords: list[str]) -> io.BytesIO:
    data = _read_file_raw(xlsx_path)

    try:
        bio_check = io.BytesIO(data)
        office = msoffcrypto.OfficeFile(bio_check)
        is_encrypted_attr = getattr(office, "is_encrypted", None)
        if callable(is_encrypted_attr):
            is_encrypted = is_encrypted_attr()
        else:
            is_encrypted = bool(is_encrypted_attr) if is_encrypted_attr is not None else True

        if not is_encrypted:
            bio = io.BytesIO(data)
            bio.seek(0)
            log.info("'%s' no esta protegido, se lee sin contrasena", xlsx_path.name)
            return bio
    except Exception as e:
        msg = str(e).lower()
        if "not encrypted" in msg or "file is not encrypted" in msg:
            bio = io.BytesIO(data)
            bio.seek(0)
            log.info("'%s' no esta protegido", xlsx_path.name)
            return bio

    last_error = None
    for password in passwords:
        try:
            bio_data = io.BytesIO(_read_file_raw(xlsx_path))
            bio = io.BytesIO()
            office = msoffcrypto.OfficeFile(bio_data)
            office.load_key(password=password)
            office.decrypt(bio)
            bio.seek(0)
            log.info("'%s' desencriptado con contrasena %s", xlsx_path.name, huella(password))
            return bio
        except Exception as e:
            last_error = e
            msg = str(e).lower()
            if "not encrypted" in msg or "file is not encrypted" in msg:
                bio = io.BytesIO(_read_file_raw(xlsx_path))
                bio.seek(0)
                log.info("'%s' no esta protegido", xlsx_path.name)
                return bio
            continue

    raise RuntimeError(
        f"No se pudo desencriptar '{xlsx_path.name}' con ninguna contrasena. "
        f"Ultimo error: {last_error}"
    )


def _read_file_raw(path: Path) -> bytes:
    try:
        fd = os.open(str(path), os.O_RDONLY | os.O_BINARY)
        try:
            import msvcrt
            size = os.path.getsize(str(path))
            msvcrt.locking(fd, msvcrt.LK_NBRLCK, size)
            data = os.read(fd, size)
            return data
        finally:
            os.close(fd)
    except Exception:
        with open(path, "rb") as f:
            return f.read()


def read_excel_any(xlsx: Path | io.BytesIO, **kwargs) -> pd.DataFrame:
    if isinstance(xlsx, io.BytesIO):
        return pd.read_excel(xlsx, engine="openpyxl", **kwargs)

    p = Path(xlsx)
    ext = p.suffix.lower()

    if ext == ".csv":
        kwargs.pop("sheet_name", None)
        return pd.read_csv(p, **kwargs)

    try:
        with open(p, "rb") as f:
            head = f.read(8)
    except Exception as e:
        log.error("No se pudo abrir '%s': %s", p, e)
        raise

    is_zip = head.startswith(b"PK")
    is_ole = head.startswith(b"\xD0\xCF\x11\xE0")

    if is_zip:
        try:
            return pd.read_excel(p, engine="openpyxl", **kwargs)
        except Exception as e:
            msg = str(e).lower()
            if "not a zip file" in msg:
                raise RuntimeError(
                    f"El archivo '{p.name}' tiene extension {ext} y empieza como ZIP, "
                    "pero openpyxl indica 'File is not a zip file'."
                ) from e
            raise

    if is_ole:
        try:
            import xlrd
        except ImportError as e:
            raise RuntimeError(
                f"El archivo '{p.name}' parece un Excel antiguo (.xls). Instala xlrd: pip install xlrd"
            ) from e
        kwargs.pop("engine", None)
        return pd.read_excel(p, engine="xlrd", **kwargs)

    raise RuntimeError(f"El archivo '{p.name}' tiene extension {ext}, pero no es ni un ZIP ni un OLE.")

