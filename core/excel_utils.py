from __future__ import annotations

import io
import os
import re
import difflib
from pathlib import Path

import msoffcrypto
import pandas as pd
from unidecode import unidecode

from config.settings import ExcelConfig, MESES_ES, MESES_ES_NOMBRE, MESES_ES_INVERTIDO
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


def norm_sheet(s: str) -> str:
    s = unidecode(str(s)).lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s


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
            log.info("'%s' desencriptado con contrasena '%s'", xlsx_path.name, password)
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


def find_sheet_name(xlsx_stream_or_path, targets: tuple[str, ...] = ("inventario", "inventario general", "inv")) -> str:
    xf = pd.ExcelFile(xlsx_stream_or_path, engine="openpyxl")
    names = xf.sheet_names
    norm_map = {norm_sheet(n): n for n in names}
    tnorms = [norm_sheet(t) for t in targets]
    for t in tnorms:
        if t in norm_map:
            return norm_map[t]
    for t in tnorms:
        for nn, real in norm_map.items():
            if t in nn:
                return real
    if names:
        return names[0]
    raise ValueError("El libro no tiene hojas.")


def find_sheet_by_pattern(workbook, patron: str, ignorar_dolares: bool = True) -> str | None:
    for sheet_name in workbook.sheetnames:
        nombre_limpio = sheet_name
        if ignorar_dolares:
            nombre_limpio = re.sub(r"^\$+", "", sheet_name).strip()
        if patron in nombre_limpio:
            return sheet_name
    return None


def find_file_by_pattern(directorio: str | Path, patron: str, ignorar_dolares: bool = True) -> Path | None:
    dir_path = Path(directorio)
    for archivo in dir_path.glob("*.xlsx"):
        nombre_limpio = archivo.stem
        if ignorar_dolares:
            nombre_limpio = re.sub(r"^\$+", "", nombre_limpio).strip()
        if patron in nombre_limpio:
            return archivo
    return None


COL_SYNONYMS: dict[str, list[str]] = {}


def resolve_cols(df: pd.DataFrame, cols_map: dict) -> dict:
    idx = {norm_colname(c): c for c in df.columns}
    resolved = {}
    for key, target in cols_map.items():
        wanted = norm_colname(target)
        cands = [target] + COL_SYNONYMS.get(target, [])
        found = None
        for c in cands:
            cn = norm_colname(c)
            if cn in idx:
                found = idx[cn]
                break
        if not found:
            for kn, real in idx.items():
                if wanted and wanted in kn:
                    found = real
                    break
        if not found:
            best = difflib.get_close_matches(wanted, list(idx.keys()), n=1, cutoff=0.7)
            if best:
                found = idx[best[0]]
        if not found:
            raise KeyError(f"No encuentro la columna '{target}'. Encabezados: {list(df.columns)}")
        resolved[key] = found
    return resolved


def write_excel(df_or_dict, path_out: Path, sheetname: str | None = None) -> None:
    if isinstance(df_or_dict, dict):
        with pd.ExcelWriter(path_out, engine="openpyxl") as xlw:
            for sh, dfo in df_or_dict.items():
                dfo.to_excel(xlw, sheet_name=sh, index=False)
    else:
        df_or_dict.to_excel(path_out, sheet_name=sheetname or "Sheet1", index=False)
