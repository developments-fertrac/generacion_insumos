"""
Reducción de tamaño de archivos .xlsx (lógica extraída de limpiar.py).

Elimina los dibujos/formas (xl/drawings/drawingN.xml) que inflan el archivo,
conservando los que se indiquen en `keep` (por defecto drawing1.xml), sin
alterar celdas, formatos ni tablas dinámicas.

Soporta archivos protegidos con contraseña (cifrado Office / CFB):
descifra -> limpia -> vuelve a cifrar con la misma contraseña.
"""
from __future__ import annotations

import io
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import msoffcrypto

_RE_DRAWING_PART = re.compile(r"^xl/drawings/(drawing\d+\.xml)$")
_RE_DRAWING_RELS = re.compile(r"^xl/drawings/_rels/(drawing\d+\.xml)\.rels$")
_RE_SHEET = re.compile(r"^xl/worksheets/(sheet\d+\.xml)$")
_RE_SHEET_RELS = re.compile(r"^xl/worksheets/_rels/(sheet\d+\.xml)\.rels$")
_RE_REL = re.compile(r"<Relationship\b[^>]*?/>")
_RE_ATTR = lambda name: re.compile(rf'\b{name}="([^"]*)"')  # noqa: E731


def _rels_path(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", f"{b}.rels")


def _partes_alcanzables(zin: zipfile.ZipFile, nombres: set[str], cortar: set[str]) -> set[str]:
    """Recorre el grafo de relaciones OPC desde la raíz, sin entrar en `cortar`."""
    vistos: set[str] = set()
    pila = [""]  # "" = paquete raíz (_rels/.rels)
    while pila:
        part = pila.pop()
        rels = "_rels/.rels" if part == "" else _rels_path(part)
        if rels not in nombres:
            continue
        base = posixpath.dirname(part)
        txt = zin.read(rels).decode("utf-8", errors="ignore")
        for rel in _RE_REL.findall(txt):
            if 'TargetMode="External"' in rel:
                continue
            t = _RE_ATTR("Target").search(rel)
            if not t:
                continue
            tgt = t.group(1)
            tgt = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
            if tgt in cortar or tgt in vistos:
                continue
            vistos.add(tgt)
            pila.append(tgt)
    return vistos


@dataclass
class ResultadoLimpieza:
    bytes_in: int = 0
    bytes_out: int = 0
    drawings_eliminados: list[str] = field(default_factory=list)
    partes_huerfanas: list[str] = field(default_factory=list)
    top_partes: list[tuple[str, int]] = field(default_factory=list)

    @property
    def ahorro_pct(self) -> float:
        return 0.0 if not self.bytes_in else (1 - self.bytes_out / self.bytes_in) * 100


def _es_cifrado(data_or_path) -> bool:
    if isinstance(data_or_path, (bytes, bytearray)):
        head = bytes(data_or_path[:8])
    else:
        with open(data_or_path, "rb") as f:
            head = f.read(8)
    return head == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def limpiar_drawings_zip(src: io.BytesIO | str | Path, dst: io.BytesIO | str | Path,
                         keep: tuple[str, ...] = ("drawing1.xml",)) -> ResultadoLimpieza:
    """Limpia un .xlsx NO cifrado (zip). `src`/`dst` pueden ser rutas o BytesIO."""
    res = ResultadoLimpieza()
    with zipfile.ZipFile(src, "r") as zin:
        infos = zin.infolist()
        res.top_partes = sorted(((i.filename, i.compress_size) for i in infos),
                                key=lambda x: x[1], reverse=True)[:10]

        eliminar = {m.group(1) for i in infos
                    if (m := _RE_DRAWING_PART.match(i.filename)) and m.group(1) not in keep}
        res.drawings_eliminados = sorted(eliminar)

        # Partes huérfanas (imágenes/gráficos que solo usaban los drawings eliminados)
        nombres = {i.filename for i in infos}
        cortar = {f"xl/drawings/{d}" for d in eliminar}
        alcanzables = _partes_alcanzables(zin, nombres, cortar)
        huerfanas = {n for n in nombres
                     if n.startswith(("xl/media/", "xl/charts/", "xl/drawings/"))
                     and not n.endswith(".rels") and n not in alcanzables}
        huerfanas |= {_rels_path(n) for n in huerfanas if _rels_path(n) in nombres}
        res.partes_huerfanas = sorted(huerfanas - cortar - {_rels_path(c) for c in cortar})

        # rId de cada hoja que apunta a un drawing eliminado
        rids_por_hoja: dict[str, set[str]] = {}
        for i in infos:
            m = _RE_SHEET_RELS.match(i.filename)
            if not m:
                continue
            txt = zin.read(i.filename).decode("utf-8", errors="ignore")
            for rel in _RE_REL.findall(txt):
                tgt = _RE_ATTR("Target").search(rel)
                rid = _RE_ATTR("Id").search(rel)
                if tgt and rid and Path(tgt.group(1)).name in eliminar \
                        and "/drawings/" in tgt.group(1):
                    rids_por_hoja.setdefault(m.group(1), set()).add(rid.group(1))

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for item in infos:
                fn = item.filename

                m = _RE_DRAWING_PART.match(fn) or _RE_DRAWING_RELS.match(fn)
                if (m and m.group(1) in eliminar) or fn in huerfanas:
                    continue

                data = zin.read(fn)

                if fn == "[Content_Types].xml" and (eliminar or huerfanas):
                    txt = data.decode("utf-8")
                    for p in cortar | huerfanas:
                        txt = re.sub(rf'<Override[^>]*PartName="/{re.escape(p)}"[^>]*/>', "", txt)
                    data = txt.encode("utf-8")

                elif (m := _RE_SHEET.match(fn)) and m.group(1) in rids_por_hoja:
                    txt = data.decode("utf-8")
                    for rid in rids_por_hoja[m.group(1)]:
                        txt = re.sub(rf'<drawing\b[^>]*\br:id="{re.escape(rid)}"[^>]*/>', "", txt)
                    data = txt.encode("utf-8")

                elif (m := _RE_SHEET_RELS.match(fn)) and m.group(1) in rids_por_hoja:
                    txt = data.decode("utf-8")
                    for rel in _RE_REL.findall(txt):
                        rid = _RE_ATTR("Id").search(rel)
                        if rid and rid.group(1) in rids_por_hoja[m.group(1)]:
                            txt = txt.replace(rel, "")
                    data = txt.encode("utf-8")

                zi = zipfile.ZipInfo(fn, date_time=item.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = item.external_attr
                zout.writestr(zi, data)
    return res


def reducir_tamano_xlsx(path_in: str | Path, path_out: str | Path | None = None,
                        password: str | None = None,
                        keep: tuple[str, ...] = ("drawing1.xml",)) -> ResultadoLimpieza:
    """
    Reduce el tamaño de un .xlsx (cifrado o no). Si `path_out` es None
    sobrescribe `path_in` de forma atómica (solo si todo salió bien).
    """
    path_in = Path(path_in)
    path_out = Path(path_out) if path_out else path_in
    bytes_in = path_in.stat().st_size
    cifrado = _es_cifrado(path_in)

    if cifrado and not password:
        raise ValueError(f"{path_in.name} está cifrado y no se indicó contraseña")

    # 1) Obtener zip plano
    plano = io.BytesIO()
    if cifrado:
        with open(path_in, "rb") as f:
            of = msoffcrypto.OfficeFile(f)
            of.load_key(password=password)
            of.decrypt(plano)
    else:
        plano.write(path_in.read_bytes())
    plano.seek(0)

    # 2) Limpiar
    limpio = io.BytesIO()
    res = limpiar_drawings_zip(plano, limpio, keep=keep)
    del plano
    limpio.seek(0)

    # 3) Escribir (re-cifrando si aplica) en temporal y reemplazar
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=path_out.parent)
    os.close(fd)
    tmp = Path(tmp)
    try:
        with open(tmp, "wb") as f:
            if cifrado:
                from msoffcrypto.format.ooxml import OOXMLFile  # requiere msoffcrypto-tool >= 5.1
                OOXMLFile(limpio).encrypt(password, f)
            else:
                f.write(limpio.getvalue())
        # Validación mínima: el resultado se puede abrir/descifrar
        if cifrado:
            with open(tmp, "rb") as f:
                chk = msoffcrypto.OfficeFile(f)
                chk.load_key(password=password)
                bio = io.BytesIO()
                chk.decrypt(bio)
                zipfile.ZipFile(bio).testzip()
        else:
            zipfile.ZipFile(tmp).testzip()
        os.replace(tmp, path_out)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise

    res.bytes_in = bytes_in
    res.bytes_out = path_out.stat().st_size
    return res
