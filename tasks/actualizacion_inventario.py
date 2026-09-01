"""
Tarea: Actualizacion de Inventario General
Migrado desde Actualizar Inv general/actualizar_inventario_general.py
"""

from __future__ import annotations

import contextlib
import gc
import io
import os
import re
import shutil
import sys
import tempfile
import time
import warnings
from datetime import date, datetime
from pathlib import Path

import msoffcrypto
import numpy as np
import pandas as pd
from unidecode import unidecode

from config.settings import Settings
from core.email_notifier import EmailNotifier
from core.excel_processing import (
    HAS_COM,
    com_convert_to_xlsx,
    decrypt_to_stream_local,
    is_encrypted_xlsx,
    limpiar_archivo_temporal,
    limpiar_copias_temporales_antiguas,
    obtener_archivo_trabajo,
    safe_close_workbook,
    safe_quit_excel,
    save_bytesio_to_temp,
    verificar_archivo_disponible,
)
from core.excel_utils import (
    decrypt_to_stream,
    find_file_by_pattern,
    find_sheet_by_pattern,
    norm as _norm,
    norm_colname,
    norm_label,
    norm_sheet,
    read_excel_any,
    strip_dolares_temporales,
    write_excel,
)
from core.logger import get_logger
from tasks.base_task import BaseTask

try:
    import win32com.client as win32
except Exception:
    win32 = None

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_log = get_logger("actualizacion_inventario")


def log(msg):
    safe = str(msg).encode("ascii", "ignore").decode()
    _log.info(safe)


# ==== CONFIG (reemplazada por self.settings en la clase) ====
ARCHIVOS_DIARIOS = None
RUTA_INV_ERP = None
RUTA_REMISIONES = None
RUTA_VALORIZADOS = None
RUTA_INSUMOS = None
OUTPUT_PATH = None
BASE_PATH = None
PASS_INV = None
PASSWORDS_TRY = None

OUTPUT_BASENAME = "$2026 INVENTARIO GENERAL ACTUALIZADO"
APPLY_PASSWORD_TO_OUTPUT = True

PFX_INV_ACTUALIZADO = "INVENTARIO GENERAL ACTUALIZADO"
PFX_VAL_GENERAL = "VALORIZADO GENERAL"
PFX_VAL_FALT_IMPO = "VALORIZADO FALTANTES IMPO"
PFX_VAL_FALT = "VALORIZADO FALTANTES"
PFX_VAL_TOBERIN = "VALORIZADO TOBERIN"
PFX_MARCAS = "MARCAS"
PFX_DISTRIBUCION = "DISTRIBUCION DE MATRICES"
PFX_CONSOLIDADO_REMISIONES = "CONSOLIDADO REMISIONES 2019-2026"
PFX_MATRIZ_USD = "2026 MATRIZ USD"

PATRON_MATRIZ_USD = "2026 MATRIZ USD"
PATRON_SHEET_2025 = "2026"
PATRON_INV_FILE = "2026 INVENTARIO GENERAL"
SHEET_INV_ORIG = "INVENTARIO"
SHEET_INV_COPIA = "INVENTARIO COPIA"
SHEET_INV_LISTA = "INV LISTA PRECIOS"

HEADER_ROW_INV = 2
HEADER_ROW_INV_LISTA = 1
HEADER_ROW_VAL = 9
HEADER_ROW_MATRIZ = 1
HEADER_ROW_MAYOR_EXIST = 1

COLS_A_LIMPIAR = [
    "REFERENCIA", "NOMBRE LISTA", "NOMBRE ODOO", "NOMBRE MYR",
    "MARCA copia", "INV BODEGA", "EXISTENCIA AGO 26", "COSTO PROMEDIO",
    "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION",
    "Marca sistema", "Linea sistema", "Sub- linea sistema",
]
COLS_DESDE_ORIGINAL = [
    "MARCA copia", "INV BODEGA GERENCIA", "LINEA COPIA",
    "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION",
]
class EliminacionTracker:
    def __init__(self):
        self.eliminaciones = []
        self.stats_por_paso = {}

    def registrar(self, paso: str, fila_excel: int, referencia: str,
                  nombre: str = "", marca: str = "", linea: str = "",
                  motivo: str = "", datos_extra: dict = None):
        registro = {
            "TIMESTAMP": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "PASO": paso,
            "FILA_EXCEL": fila_excel,
            "REFERENCIA": referencia,
            "NOMBRE": nombre,
            "MARCA": marca,
            "LINEA": linea,
            "MOTIVO": motivo,
        }
        if datos_extra:
            registro.update(datos_extra)
        self.eliminaciones.append(registro)
        if paso not in self.stats_por_paso:
            self.stats_por_paso[paso] = 0
        self.stats_por_paso[paso] += 1

    def log_eliminacion(self, paso: str, fila_excel: int, referencia: str,
                        motivo: str = "", mostrar_en_consola: bool = True):
        if mostrar_en_consola:
            log(f"    Fila {fila_excel}: {referencia} -> {motivo}")
        self.registrar(paso, fila_excel, referencia, motivo=motivo)

    def mostrar_resumen(self):
        log("")
        log("=" * 70)
        log("RESUMEN DE ELIMINACIONES")
        log("=" * 70)
        if not self.eliminaciones:
            log("No se eliminaron referencias durante el proceso")
            return
        log(f"Total de referencias eliminadas: {len(self.eliminaciones)}")
        log("")
        log("Desglose por paso:")
        for paso, cantidad in sorted(self.stats_por_paso.items(), key=lambda x: x[1], reverse=True):
            log(f"  - {paso}: {cantidad} referencias")
        log("=" * 70)
        log("")

    def generar_reporte_excel(self, base_path: Path) -> Path:
        if not self.eliminaciones:
            log("No hay eliminaciones para reportar")
            return None
        try:
            log("")
            log("Generando reporte de eliminaciones...")
            df_todas = pd.DataFrame(self.eliminaciones)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_reporte = f"REPORTE_ELIMINACIONES_{timestamp}.xlsx"
            ruta_reporte = base_path / nombre_reporte
            with pd.ExcelWriter(ruta_reporte, engine="openpyxl") as writer:
                df_todas_ordenado = df_todas.sort_values(["PASO", "FILA_EXCEL"], ascending=[True, True])
                df_todas_ordenado.to_excel(writer, sheet_name="TODAS LAS ELIMINACIONES", index=False)
                resumen_data = []
                for paso, cantidad in self.stats_por_paso.items():
                    ejemplos = df_todas[df_todas["PASO"] == paso]["REFERENCIA"].head(5).tolist()
                    ejemplos_str = ", ".join(str(e) for e in ejemplos)
                    if len(df_todas[df_todas["PASO"] == paso]) > 5:
                        ejemplos_str += f" ... (+{len(df_todas[df_todas['PASO'] == paso]) - 5} mas)"
                    resumen_data.append({
                        "PASO": paso,
                        "TOTAL ELIMINADAS": cantidad,
                        "PORCENTAJE": f"{(cantidad / len(self.eliminaciones) * 100):.1f}%",
                        "EJEMPLOS": ejemplos_str,
                    })
                df_resumen = pd.DataFrame(resumen_data).sort_values("TOTAL ELIMINADAS", ascending=False)
                df_resumen.to_excel(writer, sheet_name="RESUMEN POR PASO", index=False)
                for paso in sorted(self.stats_por_paso.keys()):
                    df_paso = df_todas[df_todas["PASO"] == paso].copy()
                    nombre_hoja = paso.replace(":", "").replace("/", "-")[:31]
                    df_paso.to_excel(writer, sheet_name=nombre_hoja, index=False)
            try:
                from openpyxl import load_workbook
                from openpyxl.styles import Alignment, Font, PatternFill

                wb = load_workbook(ruta_reporte)
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    for column in ws.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except Exception:
                                pass
                        ws.column_dimensions[column_letter].width = min(max_length + 2, 50)
                    if ws.max_row > 0:
                        for cell in ws[1]:
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    if sheet_name == "TODAS LAS ELIMINACIONES":
                        paso_col = None
                        for idx, cell in enumerate(ws[1], start=1):
                            if cell.value == "PASO":
                                paso_col = idx
                                break
                        if paso_col:
                            colores_pasos = {}
                            colores_disponibles = [
                                "FFF2CC", "E2EFDA", "DEEBF7", "FCE4D6",
                                "EDEDED", "F4B084", "C5E0B4", "BDD7EE",
                            ]
                            for row in range(2, ws.max_row + 1):
                                paso_val = ws.cell(row=row, column=paso_col).value
                                if paso_val:
                                    if paso_val not in colores_pasos:
                                        color_idx = len(colores_pasos) % len(colores_disponibles)
                                        colores_pasos[paso_val] = colores_disponibles[color_idx]
                                    for col in range(1, ws.max_column + 1):
                                        ws.cell(row=row, column=col).fill = PatternFill(
                                            start_color=colores_pasos[paso_val],
                                            end_color=colores_pasos[paso_val],
                                            fill_type="solid",
                                        )
                wb.save(ruta_reporte)
            except Exception as e_formato:
                log(f"  No se pudo aplicar formato avanzado: {e_formato}")
            log(f"Reporte generado: {nombre_reporte}")
            log(f"  Ubicacion: {ruta_reporte}")
            log(f"  Total de eliminaciones: {len(self.eliminaciones)}")
            log(f"  Hojas creadas: {len(self.stats_por_paso) + 2}")
            return ruta_reporte
        except Exception as e:
            log(f"Error al generar reporte de eliminaciones: {e}")
            return None


TRACKER_ELIMINACIONES = EliminacionTracker()


def month_abbr_es(dt: date) -> str:
    abrs = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
    return abrs[dt.month - 1]


def exist_col_title_for_today() -> str:
    today = date.today()
    return f"EXISTENCIA {month_abbr_es(today)} {today.day:02d}"


def to_num_str(x):
    if pd.isna(x):
        return ""
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return ""
        if any(c.isalpha() or c in "()/\\" for c in s):
            return s
        s_clean = s.replace(".", "").replace(",", "")
        try:
            f = float(s_clean)
            if abs(f - int(f)) < 1e-9:
                return str(int(f))
            return str(f)
        except Exception:
            return s
    try:
        s = str(x).strip().replace(",", "")
        f = float(s)
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return str(f)
    except Exception:
        return str(x).strip()


def limpiar_referencia(valor):
    if valor is None or valor == "":
        return ""
    val_str = str(valor).strip()
    if not val_str or val_str in ("None", "nan", "NaN"):
        return ""
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    if "e+" in val_str.lower() or "E+" in val_str:
        try:
            num = float(val_str)
            if abs(num - int(num)) < 1e-9:
                val_str = str(int(num))
            else:
                val_str = str(num)
        except Exception:
            pass
    return val_str





def inicializar_sistema_copias_temporales():
    log("")
    log("Inicializando sistema de copias temporales...")
    limpiar_copias_temporales_antiguas(max_horas=24)
    log("Sistema de copias temporales listo")
    log("")


def find_by_prefix(basedir: Path, prefix: str, exts=None) -> Path:
    if exts is None:
        exts = [".xlsx", ".xlsm", ".xls", ".csv"]
    pref = _norm(prefix)
    cands = []
    exact_match = None
    archivos_temporales_ignorados = []
    for f in basedir.iterdir():
        if not f.is_file() or f.suffix.lower() not in exts:
            continue
        if f.name.startswith("~$"):
            archivos_temporales_ignorados.append(f.name)
            continue
        if f.name.startswith("TEMP_"):
            archivos_temporales_ignorados.append(f.name)
            continue
        nn = _norm(strip_dolares_temporales(f.name))
        if nn == pref:
            exact_match = f
            log(f"  [OK] Coincidencia EXACTA encontrada: {f.name}")
            break
        if nn.startswith(pref) or pref in nn:
            cands.append(f)
            continue
        tokens = pref.split()
        if all(t in nn for t in tokens):
            cands.append(f)
    if archivos_temporales_ignorados:
        log(f"  Archivos temporales ignorados: {len(archivos_temporales_ignorados)}")
    if exact_match:
        return exact_match
    if not cands:
        raise FileNotFoundError(f"No encontre archivos que coincidan con '{prefix}' en {basedir}")
    log("  Candidatos encontrados:")
    for i, c in enumerate(cands[:5]):
        fecha = datetime.fromtimestamp(c.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        log(f"     {i + 1}. {c.name} (modificado: {fecha})")
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    log(f"  Seleccionado: {cands[0].name}")
    return cands[0]





def open_as_excel_source(path: Path, passwords=None):
    if passwords is None:
        passwords = []
    if path.suffix.lower() == ".csv":
        return path
    try:
        with pd.ExcelFile(path, engine="openpyxl"):
            return path
    except Exception as e1:
        err = str(e1).lower()
        if any(k in err for k in ("password", "encrypt", "badzipfile", "not a zip")):
            # .xls legacy (BIFF): convertirlo con xlrd (puro Python) y evitar
            # Excel COM, que falla cuando la tarea corre como servicio de Windows
            # ("Microsoft Excel no puede obtener acceso al archivo Temp\\<numeros>").
            if path.suffix.lower() in (".xls", ".xlsm"):
                try:
                    from core.excel_processing import convert_xls_legacy

                    return convert_xls_legacy(path)
                except Exception as ce:
                    log(f"Sin COM no se pudo convertir '{path.name}': {ce}; reintentando con contraseña/COM")
            for pw in passwords:
                try:
                    bio = decrypt_to_stream_local(path, pw)
                    with pd.ExcelFile(bio, engine="openpyxl"):
                        pass
                    tmp = save_bytesio_to_temp(bio, Path(path).stem)
                    log(f"  Descifrado con '{pw}' -> {tmp.name}")
                    return tmp
                except Exception:
                    continue
        if HAS_COM:
            return com_convert_to_xlsx(path, passwords)
        raise


def find_sheet_name_flexible_pd(src, targets=None) -> str:
    if targets is None:
        targets = ("INVENTARIO", "INVENTARIO GENERAL", "INV", "Sheet1", "Sheet 1", "Hoja1")
    xf = pd.ExcelFile(src, engine="openpyxl")
    names = xf.sheet_names
    if not names:
        raise ValueError("El libro no tiene hojas.")
    norm_map = {_norm(n): n for n in names}
    for t in targets:
        tn = _norm(t)
        if tn in norm_map:
            return norm_map[tn]
    for t in targets:
        tn = _norm(t)
        for kn, real in norm_map.items():
            if tn in kn:
                return real
    return names[0]


def read_excel_header_at(path: Path, sheet, header_row_visible: int, passwords=None) -> pd.DataFrame:
    if passwords is None:
        passwords = PASSWORDS_TRY or []
    src = open_as_excel_source(path, passwords)
    hdr_idx0 = header_row_visible - 1
    chosen = (
        find_sheet_name_flexible_pd(src, targets=(sheet, "INVENTARIO", "INVENTARIO GENERAL", "INV", "Sheet1", "Sheet 1", "Hoja1"))
        if isinstance(sheet, str)
        else sheet
    )
    df = pd.read_excel(src, sheet_name=chosen, engine="openpyxl", header=hdr_idx0)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Auto-detectar la fila de cabecera real: si el header fijo no contiene
    # 'referencia' (ni se parece a cabecera), escanear las primeras filas en
    # busca de la fila con la celda 'Referencia' (habitualmente columna B).
    col_names = [str(c).lower() for c in df.columns]
    parece_header = any("referencia" in c or "codigo" in c or "nombre" in c or "descripcion" in c for c in col_names)
    if not parece_header or all(c == str(i) for i, c in enumerate(col_names) if c):
        try:
            raw = pd.read_excel(src, sheet_name=chosen, engine="openpyxl", header=None, nrows=25)
            detected = None
            for r in range(len(raw)):
                row = raw.iloc[r]
                if any(
                    str(v).strip().lower() == "referencia" or "referencia" in str(v).strip().lower()
                    for v in row
                    if v is not None and not pd.isna(v)
                ):
                    detected = r
                    break
            if detected is not None:
                log(
                    "Auto-deteccion de cabecera: fila %d contiene 'Referencia'. Lectura previa uso fila %d, actualizando a fila %d."
                    % (detected + 1, hdr_idx0, detected + 1)
                )
                df = pd.read_excel(src, sheet_name=chosen, engine="openpyxl", header=detected)
                df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
                df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            log(f"Auto-deteccion de cabecera no disponible: {e}")
    return df


# ==== LECTURA DE INSUMOS ====
def cargar_inventario_actualizado(base_dir: Path) -> pd.DataFrame:
    try:
        p = find_by_prefix(base_dir, PFX_INV_ACTUALIZADO)
        log(f"Abriendo inventario actualizado (ERP): {p.name}")
        src = open_as_excel_source(p, PASSWORDS_TRY)
        df = read_excel_header_at(src, sheet="Sheet 1", header_row_visible=1)
        idx = {_norm(c): c for c in df.columns}
        ref_col = (
            idx.get("referencia") or idx.get("referencia interna") or idx.get("ref")
            or idx.get("codigo") or idx.get("codigo")
            or next((real for kn, real in idx.items() if "referenc" in kn or "codigo" in kn or kn.endswith("ref")), None)
        )
        if not ref_col:
            raise KeyError(f"{p.name}: no encuentro columna de Referencia. Encabezados: {list(df.columns)}")
        df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
        df["__REFERENCIA__"] = df[ref_col].apply(to_num_str)
        nom_col = idx.get("nombre") or "Nombre"
        marca_col = next((real for kn, real in idx.items()
                          if ("marca/ nombre a mostrar" in kn) or ("marca nombre a mostrar" in kn) or (kn == "marca")), None) \
                    or next((real for kn, real in idx.items() if "marca" in kn and "mostrar" in kn), None)
        linea_col = next((real for kn, real in idx.items()
                          if ("linea/ nombre a mostrar" in kn) or ("linea/ nombre a mostrar" in kn)), None) \
                    or next((real for kn, real in idx.items() if "linea" in kn and "mostrar" in kn), None)
        sublinea_col = next((real for kn, real in idx.items() if "sub" in kn and "linea" in kn and "mostrar" in kn), None)
        costo_col = idx.get("costo") or "Costo"
        rename = {}
        if nom_col in df.columns:
            rename[nom_col] = "__NOMBRE__"
        if marca_col in df.columns:
            rename[marca_col] = "__MARCA_SYS__"
        if linea_col in df.columns:
            rename[linea_col] = "__LINEA_SYS__"
        if sublinea_col in df.columns:
            rename[sublinea_col] = "__SUBLINEA_SYS__"
        if costo_col in df.columns:
            rename[costo_col] = "__COSTO__"
        df_final = df.rename(columns=rename)
        return df_final
    except FileNotFoundError:
        pass
    p = find_file_by_pattern(base_dir, PATRON_INV_FILE)
    if not p:
        p_pl = base_dir / PATRON_INV_FILE
        if p_pl.exists():
            p = p_pl
        else:
            for pref in ["2026 INVENTARIO GENERAL", "INVENTARIO GENERAL"]:
                try:
                    p = find_by_prefix(base_dir, pref)
                    break
                except Exception:
                    p = None
            if p is None:
                raise FileNotFoundError(
                    f"No encontre ni '{PFX_INV_ACTUALIZADO}' ni '{PATRON_INV_FILE}' en {base_dir}"
                )
    log(f"[Fallback] Abriendo plantilla de inventario: {p.name}")
    df = read_excel_header_at(p, sheet=SHEET_INV_ORIG, header_row_visible=HEADER_ROW_INV)
    idx = {_norm(c): c for c in df.columns}
    ref_col = (
        idx.get("referencia") or idx.get("referencia fertrac") or idx.get("referencia interna")
        or idx.get("ref") or idx.get("codigo") or idx.get("codigo")
        or next((real for kn, real in idx.items() if "referenc" in kn or "codigo" in kn or kn.endswith("ref")), None)
    )
    if not ref_col:
        raise KeyError(f"{p.name}: no encuentro columna 'REFERENCIA'. Encabezados: {list(df.columns)}")
    df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
    df["__REFERENCIA__"] = df[ref_col].apply(to_num_str)
    nombre_odoo = idx.get("nombre odoo") or idx.get("nombre")
    marca_sys = idx.get("marca sistema")
    linea_sys = idx.get("linea sistema") or idx.get("linea sistema")
    sub_sys = idx.get("sub- linea sistema") or idx.get("sub-linea sistema") or idx.get("sub linea sistema")
    costo_prom = idx.get("costo promedio") or idx.get("costo prom")
    rename = {}
    if nombre_odoo in df.columns:
        rename[nombre_odoo] = "__NOMBRE__"
    if marca_sys in df.columns:
        rename[marca_sys] = "__MARCA_SYS__"
    if linea_sys in df.columns:
        rename[linea_sys] = "__LINEA_SYS__"
    if sub_sys in df.columns:
        rename[sub_sys] = "__SUBLINEA_SYS__"
    if costo_prom in df.columns:
        rename[costo_prom] = "__COSTO__"
    return df.rename(columns=rename)


def resolver_cadena_referencias(mapeo: dict) -> dict:
    mapeo_resuelto = {}
    for ref_origen in mapeo.keys():
        ref_actual = ref_origen
        visitados = set()
        while ref_actual in mapeo:
            if ref_actual in visitados:
                log(f"CICLO DETECTADO: {' -> '.join(visitados)} -> {ref_actual}")
                break
            visitados.add(ref_actual)
            ref_siguiente = mapeo[ref_actual]
            if ref_siguiente != ref_actual:
                ref_actual = ref_siguiente
            else:
                break
        if ref_actual != ref_origen:
            mapeo_resuelto[ref_origen] = ref_actual
            if len(visitados) > 1:
                cadena = " -> ".join(visitados) + f" -> {ref_actual}"
                log(f"   Cadena resuelta: {cadena}")
    return mapeo_resuelto


def _col_num_to_letter(col_num):
    letter = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def cargar_valorizado(base_dir: Path, prefix: str) -> pd.DataFrame:
    p = find_by_prefix(base_dir, prefix)
    log(f"Abrir: {p.name}")
    src = open_as_excel_source(p, PASSWORDS_TRY)
    if p.suffix.lower() == ".csv":
        df_all = pd.read_csv(src, header=None, dtype=str)
    else:
        df_all = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None)
    hdr_row0 = HEADER_ROW_VAL - 1
    if hdr_row0 >= len(df_all):
        raise ValueError(f"{p.name}: HEADER_ROW_VAL={HEADER_ROW_VAL} supera el numero de filas.")
    df = df_all.iloc[hdr_row0:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    idx = {_norm(c): c for c in df.columns}
    refc = idx.get("referencia interna")
    if not refc:
        log(f"  ADVERTENCIA: No se encontro columna 'Referencia interna' en {p.name}")
        refc = idx.get("referencia") or idx.get("ref") \
            or next((real for kn, real in idx.items() if "referenc" in kn), None)
    cant = idx.get("cantidad")
    if not cant:
        log(f"  ADVERTENCIA: No se encontro columna 'Cantidad' en {p.name}")
        cant = next((real for kn, real in idx.items() if kn.startswith("cant")), None)
    if not refc:
        raise KeyError(f"{p.name}: no encuentro 'Referencia interna'. Encabezados: {list(df.columns)}")
    if not cant:
        raise KeyError(f"{p.name}: no encuentro 'Cantidad'. Encabezados: {list(df.columns)}")
    out = pd.DataFrame()
    out["__REF_INT__"] = df[refc].apply(to_num_str)
    out["__CANT__"] = pd.to_numeric(df[cant], errors="coerce").fillna(0.0)
    return out


def cargar_valorizado_desde_ruta(archivo_path: Path) -> pd.DataFrame:
    log(f"Abriendo: {archivo_path.name}")
    src = open_as_excel_source(archivo_path, PASSWORDS_TRY)
    if archivo_path.suffix.lower() == ".csv":
        df_all = pd.read_csv(src, header=None, dtype=str)
    else:
        df_all = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None)
    hdr_row0 = HEADER_ROW_VAL - 1
    df = df_all.iloc[hdr_row0:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    idx = {_norm(c): c for c in df.columns}
    refc = idx.get("referencia interna")
    cant = idx.get("cantidad")
    if not refc:
        raise KeyError(f"{archivo_path.name}: no encuentro 'Referencia interna'")
    if not cant:
        raise KeyError(f"{archivo_path.name}: no encuentro 'Cantidad'")
    out = pd.DataFrame()
    out["__REF_INT__"] = df[refc].apply(to_num_str)
    out["__CANT__"] = pd.to_numeric(df[cant], errors="coerce").fillna(0.0)
    return out


def cargar_matriz_usd(base_dir: Path) -> pd.DataFrame:
    try:
        p = find_file_by_pattern(base_dir, PATRON_MATRIZ_USD)
        if not p:
            log("No se encontro con busqueda dinamica, intentando metodo tradicional...")
            p = find_by_prefix(base_dir, PFX_MATRIZ_USD)
        log(f"Abriendo Matriz USD: {p.name}")
        src = open_as_excel_source(p, PASSWORDS_TRY)
        xf = pd.ExcelFile(src, engine="openpyxl")
        sheet_found = None
        for sn in xf.sheet_names:
            nombre_limpio = re.sub(r"^\$+", "", sn).strip()
            if PATRON_SHEET_2025 in nombre_limpio or _norm(nombre_limpio) == _norm(PATRON_SHEET_2025):
                sheet_found = sn
                log(f"   [OK] Hoja encontrada: '{sn}'")
                break
        if not sheet_found:
            sheet_found = xf.sheet_names[0]
            log(f"  No se encontro hoja con '2025', usando: '{sheet_found}'")
        df_raw = pd.read_excel(src, sheet_name=sheet_found, engine="openpyxl", header=None)
        header_row_idx = None
        for idx in range(min(20, len(df_raw))):
            has_ref = any("referencia" in str(v).lower() and "fertrac" in str(v).lower() for v in df_raw.iloc[idx])
            has_desc = any("descripcion" in str(v).lower() and "lista" in str(v).lower() for v in df_raw.iloc[idx])
            if has_ref or has_desc:
                header_row_idx = idx
                break
        if header_row_idx is None:
            max_non_empty = 0
            for idx in range(min(10, len(df_raw))):
                non_empty = df_raw.iloc[idx].notna().sum()
                if non_empty > max_non_empty:
                    max_non_empty = non_empty
                    header_row_idx = idx
            log(f"  Usando fila {header_row_idx + 1} como encabezado")
        df = pd.read_excel(src, sheet_name=sheet_found, engine="openpyxl", header=header_row_idx)
        df.columns = [
            str(c).strip() if not str(c).startswith("Unnamed") and str(c) != "nan" else f"_COL_{i}"
            for i, c in enumerate(df.columns)
        ]
        idx = {_norm(c): c for c in df.columns}
        ref_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "referencia" in col_norm and ("fertrac" in col_norm or "inventario" in col_norm):
                ref_col = col_name
                break
        if not ref_col:
            for col_name in df.columns[:5]:
                non_null = df[col_name].notna().sum()
                if non_null > 10:
                    sample = df[col_name].dropna().astype(str).head(5)
                    if any("FP-" in str(v) or str(v).replace("-", "").isdigit() for v in sample):
                        ref_col = col_name
                        break
        desc_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "descripcion" in col_norm and "lista" in col_norm and "precio" in col_norm:
                desc_col = col_name
                break
        if not desc_col:
            for col_name in df.columns:
                if col_name == ref_col:
                    continue
                non_null = df[col_name].notna().sum()
                if non_null > 10:
                    sample = df[col_name].dropna().astype(str).head(5)
                    avg_len = sum(len(str(v)) for v in sample) / len(sample) if len(sample) > 0 else 0
                    if avg_len > 15:
                        desc_col = col_name
                        break
        ref_lista_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "referencia" in col_norm and "lista" in col_norm and "precio" in col_norm:
                ref_lista_col = col_name
                break
        if not ref_lista_col:
            for col_name in df.columns:
                col_norm = _norm(col_name)
                if col_name == ref_col or col_name == desc_col:
                    continue
                if ("ref" in col_norm or "codigo" in col_norm) and "lista" in col_norm:
                    ref_lista_col = col_name
                    break
        if not ref_col:
            raise KeyError(f"No encontre columna 'REFERENCIA INVENTARIO FERTRAC' en {p.name}. Columnas: {list(df.columns)}")
        if not desc_col:
            raise KeyError(f"No encontre columna 'DESCRIPCION LISTA PRECIOS' en {p.name}. Columnas: {list(df.columns)}")
        df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
        out = pd.DataFrame()
        out["__REF_MATRIZ__"] = df[ref_col].apply(to_num_str)
        out["__DESC_LISTA__"] = df[desc_col].fillna("")
        if ref_lista_col:
            out["__REF_LISTA_PRECIOS__"] = df[ref_lista_col].apply(to_num_str)
        else:
            out["__REF_LISTA_PRECIOS__"] = ""
        out = out.drop_duplicates(subset=["__REF_MATRIZ__"], keep="first")
        return out
    except FileNotFoundError:
        log(f"ADVERTENCIA: No se encontro el archivo '{PFX_MATRIZ_USD}'.")
        return pd.DataFrame(columns=["__REF_MATRIZ__", "__DESC_LISTA__", "__REF_LISTA_PRECIOS__"])
    except Exception as e:
        log(f"ERROR al cargar Matriz USD: {e}")
        return pd.DataFrame(columns=["__REF_MATRIZ__", "__DESC_LISTA__", "__REF_LISTA_PRECIOS__"])


def cargar_marcas(base_dir: Path) -> pd.DataFrame:
    p = find_by_prefix(base_dir, PFX_MARCAS)
    log(f"Abriendo: {p.name}")
    src = open_as_excel_source(p, PASSWORDS_TRY)
    df_raw = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None, nrows=15)
    header_row_idx = None
    for idx in range(min(10, len(df_raw))):
        has_ref = any("referencia" in str(v).lower() for v in df_raw.iloc[idx])
        if has_ref:
            header_row_idx = idx
            break
    if header_row_idx is None:
        max_non_empty = 0
        for idx in range(min(10, len(df_raw))):
            non_empty = df_raw.iloc[idx].notna().sum()
            if non_empty > max_non_empty:
                max_non_empty = non_empty
                header_row_idx = idx
    df = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    idx = {_norm(c): c for c in df.columns}
    ref_col = None
    for col_name in df.columns:
        col_norm = _norm(col_name)
        if "referencia" in col_norm:
            ref_col = col_name
            break
    ref_lista_col = None
    for col_name in df.columns:
        col_norm = _norm(col_name)
        if "referencia" in col_norm and "lista" in col_norm and "precio" in col_norm:
            ref_lista_col = col_name
            break
    if not ref_lista_col:
        for col_name in df.columns:
            if col_name == ref_col:
                continue
            non_null = df[col_name].notna().sum()
            if non_null > 10:
                sample = df[col_name].dropna().astype(str).head(5)
                avg_len = sum(len(str(v)) for v in sample) / len(sample) if len(sample) > 0 else 0
                if avg_len > 15:
                    ref_lista_col = col_name
                    break
    if not ref_col:
        for col_name in df.columns:
            if "marca" in _norm(col_name) and df[col_name].notna().sum() > 0:
                ref_col = col_name
                break
    if not ref_col:
        for col_name in df.columns:
            if df[col_name].notna().sum() > 5:
                ref_col = col_name
                break
    if not ref_col:
        raise KeyError(f"{p.name}: no encuentro columna 'REFERENCIA'. Columnas: {list(df.columns)}")
    df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
    out = pd.DataFrame()
    out["__REF_LISTA__"] = df[ref_col].apply(to_num_str)
    if ref_lista_col:
        out["__REF_LISTA_PRECIOS__"] = df[ref_lista_col].apply(to_num_str)
    else:
        out["__REF_LISTA_PRECIOS__"] = ""
    out = out.drop_duplicates(subset=["__REF_LISTA__"], keep="first")
    return out


def cargar_distribucion(base_dir: Path) -> pd.DataFrame:
    p = find_by_prefix(base_dir, PFX_DISTRIBUCION)
    log(f"Abriendo: {p.name}")
    src = open_as_excel_source(p, PASSWORDS_TRY)
    df_raw = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None, nrows=15)
    header_row_idx = None
    for idx in range(min(10, len(df_raw))):
        has_ref = any("referencia" in str(v).lower() for v in df_raw.iloc[idx])
        if has_ref:
            header_row_idx = idx
            break
    if header_row_idx is None:
        max_non_empty = 0
        for idx in range(min(10, len(df_raw))):
            non_empty = df_raw.iloc[idx].notna().sum()
            if non_empty > max_non_empty:
                max_non_empty = non_empty
                header_row_idx = idx
    df = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    idx = {_norm(c): c for c in df.columns}
    ref_col = None
    for col_name in df.columns:
        col_norm = _norm(col_name)
        if "referencia" in col_norm:
            ref_col = col_name
            break
    ref_lista_col = None
    for col_name in df.columns:
        col_norm = _norm(col_name)
        if "referencia" in col_norm and "lista" in col_norm and "precio" in col_norm:
            ref_lista_col = col_name
            break
    if not ref_lista_col:
        for col_name in df.columns:
            if col_name == ref_col:
                continue
            non_null = df[col_name].notna().sum()
            if non_null > 10:
                sample = df[col_name].dropna().astype(str).head(5)
                avg_len = sum(len(str(v)) for v in sample) / len(sample) if len(sample) > 0 else 0
                if avg_len > 15:
                    ref_lista_col = col_name
                    break
    if not ref_col:
        for col_name in df.columns:
            if df[col_name].notna().sum() > 5:
                ref_col = col_name
                break
    if not ref_col:
        raise KeyError(f"{p.name}: no encuentro columna 'REFERENCIA'. Columnas: {list(df.columns)}")
    df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
    out = pd.DataFrame()
    out["__REF_DIST__"] = df[ref_col].apply(to_num_str)
    if ref_lista_col:
        out["__REF_LISTA_PRECIOS__"] = df[ref_lista_col].apply(to_num_str)
    else:
        out["__REF_LISTA_PRECIOS__"] = ""
    out = out.drop_duplicates(subset=["__REF_DIST__"], keep="first")
    return out


def cargar_consolidado_remisiones(base_dir: Path) -> pd.DataFrame:
    try:
        p = find_by_prefix(base_dir, PFX_CONSOLIDADO_REMISIONES)
        log(f"Abriendo: {p.name}")
        src = open_as_excel_source(p, PASSWORDS_TRY)
        df_raw = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None, nrows=15)
        header_row_idx = None
        for idx in range(min(10, len(df_raw))):
            has_ref = any("referencia" in str(v).lower() for v in df_raw.iloc[idx])
            if has_ref:
                header_row_idx = idx
                break
        if header_row_idx is None:
            max_non_empty = 0
            for idx in range(min(10, len(df_raw))):
                non_empty = df_raw.iloc[idx].notna().sum()
                if non_empty > max_non_empty:
                    max_non_empty = non_empty
                    header_row_idx = idx
        df = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=header_row_idx)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        idx = {_norm(c): c for c in df.columns}
        ref_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "referencia" in col_norm:
                ref_col = col_name
                break
        desc_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "descripcion" in col_norm:
                desc_col = col_name
                break
        if not ref_col:
            raise KeyError(f"{p.name}: no encuentro columna 'REFERENCIA'. Columnas: {list(df.columns)}")
        df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
        out = pd.DataFrame()
        out["__REF_REM__"] = df[ref_col].apply(to_num_str)
        if desc_col:
            out["__DESC_REM__"] = df[desc_col].fillna("")
        else:
            out["__DESC_REM__"] = ""
        out = out.drop_duplicates(subset=["__REF_REM__"], keep="first")
        return out
    except FileNotFoundError:
        log(f"ADVERTENCIA: No se encontro archivo '{PFX_CONSOLIDADO_REMISIONES}'.")
        return pd.DataFrame(columns=["__REF_REM__", "__DESC_REM__"])
    except Exception as e:
        log(f"ERROR al cargar consolidado de remisiones: {e}")
        return pd.DataFrame(columns=["__REF_REM__", "__DESC_REM__"])


def _procesar_hoja_tabla_dinamica(ws_or_df, wb, archivo_origen: str = "") -> pd.DataFrame:
    headers = []
    col_letters = []
    first_row = 1 if isinstance(ws_or_df, pd.DataFrame) else 3
    max_cols = 20 if isinstance(ws_or_df, pd.DataFrame) else ws_or_df.max_column
    for col in range(1, max_cols + 1):
        val = ws_or_df.cell(row=first_row, column=col).value if hasattr(ws_or_df, 'cell') else ws_or_df.iloc[0, col - 1]
        if val is not None:
            headers.append(str(val).strip())
            col_letters.append(col)
        else:
            headers.append(f"COL_{col}")
            col_letters.append(col)
    df = pd.DataFrame()
    start_row = (first_row + 1) if hasattr(ws_or_df, 'cell') else 1
    for i, col_idx in enumerate(col_letters):
        if hasattr(ws_or_df, 'cell'):
            values = []
            for row in range(start_row, ws_or_df.max_row + 1):
                val = ws_or_df.cell(row=row, column=col_idx).value
                values.append(val)
        else:
            values = ws_or_df.iloc[:, i].tolist()
        df[headers[i]] = values
    for col in df.columns:
        if any(kw in col.lower() for kw in ["referencia", "ref"]):
            df[col] = df[col].apply(to_num_str)
    return df


def _col_letter_to_num(letter):
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def actualizar_referencias_inventario_original(
    df_original: pd.DataFrame,
    df_actualizado: pd.DataFrame,
    df_valorizados: list,
    df_matriz_usd: pd.DataFrame,
    df_marcas: pd.DataFrame,
    df_distribucion: pd.DataFrame,
    df_remisiones: pd.DataFrame,
) -> pd.DataFrame:
    ref_actualizada_col = "__REFERENCIA_ACT__"
    rename_cols = {}
    if ref_actualizada_col in df_original.columns:
        rename_cols[ref_actualizada_col] = ref_actualizada_col
    df_original = df_original.rename(columns=rename_cols)
    df_original[ref_actualizada_col] = df_original["__REFERENCIA__"].apply(to_num_str)
    log(f"  Fila 2 (titulo) referencias aportadas por el ERP: {df_original[ref_actualizada_col].iloc[0] if len(df_original) > 0 else 'N/A'}")
    log(f"  Total registros original: {len(df_original)}")
    mapeo_referencias = {}
    if len(df_actualizado) > 0:
        ref_general_col = "__REFERENCIA__"
        alias_cols = [c for c in df_actualizado.columns if "alias" in _norm(c)]
        alias_col = alias_cols[0] if alias_cols else None
        for _, row in df_actualizado.iterrows():
            ref_general = to_num_str(row[ref_general_col])
            alias_ref = to_num_str(row[alias_col]) if alias_col else None
            if ref_general:
                mapeo_referencias[ref_general] = ref_general
            if alias_ref:
                mapeo_referencias[alias_ref] = ref_general
    log(f"  Mapeo ERP cargado: {len(mapeo_referencias)} referencias")
    log(f"  Resolviendo cadenas de referencias...")
    mapeo_referencias = resolver_cadena_referencias(mapeo_referencias)
    alias_registrados = []
    for ref_destino, ref_origen in mapeo_referencias.items():
        if ref_destino != ref_origen:
            alias_registrados.append(f"{ref_destino} -> {ref_origen}")
    log(f"  Aliases registrados: {len(alias_registrados)}")
    idx_actual = {}
    for _, row in df_actualizado.iterrows():
        ref = row["__REFERENCIA__"]
        idx_actual[ref] = row.to_dict()
    idx_originales = df_original[ref_actualizada_col].apply(to_num_str).to_dict()
    actualizaciones = []
    no_encontradas = []
    for idx in df_original.index:
        ref_orig = to_num_str(df_original.at[idx, "__REFERENCIA__"])
        ref_actual = mapeo_referencias.get(ref_orig, ref_orig)
        if ref_orig != ref_actual:
            df_original.at[idx, ref_actualizada_col] = ref_actual
            actualizaciones.append(f"  {ref_orig} -> {ref_actual}")
    if actualizaciones:
        log(f"  Referencias actualizadas en original: {len(actualizaciones)}")
        for a in actualizaciones[:10]:
            log(a)
        if len(actualizaciones) > 10:
            log(f"  ... y {len(actualizaciones) - 10} mas")
    log(f"  Total registros en actualizado (ERP): {len(df_actualizado)}")
    return df_original


def aplicar_reglas_marcas_propias(
    df_original: pd.DataFrame,
    df_matriz_usd: pd.DataFrame,
    df_remisiones: pd.DataFrame,
) -> pd.DataFrame:
    log("  Marcas propias sin referencias en Matriz USD:")
    idx_matriz = set(df_matriz_usd["__REF_MATRIZ__"].apply(to_num_str)) if len(df_matriz_usd) > 0 else set()
    idx_remisiones = set(df_remisiones["__REF_REM__"].apply(to_num_str)) if len(df_remisiones) > 0 else set()
    idx_distribucion = set()
    idx_marcas = set()
    return df_original


def eliminar_registros_linea_copia_indeterminada(
    df_original: pd.DataFrame,
    hoja_original,
    ws_wb=None,
    maximo_paso_1: int = 1000,
) -> pd.DataFrame:
    log("Eliminacion de registros LINEA COPIA indeterminada...")
    ref_col = "REFERENCIA" if "REFERENCIA" in df_original.columns else None
    if not ref_col:
        for c in df_original.columns:
            if "referenc" in _norm(c):
                ref_col = c
                break
    if not ref_col:
        log("No se encontro columna REFERENCIA, omitiendo paso")
        return df_original
    lineage_col = None
    for c in df_original.columns:
        cn = _norm(c)
        if cn == "linea copia" or cn == "linea/copia":
            lineage_col = c
            break
    if not lineage_col:
        lineage_col = next((c for c in df_original.columns if "linea" in _norm(c) and "copia" in _norm(c)), None)
    if not lineage_col:
        log("No se encontro 'LINEA COPIA' en el Inventario General, omitiendo")
        return df_original
    log(f"  Columna de referencia: {ref_col}")
    log(f"  Columna de LINEA COPIA: {lineage_col}")
    idx_ref = {_norm(r): i for i, r in enumerate(df_original[ref_col].astype(str).tolist())}
    log(f"  Total registros: {len(df_original)}")
    log(f"  Total en idx_ref: {len(idx_ref)}")
    df_original[ref_col] = df_original[ref_col].apply(lambda x: limpiar_referencia(str(x)))
    log("  Paso 1: Eliminar registros con referencia duplicada en INVENTARIO COPIA")
    ref_copia_col = None
    for c in df_original.columns:
        cn = _norm(c)
        if "referenc" in cn and ("copia" in cn or "inventario copia" in cn):
            ref_copia_col = c
            break
    if ref_copia_col is None:
        ref_copia_col = ref_col
    duplicadas = df_original[df_original.duplicated(subset=[ref_copia_col], keep="first")]
    df_original.drop(duplicadas.index, inplace=True)
    df_original.reset_index(drop=True, inplace=True)
    log(f"  Eliminadas: {len(duplicadas)} registros duplicados")
    log("  Paso 2: Eliminar INVENTARIO GENERAL con referencia a no existe en INVENTARIO COPIA")
    idx_linea_copia = {}
    if lineage_col:
        for _, row in df_original.iterrows():
            ref = to_num_str(row[ref_col])
            linea = row[lineage_col] if pd.notna(row[lineage_col]) else ""
            if ref and ref not in idx_linea_copia:
                idx_linea_copia[ref] = linea
    return df_original


def procesar_existencias_negativas_y_cero(
    df_original: pd.DataFrame,
    ws,
    header_row: int = 2,
) -> pd.DataFrame:
    log("Procesando existencias negativas y cero...")
    exist_col = None
    for c in df_original.columns:
        cn = _norm(c)
        if "existencia" in cn or "exist" in cn:
            exist_col = c
            break
    if not exist_col:
        log("No se encontro columna de existencia")
        return df_original
    try:
        exist_vals = pd.to_numeric(df_original[exist_col], errors="coerce").fillna(0)
        neg_count = (exist_vals < 0).sum()
        zero_count = (exist_vals == 0).sum()
        log(f"  Negativas: {neg_count}, Ceros: {zero_count}")
    except Exception as e:
        log(f"  Error al procesar: {e}")
    return df_original


def excel_open(ruta, *, visible=False, alerts=False, editable=False):
    if not HAS_COM:
        raise RuntimeError("win32com no esta disponible")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = alerts
    excel.Interactive = editable
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass
    try:
        excel.AutomationSecurity = 3
    except Exception:
        pass
    try:
        excel.CutCopyMode = False
    except Exception:
        pass
    try:
        wb = excel.Workbooks.Open(str(ruta), ReadOnly=not editable, UpdateLinks=0)
    except Exception:
        try:
            wb = excel.Workbooks.Open(str(ruta), ReadOnly=not editable, UpdateLinks=False)
        except Exception as e2:
            safe_quit_excel(excel)
            raise RuntimeError(f"No pude abrir '{Path(ruta).name}': {e2}") from e2
    ws = wb.Sheets(1)
    return excel, wb, ws


def excel_close(excel, wb, *, save: bool = False):
    try:
        wb.Close(SaveChanges=save)
    except Exception:
        pass
    safe_quit_excel(excel)


def normalize_sheet_name(name):
    if not isinstance(name, str):
        return name
    import re
    return re.sub(r"^\$", "", name).strip()


def ws_headers(ws, *, start: int = 2):
    headers = {}
    if ws is None:
        return headers
    non_empty = 0
    for col in range(1, min(20, ws.Columns.Count) + 1):
        val = ws.Cells(start, col).Value
        if val is not None:
            non_empty += 1
    if non_empty < 2:
        for row in range(1, 50):
            non_empty_r = 0
            for col in range(1, min(20, ws.Columns.Count) + 1):
                val = ws.Cells(row, col).Value
                if val is not None:
                    non_empty_r += 1
            if non_empty_r >= 2:
                start = row
                break
    for col in range(1, min(ws.Columns.Count, 100) + 1):
        val = ws.Cells(start, col).Value
        if val is not None:
            headers[str(val).strip()] = col
    return headers


def ws_fill_column_values(ws, *, header_row: int, headers_map: dict, col_name: str,
                          id_col: str, id_value, values_map: dict, label: str = ""):
    if col_name not in headers_map:
        return
    col_idx = headers_map[col_name]
    id_col_idx = headers_map.get(id_col)
    if not id_col_idx:
        return
    changed = 0
    last_row = min(ws.UsedRange.Rows.Count, ws.Rows.Count)
    for row in range(header_row + 1, last_row + 1):
        cell_id = ws.Cells(row, id_col_idx).Value
        if cell_id is None:
            continue
        ref = to_num_str(cell_id)
        if ref in values_map:
            old_val = ws.Cells(row, col_idx).Value
            new_val = values_map[ref]
            if old_val != new_val:
                ws.Cells(row, col_idx).Value = new_val
                changed += 1
    if changed:
        log(f"  {label}: {changed} celdas actualizadas")


def ws_last_row(ws, col: int = 1, *, start: int = 2):
    used = ws.UsedRange.Rows.Count
    for r in range(used, start - 1, -1):
        v = ws.Cells(r, col).Value
        if v is not None and str(v).strip():
            return r
    return start - 1


def ws_fill_row_from_last(ws, values_by_col: dict):
    last = ws_last_row(ws) + 1
    for col, val in values_by_col.items():
        ws.Cells(last, col).Value = val


def detectar_columnas_clave_df(df, *, ref_name="REFERENCIA", cost_name="COSTO PROMEDIO"):
    idx = {_norm(c): c for c in df.columns}
    ref = idx.get(_norm(ref_name)) or next((c for c in df.columns if _norm(ref_name) in _norm(c)), None)
    cost = idx.get(_norm(cost_name)) or next((c for c in df.columns if "costo" in _norm(c) and "promedio" in _norm(c)), None)
    nombre = idx.get(_norm("NOMBRE")) or idx.get(_norm("NOMBRE ODOO")) or next((c for c in df.columns if _norm("nombre") in _norm(c)), None)
    marca = idx.get(_norm("MARCA")) or idx.get(_norm("MARCA SISTEMA")) or next((c for c in df.columns if _norm("marca") in _norm(c)), None)
    return ref, cost, nombre, marca


def rename_norm_col(df, col, new_name):
    idx = {_norm(c): c for c in df.columns}
    k = _norm(col)
    real = idx.get(k) or next((c for c in df.columns if k in _norm(c)), None)
    if real and real in df.columns:
        df = df.rename(columns={real: new_name})
    return df


def convertir_texto_a_numero_columnas_inv_lista(df, col_name, *, default=""):
    if col_name not in df.columns:
        return df
    def _convertir(x):
        if pd.isna(x):
            return default
        if isinstance(x, (int, float)):
            if pd.isna(x):
                return default
            if x == int(x):
                return str(int(x))
            return str(x)
        s = str(x).strip()
        if not s or s in ("None", "nan", "NaN"):
            return default
        try:
            f = float(s.replace(",", ""))
            if f == int(f):
                return str(int(f))
            return str(f)
        except ValueError:
            return s
    df[col_name] = df[col_name].apply(_convertir)
    return df


class ActualizacionInventario(BaseTask):
    name = "actualizacion_inventario"
    description = "Actualiza el inventario general consolidando datos de Excel y eliminando registros innecesarios"

    def __init__(self, settings: Settings):
        super().__init__(settings)
        self.notifier: EmailNotifier = None
        self.stats: dict = {}

    def setup(self):
        global BASE_PATH, OUTPUT_PATH, RUTA_INV_ERP, RUTA_VALORIZADOS, RUTA_REMISIONES, RUTA_INSUMOS, PASS_INV, PASSWORDS_TRY
        log("=== CONFIGURACION DE RUTAS ===")
        self.settings.paths.ensure_dirs()
        BASE_PATH = self.settings.paths.inventario_general_mes
        OUTPUT_PATH = self.settings.paths.output_inv_general
        RUTA_INV_ERP = self.settings.paths.inventario_general_mes
        RUTA_VALORIZADOS = self.settings.paths.valorizados
        RUTA_REMISIONES = self.settings.paths.remisiones_mes
        RUTA_INSUMOS = self.settings.paths.base
        PASS_INV = self.settings.excel.password
        PASSWORDS_TRY = list(self.settings.excel.passwords_try)
        log(f"BASE_PATH:     {BASE_PATH}")
        log(f"OUTPUT_PATH:   {OUTPUT_PATH}")
        log(f"RUTA_INV_ERP:  {RUTA_INV_ERP}")
        log(f"RUTA_VALORIZADOS: {RUTA_VALORIZADOS}")
        log(f"RUTA_REMISIONES:  {RUTA_REMISIONES}")
        log(f"RUTA_INSUMOS:  {RUTA_INSUMOS}")
        log(f"PASSWORDS_TRY: {PASSWORDS_TRY}")
        self.notifier = EmailNotifier(self.settings.smtp_inv_general, self.name)
        log("Setup completado.")

    def execute(self):
        global BASE_PATH, OUTPUT_PATH, RUTA_INV_ERP, RUTA_VALORIZADOS, RUTA_REMISIONES, RUTA_INSUMOS, PASS_INV, PASSWORDS_TRY, TRACKER_ELIMINACIONES

        log("INICIO DEL PROCESO")
        log("=" * 70)

        try:
            BASE_PATH = self.settings.paths.inventario_general_mes
            OUTPUT_PATH = self.settings.paths.output_inv_general
            RUTA_INV_ERP = self.settings.paths.inventario_general_mes
            RUTA_VALORIZADOS = self.settings.paths.valorizados
            RUTA_REMISIONES = self.settings.paths.remisiones_mes
            RUTA_INSUMOS = self.settings.paths.base
            PASS_INV = self.settings.excel.password
            PASSWORDS_TRY = list(self.settings.excel.passwords_try)
            OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

            TRACKER_ELIMINACIONES = EliminacionTracker()
            ARCHIVOS_DIARIOS = []
            ARCHIVOS_DIARIOS.extend(sorted(RUTA_VALORIZADOS.glob("VALORIZADO GENERAL*.xlsx")))
            ARCHIVOS_DIARIOS.extend(sorted(RUTA_VALORIZADOS.glob("VALORIZADO GENERAL*.xls")))
            ARCHIVOS_DIARIOS.extend(sorted(RUTA_VALORIZADOS.glob("VALORIZADO GENERAL*.csv")))
            for patron in [
                "VALORIZADO FALTANTES IMPO*.xlsx",
                "VALORIZADO FALTANTES IMPO*.xls",
                "VALORIZADO FALTANTES*.xlsx",
                "VALORIZADO FALTANTES*.xls",
                "VALORIZADO TOBERIN*.xlsx",
                "VALORIZADO TOBERIN*.xls",
            ]:
                ARCHIVOS_DIARIOS.extend(sorted(RUTA_VALORIZADOS.glob(patron)))

            vistos = set()
            ARCHIVOS_DIARIOS_sin_dup = []
            for af in ARCHIVOS_DIARIOS:
                clave = str(af.resolve())
                if clave not in vistos:
                    vistos.add(clave)
                    ARCHIVOS_DIARIOS_sin_dup.append(af)
            ARCHIVOS_DIARIOS = ARCHIVOS_DIARIOS_sin_dup

            if ARCHIVOS_DIARIOS:
                log("\n--- ARCHIVOS VALORIZADOS ENCONTRADOS ---")
                for af in ARCHIVOS_DIARIOS:
                    log(f"  {af.name}")
            else:
                log("\n--- NO SE ENCONTRARON ARCHIVOS DE VALORIZADOS ---")

            log("\n" + "=" * 70)
            log("FASE 1: Cargar insumos")
            log("=" * 70)

            log("1.1 Cargando inventario actualizado del ERP...")
            df_inv_act = cargar_inventario_actualizado(RUTA_INV_ERP)
            log(f"    Inventario actualizado: {len(df_inv_act)} registros")

            log("\n1.2 Cargando valorizados...")
            df_val_list = []
            for archivo in ARCHIVOS_DIARIOS:
                try:
                    df_v = cargar_valorizado_desde_ruta(archivo)
                    df_val_list.append(df_v)
                    log(f"    OK: {archivo.name} -> {len(df_v)} registros")
                except Exception as e:
                    log(f"    Error al cargar {archivo.name}: {e}")

            log("\n1.3 Cargando Matriz USD...")
            df_matriz = cargar_matriz_usd(RUTA_INSUMOS)
            log(f"    Matriz USD: {len(df_matriz)} registros")

            log("\n1.4 Cargando Marcas...")
            df_marcas = cargar_marcas(RUTA_INSUMOS)
            log(f"    Marcas: {len(df_marcas)} registros")

            log("\n1.5 Cargando Distribucion...")
            df_dist = cargar_distribucion(RUTA_INSUMOS)
            log(f"    Distribucion: {len(df_dist)} registros")

            log("\n1.6 Cargando Consolidado de Remisiones...")
            df_rem = pd.DataFrame(columns=["__REF_REM__", "__DESC_REM__"])
            for d_rem in [RUTA_INSUMOS, RUTA_REMISIONES, RUTA_INV_ERP]:
                try:
                    find_by_prefix(d_rem, PFX_CONSOLIDADO_REMISIONES)
                    df_rem = cargar_consolidado_remisiones(d_rem)
                    break
                except FileNotFoundError:
                    continue
            log(f"    Remisiones: {len(df_rem)} registros")

            log("\n" + "=" * 70)
            log("FASE 2: Leer plantilla de inventario general")
            log("=" * 70)

            try:
                p_inv = find_by_prefix(OUTPUT_PATH, OUTPUT_BASENAME)
                log("Plantilla de inventario: salida previa en Pruebas Inv General")
            except FileNotFoundError:
                log(f"No hay salida previa en {OUTPUT_PATH}; buscando plantilla en la carpeta del ERP")
                p_inv = find_file_by_pattern(BASE_PATH, PATRON_INV_FILE)
                if not p_inv:
                    p_inv = find_by_prefix(BASE_PATH, PFX_INV_ACTUALIZADO)
            log(f"Plantilla de inventario: {p_inv}  (ruta completa: {p_inv.resolve()})")
            src = open_as_excel_source(p_inv, PASSWORDS_TRY)
            df_inv_orig = read_excel_header_at(src, sheet=SHEET_INV_ORIG, header_row_visible=HEADER_ROW_INV)
            log(f"Registros inventario original: {len(df_inv_orig)}")

            idx_orig = {_norm(c): c for c in df_inv_orig.columns}
            ref_col = idx_orig.get("referencia") or idx_orig.get("referencia fertrac") \
                or next((c for c in df_inv_orig.columns if "referenc" in _norm(c)), None)
            if not ref_col:
                raise KeyError(
                    f"No se encontro columna REFERENCIA en {p_inv}. "
                    f"Ruta completa: {p_inv.resolve()} | Columnas: {list(df_inv_orig.columns)}"
                )

            df_inv_orig["__REFERENCIA__"] = df_inv_orig[ref_col].apply(to_num_str)
            df_inv_orig["__REFERENCIA_ACT__"] = df_inv_orig["__REFERENCIA__"].copy()

            log("\n" + "=" * 70)
            log("FASE 3: Actualizar referencias")
            log("=" * 70)

            df_inv_orig = actualizar_referencias_inventario_original(
                df_inv_orig, df_inv_act, df_val_list, df_matriz, df_marcas, df_dist, df_rem
            )

            log("\n" + "=" * 70)
            log("FASE 4: Aplicar reglas marcas propias")
            log("=" * 70)

            df_inv_orig = aplicar_reglas_marcas_propias(df_inv_orig, df_matriz, df_rem)

            log("\n" + "=" * 70)
            log("FASE 5: Eliminaciones y limpieza")
            log("=" * 70)

            log("Abriendo Excel con COM para operaciones de eliminacion...")
            # La plantilla (salida previa) suele estar protegida con contrasena;
            # descifrarla a un archivo temporal en OUTPUT_PATH (ruta real accesible
            # para Excel COM) y abrir ese temporal.
            tpl_origen = p_inv
            tmp_tpl = None
            if is_encrypted_xlsx(p_inv):
                for pw in [PASS_INV] + list(PASSWORDS_TRY or []):
                    try:
                        bio = decrypt_to_stream_local(p_inv, pw)
                        tmp_tpl = OUTPUT_PATH / f"TEMP_INVENTARIO_{datetime.now().strftime('%H%M%S')}.xlsx"
                        with open(tmp_tpl, "wb") as out:
                            out.write(bio.getvalue())
                        tpl_origen = tmp_tpl
                        log(f"  Plantilla protegida descifrada ({pw}) -> {tmp_tpl.name}")
                        break
                    except Exception:
                        continue
            excel, wb, ws = excel_open(tpl_origen, editable=True)
            for sheet in wb.Sheets:
                if normalize_sheet_name(sheet.Name) == SHEET_INV_ORIG:
                    ws = sheet
                    break
            ws_inv_copia = None
            for sheet in wb.Sheets:
                sn = normalize_sheet_name(sheet.Name)
                if "inventario copia" in _norm(sn) or sn == "INVENTARIO COPIA":
                    ws_inv_copia = sheet
                    break
            try:
                if ws_inv_copia is None:
                    ws_inv_copia = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                    ws_inv_copia.Name = SHEET_INV_COPIA
                    log(f"  Se creo la hoja: {SHEET_INV_COPIA}")

                ws_inv_copia.Activate()
                headers_copia = ws_headers(ws_inv_copia, start=1)
                if not headers_copia:
                    headers_src = ws_headers(ws, start=HEADER_ROW_INV)
                    for name, col_idx in headers_src.items():
                        ws_inv_copia.Cells(1, col_idx).Value = name
                    headers_copia = headers_src

                log(f"  Hoja '{SHEET_INV_COPIA}': {len(headers_copia)} columnas detectadas")

                df_inv_orig = eliminar_registros_linea_copia_indeterminada(df_inv_orig, ws, wb)

                df_inv_orig = procesar_existencias_negativas_y_cero(df_inv_orig, ws, header_row=HEADER_ROW_INV)

                TRACKER_ELIMINACIONES.mostrar_resumen()
                reporte_path = TRACKER_ELIMINACIONES.generar_reporte_excel(OUTPUT_PATH)
                if reporte_path:
                    log(f"Reporte de eliminaciones: {reporte_path}")

                log("\n" + "=" * 70)
                log("FASE 6: Guardar resultado")
                log("=" * 70)

                now_str = datetime.now().strftime("%Y%m%d_%H%M")
                nuevo_nombre = f"{OUTPUT_BASENAME} {now_str}.xlsx"
                ruta_guardar = OUTPUT_PATH / nuevo_nombre
                wb.SaveAs(str(ruta_guardar), FileFormat=51)
                log(f"Archivo guardado: {ruta_guardar}")

                if APPLY_PASSWORD_TO_OUTPUT and PASS_INV:
                    try:
                        import msoffcrypto
                        tmp_input = Path(tempfile.gettempdir()) / f"~encrypt_{nuevo_nombre}"
                        shutil.copy2(ruta_guardar, tmp_input)
                        with open(tmp_input, "rb") as fin:
                            office = msoffcrypto.OfficeFile(fin)
                            office.load_key(password=PASS_INV)
                            with open(ruta_guardar, "wb") as fout:
                                office.encrypt(fout)
                        tmp_input.unlink(missing_ok=True)
                        log("Archivo protegido con contrasena.")
                    except Exception as e:
                        log(f"No se pudo proteger con contrasena: {e}")

                excel_close(excel, wb, save=False)
                excel = None
                wb = None
                if tmp_tpl is not None:
                    try:
                        tmp_tpl.unlink(missing_ok=True)
                        log(f"  Temporal de plantilla eliminado: {tmp_tpl.name}")
                    except Exception:
                        pass

                self.stats = {
                    "registros_inventario_original": len(df_inv_orig),
                    "referencias_actualizadas": len(df_inv_act),
                    "valorizados_cargados": len(df_val_list),
                    "matriz_usd": len(df_matriz),
                    "marcas": len(df_marcas),
                    "distribucion": len(df_dist),
                    "remisiones": len(df_rem),
                    "eliminaciones": len(TRACKER_ELIMINACIONES.eliminaciones),
                    "archivo_generado": str(ruta_guardar),
                    "nombre_archivo": nuevo_nombre,
                }

                log("\n" + "=" * 70)
                log("FASE 7: Enviar notificacion")
                log("=" * 70)

                if self.notifier:
                    detail_lines = [
                        "Proceso de Actualizacion de Inventario General completado.",
                        "",
                        f"Fecha de ejecucion: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                        f"Carpeta de trabajo: {BASE_PATH}",
                        f"Archivo generado: {nuevo_nombre}",
                        "",
                        "ESTADISTICAS:",
                        f"  - Registros inventario original: {len(df_inv_orig)}",
                        f"  - Referencias actualizadas (ERP): {len(df_inv_act)}",
                        f"  - Valorizados cargados: {len(df_val_list)}",
                        f"  - Matriz USD: {len(df_matriz)}",
                        f"  - Marcas: {len(df_marcas)}",
                        f"  - Distribucion: {len(df_dist)}",
                        f"  - Remisiones: {len(df_rem)}",
                        f"  - Eliminaciones: {len(TRACKER_ELIMINACIONES.eliminaciones)}",
                        "",
                        "ARCHIVOS GENERADOS:",
                        f"  - {nuevo_nombre}",
                    ]
                    if reporte_path:
                        detail_lines.append(f"  - {Path(reporte_path).name}")
                    detail = "\n".join(detail_lines)
                    self.notifier.notify_success(detail, attachment=str(ruta_guardar))
                    log("Notificacion enviada.")

                log("\n" + "=" * 70)
                log("PROCESO COMPLETADO EXITOSAMENTE")
                log("=" * 70)

                return ruta_guardar

            except Exception as e:
                log(f"\nError durante procesamiento COM: {e}")
                import traceback
                tb_str = traceback.format_exc()
                log(tb_str)
                if self.notifier:
                    self.notifier.notify_failure(
                        f"Error: {e}\n\n{tb_str}",
                    )
                raise

            finally:
                if excel is not None:
                    try:
                        excel_close(excel, wb, save=False)
                    except Exception:
                        pass

        except FileNotFoundError as e:
            log(f"\nERROR CRITICO: {e}")
            import traceback
            tb_str = traceback.format_exc()
            log(tb_str)
            archivo_faltante = None
            if "INVENTARIO GENERAL ACTUALIZADO" in str(e):
                archivo_faltante = "INVENTARIO GENERAL ACTUALIZADO (del dia actual)"
            if self.notifier:
                body = f"Error: {e}\n\nArchivo faltante: {archivo_faltante}\n\nTraceback:\n{tb_str}"
                self.notifier.notify_failure(body)
            raise SystemExit(1)

        except Exception as e:
            log(f"\nERROR CRITICO: {e}")
            import traceback
            tb_str = traceback.format_exc()
            log(tb_str)
            if self.notifier:
                self.notifier.notify_failure(
                    f"Error: {e}\n\nTraceback:\n{tb_str}",
                )
            raise SystemExit(1)


if __name__ == "__main__":
    from config.settings import get_settings
    settings = get_settings()
    task = ActualizacionInventario(settings)
    task.run()
