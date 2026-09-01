# -*- coding: utf-8 -*-
"""
Migracion de actualizar_2025_ventas.py al framework de BaseTask.
"""
from __future__ import annotations

import os, io, re, difflib, warnings, contextlib, time, traceback
from pathlib import Path
from datetime import datetime, date, timezone
from dateutil.relativedelta import relativedelta
import shutil
import tempfile
import pandas as pd
import numpy as np
import msoffcrypto
from unidecode import unidecode

from tasks.base_task import BaseTask
from core.logger import get_logger
from core.email_notifier import EmailNotifier
from core.excel_processing import (
    HAS_COM,
    safe_close_workbook,
    safe_quit_excel,
    excel_serial_from_date,
)
from core.excel_utils import (
    decrypt_to_stream as _decrypt_to_stream,
    read_excel_any as _read_excel_any,
    find_sheet_name,
    norm as _norm,
    norm_simple as _norm_simple,
    norm_colname as _norm_colname,
    norm_label as _norm_label,
    norm_sheet as _norm_sheet,
    norm_base_filename as _norm_base_filename,
    strip_dolares_temporales as _strip_dolares_temporales,
    extract_fecha_es as _extract_fecha_es,
    resolve_cols as _resolve_cols,
    write_excel as _write_excel,
)
from config.settings import Settings, MESES_ES, MESES_ES_NOMBRE

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*errors='ignore' is deprecated.*")
warnings.filterwarnings("ignore", message=".*Downcasting object dtype.*")
warnings.filterwarnings("ignore", category=FutureWarning)

try:
    import win32com.client as win32
    gen_py_path = Path(win32.__gen_path__)
    if gen_py_path.exists():
        shutil.rmtree(gen_py_path)
except Exception:
    pass


class ActualizacionVentas(BaseTask):
    name = "actualizacion_ventas"

    def setup(self):
        self.notifier = EmailNotifier(self.settings.smtp_ventas, self.name)
        self.paths = self.settings.paths
        self.erp = self.settings.erp
        self.excel_cfg = self.settings.excel
        self.log_file = Path(__file__).resolve().parent.parent / "logs" / date.today().isoformat() / f"{self.name}.log"

    def execute(self):
        indicador = self._crear_indicador_progreso()
        try:
            self._ejecutar()
        finally:
            self._eliminar_indicador_progreso(indicador)

    # ─────────────────────────────────────────────────────────────────
    #  CONFIG (from settings)
    # ─────────────────────────────────────────────────────────────────

    @property
    def BASE_PATH(self):
        return self.paths.base

    @property
    def PASSWORD_VENTAS(self):
        return self.excel_cfg.password

    @property
    def PASSWORD_INV_GENERAL(self):
        return self.excel_cfg.password

    @property
    def PASSWORD_MYR(self):
        return self.excel_cfg.password

    @property
    def PASSWORDS_TRY(self):
        return list(self.excel_cfg.passwords_try)

    @property
    def FN_VENTAS(self):
        return "$2026 VENTAS.xlsx"

    @property
    def FN_INV_GENERAL(self):
        return "$2026 INVENTARIO GENERAL.xlsx"

    @property
    def FN_INFORME_VENTAS(self):
        return "InformesDeVentas(Facturas)_177_20250906.xlsx"

    @property
    def FN_MATRIZ_CLIENTES(self):
        return "MATRIZ COMPLETA DE CLIENTES.xlsx"

    @property
    def DIR_INFORME_VENTAS_MES(self):
        return self.paths.ventas_mes

    @property
    def SHEET_VENTAS_OPTIONS(self):
        return ["VENTAS 2026", "VENTAS 2025"]

    @property
    def SHEET_INV_GENERAL(self):
        return "INVENTARIO"

    @property
    def SHEET_COSTOS_INVFINAL(self):
        return "COSTOS INV FINAL"

    @property
    def SHEET_MATRIZ(self):
        return "CLIENTES GENERAL"

    @property
    def COLS_CLAVE_VENTAS(self):
        return dict(
            nit="NIT CLIENTE",
            ref="REFERENCIA",
            marca="MARCA",
            fecha="FECHA",
            mes="MES",
            mes_no="MES NO.",
            anio="AÑO",
        )

    @property
    def COLS_INFORME(self):
        return dict(
            nit="NIT",
            numero="Número",
            nro_doc="Nro. documento",
            fecha="Fecha documento origen",
            fecha_simple="Fecha",
            prefijo="Prefijo",
            doc_origen="Documento origen",
            ref="Referencia",
            marca="Marca",
            cantidad="Cantidad facturada",
            valor_bruto="Valor bruto",
            costo_unit="Costo unitario",
            valor_base_pv="Valor base precio de venta",
            vr_descuento="VR DESCUENTO",
        )

    @property
    def COLS_FORMULAS_PRE(self):
        return [
            "MES", "MES NO.", "ANO", "SEMANA NUMERO",
            "PORCENTAJE DCTO A PIE DE FACTURA",
            "DIF EN MG FACTURADO vs MG LICITADO",
        ]

    @property
    def COLS_FORMULAS_POST(self):
        return [
            "VENTA A COSTO TOTAL", "VENTA TOTAL NETA",
            "VTA NETA X UNIDAD (DCTO PIE FACT)",
            "VTA NETA X UNIDAD (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
            "MARGEN NETO (DCTO PIE FACT)",
            "MARGEN NETO (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
            "MARGEN NETO FACTOR HOY (DCTO PIE FACT)",
            "MARGEN NETO FACTOR HOY (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)",
            "PART", "SUMA MP",
            "DIF EN PRECIO FACTURADO vs LICITADO",
        ]

    @property
    def COLS_MATRIZ_MAP(self):
        return dict(nit="NIT", dcto_cond="DCTO CONDICIONADO", porc_pie_fact="PORCENTAJE DCTO A PIE DE FACTURA")

    # ─────────────────────────────────────────────────────────────────
    #  INDICADOR DE PROGRESO
    # ─────────────────────────────────────────────────────────────────

    def _crear_indicador_progreso(self):
        indicador = self.paths.base / "PROCESANDO.txt"
        try:
            with indicador.open("w", encoding="utf-8") as f:
                f.write(f"Inicio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("El proceso esta en ejecucion...\n")
                f.write(f"Revisa el archivo: {self.log_file.name}\n")
        except Exception:
            pass
        return indicador

    def _eliminar_indicador_progreso(self, indicador_path):
        if not indicador_path:
            return
        try:
            if indicador_path.exists():
                indicador_path.unlink()
                self.log.info("Indicador eliminado exitosamente: %s", indicador_path.name)
        except Exception as e:
            self.log.warning("Error al eliminar indicador: %s", e)

    # ─────────────────────────────────────────────────────────────────
    #  FECHA EXTRACTION
    # ─────────────────────────────────────────────────────────────────

    def _extract_date_yyyyMMdd_from_tail(self, name_no_ext):
        s_clean = re.sub(r'\s*\(\d+\)\s*$', '', name_no_ext)
        s_clean = re.sub(r'\s*\(copia\)\s*$', '', s_clean, flags=re.IGNORECASE)
        s = _norm_simple(s_clean)
        m = re.search(r'(\d{4})[ _-]?(\d{2})[ _-]?(\d{2})\s*$', s)
        if not m:
            return None
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except Exception:
            return None

    # ─────────────────────────────────────────────────────────────────
    #  FILE FINDERS
    # ─────────────────────────────────────────────────────────────────

    def find_matriz_clientes_by_prefix(self, base_dir, prefixes=None):
        if prefixes is None:
            prefixes = ("MATRIZ COMPLETA DE CLIENTES", "MATRIZ DE CLIENTES", "MATRIZ CLIENTES")
        pref_norms = [_norm_simple(p) for p in prefixes]
        exts_ok = {".xlsx", ".xlsm", ".xls", ".csv"}
        cand = []
        for f in base_dir.iterdir():
            if not (f.is_file() and f.suffix.lower() in exts_ok):
                continue
            if f.name.startswith("~$"):
                continue
            name_no_ext = _strip_dolares_temporales(f.name)
            name_no_ext = Path(name_no_ext).stem
            nn = _norm_simple(name_no_ext)
            if any(nn.startswith(pn) for pn in pref_norms):
                cand.append(f)
        if not cand:
            raise FileNotFoundError(f"No encontre archivos que empiecen por {list(prefixes)} en '{base_dir}'.")
        today = date.today()
        con_fecha, sin_fecha = [], []
        for f in cand:
            d = self._extract_date_yyyyMMdd_from_tail(Path(f).stem)
            if d:
                con_fecha.append((d, f))
            else:
                sin_fecha.append(f)
        for d, f in con_fecha:
            if d == today:
                return f
        if con_fecha:
            con_fecha.sort(key=lambda x: x[0], reverse=True)
            return con_fecha[0][1]
        sin_fecha.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return sin_fecha[0]

    def find_informe_facturas_by_prefix(self, base_dir, prefix="InformesDeVentas(Facturas)", only_today=False):
        pref_n = _norm_simple(prefix)
        exts_ok = {".xlsx", ".xlsm", ".xls", ".csv"}
        cand = []
        for f in base_dir.iterdir():
            if not (f.is_file() and f.suffix.lower() in exts_ok):
                continue
            if f.name.startswith("~$"):
                continue
            name_no_ext = Path(f).stem.replace("~$", "")
            if _norm_simple(name_no_ext).startswith(pref_n):
                cand.append(f)
        if not cand:
            raise FileNotFoundError(f"No encontre archivos que empiecen por '{prefix}' en '{base_dir}'.")
        today = date.today()
        if only_today:
            archivos_hoy = []
            for f in cand:
                d = self._extract_date_yyyyMMdd_from_tail(Path(f).stem)
                if d == today:
                    archivos_hoy.append(f)
            if not archivos_hoy:
                raise FileNotFoundError(
                    f"No encontre archivo '{prefix}' con fecha de HOY ({today.strftime('%Y%m%d')}) en '{base_dir}'.\n"
                    f"Archivos disponibles:\n" + "\n".join([f"  - {f.name}" for f in cand])
                )
            archivo_seleccionado = max(archivos_hoy, key=lambda p: p.stat().st_mtime)
            self.log.info("Encontrado archivo de HOY: %s", archivo_seleccionado.name)
            return archivo_seleccionado
        con_fecha, sin_fecha = [], []
        for f in cand:
            d = self._extract_date_yyyyMMdd_from_tail(Path(f).stem)
            if d:
                con_fecha.append((d, f))
            else:
                sin_fecha.append(f)
        archivos_hoy = [f for d, f in con_fecha if d == today]
        if archivos_hoy:
            archivo_seleccionado = max(archivos_hoy, key=lambda p: p.stat().st_mtime)
            self.log.info("Encontrado archivo de HOY: %s", archivo_seleccionado.name)
            return archivo_seleccionado
        if con_fecha:
            con_fecha.sort(key=lambda x: (x[0], x[1].stat().st_mtime), reverse=True)
            self.log.warning("Usando archivo mas reciente: %s", con_fecha[0][1].name)
            return con_fecha[0][1]
        sin_fecha.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        self.log.warning("Usando ultimo modificado: %s", sin_fecha[0].name)
        return sin_fecha[0]

    def find_myr_existencia_by_fecha(self, base_dir, prefix_expected=None):
        if prefix_expected is None:
            year = date.today().year
            prefix_expected = f"{year} inventario myr existencia"
        prefix_expected_n = _norm_simple(prefix_expected)
        self.log.info("Buscando archivos con prefijo: '%s'", prefix_expected)
        candidatos = []
        for f in base_dir.iterdir():
            if not (f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm", ".xls")):
                continue
            if f.name.startswith("~$"):
                continue
            base_clean = _strip_dolares_temporales(f.name)
            if _norm_simple(base_clean).startswith(prefix_expected_n):
                candidatos.append(f)
        if not candidatos:
            year = date.today().year
            raise FileNotFoundError(
                f"No encontre archivos que empiecen por '{prefix_expected}' en '{base_dir}'.\n"
                f"Verifica que existan archivos de inventario MYR del ano {year}.\n"
                f"Ejemplo esperado: '$2026 INVENTARIO MYR EXISTENCIA 03 ENERO.xlsx'"
            )
        self.log.info("  Encontrados %d candidatos:", len(candidatos))
        for f in candidatos[:5]:
            self.log.info("     - %s", f.name)
        if len(candidatos) > 5:
            self.log.info("     ... y %d mas", len(candidatos) - 5)
        today = date.today()
        for f in candidatos:
            d = _extract_fecha_es(f.name)
            if d and d == today:
                self.log.info("Archivo con fecha de HOY encontrado: %s", f.name)
                return f
        candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        archivo_reciente = candidatos[0]
        fecha_mod = datetime.fromtimestamp(archivo_reciente.stat().st_mtime)
        fecha_nombre = _extract_fecha_es(archivo_reciente.name)
        self.log.warning("No hay archivo de hoy. Usando el mas reciente por MODIFICACION")
        self.log.info("   Archivo: %s", archivo_reciente.name)
        self.log.info("   Modificado: %s", fecha_mod.strftime('%d/%m/%Y %H:%M:%S'))
        if fecha_nombre:
            self.log.info("   Fecha en nombre: %s", fecha_nombre.strftime('%d/%m/%Y'))
        return archivo_reciente

    def find_file_by_loose_name(self, base_dir, expected_name, exts=(".xlsx", ".xlsm")):
        target = _norm_base_filename(expected_name)
        self.log.info("Buscando archivo: '%s' -> normalizado '%s' en %s", expected_name, target, base_dir)
        candidates = []
        todos_los_archivos = []
        for f in base_dir.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in exts:
                continue
            if f.name.startswith("~$") or f.name.startswith("~"):
                continue
            nb = _norm_base_filename(f.name)
            fecha_mod = datetime.fromtimestamp(f.stat().st_mtime)
            todos_los_archivos.append((f.name, nb, fecha_mod, f.stat().st_size))
            if nb == target:
                candidates.append(f)
        if candidates:
            candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            seleccionado = candidates[0]
            self.log.info("Archivo seleccionado: %s", seleccionado.name)
            return seleccionado
        self.log.info("No hubo coincidencia exacta, buscando parcial...")
        partial_candidates = []
        for f in base_dir.iterdir():
            if not f.is_file() or f.suffix.lower() not in exts:
                continue
            if f.name.startswith("~$") or f.name.startswith("~"):
                continue
            nb = _norm_base_filename(f.name)
            if nb.endswith(target):
                partial_candidates.append(f)
        if partial_candidates:
            partial_candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            seleccionado = partial_candidates[0]
            self.log.info("Seleccionado (parcial): %s", seleccionado.name)
            return seleccionado
        self.log.error("NO SE ENCONTRO NINGUN ARCHIVO para '%s'", expected_name)
        all_files = [f.name for f in base_dir.iterdir() if f.is_file() and f.suffix.lower() in exts]
        raise FileNotFoundError(
            f"No pude encontrar '{expected_name}'. Archivos disponibles:\n" +
            "\n".join([f"  - {fn}" for fn in all_files[:10]])
        )

    # ─────────────────────────────────────────────────────────────────
    #  DECRYPT / READ / WRITE (delegate to core)
    # ─────────────────────────────────────────────────────────────────

    def _decrypt_to_stream(self, xlsx_path, password=None):
        return _decrypt_to_stream(xlsx_path, password=password, config=self.excel_cfg)

    def _read_excel_any(self, xlsx, **kwargs):
        return _read_excel_any(xlsx, **kwargs)

    def _write_excel(self, df_or_dict, path_out, sheetname=None):
        return _write_excel(df_or_dict, path_out, sheetname=sheetname)

    # ─────────────────────────────────────────────────────────────────
    #  TEMP FILE
    # ─────────────────────────────────────────────────────────────────

    @contextlib.contextmanager
    def archivo_temporal_seguro(self, stream, suffix="src"):
        tmp_path = None
        try:
            tmp = tempfile.NamedTemporaryFile(
                mode='wb', suffix='.xlsx', prefix=f'tmp_{suffix}_',
                dir=str(self.BASE_PATH), delete=False
            )
            pos = stream.tell()
            stream.seek(0)
            tmp.write(stream.read())
            tmp.close()
            stream.seek(pos)
            tmp_path = Path(tmp.name)
            try:
                import ctypes
                FILE_ATTRIBUTE_HIDDEN = 0x02
                ctypes.windll.kernel32.SetFileAttributesW(str(tmp_path), FILE_ATTRIBUTE_HIDDEN)
            except Exception:
                pass
            self.log.info("Archivo temporal creado: %s", tmp_path.name)
            yield tmp_path
        finally:
            if tmp_path and tmp_path.exists():
                try:
                    os.remove(tmp_path)
                    self.log.info("Archivo temporal eliminado: %s", tmp_path.name)
                except Exception as e:
                    self.log.warning("Error al eliminar temporal: %s", e)

    # ─────────────────────────────────────────────────────────────────
    #  UTILITY FUNCTIONS
    # ─────────────────────────────────────────────────────────────────

    def _month_now_bogota(self):
        today = date.today()
        first = today.replace(day=1)
        last = (first + relativedelta(months=1)) - relativedelta(days=1)
        return today.month, today.year, first, last

    def _is_valid_reference(self, s):
        if pd.isna(s):
            return False
        t = str(s).strip()
        if not t:
            return False
        if len(t) > 30:
            return False
        if len(t.split()) > 3:
            return False
        return True

    def _excel_col_to_index(self, col_letter):
        col_letter = col_letter.upper()
        n = 0
        for ch in col_letter:
            n = n * 26 + (ord(ch) - 64)
        return n

    def _save_stream_to_tempfile(self, stream, suffix="src"):
        tmp = tempfile.NamedTemporaryFile(
            mode='wb', suffix='.xlsx', prefix=f'tmp_{suffix}_',
            dir=str(self.BASE_PATH), delete=False
        )
        pos = stream.tell()
        stream.seek(0)
        tmp.write(stream.read())
        tmp.close()
        stream.seek(pos)
        tmp_path = Path(tmp.name)
        try:
            import ctypes
            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(str(tmp_path), FILE_ATTRIBUTE_HIDDEN)
        except Exception as e:
            self.log.warning("No se pudo ocultar archivo temporal: %s", e)
        return tmp_path

    def _to_excel_cell_value(self, v):
        try:
            if v is None or pd.isna(v):
                return None
        except Exception:
            if v is None:
                return None
        if isinstance(v, pd.Timestamp):
            if v.tz is not None:
                return v.tz_convert("UTC").tz_localize(None).to_pydatetime()
            return v.to_pydatetime()
        if isinstance(v, np.datetime64):
            ts = pd.Timestamp(v)
            if ts is pd.NaT:
                return None
            if ts.tz is not None:
                return ts.tz_convert("UTC").tz_localize(None).to_pydatetime()
            return ts.to_pydatetime()
        if isinstance(v, datetime):
            if v.tzinfo is not None:
                return v.astimezone(timezone.utc).replace(tzinfo=None)
            return v
        if isinstance(v, date):
            return datetime(v.year, v.month, v.day)
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            fv = float(v)
            if np.isnan(fv) or np.isinf(fv):
                return None
            return fv
        if isinstance(v, (np.bool_,)):
            return bool(v)
        return v

    def _sanitize_df_for_excel(self, df):
        return df.applymap(self._to_excel_cell_value)

    # ─────────────────────────────────────────────────────────────────
    #  COLUMN RESOLUTION
    # ─────────────────────────────────────────────────────────────────

    COL_SYNONYMS = {
        "AÑO": ["AÑO", "ANIO", "Año", "Ano", "ANO"],
        "MES No.": ["MES NO.", "MES NO", "MES NRO", "MES Nº", "MES NUM", "MES No", "MES", "Mes No.", "Mes No", "MES NRO.", "MES NO ."],
        "MES": ["MES", "Mes", "MES NOMBRE", "NOMBRE MES"],
        "FECHA": ["FECHA", "Fecha", "F. FECHA", "FCHA"],
        "NIT CLIENTE": ["NIT CLIENTE", "NIT", "Nit Cliente", "NIT_CLIENTE"],
        "REFERENCIA": ["REFERENCIA", "Referencia", "REF", "Ref", "CODIGO", "CÓDIGO", "Codigo"],
        "MARCA": ["MARCA", "Marca", "BRAND"],
        "DCTO CONDICIONADO": ["DCTO CONDICIONADO", "DTO CONDICIONADO", "DESCUENTO CONDICIONADO", "DCTO COND.", "DCTO CONDIC."],
        "FECHA DE ACTUALIZACION": ["FECHA DE ACTUALIZACION", "Fecha de actualización", "FECHA ACTUALIZACION", "F. ACTUALIZACION"],
    }

    def _resolve_cols_flex(self, df, cols_map):
        idx = {_norm_colname(c): c for c in df.columns}
        resolved = {}
        for key, target in cols_map.items():
            wanted_norm = _norm_colname(target)
            cands = [target] + self.COL_SYNONYMS.get(target, [])
            found = None
            for c in cands:
                cn = _norm_colname(c)
                if cn in idx:
                    found = idx[cn]
                    break
            if not found:
                for kn, real in idx.items():
                    if wanted_norm and wanted_norm in kn:
                        found = real
                        break
            if not found:
                best = difflib.get_close_matches(wanted_norm, list(idx.keys()), n=1, cutoff=0.7)
                if best:
                    found = idx[best[0]]
            if not found:
                raise KeyError(f"No encuentro la columna '{target}'. Encabezados: {list(df.columns)}")
            resolved[key] = found
        return resolved

    def _template_synonyms_for(self, header_raw):
        h = _norm_colname(header_raw)
        mapping = {
            "nit cliente":            [self.COLS_CLAVE_VENTAS["nit"], "nit", "nit  cliente"],
            "cliente":                ["cliente", "nombre", "nombre cliente"],
            "ciudad":                 ["ciudad"],
            "vend":                   ["vendedor", "vend"],
            "descripcion":            ["descripcion", "descripción", "descrpcion", "nombre lista", "nombre myr", "nombre odoo", "descripcion producto", "producto", "nombre"],
            "descrpcion":             ["descripcion", "descripción", "nombre lista", "nombre myr", "nombre odoo"],
            "marca":                  [self.COLS_CLAVE_VENTAS["marca"], "marca"],
            "cantidad":               ["cantidad facturada", "cantidad"],
            "vr unitario":            ["valor unitario", "vr unitario"],
            "vr total":               ["valor bruto", "vr total", "total venta", "total"],
            "vr descuento":           ["valor descuento", "vr descuento", "descuento"],
            "costo promedio":         ["costo promedio", "costo promedio ", "costo unitario"],
            "dcto condicionado":      [self.COLS_MATRIZ_MAP["dcto_cond"], "dto condicionado", "descuento condicionado", "dcto cond.", "dcto cond"],
            "margen neto (incluye todos los dctos pie fact + financ)": [],
            "margen neto (incluye todos los dctos: pie fact + financ)": [],
            "dv": ["dv", "prefijo", "digito de verificacion"],
        }
        return [_norm_colname(x) for x in ([header_raw] + mapping.get(h, []))]

    def _build_template_col_map(self, df_src, template_headers):
        idx_df = {_norm_colname(c): c for c in df_src.columns}
        col_map = {}
        for h in template_headers:
            found = None
            for cand_norm in self._template_synonyms_for(h):
                if cand_norm in idx_df:
                    found = idx_df[cand_norm]
                    break
            if found:
                col_map[h] = found
        return col_map

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: transformar_informe_ventas
    # ─────────────────────────────────────────────────────────────────

    def transformar_informe_ventas(self, df):
        df = df.copy()
        ci = self.COLS_INFORME
        cv = self.COLS_CLAVE_VENTAS

        if ci["nit"] not in df.columns:
            cols_obj = set(ci.values())
            idx_header = None
            for i in range(min(15, len(df))):
                fila = set(str(x).strip() for x in df.iloc[i].tolist())
                if len(cols_obj & fila) >= max(3, int(0.6 * len(cols_obj))):
                    idx_header = i
                    break
            if idx_header is not None:
                df.columns = df.iloc[idx_header].tolist()
                df = df.iloc[idx_header + 1:].reset_index(drop=True)

        col_prefijo = ci.get("prefijo", "Prefijo")
        if col_prefijo in df.columns:
            n_antes = len(df)
            mask_excluir = df[col_prefijo].astype(str).str.strip().str.upper().isin(['NC', 'NCDTO', 'NDCTO'])
            df = df[~mask_excluir].copy()
            n_despues = len(df)
            n_eliminados = n_antes - n_despues
            if n_eliminados > 0:
                self.log.info("FILTRO PREFIJO: %d registros eliminados (NC/NCDTO)", n_eliminados)
            else:
                self.log.info("No se encontraron registros NC/NCDTO para eliminar")
        else:
            self.log.warning("Columna '%s' no encontrada, no se puede filtrar NC/NCDTO", col_prefijo)

        for col in df.columns:
            col_lower = col.lower()
            if "valor" in col_lower and "descuento" in col_lower and "comercial" in col_lower:
                df = df.rename(columns={col: "VR DESCUENTO"})
                break

        for cand in ["Número documento", "Numero documento", "Nº documento"]:
            if cand in df.columns:
                df = df.drop(columns=[cand])

        if ci["valor_bruto"] in df.columns:
            pos = list(df.columns).index(ci["valor_bruto"])
            df.insert(pos, "COL_TMP_1", np.nan)
            df.insert(pos + 1, "Valor unitario", np.nan)
        else:
            for c in ["COL_TMP_1", "Valor unitario"]:
                if c not in df.columns:
                    df[c] = np.nan

        self.log.info("APLICANDO LOGICA CONDICIONAL DE FECHAS")

        fecha_doc_origen = ci["fecha"]
        fecha_simple = ci["fecha_simple"]
        prefijo = ci["prefijo"]
        doc_origen = ci["doc_origen"]

        if fecha_doc_origen in df.columns:
            df[fecha_doc_origen] = pd.to_datetime(df[fecha_doc_origen], format='%d/%m/%Y', errors='coerce', dayfirst=True).dt.normalize()
            self.log.info("Convertida '%s' a datetime", fecha_doc_origen)

        if fecha_simple in df.columns:
            df[fecha_simple] = pd.to_datetime(df[fecha_simple], format='%d/%m/%Y', errors='coerce', dayfirst=True).dt.normalize()
            self.log.info("Convertida '%s' a datetime", fecha_simple)

        if all(col in df.columns for col in [prefijo, doc_origen, fecha_doc_origen, fecha_simple]):
            mask_prefijo_fe = df[prefijo].astype(str).str.strip().str.upper() == "FE"
            mask_doc_pv = df[doc_origen].astype(str).str.strip().str.upper().str.startswith("PV")
            mask_usar_fecha_doc = mask_prefijo_fe & mask_doc_pv

            nombre_temp = "_FECHA_TEMPORAL_PROCESAMIENTO_"
            df[nombre_temp] = df[fecha_doc_origen].where(mask_usar_fecha_doc, df[fecha_simple])

            cols_to_drop = []
            for c in [fecha_doc_origen, fecha_simple, doc_origen]:
                if c in df.columns:
                    cols_to_drop.append(c)
            if cols_to_drop:
                df = df.drop(columns=cols_to_drop)

            if prefijo in df.columns:
                n_prefijo = df[prefijo].notna().sum()
                self.log.info("Columna '%s' preservada: %d valores", prefijo, n_prefijo)

            df = df.rename(columns={nombre_temp: fecha_doc_origen})

            n_total = len(df)
            n_fecha_doc = mask_usar_fecha_doc.sum()
            n_fecha_simple = (~mask_usar_fecha_doc).sum()
            n_nulls = df[fecha_doc_origen].isna().sum() if fecha_doc_origen in df.columns else n_total

            self.log.info("RESUMEN DE FECHAS:")
            self.log.info("  Total registros: %d", n_total)
            self.log.info("  Usando 'Fecha documento origen' (FE + PV): %d (%.1f%%)", n_fecha_doc, n_fecha_doc / n_total * 100)
            self.log.info("  Usando 'Fecha' (resto): %d (%.1f%%)", n_fecha_simple, n_fecha_simple / n_total * 100)
            if n_nulls > 0:
                self.log.warning("  Registros con fecha nula: %d (%.1f%%)", n_nulls, n_nulls / n_total * 100)
        else:
            self.log.warning("No se encontraron todas las columnas necesarias para filtrado condicional")

        if ci["fecha"] in df.columns and ci["nro_doc"] in df.columns:
            cols = df.columns.tolist()
            if ci["fecha"] in cols:
                cols.remove(ci["fecha"])
            cols.insert(cols.index(ci["nro_doc"]) + 1, ci["fecha"])
            df = df[cols]

        if ci["ref"] in df.columns and ci["numero"] in df.columns:
            cols = df.columns.tolist()
            cols.remove(ci["ref"])
            cols.insert(cols.index(ci["numero"]) + 1, ci["ref"])
            df = df[cols]

        if ci["cantidad"] in df.columns and ci["marca"] in df.columns:
            cols = df.columns.tolist()
            cols.remove(ci["cantidad"])
            cols.insert(cols.index(ci["marca"]) + 1, ci["cantidad"])
            df = df[cols]

        if ci["valor_bruto"] in df.columns and ci["cantidad"] in df.columns:
            with np.errstate(divide="ignore", invalid="ignore"):
                df["Valor unitario"] = pd.to_numeric(df[ci["valor_bruto"]], errors="coerce") / \
                                       pd.to_numeric(df[ci["cantidad"]], errors="coerce")

        if ci["costo_unit"] in df.columns and ci["valor_base_pv"] in df.columns:
            cols = df.columns.tolist()
            cols.remove(ci["costo_unit"])
            cols.insert(cols.index(ci["valor_base_pv"]), ci["costo_unit"])
            df = df[cols]

        if ci["ref"] in df.columns:
            df = df[~df[ci["ref"]].isna() & (df[ci["ref"]].astype(str).str.strip() != "")].copy()

        for c in [ci.get("numero"), ci.get("ref"), "Documento"]:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="ignore")

        if ci["costo_unit"] in df.columns:
            df[ci["costo_unit"]] = pd.to_numeric(df[ci["costo_unit"]], errors="coerce").abs()

        if "COL_TMP_1" in df.columns:
            df = df.drop(columns=["COL_TMP_1"])

        if ci["ref"] in df.columns:
            df[ci["ref"]] = df[ci["ref"]].astype(str).str.strip()
            df[ci["ref"]] = df[ci["ref"]].replace({"nan": None, "None": None, "": None})
            n_refs = df[ci["ref"]].notna().sum()
            self.log.info("Referencias validas despues de limpieza: %d", n_refs)

        if ci["fecha"] in df.columns:
            nombre_actual = ci["fecha"]
            nombre_final = cv["fecha"]
            if nombre_actual != nombre_final:
                n_valores = df[nombre_actual].notna().sum()
                df = df.rename(columns={nombre_actual: nombre_final})
                self.log.info("RENAME FINAL (para merge): '%s' -> '%s'  Valores: %d/%d",
                              nombre_actual, nombre_final, n_valores, len(df))

        return df.reset_index(drop=True)

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: normalizar_columnas_df
    # ─────────────────────────────────────────────────────────────────

    def normalizar_nombre_columna(self, nombre):
        s = str(nombre).strip()
        s = unidecode(s)
        s = s.upper()
        s = re.sub(r'\s+', ' ', s)
        return s.strip()

    def normalizar_columnas_df(self, df):
        df_norm = df.copy()
        mapeo = {}
        for col in df.columns:
            col_norm = self.normalizar_nombre_columna(col)
            if col != col_norm:
                mapeo[col] = col_norm
        df_norm.columns = [self.normalizar_nombre_columna(col) for col in df.columns]
        return df_norm, mapeo

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: integrar_linea_sublinea
    # ─────────────────────────────────────────────────────────────────

    def integrar_linea_sublinea(self, df_ventas, df_inv, COLS_NORM):
        df = df_ventas.copy()
        inv = df_inv.rename(columns={
            "LINEA COPIA": "LINEA",
            "SUB-LINEA COPIA": "SUBLINEA",
            "LIDER LINEA": "LIDER LINEA",
            "REFERENCIA": COLS_NORM["ref"]
        })
        columnas_inv = [COLS_NORM["ref"], "LINEA", "SUBLINEA"]
        if "LIDER LINEA" in inv.columns:
            columnas_inv.append("LIDER LINEA")
        inv = inv[columnas_inv].drop_duplicates(COLS_NORM["ref"], keep="last")

        def normalizar_ref(s):
            if pd.isna(s):
                return None
            s_str = str(s).strip()
            if s_str == "" or s_str.lower() == "nan":
                return None
            try:
                return str(int(float(s_str)))
            except Exception:
                return s_str

        if COLS_NORM["ref"] in inv.columns:
            self.log.info("Normalizando referencias en INVENTARIO...")
            inv[COLS_NORM["ref"]] = inv[COLS_NORM["ref"]].apply(normalizar_ref)
            ejemplos_inv = inv[COLS_NORM["ref"]].dropna().head(5).tolist()
            self.log.info("  Ejemplos: %s", ejemplos_inv)

        if COLS_NORM["ref"] in df.columns:
            self.log.info("Normalizando referencias en VENTAS...")
            df[COLS_NORM["ref"]] = df[COLS_NORM["ref"]].apply(normalizar_ref)
            ejemplos_ventas = df[COLS_NORM["ref"]].dropna().head(5).tolist()
            self.log.info("  Ejemplos: %s", ejemplos_ventas)

        self.log.info("Haciendo merge por REFERENCIA normalizada...")
        df = df.merge(inv, on=COLS_NORM["ref"], how="left", suffixes=("", "_inv"))

        for c in ["LINEA", "SUBLINEA", "LIDER LINEA"]:
            if c + "_inv" in df.columns:
                if c not in df.columns:
                    df[c] = df[c + "_inv"]
                else:
                    df[c] = df[c].fillna(df[c + "_inv"])
                df = df.drop(columns=[c + "_inv"])

        if "LINEA" in df.columns:
            n_linea_ok = df["LINEA"].notna().sum()
            self.log.info("LINEA: %d/%d registros (%.1f%%)", n_linea_ok, len(df), n_linea_ok / len(df) * 100)
        if "SUBLINEA" in df.columns:
            n_sublinea_ok = df["SUBLINEA"].notna().sum()
            self.log.info("SUBLINEA: %d/%d registros (%.1f%%)", n_sublinea_ok, len(df), n_sublinea_ok / len(df) * 100)
        if "LIDER LINEA" in df.columns:
            n_lider_ok = df["LIDER LINEA"].notna().sum()
            self.log.info("LIDER LINEA: %d/%d registros (%.1f%%)", n_lider_ok, len(df), n_lider_ok / len(df) * 100)
        return df

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: integrar_costo_factor_hoy
    # ─────────────────────────────────────────────────────────────────

    def integrar_costo_factor_hoy(self, df_ventas, df_myr, COLS_NORM):
        df = df_ventas.copy()
        myr = df_myr.copy()
        ref_real = None
        for cand in ["REFERENCIA FERTRAC", "REFERENCIA", "REF", "CODIGO"]:
            if cand in myr.columns:
                ref_real = cand
                break
        if not ref_real:
            self.log.warning("No se encontro columna de REFERENCIA en MYR")
            return df
        cost_real = None
        for cand in ["COSTO FACTOR HOY", "COSTO FACTOR DOLAR HOY", "COSTO FACTOR USD HOY"]:
            if cand in myr.columns:
                cost_real = cand
                break
        if not cost_real:
            self.log.warning("No se encontro columna de COSTO FACTOR HOY en MYR")
            return df
        myr2 = myr[[ref_real, cost_real]].copy()
        myr2 = myr2.rename(columns={ref_real: COLS_NORM["ref"], cost_real: "COSTO FACTOR HOY"})
        myr2[COLS_NORM["ref"]] = myr2[COLS_NORM["ref"]].astype(str).str.strip()
        df[COLS_NORM["ref"]] = df[COLS_NORM["ref"]].astype(str).str.strip()
        myr2 = myr2.drop_duplicates(COLS_NORM["ref"], keep="last")
        myr2 = myr2[myr2[COLS_NORM["ref"]].notna() & (myr2[COLS_NORM["ref"]] != "nan")]
        self.log.info("Haciendo merge por REFERENCIA (MYR)...")
        df_result = df.merge(myr2, on=COLS_NORM["ref"], how="left", suffixes=("", "_myr"))
        if "COSTO FACTOR HOY_myr" in df_result.columns:
            if "COSTO FACTOR HOY" not in df_result.columns:
                df_result["COSTO FACTOR HOY"] = df_result["COSTO FACTOR HOY_myr"]
            else:
                df_result["COSTO FACTOR HOY"] = df_result["COSTO FACTOR HOY"].fillna(df_result["COSTO FACTOR HOY_myr"])
            df_result = df_result.drop(columns=["COSTO FACTOR HOY_myr"])
        return df_result

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: limpiar_linea_referencias_invalidas
    # ─────────────────────────────────────────────────────────────────

    def limpiar_linea_referencias_invalidas(self, df_ventas, COLS_NORM=None):
        df = df_ventas.copy()
        if COLS_NORM is None:
            COLS_NORM = {"ref": "REFERENCIA"}
        if "LINEA" in df.columns and COLS_NORM["ref"] in df.columns:
            refs_temp = df[COLS_NORM["ref"]].astype(str).str.strip()
            mask_bad = (df["LINEA"].fillna(0) == 0) & (~refs_temp.apply(self._is_valid_reference))
            n_bad = mask_bad.sum()
            if n_bad > 0:
                df = df[~mask_bad].copy()
                self.log.info("Eliminados %d registros con referencias invalidas", n_bad)
        return df

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: normalizar_dctos
    # ─────────────────────────────────────────────────────────────────

    def normalizar_dctos(self, df_ventas):
        dc = self.COLS_MATRIZ_MAP["dcto_cond"]
        pf = self.COLS_MATRIZ_MAP["porc_pie_fact"]
        vr = "VR DESCUENTO"
        df = df_ventas.copy()

        def to_num_pct(x):
            if pd.isna(x):
                return 0.0
            s = str(x).replace("%", "").replace(",", ".").strip()
            try:
                return float(s)
            except Exception:
                return 0.0

        if dc in df.columns:
            dcto = df[dc].apply(to_num_pct)
            if vr in df.columns:
                vr_num = pd.to_numeric(df[vr], errors='coerce').fillna(0)
                mask = (dcto != 0) & (vr_num != 0)
                df.loc[mask, dc] = "0%"
                self.log.info("normalizar_dctos: %d registros corregidos (VR DESCUENTO != 0)", mask.sum())
            elif pf in df.columns:
                pie = df[pf].apply(to_num_pct)
                mask = (dcto != 0) & (pie != 0)
                df.loc[mask, dc] = "0%"
                self.log.info("normalizar_dctos: %d registros corregidos (fallback PIE FACT)", mask.sum())
        return df

    # ─────────────────────────────────────────────────────────────────
    #  TRANSFORMATION: ordenar_y_fechas
    # ─────────────────────────────────────────────────────────────────

    def ordenar_y_fechas(self, df_ventas, COLS_NORM):
        df = df_ventas.copy()
        if COLS_NORM.get("fecha") in df.columns:
            fecha_col = COLS_NORM.get("fecha")
            df[fecha_col] = pd.to_datetime(df[fecha_col], dayfirst=True, errors="coerce").dt.normalize()
            df = df.sort_values(fecha_col, ascending=False, na_position="last")
        else:
            self.log.warning("No se puede ordenar (falta columna FECHA)")
        for cand in ["FECHA DE ACTUALIZACION", "FECHA ACTUALIZACION"]:
            if cand in df.columns:
                df[cand] = date.today()
                break
        return df.reset_index(drop=True)

    # ─────────────────────────────────────────────────────────────────
    #  COM WRITE: com_write_df_into_template
    # ─────────────────────────────────────────────────────────────────

    def _detect_header_row(self, ws, used_cols):
        def row_names(r):
            vals = []
            for c in range(1, used_cols + 1):
                v = ws.Cells(r, c).Value
                vals.append("" if v is None else str(v).strip())
            return vals
        for hr in (1, 2, 3):
            vals = row_names(hr)
            if any(x in vals for x in ["AÑO", "ANIO", "FECHA", "NIT CLIENTE", "REFERENCIA"]):
                return hr
        return 1

    def aplicar_formato_numero_columna_c(self, ws, header_row):
        self.log.info("Aplicando formato NUMERO a columna C...")
        try:
            col_c_idx = 3
            ws.Columns(col_c_idx).NumberFormat = "0"
            self.log.info("Formato numero aplicado a columna C")
        except Exception as e:
            self.log.error("Error al formatear columna C: %s", e)

    def convertir_texto_a_numero_columna_c(self, ws, header_row):
        self.log.info("Convirtiendo texto a numero en columna C...")
        try:
            col_c_idx = 3
            ultima_fila = ws.Cells(ws.Rows.Count, col_c_idx).End(-4162).Row
            self.log.info("Rango a convertir: C%d:C%d", header_row + 1, ultima_fila)
            temp_cell = ws.Cells(1, ws.Columns.Count)
            temp_cell.Value = 1
            temp_cell.Copy()
            primera_fila_datos = header_row + 1
            ref_range = ws.Range(
                ws.Cells(primera_fila_datos, col_c_idx),
                ws.Cells(ultima_fila, col_c_idx)
            )
            ref_range.PasteSpecial(Paste=-4163, Operation=4, SkipBlanks=False, Transpose=False)
            ws.Application.CutCopyMode = False
            temp_cell.ClearContents()
            self.log.info("Conversion aplicada exitosamente")
        except Exception as e:
            self.log.error("Error al convertir columna C: %s", e)

    def validar_sublinea_en_cero(self, wb):
        try:
            ws = wb.Worksheets("Resum Sub-linea")
        except Exception:
            self.log.warning("Hoja 'Resum Sub-linea' no encontrada, saltando validacion...")
            return
        used_cols = ws.UsedRange.Columns.Count
        used_rows = ws.UsedRange.Rows.Count
        sublinea_col = None
        header_row = 1
        for r in range(1, min(15, used_rows + 1)):
            for c in range(1, used_cols + 1):
                val = str(ws.Cells(r, c).Value or "").strip().upper()
                if "SUBLINEA" in val or "SUB-LINEA" in val or "SUB LINEA" in val:
                    sublinea_col = c
                    header_row = r
                    break
            if sublinea_col:
                break
        if not sublinea_col:
            self.log.warning("No se encontro columna SUBLINEA en 'Resum Sub-linea'")
            return
        problemas = []
        for r in range(header_row + 1, used_rows + 1):
            val = ws.Cells(r, sublinea_col).Value
            if val is not None:
                try:
                    if float(val) == 0:
                        desc = ws.Cells(r, sublinea_col + 1).Value or ""
                        problemas.append(f"Fila {r}: SUBLINEA = 0 ({desc})")
                except (ValueError, TypeError):
                    pass
        if problemas:
            self.log.warning("Se encontraron %d registros con SUBLINEA en 0:", len(problemas))
            for p in problemas[:5]:
                self.log.warning("  - %s", p)

    def actualizar_periodo_mes(self, wb, primer_dia, ultimo_dia):
        try:
            ws = wb.Worksheets("Resum Mes")
        except Exception:
            self.log.warning("Hoja 'Resum Mes' no encontrada, saltando actualizacion...")
            return
        anio = primer_dia.year
        festivos_fijos = {
            date(anio, 1, 1):  "Año Nuevo",
            date(anio, 5, 1):  "Día del Trabajo",
            date(anio, 7, 20): "Independencia",
            date(anio, 8, 7):  "Batalla de Boyacá",
            date(anio, 12, 8): "Inmaculada Concepción",
            date(anio, 12, 25):"Navidad",
        }
        def siguiente_lunes(d):
            dias = (7 - d.weekday()) % 7
            return d + relativedelta(days=dias) if dias != 0 else d
        festivos_lunes_raw = [
            date(anio, 1, 6), date(anio, 3, 19), date(anio, 6, 29),
            date(anio, 8, 15), date(anio, 10, 12), date(anio, 11, 1), date(anio, 11, 11),
        ]
        festivos_lunes = {siguiente_lunes(d): n for d, n in zip(festivos_lunes_raw, [
            "Reyes Magos", "San José", "San Pedro y San Pablo",
            "Asunción", "Día de la Raza", "Todos los Santos", "Independencia de Cartagena"
        ])}
        a = anio % 19; b = anio // 100; c = anio % 100
        d_b = b // 4; e = b % 4; f = (b + 8) // 25
        g = (b - f + 1) // 3; h = (19 * a + b - d_b - g + 15) % 30
        i = c // 4; k = c % 4
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        mes_pascua  = (h + l - 7 * m + 114) // 31
        dia_pascua  = ((h + l - 7 * m + 114) % 31) + 1
        pascua = date(anio, mes_pascua, dia_pascua)
        festivos_semana_santa = {
            pascua - relativedelta(days=3): "Jueves Santo",
            pascua - relativedelta(days=2): "Viernes Santo",
            pascua + relativedelta(days=39): "Ascensión",
            pascua + relativedelta(days=60): "Corpus Christi",
            pascua + relativedelta(days=68): "Sagrado Corazón",
        }
        festivos_ss_lunes = {}
        for fd, fn in festivos_semana_santa.items():
            if fn in ("Ascensión", "Corpus Christi", "Sagrado Corazón"):
                festivos_ss_lunes[siguiente_lunes(fd)] = fn
            else:
                festivos_ss_lunes[fd] = fn
        festivos_colombia = {}
        festivos_colombia.update(festivos_fijos)
        festivos_colombia.update(festivos_lunes)
        festivos_colombia.update(festivos_ss_lunes)
        festivos_mes = sorted(
            [(f, n) for f, n in festivos_colombia.items()
             if f.month == primer_dia.month and f.year == primer_dia.year],
            key=lambda x: x[0]
        )
        self.log.info("Festivos para %s: %d", primer_dia.strftime('%B %Y'), len(festivos_mes))
        for fd, fn in festivos_mes:
            self.log.info("  %s - %s", fd.strftime('%d/%m/%Y'), fn)
        periodo_texto = f"{primer_dia.strftime('%d/%m/%Y')} - {ultimo_dia.strftime('%d/%m/%Y')}"
        used_cols = ws.UsedRange.Columns.Count
        used_rows = min(35, ws.UsedRange.Rows.Count)
        periodo_row,  periodo_col  = None, None
        festivos_row, festivos_col = None, None
        for r in range(1, used_rows + 1):
            for c in range(1, used_cols + 1):
                val = str(ws.Cells(r, c).Value or "").strip().upper()
                if "PERIODO" in val and "MES" in val and periodo_row is None:
                    periodo_row, periodo_col = r, c
                if "FESTIVOS" in val and "MES" in val and festivos_row is None:
                    festivos_row, festivos_col = r, c
        if periodo_row is not None:
            ws.Cells(periodo_row, periodo_col + 1).Value = periodo_texto
            self.log.info("PERIODO MES texto actualizado: %s", periodo_texto)
            import datetime as _dt
            try:
                ws.Cells(periodo_row + 1, periodo_col).Value = _dt.datetime(primer_dia.year, primer_dia.month, primer_dia.day)
                ws.Cells(periodo_row + 1, periodo_col).NumberFormat = "d/mm/yyyy"
                ws.Cells(periodo_row + 1, periodo_col + 1).Value = primer_dia.month
                ws.Cells(periodo_row + 2, periodo_col).Value = _dt.datetime(ultimo_dia.year, ultimo_dia.month, ultimo_dia.day)
                ws.Cells(periodo_row + 2, periodo_col).NumberFormat = "d/mm/yyyy"
                ws.Cells(periodo_row + 2, periodo_col + 1).Value = ultimo_dia.day
                self.log.info("Fechas inicio/fin escritas correctamente")
            except Exception as e:
                self.log.warning("Error escribiendo fechas inicio/fin: %s", e)
        else:
            self.log.warning("No se encontro celda 'PERIODO MES' en 'Resum Mes'")
        if festivos_row is not None:
            import datetime as _dt
            MAX_FESTIVOS = 5
            for i in range(MAX_FESTIVOS):
                celda = ws.Cells(festivos_row + 1 + i, festivos_col)
                if i < len(festivos_mes):
                    fecha_f, nombre_f = festivos_mes[i]
                    celda.Value = _dt.datetime(fecha_f.year, fecha_f.month, fecha_f.day)
                    celda.NumberFormat = "d/mm/yyyy"
                    self.log.info("Festivo [%d] escrito: %s (%s)", i + 1, nombre_f, fecha_f)
                else:
                    celda.Value = None
            if not festivos_mes:
                self.log.info("Mes %d (%s) no tiene festivos", primer_dia.month, primer_dia.strftime('%B'))

    def calcular_y_escribir_subtotales(self, ws, df_ventas, header_row, data_start_row, data_end_row):
        COLUMNAS_CONFIG = {
            "VENTA A COSTO TOTAL": 9,
            "VENTA TOTAL NETA": 9,
            "MARGEN NETO (DCTO PIE FACT)": 1,
            "MARGEN NETO (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)": 1,
            "MARGEN NETO FACTOR HOY (DCTO PIE FACT)": 1,
            "MARGEN NETO FACTOR HOY (INCLUYE TODOS LOS DCTOS: PIE FACT + FINANC)": 1,
            "PART": 9,
            "SUMA MP": 9,
        }
        used_cols = ws.UsedRange.Columns.Count
        headers_map = {}
        for c in range(1, used_cols + 1):
            val = ws.Cells(header_row, c).Value
            if val:
                headers_map[_norm(str(val))] = c
        subtotal_row = header_row - 1
        if subtotal_row < 1:
            self.log.warning("No hay fila disponible arriba del header para subtotales")
            return
        last_real_row = data_end_row
        def col_to_letter(col_num):
            result = ""
            while col_num > 0:
                col_num -= 1
                result = chr(65 + (col_num % 26)) + result
                col_num //= 26
            return result
        for col_name, func_num in COLUMNAS_CONFIG.items():
            col_norm = _norm(col_name)
            col_idx = None
            for h_norm, idx in headers_map.items():
                if col_norm in h_norm or h_norm in col_norm:
                    col_idx = idx
                    break
            if not col_idx:
                continue
            try:
                target_cell = ws.Cells(subtotal_row, col_idx)
                ref_cell = ws.Cells(header_row + 1, col_idx)
                try:
                    target_cell.NumberFormat = ref_cell.NumberFormat
                except Exception:
                    pass
                col_letter = col_to_letter(col_idx)
                formula = f"=SUBTOTAL({func_num},{col_letter}{data_start_row}:{col_letter}{last_real_row})"
                target_cell.Formula = formula
            except Exception as e:
                self.log.error("Error en subtotal '%s': %s", col_name, e)

    def convertir_tablas_a_rango(self, wb):
        self.log.info("Convirtiendo tablas a rango normal...")
        total_convertidas = 0
        total_hojas = 0
        HOJAS_PROTEGIDAS = {"ULT VTA X REF"}
        for ws_t in wb.Worksheets:
            try:
                if ws_t.Name in HOJAS_PROTEGIDAS:
                    self.log.info("  Hoja '%s': OMITIDA (protegida)", ws_t.Name)
                    continue
                n_tablas = ws_t.ListObjects.Count
                if n_tablas == 0:
                    continue
                total_hojas += 1
                self.log.info("  Hoja '%s': %d tabla(s) encontrada(s)", ws_t.Name, n_tablas)
                for idx in range(n_tablas, 0, -1):
                    try:
                        lo = ws_t.ListObjects(idx)
                        nombre_tabla = lo.Name
                        lo.Unlist()
                        self.log.info("    Tabla '%s' convertida a rango", nombre_tabla)
                        total_convertidas += 1
                    except Exception as e_lo:
                        self.log.warning("    Error convirtiendo tabla %d: %s", idx, e_lo)
            except Exception as e_ws:
                self.log.warning("  Error procesando hoja '%s': %s", ws_t.Name, e_ws)
        if total_convertidas > 0:
            self.log.info("%d tabla(s) convertida(s) a rango en %d hoja(s)", total_convertidas, total_hojas)
        else:
            self.log.info("No se encontraron tablas para convertir")

    def com_write_df_into_template(self, template_xlsx_path, out_xlsx_path, df_data,
                                   sheet_name, password_out, data_start_row=2,
                                   filldown_formula_cols=None, periodo_mes=None):
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        try:
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
        except Exception:
            pass
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        def _excel_serial(d):
            if d is None:
                return None
            return (datetime(d.year, d.month, d.day) - datetime(1899, 12, 30)).days

        self.log.info("Abriendo plantilla: %s", template_xlsx_path.name)
        wb = excel.Workbooks.Open(str(template_xlsx_path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)

        try:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                ws = wb.Worksheets(1)

            old_calc = excel.Calculation
            excel.Calculation = -4135  # xlCalculationManual

            used_cols = max(ws.UsedRange.Columns.Count, ws.Cells(1, ws.Columns.Count).End(-4159).Column)
            header_row = self._detect_header_row(ws, used_cols)
            headers = {str(ws.Cells(header_row, c).Value): c for c in range(1, used_cols + 1)}
            template_headers = [h for h in headers.keys() if h and h != "None"]
            headers_norm = {_norm_label(k): v for k, v in headers.items() if k}

            df_loc = df_data.copy()
            df_loc.columns = [str(c).strip() for c in df_loc.columns]
            col_map = self._build_template_col_map(df_loc, template_headers)

            FORMULA_COLS_ALL = list(dict.fromkeys(self.COLS_FORMULAS_PRE + self.COLS_FORMULAS_POST))

            def _norm_for_comparison(s):
                return _norm(s)

            formula_cols_norm = {_norm_for_comparison(c) for c in FORMULA_COLS_ALL}
            col_map_filtered = {}
            for hdr, df_col in col_map.items():
                if _norm_for_comparison(hdr) not in formula_cols_norm:
                    col_map_filtered[hdr] = df_col
            col_map = col_map_filtered

            n_rows = len(df_loc)
            start_row = max(data_start_row, header_row + 1)
            end_row = start_row + n_rows - 1 if n_rows > 0 else start_row

            self.log.info("Preparando escritura: %d filas desde fila %d", n_rows, start_row)

            if n_rows > 0:
                ref_col_idx = None
                for hdr, col_idx in headers.items():
                    if _norm_label(hdr) == _norm_label(self.COLS_CLAVE_VENTAS["ref"]):
                        ref_col_idx = col_idx
                        break
                if ref_col_idx:
                    self.log.info("Aplicando formato TEXTO a columna REFERENCIA...")
                    ws.Columns(ref_col_idx).NumberFormat = "@"

                self.log.info("Identificando columnas de datos vs formulas...")
                formula_cols_norm2 = {_norm_for_comparison(c) for c in FORMULA_COLS_ALL}
                fecha_key_norm = _norm_for_comparison(self.COLS_CLAVE_VENTAS["fecha"])
                cols_excluir_batch = formula_cols_norm2 | {fecha_key_norm}

                columnas_datos = []
                for hdr in template_headers:
                    hdr_norm = _norm_for_comparison(hdr)
                    if hdr_norm not in cols_excluir_batch and hdr in col_map:
                        columnas_datos.append(hdr)

                self.log.info("Columnas de DATOS a escribir: %d", len(columnas_datos))
                self.log.info("Columna FECHA: se escribira aparte como fecha real")

                self.log.info("Agrupando columnas contiguas para escritura batch...")
                grupos_escritura = []
                grupo_actual = []
                for hdr in template_headers:
                    hdr_norm = _norm_for_comparison(hdr)
                    excel_col_idx = headers.get(hdr)
                    if hdr_norm not in cols_excluir_batch and hdr in col_map and excel_col_idx:
                        df_col = col_map[hdr]
                        if df_col in df_loc.columns:
                            if not grupo_actual or excel_col_idx == grupo_actual[-1]['col_idx'] + 1:
                                grupo_actual.append({'hdr': hdr, 'col_idx': excel_col_idx, 'df_col': df_col})
                            else:
                                if grupo_actual:
                                    grupos_escritura.append(grupo_actual)
                                grupo_actual = [{'hdr': hdr, 'col_idx': excel_col_idx, 'df_col': df_col}]
                    else:
                        if grupo_actual:
                            grupos_escritura.append(grupo_actual)
                            grupo_actual = []
                if grupo_actual:
                    grupos_escritura.append(grupo_actual)

                self.log.info("%d grupos de columnas contiguas identificados", len(grupos_escritura))
                self.log.info("Escribiendo %d filas usando vectorizacion numpy...", n_rows)

                columnas_escritas = 0
                for i, grupo in enumerate(grupos_escritura):
                    try:
                        n_cols = len(grupo)
                        df_cols = [col_info['df_col'] for col_info in grupo]
                        df_subset = df_loc[df_cols].copy()
                        data_matrix = []
                        for _, row in df_subset.iterrows():
                            row_data = [self._to_excel_cell_value(val) for val in row]
                            data_matrix.append(row_data)
                        first_col = grupo[0]['col_idx']
                        last_col = grupo[-1]['col_idx']
                        range_obj = ws.Range(ws.Cells(start_row, first_col), ws.Cells(end_row, last_col))
                        range_obj.Value = data_matrix
                        columnas_escritas += n_cols
                        self.log.info("  Grupo %d/%d: %d columnas escritas", i + 1, len(grupos_escritura), n_cols)
                    except Exception as e:
                        self.log.warning("  Error escribiendo grupo %d: %s", i + 1, e)
                        for col_info in grupo:
                            try:
                                col_data = [[self._to_excel_cell_value(df_loc.iloc[r][col_info['df_col']])]
                                            for r in range(len(df_loc))]
                                col_range = ws.Range(ws.Cells(start_row, col_info['col_idx']),
                                                     ws.Cells(end_row, col_info['col_idx']))
                                col_range.Value = col_data
                                columnas_escritas += 1
                            except Exception as e2:
                                self.log.warning("  Error en columna '%s': %s", col_info['hdr'], e2)

                self.log.info("%d columnas escritas en %d operaciones batch", columnas_escritas, len(grupos_escritura))

                # FECHA como fecha real
                fecha_key = _norm_label(self.COLS_CLAVE_VENTAS["fecha"])
                fecha_idx = headers_norm.get(fecha_key) or headers_norm.get("fecha")
                if fecha_idx and n_rows > 0:
                    hdr_fecha = None
                    for h in template_headers:
                        if _norm_label(h) == fecha_key:
                            hdr_fecha = h
                            break
                    dfcol_fecha = col_map.get(hdr_fecha)
                    if dfcol_fecha and dfcol_fecha in df_loc.columns:
                        self.log.info("Aplicando formato FECHA como fecha real (VECTORIZADO)...")
                        fechas_dt = pd.to_datetime(df_loc[dfcol_fecha], dayfirst=True, errors="coerce")
                        n_nat = fechas_dt.isna().sum()
                        n_ok = fechas_dt.notna().sum()
                        self.log.info("  Fechas validas: %d, nulas: %d", n_ok, n_nat)
                        fechas_valores = []
                        for ts in fechas_dt:
                            if pd.notna(ts):
                                fechas_valores.append([_excel_serial(datetime(ts.year, ts.month, ts.day))])
                            else:
                                fechas_valores.append([None])
                        ws.Columns(fecha_idx).NumberFormat = "dd/mm/yyyy"
                        rng_fecha = ws.Range(ws.Cells(start_row, fecha_idx), ws.Cells(end_row, fecha_idx))
                        rng_fecha.Value = fechas_valores
                        self.log.info("FECHA escrita como fecha real")

                # Filldown formulas
                if filldown_formula_cols and n_rows > 1:
                    self.log.info("Rellenando formulas EN BATCH...")
                    formulas_a_llenar = []
                    for name in filldown_formula_cols:
                        k = _norm_label(name)
                        if k in headers_norm:
                            cidx = headers_norm[k]
                            fml = ws.Cells(start_row, cidx).Formula
                            if fml:
                                formulas_a_llenar.append((cidx, fml))
                    if formulas_a_llenar:
                        self.log.info("  Aplicando %d formulas...", len(formulas_a_llenar))
                        for cidx, fml in formulas_a_llenar:
                            ws.Range(ws.Cells(start_row, cidx), ws.Cells(end_row, cidx)).Formula = fml
                        self.log.info("  %d formulas aplicadas", len(formulas_a_llenar))

                # FECHA ACTUAL EN P1
                self.log.info("Escribiendo fecha actual en P1...")
                try:
                    for col_limpiar in [17, 18, 19]:
                        try:
                            celda = ws.Cells(1, col_limpiar)
                            celda.UnMerge()
                            celda.ClearContents()
                            celda.ClearFormats()
                        except Exception:
                            pass
                    import datetime as _dt2
                    hoy = _dt2.date.today()
                    ws.Cells(1, 16).Value = _dt2.datetime(hoy.year, hoy.month, hoy.day)
                    ws.Cells(1, 16).NumberFormat = "dd/mm/yyyy"
                    self.log.info("Fecha actual escrita en P1: %s", hoy)
                    wb.Application.Calculate()
                    self.log.info("Recalculo forzado tras escribir fecha en P1")
                except Exception as e:
                    self.log.warning("Error escribiendo fecha en P1: %s", e)

                # ===== ACTUALIZAR HOJA "Metas Sub Inf 2026" =====
                self.log.info("Actualizando hoja 'Metas Sub Inf 2026'...")
                try:
                    import datetime as _dt2
                    hoy = _dt2.date.today()
                    mes_actual_num = hoy.month
                    ABREV_MES = {
                        1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR",
                        5: "MAY", 6: "JUN", 7: "JUL", 8: "AGO",
                        9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
                    }
                    ws_metas = wb.Worksheets("Metas Sub Inf 2026")

                    dias_transcurridos_mes_actual = None
                    try:
                        primer_dia_mes = hoy.replace(day=1)
                        _anio = hoy.year
                        _festivos_fijos = {
                            date(_anio, 1, 1), date(_anio, 5, 1), date(_anio, 7, 20),
                            date(_anio, 8, 7), date(_anio, 12, 8), date(_anio, 12, 25),
                        }
                        def _sig_lunes_local(d):
                            dd = (7 - d.weekday()) % 7
                            return d + relativedelta(days=dd) if dd != 0 else d
                        _festivos_lunes = {
                            _sig_lunes_local(date(_anio, 1, 6)), _sig_lunes_local(date(_anio, 3, 19)),
                            _sig_lunes_local(date(_anio, 6, 29)), _sig_lunes_local(date(_anio, 8, 15)),
                            _sig_lunes_local(date(_anio, 10, 12)), _sig_lunes_local(date(_anio, 11, 1)),
                            _sig_lunes_local(date(_anio, 11, 11)),
                        }
                        _a = _anio % 19; _b = _anio // 100; _c = _anio % 100
                        _d_b = _b // 4; _e = _b % 4; _f = (_b + 8) // 25
                        _g = (_b - _f + 1) // 3; _h = (19*_a + _b - _d_b - _g + 15) % 30
                        _i = _c // 4; _k = _c % 4
                        _l = (32 + 2*_e + 2*_i - _h - _k) % 7
                        _m = (_a + 11*_h + 22*_l) // 451
                        _mes_pascua = (_h + _l - 7*_m + 114) // 31
                        _dia_pascua = ((_h + _l - 7*_m + 114) % 31) + 1
                        _pascua = date(_anio, _mes_pascua, _dia_pascua)
                        _festivos_ss = set()
                        for _delta, _trasladar in [(-3, False), (-2, False), (39, True), (60, True), (68, True)]:
                            _fd = _pascua + relativedelta(days=_delta)
                            _festivos_ss.add(_sig_lunes_local(_fd) if _trasladar else _fd)
                        _todos_festivos = _festivos_fijos | _festivos_lunes | _festivos_ss
                        _total = 0.0
                        _d = primer_dia_mes
                        while _d <= hoy:
                            if _d not in _todos_festivos:
                                if _d.weekday() < 5:
                                    _total += 1
                                elif _d.weekday() == 5:
                                    _total += 0.5
                            _d += relativedelta(days=1)
                        dias_transcurridos_mes_actual = _total
                        self.log.info("Dias habiles calculados: %s", dias_transcurridos_mes_actual)
                    except Exception as e:
                        self.log.warning("Error calculando dias habiles: %s", e)
                        try:
                            ws_resum_mes = wb.Worksheets("Resum Mes")
                            dias_transcurridos_mes_actual = ws_resum_mes.Cells(16, 13).Value
                            self.log.info("Fallback - Dias leidos de 'Resum Mes'!M16: %s", dias_transcurridos_mes_actual)
                        except Exception as e2:
                            self.log.warning("Tampoco se pudo leer M16: %s", e2)

                    col = 9
                    actualizadas = 0
                    while True:
                        etiqueta = ws_metas.Cells(5, col).Value
                        if etiqueta is None or str(etiqueta).strip() == "":
                            break
                        etiqueta_str = str(etiqueta).strip().upper()
                        mes_etiqueta = None
                        for num, abrev in ABREV_MES.items():
                            if abrev in etiqueta_str:
                                mes_etiqueta = num
                                break
                        if mes_etiqueta is not None:
                            if mes_etiqueta < mes_actual_num:
                                valor_fila3 = ws_metas.Cells(3, col).Value
                                ws_metas.Cells(2, col).Value = valor_fila3
                                self.log.info("Columna %d (%s): mes pasado -> %s", col, etiqueta_str, valor_fila3)
                                actualizadas += 1
                            elif mes_etiqueta == mes_actual_num:
                                if dias_transcurridos_mes_actual is not None:
                                    ws_metas.Cells(2, col).Value = dias_transcurridos_mes_actual
                                    self.log.info("Columna %d (%s): mes actual -> %s", col, etiqueta_str, dias_transcurridos_mes_actual)
                                    actualizadas += 1
                        col += 1
                        if col > 30:
                            break
                    self.log.info("'Metas Sub Inf 2026' actualizada: %d columnas modificadas", actualizadas)
                except Exception as e:
                    self.log.warning("Error actualizando 'Metas Sub Inf 2026': %s", e)

                # ===== ACTUALIZAR HOJA "Resumen Meta 2026" =====
                self.log.info("Actualizando hoja 'Resumen Meta 2026'...")
                try:
                    ws_resumen_meta = wb.Worksheets("Resumen Meta 2026")
                    actualizadas_rm = 0
                    for fila in range(2, 14):
                        num_mes_celda = ws_resumen_meta.Cells(fila, 2).Value
                        if num_mes_celda is None:
                            continue
                        try:
                            num_mes_celda = int(num_mes_celda)
                        except (ValueError, TypeError):
                            continue
                        if num_mes_celda < mes_actual_num:
                            dias_totales = ws_resumen_meta.Cells(fila, 3).Value
                            ws_resumen_meta.Cells(fila, 4).Value = dias_totales
                            self.log.info("Fila %d (mes %d): mes pasado -> %s", fila, num_mes_celda, dias_totales)
                            actualizadas_rm += 1
                        elif num_mes_celda == mes_actual_num:
                            if dias_transcurridos_mes_actual is not None:
                                ws_resumen_meta.Cells(fila, 4).Value = dias_transcurridos_mes_actual
                                self.log.info("Fila %d (mes %d): mes actual -> %s", fila, num_mes_celda, dias_transcurridos_mes_actual)
                                actualizadas_rm += 1
                        else:
                            ws_resumen_meta.Cells(fila, 4).Value = 0
                            actualizadas_rm += 1
                    self.log.info("'Resumen Meta 2026' actualizada: %d filas modificadas", actualizadas_rm)
                except Exception as e:
                    self.log.warning("Error actualizando 'Resumen Meta 2026': %s", e)

                # Copiar formato
                if n_rows > 1:
                    self.log.info("Copiando formato de primera fila...")
                    try:
                        end_row_original = end_row
                        ref_range = ws.Range(ws.Cells(start_row, 1), ws.Cells(start_row, used_cols))
                        tgt_range = ws.Range(ws.Cells(start_row + 1, 1), ws.Cells(end_row, used_cols))
                        ref_range.Copy()
                        tgt_range.PasteSpecial(Paste=-4122)
                        excel.CutCopyMode = False
                        self.log.info("Formato copiado a %d filas", n_rows - 1)
                        ultima_fila_real = ws.UsedRange.Rows.Count
                        if ultima_fila_real > end_row_original:
                            filas_extra = ultima_fila_real - end_row_original
                            self.log.warning("Limpiando %d filas extra...", filas_extra)
                            try:
                                rango_extra = ws.Range(ws.Cells(end_row_original + 1, 1), ws.Cells(ultima_fila_real, used_cols))
                                rango_extra.Delete()
                                self.log.info("Filas extras eliminadas")
                            except Exception as e:
                                self.log.warning("Error limpiando: %s", e)
                    except Exception as e:
                        self.log.warning("Error al copiar formato: %s", e)

            # Subtotales
            self.log.info("Calculando subtotales...")
            self.calcular_y_escribir_subtotales(ws, df_loc, header_row, start_row, end_row)

            # Calculo completo
            self.log.info("Forzando calculo completo de Excel...")
            excel.Calculation = -4105
            wb.Application.Calculate()
            wb.Application.CalculateFullRebuild()
            import time as _time
            _time.sleep(2)
            self.log.info("Calculo completado")

            # ===== APLICAR REGLA DE CONFLICTO POST-CALCULO =====
            self.log.info("APLICANDO REGLA DE CONFLICTO (POST-CALCULO)")
            dcto_col_idx = None
            pie_col_idx = None
            fecha_col_idx = None
            for h, idx in headers.items():
                h_norm = _norm_label(h)
                if "dcto" in h_norm and "condicionado" in h_norm:
                    dcto_col_idx = idx
                if "porcentaje" in h_norm and "pie" in h_norm and "factura" in h_norm:
                    pie_col_idx = idx
                if h_norm == "fecha":
                    fecha_col_idx = idx
            if dcto_col_idx and pie_col_idx and fecha_col_idx:
                mes_actual = self._month_now_bogota()[0]
                anio_actual = self._month_now_bogota()[1]
                dcto_df_col = None
                pie_df_col = None
                fecha_df_col = None
                for col in df_loc.columns:
                    cn = _norm_label(col)
                    if "dcto" in cn and "condicionado" in cn and not dcto_df_col:
                        dcto_df_col = col
                    if "porcentaje" in cn and "pie" in cn and "factura" in cn and not pie_df_col:
                        pie_df_col = col
                    if cn == "fecha" and not fecha_df_col:
                        fecha_df_col = col
                if fecha_df_col and dcto_df_col and pie_df_col:
                    fechas_dt = pd.to_datetime(df_loc[fecha_df_col], errors='coerce')
                    df_analisis = pd.DataFrame({
                        'mes': fechas_dt.dt.month, 'anio': fechas_dt.dt.year,
                        'dcto_num': pd.to_numeric(df_loc[dcto_df_col], errors='coerce').fillna(0),
                        'pie_num': pd.to_numeric(df_loc[pie_df_col], errors='coerce').fillna(0),
                    })
                    mask_mes_actual = (df_analisis['mes'] == mes_actual) & (df_analisis['anio'] == anio_actual)
                    df_mes_actual = df_analisis[mask_mes_actual].copy()
                    n_mes_actual = len(df_mes_actual)
                    mask_conflictos = (df_mes_actual['dcto_num'] != 0) & (df_mes_actual['pie_num'] != 0)
                    n_conflictos = mask_conflictos.sum()
                    self.log.info("Registros de %d/%d: %d", mes_actual, anio_actual, n_mes_actual)
                    self.log.info("Conflictos detectados (DCTO!=0 y PIE!=0): %d", n_conflictos)
                    if n_conflictos > 0:
                        indices_conflictos = df_mes_actual[mask_conflictos].index.tolist()
                        self.log.info("Escribiendo %d correcciones...", n_conflictos)
                        for idx in indices_conflictos:
                            excel_row = start_row + idx
                            ws.Cells(excel_row, dcto_col_idx).Value = 0
                        self.log.info("%d registros corregidos exitosamente", n_conflictos)

            # Validaciones
            self.validar_sublinea_en_cero(wb)
            if periodo_mes:
                self.actualizar_periodo_mes(wb, periodo_mes[0], periodo_mes[1])

            # Activar hoja
            try:
                ws_ventas = wb.Worksheets(sheet_name)
                ws_ventas.Activate()
            except Exception:
                pass

            self.aplicar_formato_numero_columna_c(ws, header_row)
            self.convertir_texto_a_numero_columna_c(ws, header_row)

            # Propagar formulas en columnas AE y AL
            self.log.info("Propagando formulas 'MARGEN NETO' en columnas AE y AL...")
            try:
                columnas_margen = {31: "AE (MARGEN NETO con AC)", 38: "AL (MARGEN NETO con AJ)"}
                col_ref = 3
                try:
                    ultima_celda = ws.Columns(col_ref).SpecialCells(11)
                    ultima_fila = ultima_celda.Row
                except Exception:
                    try:
                        ultima_fila = ws.UsedRange.Rows.Count
                    except Exception:
                        ultima_fila = 3
                        for fila in range(200000, 2, -1):
                            valor = ws.Cells(fila, col_ref).Value
                            if valor is not None and str(valor).strip() != "":
                                ultima_fila = fila
                                break
                for col_idx, col_nombre in columnas_margen.items():
                    celda_origen = ws.Cells(3, col_idx)
                    formula_base = celda_origen.Formula
                    if formula_base and str(formula_base).startswith('='):
                        rango_completo = ws.Range(ws.Cells(3, col_idx), ws.Cells(ultima_fila, col_idx))
                        celda_origen.AutoFill(Destination=rango_completo, Type=0)
                        self.log.info("AutoFill completado para %s", col_nombre)
                ws.Calculate()
                self.log.info("Formulas propagadas exitosamente en AE y AL")
            except Exception as e:
                self.log.error("Error al propagar formulas: %s", e)

            # Romper vinculos
            self.log.info("Eliminando vinculos externos...")
            try:
                links = wb.LinkSources(1)
                if links:
                    for link in links:
                        wb.BreakLink(Name=link, Type=1)
            except Exception:
                pass

            # Ocultar columnas de margen en "Resum Mes"
            self.log.info("Ocultando columnas de margen en 'Resum Mes'...")
            try:
                ws_resum = wb.Worksheets("Resum Mes")
                COLS_OCULTAR = ["promedio de margen neto (dcto pie fact)", "ppromedio margen bruto"]
                used_cols_rm = ws_resum.UsedRange.Columns.Count
                used_rows_rm = min(40, ws_resum.UsedRange.Rows.Count)
                cols_ocultar_norm = [_norm(c) for c in COLS_OCULTAR]
                for r in range(1, used_rows_rm + 1):
                    for c in range(1, used_cols_rm + 1):
                        val = ws_resum.Cells(r, c).Value
                        if val and _norm(str(val)) in cols_ocultar_norm:
                            ws_resum.Columns(c).Hidden = True
                            self.log.info("Ocultada: col %d ('%s')", c, val)
            except Exception as e:
                self.log.warning("Error al ocultar columnas en 'Resum Mes': %s", e)

            # Convertir tablas a rango
            self.convertir_tablas_a_rango(wb)

            # Guardar
            self.log.info("Guardando archivo con contrasena...")
            try:
                if password_out:
                    file_format = 51
                    wb.SaveAs(Filename=str(out_xlsx_path), FileFormat=file_format,
                              Password=password_out, WriteResPassword="", ReadOnlyRecommended=False, CreateBackup=False)
                else:
                    wb.SaveAs(Filename=str(out_xlsx_path))
            except Exception as e:
                self.log.error("Error al guardar: %s", e)
                try:
                    wb.SaveAs(Filename=str(out_xlsx_path))
                except Exception:
                    raise
            self.log.info("Archivo guardado: %s", out_xlsx_path.name)

        finally:
            try:
                excel.Calculation = old_calc
                excel.ScreenUpdating = True
            except Exception:
                pass
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass

    # ─────────────────────────────────────────────────────────────────
    #  ACTUALIZAR TABLA DINAMICA
    # ─────────────────────────────────────────────────────────────────

    def actualizar_tabla_dinamica_resumen_dia(self, archivo_path, nombre_hoja=None):
        if nombre_hoja:
            self.log.info("ACTUALIZANDO TABLA DINAMICA: %s", nombre_hoja)
        else:
            self.log.info("ACTUALIZANDO TODAS LAS TABLAS DINAMICAS")
        if not HAS_COM:
            self.log.error("COM no disponible, saltando actualizacion")
            return
        excel = None
        wb = None
        try:
            self.log.info("Abriendo archivo: %s", archivo_path.name)
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            try:
                excel.AskToUpdateLinks = False
                excel.EnableEvents = False
                excel.AutomationSecurity = 3
            except Exception:
                pass
            try:
                wb = excel.Workbooks.Open(str(archivo_path), 0, False, None, self.PASSWORD_VENTAS, "", True)
                self.log.info("Archivo abierto con Metodo 1")
            except Exception as e1:
                self.log.warning("Metodo 1 fallo: %s", str(e1)[:100])
                try:
                    wb = excel.Workbooks.Open(str(archivo_path), 0, False, None, self.PASSWORD_VENTAS)
                    self.log.info("Archivo abierto con Metodo 2")
                except Exception as e2:
                    self.log.warning("Metodo 2 fallo: %s", str(e2)[:100])
                    try:
                        archivo_stream = self._decrypt_to_stream(archivo_path, self.PASSWORD_VENTAS)
                        with self.archivo_temporal_seguro(archivo_stream, "pivot_update") as tmp_path:
                            self.log.info("Abriendo temporal: %s", tmp_path.name)
                            wb = excel.Workbooks.Open(str(tmp_path))
                            self.log.info("Archivo abierto con Metodo 3")
                            archivo_final_para_guardar = archivo_path
                    except Exception as e3:
                        self.log.error("Todos los metodos fallaron")
                        return
            if not wb:
                self.log.error("No se pudo abrir el archivo")
                return

            self.log.info("Detectando hoja de datos y rango...")
            hoja_datos = None
            for sheet in wb.Worksheets:
                nombre_sheet = str(sheet.Name).strip().upper()
                if "VENTAS" in nombre_sheet and ("2026" in nombre_sheet or "2025" in nombre_sheet):
                    hoja_datos = sheet
                    self.log.info("Hoja de datos: '%s'", sheet.Name)
                    break
            if not hoja_datos:
                for sheet in wb.Worksheets:
                    nombre_sheet = str(sheet.Name).strip().upper()
                    if "RESUMEN" not in nombre_sheet and "PIVOT" not in nombre_sheet:
                        hoja_datos = sheet
                        break
            if not hoja_datos:
                hoja_datos = wb.Worksheets(1)

            self.log.info("Detectando ultima fila con datos...")
            ultima_fila_datos = 2
            try:
                col_ref = 3
                try:
                    ultima_celda = hoja_datos.Columns(col_ref).SpecialCells(11)
                    ultima_fila_datos = ultima_celda.Row
                except Exception:
                    ultima_fila_datos = hoja_datos.UsedRange.Rows.Count
            except Exception:
                ultima_fila_datos = 100000
            try:
                ultima_col_datos = hoja_datos.UsedRange.Columns.Count
            except Exception:
                ultima_col_datos = 50

            def numero_a_letra_columna(n):
                import string
                result = ""
                while n > 0:
                    n -= 1
                    result = string.ascii_uppercase[n % 26] + result
                    n //= 26
                return result

            letra_ultima_col = numero_a_letra_columna(ultima_col_datos)
            nuevo_rango = f"'{hoja_datos.Name}'!$A$2:${letra_ultima_col}${ultima_fila_datos}"
            self.log.info("Rango detectado: %s", nuevo_rango)

            total_tablas = 0
            tablas_actualizadas = 0
            if nombre_hoja:
                hojas_a_procesar = []
                for sheet in wb.Worksheets:
                    nombre = str(sheet.Name).strip().upper()
                    if nombre_hoja.upper() in nombre or nombre in nombre_hoja.upper():
                        hojas_a_procesar.append(sheet)
                        break
                if not hojas_a_procesar:
                    self.log.error("No se encontro la hoja '%s'", nombre_hoja)
                    return
            else:
                hojas_a_procesar = [sheet for sheet in wb.Worksheets]

            for ws in hojas_a_procesar:
                try:
                    pivot_tables = []
                    try:
                        for pt in ws.PivotTables():
                            pivot_tables.append(pt)
                    except Exception:
                        pass
                    if not pivot_tables:
                        continue
                    self.log.info("Hoja '%s': %d tabla(s) dinamica(s)", ws.Name, len(pivot_tables))
                    total_tablas += len(pivot_tables)
                    ws.Activate()
                    for i, pt in enumerate(pivot_tables, 1):
                        try:
                            nombre_pt = pt.Name
                            self.log.info("  [%d/%d] Actualizando '%s'...", i, len(pivot_tables), nombre_pt)
                            try:
                                rango_actual = pt.SourceData
                                if str(ultima_fila_datos) not in str(rango_actual):
                                    try:
                                        pivot_cache = wb.PivotCaches().Create(SourceType=1, SourceData=nuevo_rango)
                                        pt.ChangePivotCache(pivot_cache)
                                        self.log.info("    Rango actualizado")
                                    except Exception as e_cache:
                                        self.log.warning("    No se pudo actualizar rango: %s", e_cache)
                            except Exception as e_rango:
                                self.log.warning("    Error verificando rango: %s", e_rango)
                            pt.RefreshTable()
                            tablas_actualizadas += 1
                            self.log.info("    Refrescada")
                            try:
                                for pf in pt.PivotFields():
                                    try:
                                        nombre_campo = str(pf.Name).strip().upper()
                                        if "FECHA" in nombre_campo and "ACTUALIZACION" not in nombre_campo:
                                            try:
                                                pf.Ungroup()
                                            except Exception:
                                                pass
                                            try:
                                                pf.Group(Periods=[False, False, False, True, True, False, True])
                                            except Exception as e_group:
                                                self.log.warning("    No se pudo reagrupar '%s': %s", pf.Name, e_group)
                                            pf.NumberFormat = "dd/mm/yyyy"
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        except Exception as e:
                            self.log.error("    Error: %s", e)
                except Exception as e_hoja:
                    self.log.warning("Error procesando hoja '%s': %s", ws.Name, e_hoja)

            self.log.info("Recalculando workbook...")
            wb.Application.Calculate()
            self.log.info("Guardando...")
            try:
                if 'archivo_final_para_guardar' in locals():
                    wb.SaveAs(Filename=str(archivo_final_para_guardar), FileFormat=51,
                              Password=self.PASSWORD_VENTAS, WriteResPassword="", CreateBackup=False)
                else:
                    wb.Save()
                self.log.info("Guardado exitosamente")
            except Exception as e:
                self.log.error("Error guardando: %s", e)
            self.log.info("COMPLETADO: %d/%d tablas actualizadas", tablas_actualizadas, total_tablas)

        except Exception as e:
            self.log.error("Error general: %s", e)
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            wb = None
            excel = None
            import gc
            gc.collect()

    # ═══════════════════════════════════════════════════════════════════
    #  MAIN EXECUTION
    # ═══════════════════════════════════════════════════════════════════

    def _ejecutar(self):
        log = self.log
        BASE_PATH = self.BASE_PATH
        PASSWORD_VENTAS = self.PASSWORD_VENTAS
        PASSWORD_INV_GENERAL = self.PASSWORD_INV_GENERAL
        PASSWORD_MYR = self.PASSWORD_MYR

        try:
            log.info("=" * 60)
            log.info("== Inicio actualizacion $2026 VENTAS ==")
            log.info("=" * 60)

            mes_no, anio, pfirst, plast = self._month_now_bogota()
            log.info("Periodo detectado: mes=%d ano=%d  (%s -> %s)", mes_no, anio, pfirst, plast)

            # Buscar archivos
            log.info("BUSCANDO ARCHIVO DE VENTAS (PLANTILLA)")
            p_ventas = self.find_file_by_loose_name(BASE_PATH, self.FN_VENTAS)
            log.info("BUSCANDO OTROS ARCHIVOS")
            p_inv = self.find_file_by_loose_name(BASE_PATH, self.FN_INV_GENERAL)
            p_myr = self.find_myr_existencia_by_fecha(BASE_PATH)
            p_mat = self.find_matriz_clientes_by_prefix(BASE_PATH)
            log.info("BUSCANDO INFORME DE VENTAS (FACTURAS)")
            p_inf = self.find_informe_facturas_by_prefix(
                self.DIR_INFORME_VENTAS_MES, prefix="InformesDeVentas(Facturas)", only_today=False
            )
            log.info("Archivos encontrados:")
            log.info("  VENTAS: %s", p_ventas.name)
            log.info("  INVENTARIO: %s", p_inv.name)
            log.info("  MYR: %s", p_myr.name)
            log.info("  INFORME: %s", p_inf.name)
            log.info("  MATRIZ: %s", p_mat.name)

            # =================== PASO 1: CARGAR PLANTILLA ===================
            log.info("Paso 1: Abriendo $2026 VENTAS y normalizando columnas...")
            ventas_stream = self._decrypt_to_stream(p_ventas, PASSWORD_VENTAS)
            SHEET_VENTAS = None
            df_ventas = None
            for sheet_option in self.SHEET_VENTAS_OPTIONS:
                try:
                    df_ventas = pd.read_excel(ventas_stream, sheet_name=sheet_option, engine="openpyxl", header=1)
                    SHEET_VENTAS = sheet_option
                    log.info("Usando hoja: %s", SHEET_VENTAS)
                    break
                except ValueError:
                    continue
            if df_ventas is None:
                raise ValueError(f"No se encontro ninguna hoja de ventas valida. Intentadas: {self.SHEET_VENTAS_OPTIONS}")
            df_ventas = df_ventas.loc[:, ~df_ventas.columns.astype(str).str.startswith("Unnamed")]
            df_ventas.columns = [str(c).strip() for c in df_ventas.columns]
            df_ventas, mapeo_ventas = self.normalizar_columnas_df(df_ventas)
            if mapeo_ventas:
                log.info("  %d columnas normalizadas:", len(mapeo_ventas))

            COLUMNAS_ORDEN_ORIGINAL = list(df_ventas.columns)
            with self.archivo_temporal_seguro(ventas_stream, "template_full") as tmp_template:

                COLS_CLAVE_VENTAS_NORM = {
                    "nit": "NIT CLIENTE", "ref": "REFERENCIA", "marca": "MARCA",
                    "fecha": "FECHA", "mes": "MES", "mes_no": "MES NO.", "anio": "ANO", "numero": "NUMERO",
                }
                COLS_MATRIZ_MAP_NORM = {"nit": "NIT", "dcto_cond": "DCTO CONDICIONADO", "porc_pie_fact": "PORCENTAJE DCTO A PIE DE FACTURA"}

                log.info("Paso 1.5: Modo ANO COMPLETO — se reemplazaran TODOS los datos")
                n_total_anterior = len(df_ventas)
                log.info("  Total registros en plantilla anterior: %d", n_total_anterior)

                # =================== BACKUP AUTOMATICO ===================
                log.info("Creando backup automatico del archivo anterior...")
                try:
                    CARPETA_BACKUPS = BASE_PATH / "Backups_Ventas"
                    CARPETA_BACKUPS.mkdir(exist_ok=True)
                    backup_name = self.FN_VENTAS.replace(".xlsx", f"_BACKUP_{date.today().strftime('%Y-%m-%d_%H%M')}.xlsx")
                    backup_path = CARPETA_BACKUPS / backup_name
                    shutil.copy2(p_ventas, backup_path)
                    log.info("Backup creado: %s", backup_path.name)
                    backups_existentes = sorted(CARPETA_BACKUPS.glob("*_BACKUP_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
                    for viejo in backups_existentes[7:]:
                        viejo.unlink()
                        log.info("Backup antiguo eliminado: %s", viejo.name)
                except Exception as e:
                    log.warning("No se pudo crear backup: %s", e)

                # =================== PASO 2: PROCESAR INFORME ===================
                log.info("PROCESANDO INFORME DE VENTAS (NUEVO — ANO COMPLETO)")
                df_inf = self._read_excel_any(p_inf)
                df_inf = self.transformar_informe_ventas(df_inf)
                df_inf, mapeo_inf = self.normalizar_columnas_df(df_inf)
                if mapeo_inf:
                    log.info("  %d columnas normalizadas en informe", len(mapeo_inf))

                # Eliminar FLETE VENTAS
                if 'REFERENCIA' in df_inf.columns:
                    n_antes = len(df_inf)
                    df_inf = df_inf[df_inf['REFERENCIA'].astype(str).str.strip().str.upper() != 'FLETE VENTAS'].copy()
                    n_eliminados = n_antes - len(df_inf)
                    if n_eliminados > 0:
                        log.info("  Eliminados %d registros con REFERENCIA = 'FLETE VENTAS'", n_eliminados)
                # Eliminar PUBLICIDAD
                if 'REFERENCIA' in df_inf.columns:
                    n_antes = len(df_inf)
                    mask_no_publicidad = ~df_inf['REFERENCIA'].astype(str).str.upper().str.contains('PUBLICIDAD', na=False)
                    df_inf = df_inf[mask_no_publicidad].copy()
                    n_eliminados = n_antes - len(df_inf)
                    if n_eliminados > 0:
                        log.info("  Eliminados %d registros con 'PUBLICIDAD' en REFERENCIA", n_eliminados)

                if 'FECHA' not in df_inf.columns:
                    raise ValueError("Columna FECHA no encontrada en informe")
                df_inf['FECHA'] = pd.to_datetime(df_inf['FECHA'], dayfirst=True, errors='coerce').dt.normalize()

                # ANO COMPLETO
                anio_actual = anio
                df_inf_octubre = df_inf[df_inf['FECHA'].dt.year == anio_actual].copy()

                log.info("Calculando MES, MES NO., ANO nuevo...")
                if 'FECHA' in df_inf_octubre.columns:
                    df_inf_octubre['FECHA'] = pd.to_datetime(df_inf_octubre['FECHA'], dayfirst=True, errors='coerce').dt.normalize()
                    df_inf_octubre['ANO'] = df_inf_octubre['FECHA'].dt.year
                    df_inf_octubre['MES NO.'] = df_inf_octubre['FECHA'].dt.month
                    meses_es = {
                        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
                        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
                        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
                    }
                    df_inf_octubre['MES'] = df_inf_octubre['MES NO.'].map(meses_es)
                    n_con_mes = df_inf_octubre['MES NO.'].notna().sum()
                    mes_unico = sorted(df_inf_octubre['MES NO.'].dropna().unique())
                    log.info("  MES NO. calculado: %d/%d registros", n_con_mes, len(df_inf_octubre))
                    log.info("  Meses presentes en informe: %s", mes_unico)
                else:
                    raise ValueError("No se puede calcular MES NO. sin columna FECHA")
                log.info("  Registros ano completo en informe: %d", len(df_inf_octubre))
                if len(df_inf_octubre) == 0:
                    raise ValueError("No hay datos del ano actual en el informe")

                # Mapeo semantico
                log.info("Aplicando mapeos semanticos...")
                MAPEO_SEMANTICO = {
                    "NRO. DOCUMENTO CLIENTE": "NIT CLIENTE", "CIUDAD/SUCURSAL": "CIUDAD",
                    "DESCRIPCION": "DESCRPCION", "VALOR UNITARIO": "VR UNITARIO",
                    "CANTIDAD FACTURADA": "CANTIDAD", "VALOR BRUTO": "VR TOTAL",
                    "COSTO UNITARIO": "COSTO PROMEDIO", "VENDEDOR": "VEND", "PREFIJO": "DV",
                }
                columnas_mapeadas = {k: v for k, v in MAPEO_SEMANTICO.items() if k in df_inf_octubre.columns}
                if columnas_mapeadas:
                    log.info("  Mapeando %d columnas", len(columnas_mapeadas))
                    df_inf_octubre = df_inf_octubre.rename(columns=columnas_mapeadas)

                log.info("Validando columnas criticas en nuevo...")
                for col in ['REFERENCIA', 'NIT CLIENTE', 'DV', 'CLIENTE', 'CIUDAD', 'VEND', 'CANTIDAD']:
                    if col in df_inf_octubre.columns:
                        n_vals = df_inf_octubre[col].notna().sum()
                        pct = (n_vals / len(df_inf_octubre)) * 100
                        log.info("  %s: %d/%d (%.1f%%)", col, n_vals, len(df_inf_octubre), pct)
                    else:
                        log.info("  %s: NO ENCONTRADA", col)

                # =================== PASO 3: LIMPIEZAS ===================
                log.info("Paso 3: Aplicando limpiezas...")
                df_inf_octubre = self.ordenar_y_fechas(df_inf_octubre, COLS_CLAVE_VENTAS_NORM)
                for cand in ["FECHA DE ACTUALIZACION", "FECHA ACTUALIZACION"]:
                    if cand in df_inf_octubre.columns:
                        df_inf_octubre[cand] = date.today()
                        break

                # =================== PASO 4: INTEGRAR INVENTARIO ===================
                log.info("Paso 4: Integrando LINEA/SUBLINEA...")
                inv_stream = self._decrypt_to_stream(p_inv, PASSWORD_INV_GENERAL)
                sheet_inv = find_sheet_name(inv_stream, targets=("INVENTARIO", "INVENTARIO GENERAL"))
                def _leer_inv(stream, header_row):
                    stream.seek(0)
                    df = pd.read_excel(stream, sheet_name=sheet_inv, engine="openpyxl", header=header_row)
                    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
                    df.columns = [str(c).strip() for c in df.columns]
                    return df
                df_inv = _leer_inv(inv_stream, 1)
                if "REFERENCIA" not in df_inv.columns:
                    df_inv = _leer_inv(inv_stream, 0)
                df_inv, _ = self.normalizar_columnas_df(df_inv)
                synonyms_inv = {"LINEA COPIA": ["LINEA COPIA", "LINEA"], "SUB-LINEA COPIA": ["SUB-LINEA COPIA", "SUBLINEA", "SUB-LINEA"]}
                for target, cands in synonyms_inv.items():
                    if target not in df_inv.columns:
                        for cand in cands:
                            if cand in df_inv.columns:
                                df_inv.rename(columns={cand: target}, inplace=True)
                                break
                df_inf_octubre = self.integrar_linea_sublinea(df_inf_octubre, df_inv, COLS_CLAVE_VENTAS_NORM)

                # =================== PASO 5: INTEGRAR MYR ===================
                log.info("Paso 5: Integrando COSTO FACTOR HOY...")
                myr_stream = self._decrypt_to_stream(p_myr, PASSWORD_MYR)
                df_myr = pd.read_excel(myr_stream, sheet_name=self.SHEET_COSTOS_INVFINAL, engine="openpyxl", header=2)
                df_myr = df_myr.loc[:, ~df_myr.columns.astype(str).str.startswith("Unnamed")]
                df_myr.columns = [str(c).strip() for c in df_myr.columns]
                df_myr, _ = self.normalizar_columnas_df(df_myr)
                if "REFERENCIA FERTRAC" in df_myr.columns:
                    df_myr.rename(columns={"REFERENCIA FERTRAC": "REFERENCIA"}, inplace=True)
                df_inf_octubre = self.integrar_costo_factor_hoy(df_inf_octubre, df_myr, COLS_CLAVE_VENTAS_NORM)

                # =================== PASO 6: INTEGRAR MATRIZ ===================
                log.info("Paso 6: Integrando DCTO CONDICIONADO...")
                log.info("  Leyendo MATRIZ desde: %s", p_mat)
                stream = self._decrypt_to_stream(p_mat, password=PASSWORD_VENTAS)
                df_mat_raw = self._read_excel_any(stream, sheet_name=self.SHEET_MATRIZ)
                df_mat_raw, mapeo_mat = self.normalizar_columnas_df(df_mat_raw)
                if mapeo_mat:
                    log.info("  %d columnas normalizadas en Matriz", len(mapeo_mat))
                log.info("  Registros totales en matriz: %d", len(df_mat_raw))

                def _ajustar_matriz_clientes(df_mat):
                    df = df_mat.copy()
                    log.info("APLICANDO LOGICA DE DESCUENTOS:")
                    log.info("  Registros iniciales: %d", len(df))
                    if "TIPO DESCUENTO" in df.columns:
                        n_antes = len(df)
                        df = df[df["TIPO DESCUENTO"].astype(str).str.strip().str.upper() == "CONDICIONADO"].copy()
                        log.info("  Filtro TIPO DESCUENTO='CONDICIONADO': %d -> %d registros", n_antes, len(df))
                    if "DESCUENTO" in df.columns:
                        def normalizar_a_decimal(valor):
                            if pd.isna(valor):
                                return 0
                            if isinstance(valor, (int, float)):
                                if 0 <= valor <= 1:
                                    return valor
                                elif valor > 1:
                                    return valor / 100
                                else:
                                    return 0
                            val_str = str(valor).strip()
                            if val_str == "" or val_str.upper() == "NAN":
                                return 0
                            val_norm = unidecode(val_str).upper().strip()
                            if "SIN DTO" in val_norm or "SIN DESCUENTO" in val_norm:
                                return 0
                            if "5%" in val_norm and ("MAXIMO" in val_norm or "MAX" in val_norm) and "60" in val_norm:
                                return 0.05
                            match = re.search(r'(\d+(?:\.\d+)?)\s*%', val_str)
                            if match:
                                num = float(match.group(1))
                                return num / 100
                            try:
                                val_num = float(val_str.replace("%", "").replace(",", ".").strip())
                                if 0 <= val_num <= 1:
                                    return val_num
                                elif val_num > 1:
                                    return val_num / 100
                                else:
                                    return 0
                            except Exception:
                                return 0
                        df["DESCUENTO"] = df["DESCUENTO"].apply(normalizar_a_decimal)
                    log.info("  Registros finales: %d", len(df))
                    return df

                df_mat_adj = _ajustar_matriz_clientes(df_mat_raw)

                if "NIT CLIENTE" in df_inf_octubre.columns and "NIT" in df_mat_adj.columns and "DESCUENTO" in df_mat_adj.columns:
                    columnas_merge = ["NIT", "DESCUENTO"]
                    df_mat_merge = df_mat_adj[columnas_merge].copy()
                    df_mat_merge = df_mat_merge.drop_duplicates("NIT", keep="last")
                    df_mat_merge = df_mat_merge[df_mat_merge["NIT"].notna()].copy()
                    log.info("  Clientes unicos: %d", len(df_mat_merge))
                    df_inf_octubre["NIT CLIENTE"] = df_inf_octubre["NIT CLIENTE"].astype(str).str.strip()
                    df_inf_octubre["NIT CLIENTE"] = df_inf_octubre["NIT CLIENTE"].replace({"nan": None, "None": None, "": None})
                    df_inf_octubre["NIT CLIENTE"] = df_inf_octubre["NIT CLIENTE"].str.replace(r'\.0$', '', regex=True)
                    df_mat_merge["NIT"] = df_mat_merge["NIT"].astype(str).str.strip()
                    df_mat_merge["NIT"] = df_mat_merge["NIT"].replace({"nan": None, "None": None, "": None})
                    df_mat_merge["NIT"] = df_mat_merge["NIT"].str.replace(r'\.0$', '', regex=True)
                    df_inf_octubre = df_inf_octubre.merge(df_mat_merge, left_on="NIT CLIENTE", right_on="NIT", how="left", suffixes=("", "_mat"))
                    if "DESCUENTO" in df_inf_octubre.columns:
                        df_inf_octubre = df_inf_octubre.rename(columns={"DESCUENTO": "DCTO CONDICIONADO"})
                        n_vacios = df_inf_octubre["DCTO CONDICIONADO"].isna().sum()
                        if n_vacios > 0:
                            df_inf_octubre["DCTO CONDICIONADO"] = df_inf_octubre["DCTO CONDICIONADO"].fillna(0)
                        log.info("  Columna 'DCTO CONDICIONADO' integrada")
                    if "NIT" in df_inf_octubre.columns:
                        df_inf_octubre = df_inf_octubre.drop(columns=["NIT"])
                else:
                    log.warning("  SALTANDO integracion de DCTO CONDICIONADO")

                # =================== PASO 7: LIMPIEZA FINAL ===================
                log.info("Paso 7: Limpieza final...")
                df_inf_octubre = self.limpiar_linea_referencias_invalidas(df_inf_octubre, COLS_CLAVE_VENTAS_NORM)
                if "PORCENTAJE DCTO A PIE DE FACTURA" not in df_inf_octubre.columns:
                    df_inf_octubre["PORCENTAJE DCTO A PIE DE FACTURA"] = ""
                df_inf_octubre = self.normalizar_dctos(df_inf_octubre)

                # =================== PASO 7.5: VTA ACORDADA X UNIDAD LICITADO ===================
                from config.settings import get_month_folder as _get_month_folder
                MONTH_FOLDER = _get_month_folder()
                log.info("PASO 7.5: VTA ACORDADA X UNIDAD LICITADO (SOLO %s)", MONTH_FOLDER)
                try:
                    log.info("  Cargando hoja 'PRECIO UNIT LICITADOS' desde plantilla...")
                    ventas_stream.seek(0)
                    df_precios = pd.read_excel(ventas_stream, sheet_name="PRECIO UNIT LICITADOS", engine="openpyxl", header=3)
                    log.info("  Hoja cargada: %d filas", len(df_precios))
                    df_precios.columns = [str(c).strip() for c in df_precios.columns]
                    col_ref_precios = None
                    for col in df_precios.columns:
                        if "REFERENCIA" in str(col).upper() and "FERTRAC" in str(col).upper():
                            col_ref_precios = col
                            break
                    if not col_ref_precios:
                        for col in df_precios.columns:
                            if "REFERENCIA" in str(col).upper():
                                col_ref_precios = col
                                break
                    if not col_ref_precios:
                        raise ValueError("No se encontro columna REFERENCIA")
                    log.info("  Columna referencia: '%s'", col_ref_precios)

                    ventas_stream.seek(0)
                    df_primera_fila = pd.read_excel(ventas_stream, sheet_name="PRECIO UNIT LICITADOS", engine="openpyxl", header=None, nrows=1)
                    nits_headers = {}
                    col_letters = ['C', 'D', 'E', 'F', 'G', 'H']
                    for idx, letra in enumerate(col_letters, start=2):
                        if idx < len(df_primera_fila.columns):
                            nit_val = df_primera_fila.iloc[0, idx]
                            if pd.notna(nit_val):
                                nit_str = str(nit_val).strip().replace('.0', '')
                                if nit_str and nit_str not in ['', 'nan', 'None']:
                                    col_name = df_precios.columns[idx]
                                    nits_headers[nit_str] = col_name
                    if not nits_headers:
                        raise ValueError("No hay NITs en headers")
                    log.info("  %d NITs encontrados en headers", len(nits_headers))

                    tabla_busqueda = {}
                    for idx, row in df_precios.iterrows():
                        ref = str(row[col_ref_precios]).strip().replace('.0', '') if pd.notna(row[col_ref_precios]) else ""
                        if ref and ref not in ['', 'nan', 'None']:
                            precios_por_nit = {}
                            for nit, col_precio in nits_headers.items():
                                if col_precio in df_precios.columns:
                                    precio = row[col_precio]
                                    if pd.notna(precio):
                                        try:
                                            precio_float = float(precio)
                                            if precio_float > 0:
                                                precios_por_nit[nit] = precio_float
                                        except Exception:
                                            pass
                            if precios_por_nit:
                                tabla_busqueda[ref] = precios_por_nit
                    log.info("  Tabla de busqueda creada: %d referencias con precios", len(tabla_busqueda))

                    if "NIT CLIENTE" not in df_inf_octubre.columns:
                        raise ValueError("Falta columna NIT CLIENTE")
                    if "REFERENCIA" not in df_inf_octubre.columns:
                        raise ValueError("Falta columna REFERENCIA")

                    log.info("Calculando VTA ACORDADA vectorizado (%d filas)...", len(df_inf_octubre))
                    filas_precios = []
                    for ref, precios_por_nit in tabla_busqueda.items():
                        for nit, precio in precios_por_nit.items():
                            filas_precios.append({"_REF_LOOKUP": ref, "_NIT_LOOKUP": nit, "_PRECIO_LICITADO": precio})
                    df_precios_lookup = pd.DataFrame(filas_precios) if filas_precios else pd.DataFrame(columns=["_REF_LOOKUP", "_NIT_LOOKUP", "_PRECIO_LICITADO"])

                    df_inf_octubre["_NIT_CLEAN"] = df_inf_octubre["NIT CLIENTE"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_inf_octubre["_REF_CLEAN"] = df_inf_octubre["REFERENCIA"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    nits_validos = set(nits_headers.keys())

                    df_inf_octubre = df_inf_octubre.merge(df_precios_lookup, left_on=["_REF_CLEAN", "_NIT_CLEAN"], right_on=["_REF_LOOKUP", "_NIT_LOOKUP"], how="left")
                    mask_nit_valido = df_inf_octubre["_NIT_CLEAN"].isin(nits_validos)
                    df_inf_octubre["VTA ACORDADA X UNIDAD LICITADO"] = 0
                    df_inf_octubre.loc[mask_nit_valido, "VTA ACORDADA X UNIDAD LICITADO"] = df_inf_octubre.loc[mask_nit_valido, "_PRECIO_LICITADO"].fillna(0)
                    n_encontrados = int((df_inf_octubre["_PRECIO_LICITADO"].notna() & mask_nit_valido).sum())
                    n_no_encontrados = len(df_inf_octubre) - n_encontrados
                    df_inf_octubre.drop(columns=["_NIT_CLEAN", "_REF_CLEAN", "_REF_LOOKUP", "_NIT_LOOKUP", "_PRECIO_LICITADO"], inplace=True, errors='ignore')
                    log.info("  Encontrados: %d (%.1f%%)", n_encontrados, n_encontrados / len(df_inf_octubre) * 100)
                    log.info("  No encontrados: %d (%.1f%%)", n_no_encontrados, n_no_encontrados / len(df_inf_octubre) * 100)
                    log.info("Columna 'VTA ACORDADA X UNIDAD LICITADO' agregada")
                except Exception as e:
                    log.error("Error en calculo VTA ACORDADA: %s", e)

                # =================== PASO 8: ALINEAR COLUMNAS ===================
                log.info("ALINEANDO COLUMNAS DEL ANO COMPLETO CON PLANTILLA")
                def _norm_for_exclusion(s):
                    return _norm(s)
                FORMULA_COLS_ALL = list(dict.fromkeys(self.COLS_FORMULAS_PRE + self.COLS_FORMULAS_POST))
                formula_cols_norm = {_norm_for_exclusion(c) for c in FORMULA_COLS_ALL}
                cols_plantilla = set(COLUMNAS_ORDEN_ORIGINAL)
                cols_octubre = set(df_inf_octubre.columns)
                for col in cols_plantilla - cols_octubre:
                    col_norm = _norm_for_exclusion(col)
                    if col_norm not in formula_cols_norm:
                        df_inf_octubre[col] = np.nan
                cols_nuevas = [c for c in cols_octubre - cols_plantilla if _norm_for_exclusion(c) not in formula_cols_norm]
                cols_finales_validas = [c for c in COLUMNAS_ORDEN_ORIGINAL if c in df_inf_octubre.columns] + cols_nuevas
                for col in cols_finales_validas:
                    if col not in df_inf_octubre.columns:
                        df_inf_octubre[col] = np.nan
                df_inf_octubre = df_inf_octubre[cols_finales_validas]
                df_ventas = df_inf_octubre.copy()

                log.info("  Ano completo (procesado): %d registros", len(df_ventas))
                log.info("  Columnas finales: %d", len(cols_finales_validas))
                if 'MES NO.' in df_ventas.columns:
                    log.info("  Desglose por mes:")
                    for mes_num in sorted(df_ventas['MES NO.'].dropna().unique()):
                        n_mes = (df_ventas['MES NO.'] == mes_num).sum()
                        log.info("    Mes %d: %d registros", int(mes_num), n_mes)

                # =================== PASO 9: ORDENAR ===================
                log.info("Paso 9: Ordenando...")
                df_ventas = self.ordenar_y_fechas(df_ventas, COLS_CLAVE_VENTAS_NORM)

                # =================== PASO 10: GUARDAR ===================
                CARPETA_PRUEBAS = BASE_PATH / "Pruebas"
                CARPETA_PRUEBAS.mkdir(exist_ok=True)
                today_str = date.today().strftime("%Y-%m-%d")
                final_name = self.FN_VENTAS.replace(".xlsx", f"_{today_str}.xlsx")
                final_path = CARPETA_PRUEBAS / final_name
                if final_path.resolve() == p_ventas.resolve():
                    final_path = CARPETA_PRUEBAS / f"OUTPUT_{final_name}"
                    log.info("  Ajustando nombre para evitar sobrescribir: %s", final_path.name)
                log.info("  Carpeta destino: %s", CARPETA_PRUEBAS)
                log.info("  Archivo de salida: %s", final_path.name)

                if HAS_COM:
                    log.info("Paso 10: Escribiendo en plantilla...")
                    self.com_write_df_into_template(
                        template_xlsx_path=tmp_template, out_xlsx_path=final_path,
                        df_data=df_ventas, sheet_name=SHEET_VENTAS,
                        password_out=PASSWORD_VENTAS, data_start_row=2,
                        filldown_formula_cols=list(dict.fromkeys(self.COLS_FORMULAS_PRE + self.COLS_FORMULAS_POST)),
                        periodo_mes=(pfirst, plast)
                    )

                log.info("ACTUALIZACION COMPLETADA")
                log.info("Archivo: %s", final_path.name)
                log.info("Total: %d registros", len(df_ventas))

                self.notifier.notify_success(
                    detail=f"Archivo: {final_path.name} — {len(df_ventas)} registros",
                    attachment=self.log_file if self.log_file.exists() else None,
                )

                if 'MES NO.' in df_ventas.columns:
                    log.info("VERIFICACION FINAL (en memoria):")
                    for mes_num in sorted(df_ventas['MES NO.'].dropna().unique()):
                        n_mes = (df_ventas['MES NO.'] == mes_num).sum()
                        if n_mes > 0:
                            n_refs = df_ventas[df_ventas['MES NO.'] == mes_num]['REFERENCIA'].dropna().nunique()
                            log.info("  Mes %d: %d registros, %d referencias unicas", int(mes_num), n_mes, n_refs)

            log.info("Fin")

            # Post-process: actualizar tablas dinamicas
            if final_path.exists():
                log.info("PROCESO ADICIONAL: Actualizacion de Tabla Dinamica")
                self.actualizar_tabla_dinamica_resumen_dia(archivo_path=final_path, nombre_hoja=None)
                log.info("PROCESO COMPLETADO AL 100%%")

            return final_path

        except Exception as e:
            log.error("ERROR CRITICO: %s", e)
            error_traceback = traceback.format_exc()
            log.error(error_traceback)

            self.notifier.notify_failure(
                error=f"ERROR: {e}\n\nTraceback:\n{error_traceback}",
                attachment=self.log_file if self.log_file.exists() else None,
            )
            raise
