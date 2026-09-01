from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from datetime import date

from dotenv import load_dotenv


_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(_ENV_PATH)


def _env(key: str, default: str | None = None, required: bool = False) -> str:
    val = os.getenv(key, default)
    if required and not val:
        raise EnvironmentError(f"Variable de entorno requerida no encontrada: {key}")
    return val or ""


def _emails(key: str) -> tuple[str, ...]:
    val = os.getenv(key, "")
    return tuple(e.strip() for e in val.split(",") if e.strip())


MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

MESES_ES_NOMBRE = {v.lower(): k for k, v in MESES_ES.items()}
MESES_ES_NOMBRE["setiembre"] = 9

MESES_ES_INVERTIDO = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

MESES_REMISIONES = [
    "1. ENERO", "2. FEBRERO", "3. MARZO", "4. ABRIL", "5.MAYO", "6. JUNIO",
    "7. JULIO", "8. AGOSTO", "9. SEPTIEMBRE", "10. OCTUBRE", "11. NOVIEMBRE", "12. DICIEMBRE",
]


def get_month_folder() -> str:
    now = date.today()
    return f"{now.month:02d}. {MESES_ES[now.month]}"


def get_month_folder_remisiones() -> str:
    return MESES_REMISIONES[date.today().month - 1]


@dataclass(frozen=True)
class ErpConfig:
    usuario: str = field(default_factory=lambda: _env("FERTRAC_USER", required=True))
    clave: str = field(default_factory=lambda: _env("FERTRAC_PASS", required=True))
    url_login: str = "https://erp.fertrac.com/web/login"
    url_inventario: str = "https://erp.fertrac.com/web?#action=246&model=stock.picking.type&view_type=kanban&menu_id=174"
    url_productos: str = "https://erp.fertrac.com/web#action=278&model=product.template&view_type=kanban&menu_id=174"
    url_crm: str = "https://erp.fertrac.com/web?#action=398&model=crm.lead&view_type=list&menu_id=371"


@dataclass(frozen=True)
class SmtpConfig:
    server: str = "smtp.gmail.com"
    port: int = 587
    sender_email: str = field(default_factory=lambda: _env("SMTP_SENDER_EMAIL", required=True))
    sender_password: str = field(default_factory=lambda: _env("SMTP_PASSWORD", required=True))
    recipient_emails: tuple[str, ...] = field(default_factory=lambda: _emails("SMTP_RECIPIENT_EMAILS"))
    enabled: bool = True


@dataclass(frozen=True)
class SmtpConfigVentas(SmtpConfig):
    recipient_emails: tuple[str, ...] = field(default_factory=lambda: _emails("SMTP_RECIPIENT_EMAILS_VENTAS"))


@dataclass(frozen=True)
class SmtpConfigInvGeneral(SmtpConfig):
    recipient_emails: tuple[str, ...] = field(default_factory=lambda: _emails("SMTP_RECIPIENT_EMAILS_INV_GENERAL"))


@dataclass(frozen=True)
class ExcelConfig:
    password: str = field(default_factory=lambda: _env("EXCEL_PASSWORD", required=True))
    passwords_try: tuple[str, ...] = field(default_factory=lambda: tuple(
        p.strip() for p in _env("EXCEL_PASSWORDS_TRY", "").split(",") if p.strip()
    ))


@dataclass(frozen=True)
class PathsConfig:
    base: Path = field(default_factory=lambda: Path(_env("BASE_PATH", r"D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026")))
    remisiones: Path = field(default_factory=lambda: Path(_env("REMISIONES_BASE", r"D:\Fertrac\Usuarios\infocompras\$CARPETA COMPRAS 2026\REMISIONES")))

    @property
    def informes(self) -> Path:
        return self.base / "INFORMES"

    @property
    def inventario_general(self) -> Path:
        return self.informes / "INVENTARIO GENERAL ACTUALIZADO"

    @property
    def inventario_general_mes(self) -> Path:
        return self.inventario_general / get_month_folder()

    @property
    def ventas(self) -> Path:
        return self.informes / "VENTAS 2026"

    @property
    def ventas_mes(self) -> Path:
        return self.ventas / get_month_folder()

    @property
    def valorizados(self) -> Path:
        return self.base / "Pruebas Inv General" / "Valorizados"

    @property
    def output_inv_general(self) -> Path:
        return self.base / "Pruebas Inv General"

    @property
    def remisiones_mes(self) -> Path:
        return self.remisiones / get_month_folder_remisiones()

    def ensure_dirs(self) -> None:
        for p in [self.inventario_general_mes, self.ventas_mes, self.valorizados, self.output_inv_general]:
            p.mkdir(parents=True, exist_ok=True)


@dataclass(frozen=True)
class BrowserConfig:
    headless: bool = field(default_factory=lambda: _env("HEADLESS", "true").lower() in ("true", "1", "yes"))
    timeout_descarga: int = 60
    timeout_carga_maxima: int = 3600
    timeout_detectar_total: int = 60
    timeout_descarga_completa: int = 1500
    periodo_gracia: int = 180
    margen_snapshot: int = 5
    max_intentos_driver: int = 5
    max_intentos_login: int = 3


@dataclass(frozen=True)
class Settings:
    erp: ErpConfig = field(default_factory=ErpConfig)
    smtp: SmtpConfig = field(default_factory=SmtpConfig)
    smtp_ventas: SmtpConfigVentas = field(default_factory=SmtpConfigVentas)
    smtp_inv_general: SmtpConfigInvGeneral = field(default_factory=SmtpConfigInvGeneral)
    excel: ExcelConfig = field(default_factory=ExcelConfig)
    paths: PathsConfig = field(default_factory=PathsConfig)
    browser: BrowserConfig = field(default_factory=BrowserConfig)
