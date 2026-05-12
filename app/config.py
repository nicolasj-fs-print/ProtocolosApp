from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

from .utils.paths import project_root, runtime_dir, resource_path


def _env_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass
class Settings:
    sql_driver: str
    sql_server: str
    sql_database: str
    sql_database_trazabilidad: str
    sql_trusted_connection: bool
    sql_user: str
    sql_password: str
    sql_trust_server_cert: str
    sql_encrypt: str

    protocols_folder: Path
    default_output_folder: Path
    clients_xlsx: Path
    clients_sheet: str
    ingresos_xlsx: Path
    ingresos_sheet: str

    sumatra_path: Path

    log_level: str
    mock_data: bool

    # SharePoint / Microsoft Graph
    storage_backend: str
    ms_tenant_id: str
    ms_client_id: str
    sharepoint_hostname: str
    sharepoint_site: str
    sp_app_hostname: str
    sp_app_site: str
    sp_clients_file: str
    sp_ingresos_file: str
    sp_protocols_folder: str
    sp_output_folder: str
    sp_logs_folder: str
    sp_ejecuciones_folder: str
    sp_usuarios_file: str

    base_dir: Path = field(default_factory=project_root)
    runtime_dir: Path = field(default_factory=runtime_dir)

    @classmethod
    def load(cls) -> "Settings":
        env_file = runtime_dir() / ".env"
        if not env_file.exists():
            env_file = project_root() / ".env"
        if env_file.exists():
            load_dotenv(env_file)
        else:
            load_dotenv()

        sumatra_raw = os.getenv("SUMATRA_PATH", r"tools\SumatraPDF.exe")
        sumatra_path = Path(sumatra_raw)
        if not sumatra_path.is_absolute():
            sumatra_path = resource_path(sumatra_raw)

        return cls(
            sql_driver=os.getenv("SQL_DRIVER", "ODBC Driver 18 for SQL Server"),
            sql_server=os.getenv("SQL_SERVER", ""),
            sql_database=os.getenv("SQL_DATABASE", ""),
            sql_database_trazabilidad=os.getenv("SQL_DATABASE_TRAZABILIDAD", "FSBI"),
            sql_trusted_connection=_env_bool(os.getenv("SQL_TRUSTED_CONNECTION"), False),
            sql_user=os.getenv("SQL_USER", ""),
            sql_password=os.getenv("SQL_PASSWORD", ""),
            sql_trust_server_cert=os.getenv("SQL_TRUST_SERVER_CERT", "yes"),
            sql_encrypt=os.getenv("SQL_ENCRYPT", "yes"),
            protocols_folder=Path(os.getenv("PROTOCOLS_FOLDER", "")),
            default_output_folder=Path(os.getenv("DEFAULT_OUTPUT_FOLDER", "")),
            clients_xlsx=Path(os.getenv("CLIENTS_XLSX", "")),
            clients_sheet=os.getenv("CLIENTS_SHEET", "Clientes"),
            ingresos_xlsx=Path(os.getenv("INGRESOS_XLSX", "")),
            ingresos_sheet=os.getenv("INGRESOS_SHEET", "") or "",
            sumatra_path=sumatra_path,
            log_level=os.getenv("LOG_LEVEL", "INFO"),
            mock_data=_env_bool(os.getenv("MOCK_DATA"), False),
            storage_backend=(os.getenv("STORAGE_BACKEND", "sharepoint") or "sharepoint").strip().lower(),
            ms_tenant_id=os.getenv("MS_TENANT_ID", ""),
            ms_client_id=os.getenv("MS_CLIENT_ID", ""),
            sharepoint_hostname=os.getenv("SHAREPOINT_HOSTNAME", ""),
            sharepoint_site=os.getenv("SHAREPOINT_SITE", ""),
            sp_app_hostname=os.getenv("SP_APP_HOSTNAME", "") or os.getenv("SHAREPOINT_HOSTNAME", ""),
            sp_app_site=os.getenv("SP_APP_SITE", ""),
            sp_clients_file=os.getenv("SP_CLIENTS_FILE", ""),
            sp_ingresos_file=os.getenv("SP_INGRESOS_FILE", ""),
            sp_protocols_folder=os.getenv("SP_PROTOCOLS_FOLDER", ""),
            sp_output_folder=os.getenv("SP_OUTPUT_FOLDER", ""),
            sp_logs_folder=os.getenv("SP_LOGS_FOLDER", ""),
            sp_ejecuciones_folder=os.getenv("SP_EJECUCIONES_FOLDER", ""),
            sp_usuarios_file=os.getenv("SP_USUARIOS_FILE", ""),
        )

    def odbc_connection_string(self, database: str | None = None) -> str:
        db = database or self.sql_database
        parts = [
            f"DRIVER={{{self.sql_driver}}}",
            f"SERVER={self.sql_server}",
            f"DATABASE={db}",
            f"Encrypt={self.sql_encrypt}",
            f"TrustServerCertificate={self.sql_trust_server_cert}",
        ]
        if self.sql_trusted_connection:
            parts.append("Trusted_Connection=yes")
        else:
            parts.append(f"UID={self.sql_user}")
            parts.append(f"PWD={self.sql_password}")
        return ";".join(parts) + ";"


_settings: Settings | None = None


def get_settings(force_reload: bool = False) -> Settings:
    global _settings
    if _settings is None or force_reload:
        _settings = Settings.load()
    return _settings
