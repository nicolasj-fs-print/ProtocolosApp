from __future__ import annotations

import logging
import os
import queue
from datetime import datetime
from logging.handlers import QueueHandler
from pathlib import Path


_RUN_ID: str | None = None


def new_run_id() -> str:
    global _RUN_ID
    _RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")
    return _RUN_ID


def current_run_id() -> str:
    return _RUN_ID or new_run_id()


def logs_dir() -> Path:
    """Carpeta local de logs en %APPDATA%\\ProtocolosApp\\logs."""
    base = os.getenv("APPDATA") or str(Path.home())
    p = Path(base) / "ProtocolosApp" / "logs"
    p.mkdir(parents=True, exist_ok=True)
    return p


def log_file_path() -> Path:
    return logs_dir() / f"run_{current_run_id()}.log"


def setup_logger(
    level: str = "INFO",
    gui_queue: "queue.Queue[str] | None" = None,
) -> logging.Logger:
    """Configura el logger root con archivo + consola + (opcional) cola para GUI."""
    log_file = log_file_path()

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(getattr(logging, level.upper(), logging.INFO))

    for h in list(root.handlers):
        root.removeHandler(h)

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    root.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root.addHandler(sh)

    if gui_queue is not None:
        qh = QueueHandler(gui_queue)
        qh.setFormatter(fmt)
        root.addHandler(qh)

    root.info("Log file: %s", log_file)
    return root


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)


def upload_run_log_to_sharepoint() -> str | None:
    """Sube el log de la corrida actual a la carpeta SP_LOGS_FOLDER. Best-effort.

    Devuelve el path SharePoint si OK, None si falló.
    """
    log = get_logger(__name__)
    try:
        from ..config import get_settings
        s = get_settings()
        if s.storage_backend != "sharepoint" or not s.sp_logs_folder:
            return None
        log_path = log_file_path()
        if not log_path.exists():
            return None

        # Flush handlers para asegurar contenido completo
        for h in logging.getLogger().handlers:
            try:
                h.flush()
            except Exception:
                pass

        from ..sharepoint.client import get_app_client
        sp_path = f"{s.sp_logs_folder.rstrip('/')}/run_{current_run_id()}.log"
        get_app_client().upload_file(sp_path, log_path.read_bytes(), overwrite=True)
        log.info("Log subido a SharePoint: %s", sp_path)
        return sp_path
    except Exception as e:
        log.warning("No se pudo subir log a SharePoint: %s", e)
        return None
