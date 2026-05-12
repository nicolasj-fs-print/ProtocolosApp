from __future__ import annotations

import re
import sys
from pathlib import Path


def project_root() -> Path:
    """Raíz del proyecto en dev, raíz del bundle en .exe."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[2]


def runtime_dir() -> Path:
    """Carpeta donde corre la app (junto al .exe en bundle, raíz del proyecto en dev).
    Útil para logs/output que NO van dentro del bundle."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def resource_path(rel: str | Path) -> Path:
    """Resuelve un recurso embebido (assets, tools)."""
    return project_root() / rel


_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def sanitize_filename(name: str, max_len: int = 120) -> str:
    if name is None:
        return "_"
    s = str(name).strip()
    s = _INVALID_FILENAME_CHARS.sub("_", s)
    s = s.strip(". ")
    if not s:
        s = "_"
    return s[:max_len]
