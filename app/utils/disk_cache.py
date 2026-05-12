"""Cache en disco para archivos SharePoint.

Guarda los bytes en `%APPDATA%\\ProtocolosApp\\cache\\` junto a un .meta.json
con `lastModifiedDateTime` y `eTag` del item remoto. Antes de devolver el
archivo, se chequea si SharePoint cambió el `lastModifiedDateTime`:
- Si NO cambió → leer del disco (instantáneo).
- Si cambió → re-descargar y reemplazar el cache.
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

from .logger import get_logger

log = get_logger(__name__)


def cache_dir() -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    p = Path(base) / "ProtocolosApp" / "cache"
    p.mkdir(parents=True, exist_ok=True)
    return p


_INVALID = re.compile(r'[\\/:*?"<>|]+')


def _safe_name(sp_path: str) -> str:
    """Convierte un path SP a un filename seguro."""
    return _INVALID.sub("_", (sp_path or "").strip("/"))


def cache_paths(sp_path: str) -> tuple[Path, Path]:
    """Devuelve (data_path, meta_path) para un sp_path."""
    safe = _safe_name(sp_path)
    base = cache_dir() / safe
    meta = base.with_name(base.name + ".meta.json")
    return base, meta


def read_cached(sp_path: str, remote_modified: str) -> bytes | None:
    """Devuelve bytes si el cache existe y `lastModifiedDateTime` coincide
    con el remoto. Sino None (caller debe re-descargar)."""
    if not remote_modified:
        return None
    data_path, meta_path = cache_paths(sp_path)
    if not data_path.exists() or not meta_path.exists():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception as e:
        log.warning("disk_cache: meta corrupto para %s: %s", sp_path, e)
        return None
    if meta.get("lastModifiedDateTime") != remote_modified:
        log.info(
            "disk_cache: %s cambió (local=%s, remoto=%s) → re-descargar",
            sp_path, meta.get("lastModifiedDateTime"), remote_modified,
        )
        return None
    try:
        return data_path.read_bytes()
    except OSError as e:
        log.warning("disk_cache: read falló (%s): %s", sp_path, e)
        return None


def write_cached(
    sp_path: str,
    data: bytes,
    remote_modified: str,
    etag: str = "",
) -> None:
    data_path, meta_path = cache_paths(sp_path)
    try:
        data_path.write_bytes(data)
        meta_path.write_text(
            json.dumps({
                "sp_path": sp_path,
                "lastModifiedDateTime": remote_modified,
                "eTag": etag,
                "size": len(data),
            }),
            encoding="utf-8",
        )
        log.info("disk_cache: guardado %s (%.1f MB)", sp_path, len(data) / (1024 * 1024))
    except OSError as e:
        log.warning("disk_cache: write falló (%s): %s", sp_path, e)


def clear() -> None:
    """Borra todo el cache."""
    d = cache_dir()
    for p in d.iterdir():
        try:
            p.unlink()
        except OSError:
            pass
    log.info("disk_cache: limpio (%s)", d)
