"""Búsqueda de PDFs originales de protocolos (local o SharePoint)."""
from __future__ import annotations

import threading
from pathlib import Path

from ..config import get_settings
from ..utils.logger import get_logger

log = get_logger(__name__)

_PROTOCOL_KEY_LEN = 6  # PR0001 - ...

# Cache global del index. Se construye una vez por sesión y se reusa entre
# corridas del worker. Evita re-listar la carpeta SharePoint (que se colgaba
# cuando el thread daemon de telemetría usaba el cliente al mismo tiempo).
_GLOBAL_INDEX: dict[str, object] | None = None
_INDEX_LOCK = threading.Lock()


def reset_protocol_index() -> None:
    """Forza una recarga del index en la próxima llamada a `index()`."""
    global _GLOBAL_INDEX
    with _INDEX_LOCK:
        _GLOBAL_INDEX = None


class ProtocolFinder:
    """Indexa una sola vez la carpeta de protocolos para acelerar búsquedas.

    Soporta backend `local` (filesystem) y `sharepoint` (Graph API list folder).
    """

    def __init__(self, folder: Path | None = None):
        s = get_settings()
        self.backend = s.storage_backend
        self.folder_local = folder if folder is not None else s.protocols_folder
        self.folder_sp = s.sp_protocols_folder
        # Local: dict[key, Path]; SharePoint: dict[key, item_dict]
        self._index: dict[str, object] | None = None

    # -------------------- Index --------------------
    def _build_index_local(self) -> dict[str, Path]:
        if not self.folder_local.exists():
            log.error("Carpeta de protocolos no existe: %s", self.folder_local)
            return {}
        idx: dict[str, Path] = {}
        for f in self.folder_local.iterdir():
            if not f.is_file() or f.suffix.lower() != ".pdf":
                continue
            key = f.name[:_PROTOCOL_KEY_LEN].upper()
            idx.setdefault(key, f)
        log.info("Index protocolos (local): %d archivos", len(idx))
        return idx

    def _build_index_sharepoint(self) -> dict[str, dict]:
        from ..sharepoint.client import get_client
        client = get_client()
        try:
            items = client.list_folder(self.folder_sp)
        except Exception as e:
            log.error("No se pudo listar carpeta SP %s: %s", self.folder_sp, e)
            return {}
        idx: dict[str, dict] = {}
        for it in items:
            if "file" not in it:
                continue
            name = it.get("name", "")
            if not name.lower().endswith(".pdf"):
                continue
            key = name[:_PROTOCOL_KEY_LEN].upper()
            idx.setdefault(key, it)
        log.info("Index protocolos (SharePoint): %d archivos", len(idx))
        return idx

    def index(self) -> dict[str, object]:
        global _GLOBAL_INDEX
        # Cache global: la lista de protocolos no cambia durante la sesión,
        # así que se construye una sola vez y se reutiliza entre runs.
        with _INDEX_LOCK:
            if _GLOBAL_INDEX is not None:
                self._index = _GLOBAL_INDEX
                return _GLOBAL_INDEX

            if self.backend == "sharepoint":
                _GLOBAL_INDEX = self._build_index_sharepoint()
            else:
                _GLOBAL_INDEX = self._build_index_local()
            self._index = _GLOBAL_INDEX
            return _GLOBAL_INDEX

    # -------------------- Lookup --------------------
    def find(self, protocol_id: str | None) -> object | None:
        if not protocol_id:
            return None
        key = str(protocol_id).strip().upper()[:_PROTOCOL_KEY_LEN]
        if not key:
            return None
        return self.index().get(key)

    def read_bytes(self, protocol_id: str) -> bytes | None:
        item = self.find(protocol_id)
        if item is None:
            return None

        if self.backend == "sharepoint":
            from ..sharepoint.client import get_client
            client = get_client()
            # 1) Si trae downloadUrl pre-firmada, usarla (no necesita auth).
            url = item.get("@microsoft.graph.downloadUrl") if isinstance(item, dict) else None
            if url:
                try:
                    import requests
                    r = requests.get(url, timeout=60)
                    if r.ok:
                        return r.content
                    log.warning("downloadUrl falló (%s) → fallback path", r.status_code)
                except Exception as e:
                    log.warning("downloadUrl error: %s → fallback path", e)
            # 2) Fallback: descarga por path
            name = item.get("name") if isinstance(item, dict) else None
            if not name:
                return None
            path = f"{self.folder_sp.rstrip('/')}/{name}"
            try:
                return client.download_file(path)
            except Exception as e:
                log.warning("No se pudo descargar %s: %s", path, e)
                return None

        # Backend local
        if isinstance(item, Path):
            try:
                return item.read_bytes()
            except OSError as e:
                log.warning("No se pudo leer %s: %s", item, e)
        return None
