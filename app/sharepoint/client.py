"""Wrapper Microsoft Graph para operaciones SharePoint."""
from __future__ import annotations

import threading
import time
import urllib.parse
from typing import Iterable

import requests

from ..config import get_settings
from ..utils.logger import get_logger
from .auth import get_token

log = get_logger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"
TIMEOUT = (15, 120)  # (connect, read) — segundo elemento aplica al body
SMALL_UPLOAD_LIMIT = 4 * 1024 * 1024  # 4 MB


class SharePointError(RuntimeError):
    pass


def _quote(path: str) -> str:
    """URL-encode un path estilo `Carpeta/Archivo.xlsx` para Graph."""
    path = (path or "").strip().strip("/").replace("\\", "/")
    return urllib.parse.quote(path, safe="/")


class SharePointClient:
    """Cliente liviano sobre Graph REST. Reutiliza session HTTP y cachea site/drive id.

    Si no se pasa hostname/site, usa los defaults de Settings (sitio principal de datos).
    """

    def __init__(self, hostname: str | None = None, site: str | None = None):
        s = get_settings()
        host = hostname or s.sharepoint_hostname
        site_name = site or s.sharepoint_site
        if not host or not site_name:
            raise SharePointError(
                "Faltan SHAREPOINT_HOSTNAME / SHAREPOINT_SITE en .env "
                "(o no se especificó hostname/site al instanciar)."
            )
        self.hostname = host
        self.site_name = site_name
        self._site_id: str | None = None
        self._drive_id: str | None = None
        self._lock = threading.RLock()  # reentrant: drive_id() llama site_id() adentro del lock
        # Una Session por thread → evita pelea por el connection pool
        # cuando varios threads usan el mismo cliente (ej. register_login + Buscar).
        self._sessions = threading.local()

    # ---------- Auth headers ----------
    def _auth_headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {get_token(interactive=True)}"}

    def _session(self) -> requests.Session:
        """Devuelve la Session del thread actual (la crea si no existe)."""
        s = getattr(self._sessions, "session", None)
        if s is None:
            s = requests.Session()
            self._sessions.session = s
        return s

    def _request(
        self,
        method: str,
        url: str,
        *,
        params: dict | None = None,
        data: bytes | None = None,
        json_body: dict | None = None,
        extra_headers: dict | None = None,
        stream: bool = False,
    ) -> requests.Response:
        headers = self._auth_headers()
        if extra_headers:
            headers.update(extra_headers)
        for attempt in range(1, 4):
            try:
                resp = self._session().request(
                    method,
                    url,
                    headers=headers,
                    params=params,
                    data=data,
                    json=json_body,
                    timeout=TIMEOUT,
                    stream=stream,
                )
            except requests.RequestException as e:
                if attempt == 3:
                    raise SharePointError(f"Sin conexión a Microsoft Graph: {e}") from e
                time.sleep(1.5 * attempt)
                continue

            if resp.status_code == 429 or 500 <= resp.status_code < 600:
                retry = int(resp.headers.get("Retry-After", "2"))
                log.warning("Graph %s → reintento en %ds (intento %d/3)", resp.status_code, retry, attempt)
                time.sleep(min(retry, 10))
                continue

            return resp
        raise SharePointError("Graph no respondió tras 3 reintentos.")

    # ---------- Site / Drive ----------
    def site_id(self) -> str:
        # Double-check: si ya está cacheado, evitamos tomar el lock
        # para no contender con otros threads (telemetría, worker, etc.).
        if self._site_id:
            return self._site_id
        with self._lock:
            if self._site_id:
                return self._site_id
            url = f"{GRAPH}/sites/{self.hostname}:/sites/{urllib.parse.quote(self.site_name)}"
            resp = self._request("GET", url)
            if not resp.ok:
                raise SharePointError(
                    f"No se pudo resolver el sitio '{self.site_name}': {resp.status_code} {resp.text}"
                )
            self._site_id = resp.json()["id"]
            log.info("Site resuelto: %s", self._site_id)
            return self._site_id

    def drive_id(self) -> str:
        if self._drive_id:
            return self._drive_id
        with self._lock:
            if self._drive_id:
                return self._drive_id
            url = f"{GRAPH}/sites/{self.site_id()}/drive"
            resp = self._request("GET", url)
            if not resp.ok:
                raise SharePointError(
                    f"No se pudo resolver el drive default: {resp.status_code} {resp.text}"
                )
            self._drive_id = resp.json()["id"]
            log.info("Drive resuelto: %s", self._drive_id)
            return self._drive_id

    # ---------- Helpers de URL ----------
    def _item_url_by_path(self, path: str) -> str:
        return f"{GRAPH}/sites/{self.site_id()}/drive/root:/{_quote(path)}"

    # ---------- Files ----------
    def file_exists(self, path: str) -> bool:
        url = self._item_url_by_path(path)
        resp = self._request("GET", url)
        if resp.status_code == 200:
            return True
        if resp.status_code == 404:
            return False
        raise SharePointError(f"file_exists falló ({resp.status_code}): {resp.text}")

    def get_item(self, path: str) -> dict | None:
        url = self._item_url_by_path(path)
        resp = self._request("GET", url)
        if resp.status_code == 404:
            return None
        if not resp.ok:
            raise SharePointError(f"get_item ({path}) falló: {resp.status_code} {resp.text}")
        return resp.json()

    def download_file(self, path: str) -> bytes:
        import time as _time
        url = self._item_url_by_path(path) + ":/content"
        log.info("Descarga iniciada: %s", path)
        t0 = _time.time()
        # IMPORTANT: stream=False (default) → el body se baja respetando el timeout.
        # Con stream=True + .content había hangs sin respetar timeout.
        resp = self._request("GET", url)
        if not resp.ok:
            raise SharePointError(f"Descarga falló ({path}): {resp.status_code} {resp.text}")
        elapsed = _time.time() - t0
        size_mb = len(resp.content) / (1024 * 1024)
        log.info("Descarga OK: %s (%.1f MB en %.1fs)", path, size_mb, elapsed)
        return resp.content

    def list_folder(self, path: str) -> list[dict]:
        """Lista todos los items en una carpeta (sigue paginación)."""
        if path:
            base = self._item_url_by_path(path) + ":/children"
        else:
            base = f"{GRAPH}/sites/{self.site_id()}/drive/root/children"
        items: list[dict] = []
        next_url: str | None = base
        params: dict | None = {"$top": "200", "$select": "id,name,size,file,folder,@microsoft.graph.downloadUrl"}
        while next_url:
            resp = self._request("GET", next_url, params=params)
            if not resp.ok:
                raise SharePointError(f"list_folder ({path}) falló: {resp.status_code} {resp.text}")
            data = resp.json()
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            params = None  # nextLink ya tiene query string
        return items

    def upload_file(self, path: str, content: bytes, *, overwrite: bool = True) -> dict:
        """Sube `content` a `path`. Usa upload session si > 4MB.
        Crea las carpetas intermedias si no existen.
        """
        # Asegurar que la carpeta padre exista — Graph NO crea carpetas
        # intermedias automáticamente al hacer PUT por path.
        parent = "/".join((path or "").strip("/").split("/")[:-1])
        if parent:
            try:
                self.ensure_folder(parent)
            except Exception as e:
                log.warning("ensure_folder(%s) falló: %s", parent, e)

        if len(content) <= SMALL_UPLOAD_LIMIT:
            return self._upload_small(path, content, overwrite=overwrite)
        return self._upload_large(path, content, overwrite=overwrite)

    def ensure_folder(self, folder_path: str) -> None:
        """Crea recursivamente las carpetas que falten en `folder_path`.

        Idempotente: si ya existen no hace nada. Usa el endpoint
        POST /drive/items/root:/path:/children con `folder` body.
        """
        folder_path = (folder_path or "").strip("/").replace("\\", "/")
        if not folder_path:
            return
        # Si ya existe, salir.
        try:
            url = self._item_url_by_path(folder_path)
            resp = self._request("GET", url)
            if resp.status_code == 200:
                return
        except Exception:
            pass

        log.info("Creando carpetas faltantes en: %s", folder_path)
        site = self.site_id()
        parts = folder_path.split("/")
        accumulated: list[str] = []
        for part in parts:
            if not part:
                continue
            parent_path = "/".join(accumulated)
            accumulated.append(part)
            current = "/".join(accumulated)

            # Verificar si esta sub-carpeta ya existe
            check_url = self._item_url_by_path(current)
            r = self._request("GET", check_url)
            if r.status_code == 200:
                log.info("  - existe: %s", current)
                continue

            # No existe → crearla bajo el parent
            if parent_path:
                create_url = (
                    f"{GRAPH}/sites/{site}/drive/root:/{_quote(parent_path)}:/children"
                )
            else:
                create_url = f"{GRAPH}/sites/{site}/drive/root/children"

            body = {
                "name": part,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            }
            cr = self._request("POST", create_url, json_body=body)
            if cr.ok:
                log.info("  + creada: %s", current)
            elif cr.status_code == 409:
                # Ya existía (race condition)
                log.info("  - existe (409): %s", current)
            else:
                raise SharePointError(
                    f"No se pudo crear carpeta '{current}': "
                    f"{cr.status_code} {cr.text}"
                )

    def _upload_small(self, path: str, content: bytes, *, overwrite: bool) -> dict:
        url = self._item_url_by_path(path) + ":/content"
        params = {"@microsoft.graph.conflictBehavior": "replace" if overwrite else "fail"}
        resp = self._request(
            "PUT", url, params=params, data=content,
            extra_headers={"Content-Type": "application/octet-stream"},
        )
        if not resp.ok:
            raise SharePointError(f"upload_small ({path}) falló: {resp.status_code} {resp.text}")
        return resp.json()

    def _upload_large(self, path: str, content: bytes, *, overwrite: bool) -> dict:
        url = self._item_url_by_path(path) + ":/createUploadSession"
        body = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace" if overwrite else "fail",
            }
        }
        resp = self._request("POST", url, json_body=body)
        if not resp.ok:
            raise SharePointError(f"createUploadSession ({path}) falló: {resp.status_code} {resp.text}")
        upload_url = resp.json()["uploadUrl"]

        chunk = 5 * 1024 * 1024  # 5 MB chunks
        total = len(content)
        for start in range(0, total, chunk):
            end = min(start + chunk, total) - 1
            part = content[start:end + 1]
            headers = {
                "Content-Length": str(len(part)),
                "Content-Range": f"bytes {start}-{end}/{total}",
            }
            r = self._session().put(upload_url, data=part, headers=headers, timeout=TIMEOUT)
            if r.status_code not in (200, 201, 202):
                raise SharePointError(f"chunk upload falló: {r.status_code} {r.text}")
            if r.status_code in (200, 201):
                return r.json()
        raise SharePointError("upload_large terminó sin respuesta final.")

    def web_url(self, path: str) -> str:
        """URL amigable para abrir el archivo en el navegador."""
        try:
            item = self.get_item(path)
            if item and "webUrl" in item:
                return item["webUrl"]
        except Exception as e:
            log.warning("web_url: get_item falló (%s) → fallback", e)
        return self._build_browse_url(path)

    def folder_web_url(self, path: str) -> str:
        try:
            item = self.get_item(path)
            if item and "webUrl" in item:
                return item["webUrl"]
        except Exception as e:
            log.warning("folder_web_url: get_item falló (%s) → fallback", e)
        return self._build_browse_url(path)

    def _build_browse_url(self, path: str) -> str:
        """URL de SharePoint estilo `Forms/AllItems.aspx?id=...` que abre la
        vista de carpeta correctamente, aunque get_item haya fallado.
        """
        clean = (path or "").strip("/").replace("\\", "/")
        full_id = f"/sites/{self.site_name}/Documentos compartidos/{clean}"
        # urlencode estilo SharePoint: espacios = %20, slashes = %2F en el id.
        encoded_id = urllib.parse.quote(full_id, safe="")
        return (
            f"https://{self.hostname}/sites/{self.site_name}"
            f"/Documentos%20compartidos/Forms/AllItems.aspx?id={encoded_id}"
        )


# ---------- Singletons ----------
_CLIENT: SharePointClient | None = None
_APP_CLIENT: SharePointClient | None = None
_CLIENT_LOCK = threading.Lock()


def get_client() -> SharePointClient:
    """Cliente del sitio principal de datos (VentasPowerBI por default)."""
    global _CLIENT
    with _CLIENT_LOCK:
        if _CLIENT is None:
            _CLIENT = SharePointClient()
        return _CLIENT


def get_app_client() -> SharePointClient:
    """Cliente DEDICADO para telemetría (Logs / Ejecuciones / Usuarios).

    SIEMPRE es una instancia separada de get_client() para que un cuelgue
    en el thread daemon de telemetría NO afecte al cliente del worker.
    Si SP_APP_SITE no está configurado, apunta al mismo sitio principal.
    """
    global _APP_CLIENT
    with _CLIENT_LOCK:
        if _APP_CLIENT is None:
            s = get_settings()
            host = s.sp_app_hostname or s.sharepoint_hostname
            site = s.sp_app_site or s.sharepoint_site
            _APP_CLIENT = SharePointClient(hostname=host, site=site)
        return _APP_CLIENT


def reset_client() -> None:
    global _CLIENT, _APP_CLIENT
    with _CLIENT_LOCK:
        _CLIENT = None
        _APP_CLIENT = None
