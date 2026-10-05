"""Wrapper Microsoft Graph para LEER mails y manejar carpetas/adjuntos.

Usado por el MailBot Fedrigoni (Iteración 8). Requiere scope `Mail.ReadWrite`
(delegated, pedido on-demand) y la cuenta AGENTE logueada con `--setup-auth`
en un cache propio (`token_agente.cache`).

NO usa `SharePointClient` porque ese tiene auth con cache "default"; este
levanta su propia sesión y arma headers con un token específico de AGENTE.
"""
from __future__ import annotations

import base64
import urllib.parse
from dataclasses import dataclass
from typing import Iterable

import requests

from ..sharepoint.auth import MAIL_READWRITE_SCOPE, get_token
from ..sharepoint.client import GRAPH, SharePointError
from ..utils.logger import get_logger

log = get_logger(__name__)

# Cache name del token MSAL para AGENTE.
AGENTE_CACHE = "agente"

TIMEOUT = (15, 60)


@dataclass
class MailMessage:
    id: str
    subject: str
    from_address: str
    received_at: str
    has_attachments: bool


def _token() -> str:
    """Token con scopes Files+Sites+User.Read+Mail.ReadWrite para AGENTE."""
    return get_token(
        interactive=False,
        extra_scopes=[MAIL_READWRITE_SCOPE],
        cache_name=AGENTE_CACHE,
    )


def _headers(extra: dict | None = None) -> dict[str, str]:
    h = {"Authorization": f"Bearer {_token()}"}
    if extra:
        h.update(extra)
    return h


def whoami() -> dict:
    """Devuelve la info del usuario autenticado (GET /me).

    Útil para diagnosticar con qué cuenta quedó el token cache de AGENTE.
    """
    resp = _request("GET", f"{GRAPH}/me")
    if not resp.ok:
        raise SharePointError(f"whoami falló: {resp.status_code} {resp.text}")
    return resp.json() or {}


def _request(method: str, url: str, **kwargs) -> requests.Response:
    """GET/POST/PATCH/DELETE genérico con manejo simple de retry."""
    import time
    for attempt in range(1, 4):
        try:
            resp = requests.request(
                method, url,
                headers=_headers(kwargs.pop("extra_headers", None) or {}),
                timeout=TIMEOUT,
                **kwargs,
            )
        except requests.RequestException as e:
            if attempt == 3:
                raise SharePointError(f"Sin conexión a Graph (mail): {e}") from e
            time.sleep(1.5 * attempt)
            continue

        if resp.status_code == 429 or 500 <= resp.status_code < 600:
            retry = int(resp.headers.get("Retry-After", "2"))
            log.warning("Graph mail %s → reintento en %ds (intento %d/3)",
                        resp.status_code, retry, attempt)
            time.sleep(min(retry, 10))
            continue
        return resp
    raise SharePointError("Graph mail no respondió tras 3 reintentos.")


# ===========================================================================
# Folders
# ===========================================================================
def list_mail_folders() -> list[dict]:
    """Lista las carpetas top-level del usuario logueado (no recursivo)."""
    url = f"{GRAPH}/me/mailFolders?$top=200"
    resp = _request("GET", url)
    if not resp.ok:
        raise SharePointError(f"list_mail_folders falló: {resp.status_code} {resp.text}")
    return resp.json().get("value", []) or []


def ensure_mail_folder(name: str) -> str:
    """Devuelve el id de la carpeta `name`. La crea si no existe (top-level)."""
    name_clean = (name or "").strip()
    if not name_clean:
        raise ValueError("ensure_mail_folder: name vacío.")
    for f in list_mail_folders():
        if (f.get("displayName") or "").strip().lower() == name_clean.lower():
            return f["id"]
    log.info("Creando mail folder: %s", name_clean)
    resp = _request(
        "POST",
        f"{GRAPH}/me/mailFolders",
        json={"displayName": name_clean},
    )
    if not resp.ok:
        raise SharePointError(
            f"No se pudo crear la carpeta '{name_clean}': {resp.status_code} {resp.text}"
        )
    return resp.json()["id"]


# ===========================================================================
# Messages
# ===========================================================================
def list_unread_in_inbox(senders: Iterable[str] | None = None, top: int = 200) -> list[MailMessage]:
    """Lista mensajes NO leídos en Inbox, opcionalmente filtrados por sender.

    NOTA: el filtro combinado `isRead AND (sender1 OR sender2 ...)` + orderby
    es demasiado complejo para Graph (InefficientFilter 400). Por eso pedimos
    SOLO `isRead eq false` server-side y filtramos por sender en Python.
    Como un Inbox típicamente tiene decenas de no leídos, es performante.
    """
    senders_clean = {s.strip().lower() for s in (senders or []) if s and s.strip()}

    params = {
        "$filter": "isRead eq false",
        "$top": str(top),
        # desc: los mails más NUEVOS primero. Útil para que el bot vea los
        # recientes aunque haya ruido acumulado de mails viejos.
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
    }
    qs = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    url = f"{GRAPH}/me/mailFolders/Inbox/messages?{qs}"
    resp = _request("GET", url)
    if not resp.ok:
        raise SharePointError(f"list_unread_in_inbox falló: {resp.status_code} {resp.text}")

    raw = resp.json().get("value", []) or []
    log.info("Inbox: %d mail(s) NO leído(s) total (antes de filtrar).", len(raw))

    out: list[MailMessage] = []
    discarded_senders: dict[str, int] = {}
    discarded_no_attach = 0
    for m in raw:
        addr = (m.get("from") or {}).get("emailAddress") or {}
        from_addr = str(addr.get("address") or "").strip().lower()
        has_attach = bool(m.get("hasAttachments"))

        # Filtro 1: sin attachments → descartar inmediatamente (no es Fedrigoni).
        # Ahorra una llamada Graph adicional por mail (no hace falta listar attachments
        # para confirmar que no hay PDF).
        if not has_attach:
            discarded_no_attach += 1
            continue

        # Filtro 2: sender no en las listas configuradas.
        if senders_clean and from_addr not in senders_clean:
            discarded_senders[from_addr] = discarded_senders.get(from_addr, 0) + 1
            continue

        out.append(MailMessage(
            id=m["id"],
            subject=str(m.get("subject") or ""),
            from_address=str(addr.get("address") or ""),
            received_at=str(m.get("receivedDateTime") or ""),
            has_attachments=has_attach,
        ))

    if discarded_no_attach:
        log.info("Inbox: %d descartado(s) por NO tener adjuntos.", discarded_no_attach)
    if discarded_senders:
        log.info(
            "Inbox: %d descartado(s) por sender no listado (con adjuntos): %s",
            sum(discarded_senders.values()),
            dict(sorted(discarded_senders.items(), key=lambda x: -x[1])),
        )
    return out


def get_message_attachments(message_id: str) -> list[dict]:
    """Devuelve la lista de adjuntos (metadata) de un mensaje."""
    url = (
        f"{GRAPH}/me/messages/{message_id}/attachments"
        "?$select=id,name,contentType,size"
    )
    resp = _request("GET", url)
    if not resp.ok:
        raise SharePointError(
            f"get_message_attachments falló: {resp.status_code} {resp.text}"
        )
    return resp.json().get("value", []) or []


def download_attachment(message_id: str, attachment_id: str) -> bytes:
    """Descarga el contenido binario de un adjunto.

    Para `fileAttachment`, Graph devuelve un campo `contentBytes` (base64).
    """
    url = f"{GRAPH}/me/messages/{message_id}/attachments/{attachment_id}"
    resp = _request("GET", url)
    if not resp.ok:
        raise SharePointError(
            f"download_attachment falló: {resp.status_code} {resp.text}"
        )
    data = resp.json()
    cb = data.get("contentBytes")
    if not cb:
        raise SharePointError(
            f"El adjunto no tiene contentBytes (tipo={data.get('@odata.type')})."
        )
    return base64.b64decode(cb)


def move_message(message_id: str, dest_folder_id: str) -> None:
    """Mueve el mensaje a la carpeta destino."""
    url = f"{GRAPH}/me/messages/{message_id}/move"
    resp = _request(
        "POST", url,
        json={"destinationId": dest_folder_id},
    )
    if not resp.ok:
        raise SharePointError(f"move_message falló: {resp.status_code} {resp.text}")


def mark_as_read(message_id: str) -> None:
    """Marca como leído (best-effort; si falla no es crítico)."""
    url = f"{GRAPH}/me/messages/{message_id}"
    resp = _request("PATCH", url, json={"isRead": True})
    if not resp.ok:
        log.warning("mark_as_read falló (%s): %s", resp.status_code, resp.text)
