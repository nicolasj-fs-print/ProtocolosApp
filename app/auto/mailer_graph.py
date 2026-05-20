"""Envío de mails vía Microsoft Graph Mail API (sin Outlook desktop).

Usado por el modo automático. Manda desde la mailbox del usuario autenticado
(el bot loguea con su cuenta una vez via `--setup-auth`).
"""
from __future__ import annotations

import base64
from typing import Sequence

import requests

from ..sharepoint.auth import MAIL_SEND_SCOPE, get_token
from ..sharepoint.client import GRAPH, SharePointError
from ..utils.logger import get_logger

log = get_logger(__name__)


def _attachment(name: str, content: bytes) -> dict:
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": "application/pdf",
        "contentBytes": base64.b64encode(content).decode("ascii"),
    }


def _recipients(addresses: Sequence[str]) -> list[dict]:
    out = []
    for a in addresses:
        if not a:
            continue
        for piece in str(a).split(";"):
            piece = piece.strip()
            if piece:
                out.append({"emailAddress": {"address": piece}})
    return out


def send_mail(
    to: Sequence[str],
    subject: str,
    html_body: str,
    attachments: Sequence[tuple[str, bytes]] = (),
    cc: Sequence[str] = (),
    save_to_sent: bool = True,
) -> None:
    """Envía un mail HTML con adjuntos opcionales vía Graph.

    `to` y `cc` aceptan strings con varios mails separados por `;`.
    `attachments` es lista de (filename, bytes) — para PDFs < 3MB.
    Lanza `SharePointError` si la API falla.
    """
    to_list = _recipients(to)
    if not to_list:
        raise ValueError("send_mail: lista de destinatarios vacía.")

    message: dict = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": to_list,
    }
    cc_list = _recipients(cc)
    if cc_list:
        message["ccRecipients"] = cc_list

    if attachments:
        message["attachments"] = [_attachment(name, data) for name, data in attachments]

    body = {"message": message, "saveToSentItems": save_to_sent}

    # `Mail.Send` se pide on-demand (requiere admin consent — no está en el
    # SCOPES default para no romper el login de la GUI manual).
    token = get_token(interactive=False, extra_scopes=[MAIL_SEND_SCOPE])
    url = f"{GRAPH}/me/sendMail"
    resp = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=body,
        timeout=(15, 60),
    )
    if not resp.ok:
        raise SharePointError(
            f"sendMail falló: {resp.status_code} {resp.text}"
        )
    log.info("Mail enviado: '%s' → %s", subject, [r["emailAddress"]["address"] for r in to_list])
