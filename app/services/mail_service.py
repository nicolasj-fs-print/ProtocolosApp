"""Borrador de mail en Outlook (no envía: solo abre la ventana)."""
from __future__ import annotations

import os
import urllib.parse
from pathlib import Path
from typing import Sequence

from ..utils.logger import get_logger

log = get_logger(__name__)


def _build_html_body(
    razon_social: str,
    comprobante: str,
    encontrados: Sequence[str],
    no_encontrados: Sequence[str],
) -> str:
    encontrados_html = "".join(f"<li>{p}</li>" for p in encontrados) or "<li>(sin protocolos)</li>"
    nf_block = ""
    if no_encontrados:
        items = "".join(f"<li>{p}</li>" for p in no_encontrados)
        nf_block = (
            "<p><b>Protocolos no encontrados en archivo (revisar):</b></p>"
            f"<ul>{items}</ul>"
        )

    return (
        "<p>Estimados,</p>"
        f"<p>Se adjuntan los protocolos correspondientes al remito <b>{comprobante}</b> "
        f"del cliente <b>{razon_social}</b>.</p>"
        "<p><b>Detalle de protocolos:</b></p>"
        f"<ul>{encontrados_html}</ul>"
        f"{nf_block}"
        "<p>Quedamos a disposición ante cualquier consulta.</p>"
        "<p>Saludos cordiales.</p>"
    )


def open_outlook_draft(
    to: str,
    razon_social: str,
    comprobante: str,
    pdf_path: Path,
    protocolos_encontrados: Sequence[str],
    protocolos_no_encontrados: Sequence[str] = (),
) -> bool:
    """Abre Outlook con un borrador listo para revisar y enviar manualmente.

    Devuelve True si se abrió Outlook, False si cayó al fallback mailto:.
    """
    subject = f"Protocolos de calidad - {razon_social} | Remito {comprobante}"
    html_body = _build_html_body(razon_social, comprobante, protocolos_encontrados, protocolos_no_encontrados)

    try:
        import win32com.client  # type: ignore
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to or ""
        mail.Subject = subject
        mail.HTMLBody = html_body
        if pdf_path and Path(pdf_path).exists():
            mail.Attachments.Add(str(Path(pdf_path).resolve()))
        mail.Display(False)
        log.info("Borrador Outlook abierto para %s", comprobante)
        return True
    except Exception as e:
        log.warning("Outlook no disponible (%s) → fallback mailto:", e)

    plain_body = (
        f"Estimados,\n\n"
        f"Se adjuntan los protocolos del remito {comprobante} - {razon_social}.\n"
        f"PDF: {pdf_path}\n\n"
        f"Detalle:\n" + "\n".join(f"- {p}" for p in protocolos_encontrados)
    )
    params = urllib.parse.urlencode({
        "subject": subject,
        "body": plain_body,
    }, quote_via=urllib.parse.quote)
    url = f"mailto:{to}?{params}"
    try:
        os.startfile(url)
    except OSError as e:
        log.error("No se pudo abrir mailto: %s", e)
    return False
