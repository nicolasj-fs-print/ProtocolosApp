"""Overlay de header (logo + 'Protocolo: XXXX') sobre cada página de un PDF original."""
from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF

from ..utils.paths import resource_path

_HEADER_HEIGHT = 42      # pt — alto de la barra (NO se mueve la línea inferior)
_PAD = 10
_LOGO_W = 70
_LOGO_PAD_Y = 4
_TEXT_FONTSIZE = 13
_TEXT_OFFSET_FROM_BOTTOM = 6   # pt — distancia desde la línea inferior hasta el baseline del texto.
                               # MENOR = texto MÁS ABAJO (más cerca de la línea).
                               # MAYOR = texto más arriba.


def add_header_to_pdf(pdf_bytes: bytes, protocolo_id: str) -> bytes:
    """Devuelve un PDF nuevo donde cada página tiene un header arriba (logo FS +
    texto 'Protocolo: XXXX') y el contenido original escalado hacia abajo para
    no pisarlo. Reescribe el PDF en lugar de hacer overlay.
    """
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    logo_fs = resource_path("assets/logo_fs.png")

    try:
        for src_page in src:
            rect = src_page.rect
            new_page = out.new_page(width=rect.width, height=rect.height)

            content_rect = fitz.Rect(
                rect.x0,
                rect.y0 + _HEADER_HEIGHT,
                rect.x1,
                rect.y1,
            )
            new_page.show_pdf_page(content_rect, src, src_page.number)

            # Línea divisoria
            new_page.draw_line(
                (rect.x0, rect.y0 + _HEADER_HEIGHT),
                (rect.x1, rect.y0 + _HEADER_HEIGHT),
                color=(0.12, 0.22, 0.39),
                width=0.8,
                overlay=True,
            )

            # Texto "Protocolo: XXXX" — bajado para dejar margen arriba a la paginación.
            try:
                # baseline del texto pegado a la línea inferior del header.
                text_y = rect.y0 + _HEADER_HEIGHT - _TEXT_OFFSET_FROM_BOTTOM
                new_page.insert_text(
                    (rect.x0 + _PAD, text_y),
                    f"Protocolo: {protocolo_id}",
                    fontsize=_TEXT_FONTSIZE,
                    fontname="helv",
                    color=(0.12, 0.22, 0.39),
                    overlay=True,
                )
            except Exception:
                pass

            # Logo FS a la derecha
            if logo_fs.exists():
                logo_h = _HEADER_HEIGHT - 2 * _LOGO_PAD_Y
                r = fitz.Rect(
                    rect.x1 - _PAD - _LOGO_W,
                    rect.y0 + _LOGO_PAD_Y,
                    rect.x1 - _PAD,
                    rect.y0 + _LOGO_PAD_Y + logo_h,
                )
                try:
                    new_page.insert_image(r, filename=str(logo_fs), keep_proportion=True, overlay=True)
                except Exception:
                    pass

        return out.tobytes(garbage=4, deflate=True, clean=True)
    finally:
        src.close()
        out.close()
