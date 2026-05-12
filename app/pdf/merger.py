"""Merge de PDFs con PyMuPDF + compresión final."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

import fitz

# Paginación: "1 de N" en la esquina superior izquierda
_PAGINATION_FONTSIZE = 7
_PAGINATION_X = 10
_PAGINATION_Y = 10
_PAGINATION_COLOR = (0.4, 0.4, 0.4)


def _merge_into(parts: Iterable[bytes]) -> fitz.Document:
    merged = fitz.open()
    for chunk in parts:
        if not chunk:
            continue
        sub = fitz.open(stream=chunk, filetype="pdf")
        try:
            merged.insert_pdf(sub)
        finally:
            sub.close()
    return merged


def _paginate(doc: fitz.Document) -> None:
    """Dibuja 'X de N' en la esquina superior izquierda de cada página."""
    n = doc.page_count
    for i, page in enumerate(doc, 1):
        try:
            page.insert_text(
                (_PAGINATION_X, _PAGINATION_Y),
                f"{i} de {n}",
                fontsize=_PAGINATION_FONTSIZE,
                fontname="helv",
                color=_PAGINATION_COLOR,
                overlay=True,
            )
        except Exception:
            pass


def merge_pdfs(parts: Iterable[bytes], output_path: Path) -> Path:
    """Merge a archivo en disco (backend local)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    merged = _merge_into(parts)
    try:
        _paginate(merged)
        merged.save(
            str(output_path),
            garbage=4,
            deflate=True,
            clean=True,
        )
    finally:
        merged.close()
    return output_path


def merge_pdfs_to_bytes(parts: Iterable[bytes]) -> bytes:
    """Merge en memoria → bytes (backend SharePoint)."""
    merged = _merge_into(parts)
    try:
        _paginate(merged)
        return merged.tobytes(garbage=4, deflate=True, clean=True)
    finally:
        merged.close()
