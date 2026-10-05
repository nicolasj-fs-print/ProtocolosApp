"""Parsers de PDFs de Fedrigoni para el MailBot (Iteración 8).

Formato 1 (sender mis@fedrigoni.com):
    Nota Fiscal: 000093200
    Cod.Ref.: B1147
    Data: 22/05/2026

Formato 2 (sender larissa.lopes_ext@fedrigoni.com):
    NF: 91684
    Data: 24/04/2026
    Produto: THERMAL TOP BPA FREE FSC / P7 / YG55

Cuando el sender es ambiguo (ej. nicolas.jerez@fs-print.com usado para pruebas
en AMBAS listas), `autodetect_format(text)` decide en base al contenido del PDF.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

import fitz  # PyMuPDF

from ..utils.logger import get_logger

log = get_logger(__name__)

FormatoType = Literal["F1", "F2"]


class PdfParseError(RuntimeError):
    pass


@dataclass
class ParsedF1:
    nota_fiscal: str       # número sin ceros a la izquierda
    cod_ref: str           # código de producto directo (ej. B1147)
    fecha: str             # dd/MM/yyyy como string


@dataclass
class ParsedF2:
    nf: str                # número sin ceros a la izquierda
    fecha: str             # dd/MM/yyyy como string
    producto: str          # string completo del Produto (sin '/'), ej.
                           # "THERMAL TOP BPA FREE FSC P7 YG55" o
                           # "TINTORETTO GESSO H+O ULTRA WS FSC SH9020 WG74".
                           # El lookup SQL hace match fuzzy por palabras.


# ---------------------------------------------------------------------------
# Extracción de texto del PDF
# ---------------------------------------------------------------------------
def extract_text(pdf_bytes: bytes) -> str:
    """Devuelve TODO el texto del PDF concatenado. Lanza PdfParseError si falla.

    Usa `sort=True` para que el texto salga en orden visual de lectura
    (top-to-bottom, left-to-right), no en el orden interno del PDF. Crítico
    para PDFs con layout de tabla (como los certificados de Fedrigoni), donde
    sin sort los labels quedan separados de sus valores.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise PdfParseError(f"No se pudo abrir el PDF: {e}") from e
    try:
        chunks: list[str] = []
        for page in doc:
            try:
                chunks.append(page.get_text("text", sort=True))
            except Exception as e:
                log.warning("Error extrayendo texto de página: %s", e)
        return "\n".join(chunks)
    finally:
        doc.close()


# ---------------------------------------------------------------------------
# Detección de formato por contenido
# ---------------------------------------------------------------------------
_RE_F1_MARKERS = (
    re.compile(r"Nota\s*Fiscal\s*:", re.IGNORECASE),
    re.compile(r"Cod\.\s*Ref\.\s*:", re.IGNORECASE),
)
_RE_F2_MARKERS = (
    # NF puede venir como "NF:" o como "NF " (sin dos puntos).
    re.compile(r"\bNF\s*:?\s*\d", re.IGNORECASE),
    re.compile(r"Produto\s*:", re.IGNORECASE),
)


def autodetect_format(text: str) -> FormatoType:
    """Decide si el texto es Formato 1 o 2.

    F1 requiere "Nota Fiscal:" y "Cod.Ref.:" (ambos).
    F2 requiere "NF:" (al inicio de línea o con espacios) y "Produto:".

    Si los dos matchean o ninguno → PdfParseError (ambiguo o no reconocido).
    """
    is_f1 = all(rx.search(text) for rx in _RE_F1_MARKERS)
    is_f2 = all(rx.search(text) for rx in _RE_F2_MARKERS)
    if is_f1 and not is_f2:
        return "F1"
    if is_f2 and not is_f1:
        return "F2"
    if is_f1 and is_f2:
        raise PdfParseError("Autodetect ambiguo: el PDF tiene marcadores de F1 y F2.")
    raise PdfParseError("Autodetect falló: no se reconoce ningún formato Fedrigoni.")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _strip_leading_zeros(s: str) -> str:
    cleaned = (s or "").strip().lstrip("0")
    return cleaned or "0"


def _normalize_fecha(s: str) -> str:
    """Acepta '22/05/2026', '22-05-2026' o ISO; devuelve siempre 'dd/MM/yyyy'.
    Si no puede parsear, devuelve el string original tal cual."""
    s = (s or "").strip()
    if not s:
        return ""
    for sep in ("/", "-", "."):
        parts = s.split(sep)
        if len(parts) == 3:
            try:
                d = int(parts[0]); m = int(parts[1]); y = int(parts[2])
                if y < 100:
                    y += 2000
                return f"{d:02d}/{m:02d}/{y:04d}"
            except ValueError:
                pass
    return s


# ---------------------------------------------------------------------------
# Formato 1
# ---------------------------------------------------------------------------
_RE_F1_NF = re.compile(r"Nota\s*Fiscal\s*:\s*([0-9]+)", re.IGNORECASE)
_RE_F1_CODREF = re.compile(r"Cod\.\s*Ref\.\s*:\s*([A-Za-z0-9_\-\.]+)", re.IGNORECASE)
# Estricto dd/mm/yyyy (2-2-4) para evitar matchear el footer del PDF que tiene
# formato americano "Data:5/22/2026 - 9:43:27AM".
_RE_F1_DATA = re.compile(r"Data\s*:\s*([0-9]{2}[\-/\.][0-9]{2}[\-/\.][0-9]{4})", re.IGNORECASE)


def parse_formato_1(pdf_bytes: bytes) -> ParsedF1:
    text = extract_text(pdf_bytes)

    m_nf = _RE_F1_NF.search(text)
    if not m_nf:
        raise PdfParseError("Formato 1: no se encontró 'Nota Fiscal:'.")
    m_cr = _RE_F1_CODREF.search(text)
    if not m_cr:
        raise PdfParseError("Formato 1: no se encontró 'Cod.Ref.:'.")
    m_dt = _RE_F1_DATA.search(text)
    if not m_dt:
        raise PdfParseError("Formato 1: no se encontró 'Data:'.")

    return ParsedF1(
        nota_fiscal=_strip_leading_zeros(m_nf.group(1)),
        cod_ref=m_cr.group(1).strip(),
        fecha=_normalize_fecha(m_dt.group(1)),
    )


# ---------------------------------------------------------------------------
# Formato 2
# ---------------------------------------------------------------------------
# `NF` puede venir con `:` o sin él (`NF 91684` o `NF: 91684`). Aceptamos las dos.
_RE_F2_NF = re.compile(r"\bNF\s*:?\s*([0-9]+)", re.IGNORECASE)
_RE_F2_DATA = re.compile(r"Data\s*:\s*([0-9]{2}[\-/\.][0-9]{2}[\-/\.][0-9]{4})", re.IGNORECASE)
# `Produto:` puede venir con `/` separador o sin él (separación por espacios).
# NO splitamos por `/` — devolvemos el string completo, el lookup SQL hace
# match fuzzy por palabras.
_RE_F2_PRODUTO = re.compile(r"Produto\s*:\s*(.+)", re.IGNORECASE)


def parse_formato_2(pdf_bytes: bytes) -> ParsedF2:
    text = extract_text(pdf_bytes)

    m_nf = _RE_F2_NF.search(text)
    if not m_nf:
        raise PdfParseError("Formato 2: no se encontró 'NF'.")
    m_dt = _RE_F2_DATA.search(text)
    if not m_dt:
        raise PdfParseError("Formato 2: no se encontró 'Data:'.")
    m_pr = _RE_F2_PRODUTO.search(text)
    if not m_pr:
        raise PdfParseError("Formato 2: no se encontró 'Produto:'.")

    raw_produto = m_pr.group(1).strip()
    # Tomamos solo la primera línea (si después del produto hay más texto en
    # líneas siguientes lo ignoramos).
    raw_produto = raw_produto.split("\n", 1)[0].strip()
    # Normalizar: quitar `/` (en algunos PDFs separa, en otros no aparece) y
    # colapsar whitespace múltiple a uno solo.
    produto = raw_produto.replace("/", " ")
    produto = " ".join(produto.split())

    if not produto:
        raise PdfParseError("Formato 2: 'Produto' vacío.")

    return ParsedF2(
        nf=_strip_leading_zeros(m_nf.group(1)),
        fecha=_normalize_fecha(m_dt.group(1)),
        producto=produto,
    )
