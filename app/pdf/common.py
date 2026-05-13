"""Helpers comunes para construir PDFs (header con logos, formateo de filas)."""
from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from ..utils.format import format_serie
from ..utils.paths import resource_path
from . import styles as st

# ============================================================
# CONFIG DE LOGOS Y MÁRGENES — tocá acá para mover/redimensionar.
# Todas las unidades en mm.
# ============================================================

# Tamaños (alto). Ancho se calcula automático manteniendo proporción.
LOGO_FS_HEIGHT_MM = 20       # logo FS (izquierda) — más grande.
LOGO_FED_HEIGHT_MM = 14      # logo Fedrigoni (derecha) — más chico.

# Ancho reservado para cada celda de logo (caja contenedora).
LOGO_COLUMN_WIDTH_MM = 45

# Posición FS (izquierda):
#   X → LEFT_PADDING positivo lo empuja hacia la derecha (alejándolo del borde izquierdo).
#   Y → TOP_PADDING positivo lo empuja hacia abajo.
LOGO_FS_LEFT_PADDING_MM = 0
LOGO_FS_TOP_PADDING_MM = 0

# Posición Fedrigoni (derecha):
#   X → RIGHT_PADDING positivo lo empuja hacia la izquierda (alejándolo del borde derecho).
#   Y → TOP_PADDING positivo lo empuja hacia abajo.
LOGO_FED_RIGHT_PADDING_MM = 0
LOGO_FED_TOP_PADDING_MM = 9

# Spacing vertical del bloque header completo dentro del PDF.
TOP_SPACING_MM = 0           # margen arriba del header (más aire desde el borde superior).
BOTTOM_SPACING_MM = 4        # spacer entre header y tabla de productos.


@dataclass
class CaratulaMeta:
    cliente: str
    razon_social: str
    comprobante: str
    numero_oc: str
    fecha: str | datetime | None = None

    @property
    def fecha_str(self) -> str:
        if self.fecha is None:
            return datetime.now().strftime("%d/%m/%Y")
        if isinstance(self.fecha, datetime):
            return self.fecha.strftime("%d/%m/%Y")
        return str(self.fecha)


def _logo_image(rel_path: str, height_mm: float = 14) -> Image | None:
    path = resource_path(rel_path)
    if not path.exists():
        return None
    img = Image(str(path))
    h = height_mm * mm
    ratio = img.imageWidth / max(img.imageHeight, 1)
    img.drawHeight = h
    img.drawWidth = h * ratio
    return img


def build_header_block(title: str, meta: CaratulaMeta) -> list:
    """Bloque superior con logo FS (izquierda) + título/datos (centro) +
    logo Fedrigoni (derecha)."""
    logo_fs = _logo_image("assets/logo_fs.png", height_mm=LOGO_FS_HEIGHT_MM)
    logo_fed = _logo_image("assets/logo_fedrigoni.png", height_mm=LOGO_FED_HEIGHT_MM)

    info_html = (
        f"<b>{title}</b><br/>"
        f"<font size='9'>"
        f"<b>Cliente:</b> {_safe(meta.cliente)} - {_safe(meta.razon_social)}<br/>"
        f"<b>Remito:</b> {_safe(meta.comprobante)}<br/>"
        f"<b>OC:</b> {_safe(meta.numero_oc)}<br/>"
        f"<b>Fecha:</b> {meta.fecha_str}"
        f"</font>"
    )
    info = Paragraph(info_html, st.TITLE)

    fs_cell = logo_fs if logo_fs is not None else Paragraph("", st.SMALL)
    fed_cell = logo_fed if logo_fed is not None else Paragraph("", st.SMALL)

    # Layout: [info izquierda] [FS medio] [Fed derecha]
    info_width_mm = 170 - 2 * LOGO_COLUMN_WIDTH_MM  # ancho usable ~A4 - márgenes
    container = Table(
        [[info, fs_cell, fed_cell]],
        colWidths=[
            info_width_mm * mm,
            LOGO_COLUMN_WIDTH_MM * mm,
            LOGO_COLUMN_WIDTH_MM * mm,
        ],
    )
    container.setStyle(TableStyle([
        ("VALIGN", (0, 0), (0, 0), "MIDDLE"),  # info centrada verticalmente
        ("VALIGN", (1, 0), (2, 0), "TOP"),     # logos pegados arriba (alineados a "Protocolo de Calidad")
        ("ALIGN", (0, 0), (0, 0), "LEFT"),     # info pegada a la izquierda
        ("ALIGN", (1, 0), (1, 0), "LEFT"),     # FS al medio (alineado a la izquierda de su celda)
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),    # Fed a la derecha

        # Info (columna 0) sin padding extra.
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), 0),
        ("TOPPADDING", (0, 0), (0, 0), 2),
        ("BOTTOMPADDING", (0, 0), (0, 0), 6),

        # FS (columna 1): X = LEFTPADDING (positivo = FS más a la derecha).
        ("LEFTPADDING", (1, 0), (1, 0), LOGO_FS_LEFT_PADDING_MM * mm),
        ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (1, 0), (1, 0), LOGO_FS_TOP_PADDING_MM * mm + 2),
        ("BOTTOMPADDING", (1, 0), (1, 0), 6),

        # Fedrigoni (columna 2): X = RIGHTPADDING (positivo = Fed más a la izquierda).
        ("LEFTPADDING", (2, 0), (2, 0), 0),
        ("RIGHTPADDING", (2, 0), (2, 0), LOGO_FED_RIGHT_PADDING_MM * mm),
        ("TOPPADDING", (2, 0), (2, 0), LOGO_FED_TOP_PADDING_MM * mm + 2),
        ("BOTTOMPADDING", (2, 0), (2, 0), 6),
    ]))
    return [
        Spacer(1, TOP_SPACING_MM * mm),    # ← margen arriba (configurable).
        container,
        Spacer(1, BOTTOM_SPACING_MM * mm),
    ]


def build_products_table(rows: Iterable[dict]) -> Table:
    data = [[Paragraph(h, st.CELL_HEADER) for h in st.TABLE_HEADERS]]
    total_m2 = 0.0
    n_rollos = 0
    for r in rows:
        n_rollos += 1
        try:
            total_m2 += float(r.get("#m2") or 0)
        except (TypeError, ValueError):
            pass

        proto_val = str(r.get("#Protocolo PDF") or "").strip()
        proto_display = proto_val if proto_val else "En proceso"

        data.append([
            Paragraph(_safe(r.get("Producto")), st.CELL),
            Paragraph(_safe(r.get("Descripcion2")), st.CELL),
            Paragraph(_fmt_num(r.get("Ancho")), st.CELL_RIGHT),
            Paragraph(_fmt_num(r.get("Largo")), st.CELL_RIGHT),
            Paragraph(_fmt_int(r.get("#m2")), st.CELL_RIGHT),
            Paragraph(_safe(format_serie(r.get("Serie"))), st.CELL),
            Paragraph(_safe(proto_display), st.CELL),
        ])

    data.append([
        Paragraph("<b>Total m²</b>", st.CELL),
        "", "", "",
        Paragraph(f"<b>{_fmt_int(total_m2)}</b>", st.CELL_RIGHT),
        "", "",
    ])
    data.append([
        Paragraph("<b>Cantidad de rollos</b>", st.CELL),
        "", "", "",
        Paragraph(f"<b>{n_rollos}</b>", st.CELL_RIGHT),
        "", "",
    ])

    col_widths = [18 * mm, 50 * mm, 18 * mm, 18 * mm, 18 * mm, 34 * mm, 24 * mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    style = TableStyle(st.TABLE_STYLE_BASE + [
        ("SPAN", (0, -2), (3, -2)),     # Total m²
        ("SPAN", (0, -1), (3, -1)),     # Cantidad de rollos
        ("BACKGROUND", (0, -2), (-1, -1), colors.HexColor("#F2F2F2")),
        ("ALIGN", (4, -2), (4, -1), "RIGHT"),
        ("FONT", (0, -2), (-1, -1), "Helvetica-Bold", 8),
    ])
    table.setStyle(style)
    return table


def _safe(v) -> str:
    if v is None:
        return ""
    return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _fmt_num(v) -> str:
    if v is None or v == "":
        return ""
    try:
        f = float(v)
        if f.is_integer():
            return f"{int(f):,}".replace(",", ".")
        return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(v)


def _fmt_int(v) -> str:
    """M2 / cantidades enteras: redondea y pone separador de miles con punto."""
    if v is None or v == "":
        return ""
    try:
        return f"{int(round(float(v))):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(v)
