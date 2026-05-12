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
    """Bloque superior con título + datos cliente + logos a la derecha."""
    logo_fed = _logo_image("assets/logo_fedrigoni.png")
    logo_fs = _logo_image("assets/logo_fs.png")

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

    logos_cell = []
    if logo_fed is not None:
        logos_cell.append(logo_fed)
    if logo_fs is not None:
        logos_cell.append(logo_fs)
    if not logos_cell:
        logos_cell.append(Paragraph("", st.SMALL))

    logos_table = Table(
        [[c] for c in logos_cell],
        colWidths=[40 * mm],
        hAlign="RIGHT",
    )
    logos_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    container = Table(
        [[info, logos_table]],
        colWidths=[110 * mm, 60 * mm],
    )
    container.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return [container, Spacer(1, 4 * mm)]


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
