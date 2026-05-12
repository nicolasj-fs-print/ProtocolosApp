"""Estilos compartidos entre carátula y detalle."""
from __future__ import annotations

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm

PAGE_SIZE = A4
MARGIN = 15 * mm

PRIMARY = colors.HexColor("#1F3864")
LIGHT = colors.HexColor("#D9E1F2")
GREY = colors.HexColor("#7F7F7F")
BORDER = colors.HexColor("#404040")

_styles = getSampleStyleSheet()

TITLE = ParagraphStyle(
    "Title",
    parent=_styles["Heading1"],
    fontName="Helvetica-Bold",
    fontSize=16,
    leading=20,
    textColor=PRIMARY,
    spaceAfter=4,
)

SUBTITLE = ParagraphStyle(
    "Subtitle",
    parent=_styles["Normal"],
    fontName="Helvetica",
    fontSize=10,
    leading=13,
    textColor=colors.black,
)

LABEL = ParagraphStyle(
    "Label",
    parent=_styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=9,
    leading=12,
    textColor=PRIMARY,
)

SMALL = ParagraphStyle(
    "Small",
    parent=_styles["Normal"],
    fontName="Helvetica",
    fontSize=8,
    leading=10,
    textColor=GREY,
)

CELL = ParagraphStyle(
    "Cell",
    parent=_styles["Normal"],
    fontName="Helvetica",
    fontSize=8,
    leading=10,
)

CELL_RIGHT = ParagraphStyle(
    "CellRight",
    parent=_styles["Normal"],
    fontName="Helvetica",
    fontSize=8,
    leading=10,
    alignment=2,  # TA_RIGHT
)

CELL_HEADER = ParagraphStyle(
    "CellHeader",
    parent=_styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=8,
    leading=10,
    textColor=colors.whitesmoke,
    alignment=1,  # center
)

TABLE_STYLE_BASE = [
    ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
    ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
    ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
    ("ALIGN", (2, 1), (4, -1), "RIGHT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
    ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
]

TABLE_HEADERS = [
    "Producto",
    "Descripción",
    "Ancho",
    "Largo",
    "m²",
    "Serie",
    "Protocolo",
]
