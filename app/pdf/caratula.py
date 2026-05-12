"""Generación de la carátula consolidada por comprobante."""
from __future__ import annotations

import io
from typing import Iterable

from reportlab.platypus import SimpleDocTemplate

from . import styles as st
from .common import CaratulaMeta, build_header_block, build_products_table


def build_caratula(meta: CaratulaMeta, rows: Iterable[dict]) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=st.PAGE_SIZE,
        leftMargin=st.MARGIN,
        rightMargin=st.MARGIN,
        topMargin=st.MARGIN,
        bottomMargin=st.MARGIN,
        title=f"Protocolos {meta.comprobante}",
        author="FS Print & Projects",
    )
    story: list = []
    story.extend(build_header_block("Protocolos de Calidad", meta))
    story.append(build_products_table(list(rows)))
    doc.build(story)
    return buf.getvalue()
