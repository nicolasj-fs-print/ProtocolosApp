"""Detalle de protocolo individual: misma estructura que carátula pero filtrado por protocolo."""
from __future__ import annotations

import io
from typing import Iterable

from reportlab.platypus import SimpleDocTemplate

from . import styles as st
from .common import CaratulaMeta, build_header_block, build_products_table


def build_detalle(
    meta: CaratulaMeta,
    rows: Iterable[dict],
    protocolo_id: str,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=st.PAGE_SIZE,
        leftMargin=st.MARGIN,
        rightMargin=st.MARGIN,
        topMargin=st.MARGIN,
        bottomMargin=st.MARGIN,
        title=f"Detalle Protocolo {protocolo_id}",
        author="FS Print & Projects",
    )
    story: list = []
    story.extend(build_header_block(f"Detalle Protocolo {protocolo_id}", meta))
    story.append(build_products_table(list(rows)))
    doc.build(story)
    return buf.getvalue()
