"""Lookup de trazabilidad de papel en DB secundaria FSBI."""
from __future__ import annotations

from typing import Sequence

import pandas as pd

from ..config import get_settings
from ..db import queries as q
from ..db.connection import get_connection
from ..utils.logger import get_logger

log = get_logger(__name__)

_CHUNK_SIZE = 1500  # margen al límite ~2100 de parámetros de SQL Server


def lookup_trazabilidad(series: Sequence[str]) -> pd.DataFrame:
    """Devuelve DataFrame con cols `Serie`, `FormularioCodigo`,
    `FormularioNumero`, `ProductoTraza`. TOP 1 por Serie."""
    cols = ["Serie", "FormularioCodigo", "FormularioNumero", "ProductoTraza"]
    cleaned = sorted({str(s).strip() for s in series if s and str(s).strip()})
    if not cleaned:
        return pd.DataFrame(columns=cols)

    settings = get_settings()
    db = settings.sql_database_trazabilidad

    chunks_df: list[pd.DataFrame] = []
    with get_connection(database=db) as conn:
        for i in range(0, len(cleaned), _CHUNK_SIZE):
            chunk = cleaned[i:i + _CHUNK_SIZE]
            sql = q.build_trazabilidad_query(len(chunk))
            params: list = list(q.TRAZABILIDAD_FORMULARIO_CODES) + chunk
            log.info("Trazabilidad: chunk %d-%d (n=%d)", i, i + len(chunk), len(chunk))
            df = pd.read_sql(sql, conn, params=params)
            chunks_df.append(df)

    if not chunks_df:
        return pd.DataFrame(columns=cols)

    out = pd.concat(chunks_df, ignore_index=True)
    out = out.rename(columns={"Producto": "ProductoTraza"})
    for c in ("Serie", "FormularioCodigo", "FormularioNumero", "ProductoTraza"):
        out[c] = out[c].fillna("").astype(str).str.strip()

    out = out.drop_duplicates(subset=["Serie"], keep="first").reset_index(drop=True)
    log.info("Trazabilidad: %d series matcheadas (de %d consultadas).", len(out), len(cleaned))
    return out[cols]
