"""Lookup de trazabilidad de papel en DB secundaria FSBI."""
from __future__ import annotations

from typing import Sequence

import pandas as pd

from ..config import get_settings
from ..db import queries as q
from ..db.connection import get_connection
from ..utils.logger import get_logger

log = get_logger(__name__)


def lookup_trazabilidad_by_nf(nf: str) -> dict | None:
    """Busca un registro en `TrazabilidadPapel` por `FormularioNumero` (NF).

    Usado por el MailBot Fedrigoni: dado el NF del PDF, devuelve el LF
    correspondiente. Aplica prioridad LF > IRP (descarta IR — se filtra ya
    en la query).

    Devuelve dict `{"FormularioCodigo", "FormularioNumero", "Producto"}` o
    `None` si no hay match.
    """
    nf_clean = str(nf or "").strip().lstrip("0") or "0"
    if not nf_clean:
        return None

    settings = get_settings()
    db = settings.sql_database_trazabilidad

    with get_connection(database=db) as conn:
        df = pd.read_sql(q.SELECT_TRAZABILIDAD_BY_NF, conn, params=[nf_clean])

    if df is None or df.empty:
        return None

    for c in ("FormularioCodigo", "FormularioNumero", "Producto"):
        df[c] = df[c].fillna("").astype(str).str.strip()

    # Prioridad LF (0) > IRP (1). Stable para preservar orden ante empates.
    priority_map = {"LF": 0, "IRP": 1}
    df["_priority"] = df["FormularioCodigo"].map(priority_map).fillna(99).astype(int)
    df = df.sort_values("_priority", kind="stable").reset_index(drop=True)

    row = df.iloc[0]
    return {
        "FormularioCodigo": row["FormularioCodigo"],
        "FormularioNumero": row["FormularioNumero"],
        "Producto": row["Producto"],
    }


def _split_words(s: str) -> list[str]:
    """Tokeniza un string por whitespace y `/`. Normaliza a UPPER + strip."""
    if not s:
        return []
    s = s.replace("/", " ")
    return [w.strip().upper() for w in s.split() if w.strip()]


def _is_word_subsequence(small: list[str], big: list[str]) -> bool:
    """True si todas las palabras de `small` aparecen en `big` en el mismo
    orden (no necesariamente contiguas).

    Ej: small=['THERMAL','TOP','P7','YG55']
        big=['THERMAL','TOP','BPA','FREE','FSC','P7','YG55']
        → True (las palabras de small están como subsecuencia ordenada de big).
    """
    if not small:
        return False
    j = 0
    for word in big:
        if word == small[j]:
            j += 1
            if j == len(small):
                return True
    return j == len(small)


def find_artcod_by_producto(producto: str) -> str | None:
    """Busca el ARTCOD que matchee con el `Producto` del PDF (Formato 2).

    Estrategia: trae todos los SAFED de STMPDH y aplica match fuzzy. Las
    palabras del CONCAT(DESCRP, ADHESI, PROLIN) de la BD deben aparecer en
    el `producto` del PDF en el mismo orden (subsecuencia). El PDF puede
    tener palabras extras como 'FSC', 'BPA FREE', etc., que la BD no
    almacena en esas columnas.

    Si hay varios matches, gana el de MÁS palabras (más específico).

    Ejemplos confirmados:
    - PDF "THERMAL TOP BPA FREE FSC P7 YG55"
      → B0402 (DESCRP='THERMAL TOP', ADHESI='P7', PROLIN='YG55')
    - PDF "TINTORETTO GESSO H+O ULTRA WS FSC SH9020 WG74"
      → B1042 (DESCRP='TINTORETTO GESSO H+O ULTRA WS', ADHESI='SH9020', PROLIN='WG74')
    """
    pdf_words = _split_words(producto)
    if not pdf_words:
        return None

    log.info("find_artcod_by_producto: buscando %r (%d palabras)",
             producto, len(pdf_words))

    settings = get_settings()
    db = settings.sql_database

    with get_connection(database=db) as conn:
        df = pd.read_sql(q.SELECT_ALL_SAFED_PRODUCTOS, conn)

    if df is None or df.empty:
        log.warning("find_artcod_by_producto: 0 SAFED en STMPDH")
        return None

    best_artcod: str | None = None
    best_match_count = 0
    best_concat = ""

    for _, row in df.iterrows():
        descrp = str(row.get("DESCRP", "")).strip()
        adhesi = str(row.get("ADHESI", "")).strip()
        prolin = str(row.get("PROLIN", "")).strip()
        if not descrp:
            continue
        # Concatenamos para tokenizar todo junto en orden DESCRP→ADHESI→PROLIN.
        db_concat = f"{descrp} {adhesi} {prolin}".strip()
        db_words = _split_words(db_concat)
        if not db_words:
            continue

        if _is_word_subsequence(db_words, pdf_words):
            if len(db_words) > best_match_count:
                best_artcod = str(row["ARTCOD"]).strip()
                best_match_count = len(db_words)
                best_concat = db_concat

    if best_artcod:
        log.info(
            "find_artcod_by_producto: %r → ARTCOD=%r (BD=%r, %d palabras match)",
            producto, best_artcod, best_concat, best_match_count,
        )
    else:
        log.info("find_artcod_by_producto: sin match para %r", producto)

    return best_artcod

_CHUNK_SIZE = 1500  # margen al límite ~2100 de parámetros de SQL Server


def lookup_trazabilidad(
    series: Sequence[str],
    *,
    prioritize_lf: bool = False,
) -> pd.DataFrame:
    """Devuelve DataFrame con cols `Serie`, `FormularioCodigo`,
    `FormularioNumero`, `ProductoTraza`. TOP 1 por Serie.

    Si `prioritize_lf=True`:
      - Descarta los registros con FormularioCodigo == 'IR'.
      - Prioriza LF sobre IRF cuando hay ambos para la misma Serie.

    Default (`prioritize_lf=False`): comportamiento original — toma el primero
    que devuelve la BD (orden arbitrario entre los 3 códigos).
    """
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

    if prioritize_lf:
        # Descartamos 'IR' (no es válido en esta política) y priorizamos LF > IRP.
        before = len(out)
        out = out[out["FormularioCodigo"].isin(["LF", "IRP"])].copy()
        descartados_ir = before - len(out)
        if descartados_ir:
            log.info("Trazabilidad (prioritize_lf): %d registro(s) IR descartados.",
                     descartados_ir)
        # Sort: LF (prioridad 0) antes que IRP (prioridad 1). Stable para conservar
        # el orden original ante empates.
        priority_map = {"LF": 0, "IRP": 1}
        out["_priority"] = out["FormularioCodigo"].map(priority_map).fillna(99).astype(int)
        out = out.sort_values(["Serie", "_priority"], kind="stable")
        out = out.drop(columns="_priority")

    out = out.drop_duplicates(subset=["Serie"], keep="first").reset_index(drop=True)
    log.info("Trazabilidad: %d series matcheadas (de %d consultadas).", len(out), len(cleaned))
    return out[cols]
