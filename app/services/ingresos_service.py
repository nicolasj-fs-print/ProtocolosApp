"""Lookup en el Excel `Protocolos x Ingreso OK.xlsx`.

Cols esperadas (por POSICIÓN, no por nombre):
  F (idx 5) → LF (FormularioNumero)
  G (idx 6) → Producto
  H (idx 7) → Protocolo (código tipo PR####)
"""
from __future__ import annotations

import io

import pandas as pd

from ..config import get_settings
from ..utils.io import read_file_shared
from ..utils.logger import get_logger

log = get_logger(__name__)


def load_ingresos() -> pd.DataFrame:
    """Devuelve DataFrame con cols `LF`, `ProductoExcel`, `Protocolo`."""
    s = get_settings()
    cols = ["LF", "ProductoExcel", "Protocolo"]
    sheet: int | str = s.ingresos_sheet if s.ingresos_sheet else 0

    if s.storage_backend == "sharepoint":
        from .data_service import _sp_get_or_download
        try:
            data = _sp_get_or_download(s.sp_ingresos_file)
        except Exception as e:
            log.warning("No se pudo descargar Ingresos xlsx desde SP: %s", e)
            return pd.DataFrame(columns=cols)
        raw = pd.read_excel(
            io.BytesIO(data),
            sheet_name=sheet,
            engine="openpyxl",
            dtype=str,
            header=0,
        )
    else:
        if not s.ingresos_xlsx or not s.ingresos_xlsx.exists():
            log.warning("Ingresos xlsx no configurado o no encontrado: %s", s.ingresos_xlsx)
            return pd.DataFrame(columns=cols)
        try:
            raw = pd.read_excel(
                s.ingresos_xlsx,
                sheet_name=sheet,
                engine="openpyxl",
                dtype=str,
                header=0,
            )
        except PermissionError:
            log.warning("Ingresos xlsx bloqueado → lectura compartida.")
            data = read_file_shared(s.ingresos_xlsx)
            raw = pd.read_excel(
                io.BytesIO(data),
                sheet_name=sheet,
                engine="openpyxl",
                dtype=str,
                header=0,
            )

    if raw.shape[1] < 8:
        log.error(
            "Ingresos xlsx con %d columnas (<8). No se puede mapear F/G/H.",
            raw.shape[1],
        )
        return pd.DataFrame(columns=cols)

    out = pd.DataFrame({
        "LF": raw.iloc[:, 5],
        "ProductoExcel": raw.iloc[:, 6],
        "Protocolo": raw.iloc[:, 7],
    })
    for c in cols:
        out[c] = out[c].fillna("").astype(str).str.strip()
    out = out[(out["LF"] != "") | (out["ProductoExcel"] != "")]
    log.info("Ingresos xlsx: %d filas cargadas.", len(out))
    return out.reset_index(drop=True)
