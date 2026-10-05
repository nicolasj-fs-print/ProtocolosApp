"""Qué protocolo (`PR####`) le toca a cada material.

============================================================================
FUENTE: `app_2_db`, NO el Excel  (2026-09-11)
============================================================================
Hasta esta fecha salía de `Protocolos x Ingreso OK.xlsx`, que una persona mantenía a mano.
Ahora sale de `public.v_protocolos_ingresos`, que la app de Seguimiento de Órdenes llena
desde `/despacho/protocolos/certificados`. Era la última planilla de la que dependía una
corrida del bot.

⚠️ SI LA BASE NO RESPONDE SE CAE AL EXCEL, y se grita. Es el mismo criterio que
`load_clientes_protocolos()`: abortar la corrida dejaría sin certificado a TODOS los remitos
del día por un problema de red. El Excel queda como red de emergencia —puede estar
desactualizado— y por eso `ULTIMA_FUENTE_INGRESOS` queda en "excel-fallback" y el WARNING
sale con todas las letras.

⚠️ Lo que NO hay que hacer es al revés —leer el Excel y caer a la base—: sería volver a tener
dos fuentes vivas, que es de donde venimos.

⚠️⚠️ La vista devuelve TODAS las filas, no una por clave. 25 claves `(producto, LF)` tienen
2 o 3 protocolos DISTINTOS —un ingreso puede traer el mismo producto llegado en varias Notas
Fiscales—, y el `merge` de `build_dataframe()` **no deduplica después**, así que hoy el bot le
adjunta al rollo todos los candidatos. Deduplicar acá le cambiaría al cliente lo que recibe,
así que **no se deduplica**: se replica exactamente lo que hacía el Excel.

── El Excel, mientras siga siendo el respaldo ──────────────────────────────
Cols esperadas (por POSICIÓN, no por nombre):
  F (idx 5) → LF (FormularioNumero)
  G (idx 6) → Producto
  H (idx 7) → Protocolo (código tipo PR####)
"""
from __future__ import annotations

import io

import pandas as pd

from ..config import get_settings
from ..db import app2
from ..utils.io import read_file_shared
from ..utils.logger import get_logger

log = get_logger(__name__)

#: De dónde salió el último mapeo: "app_2_db" o "excel-fallback".
ULTIMA_FUENTE_INGRESOS = "(sin leer)"

#: Las columnas que espera `build_dataframe()` para su merge. No cambiar los nombres: el
#: merge es `right_on=["ProductoExcel", "LF"]` y el resto del pipeline los borra por nombre.
_COLS = ["LF", "ProductoExcel", "Protocolo"]


def _ingresos_desde_app2() -> pd.DataFrame:
    """El mapeo de app_2_db, con la MISMA forma que devolvía el Excel.

    ⚠️ `astype(str)` porque el camino del Excel leía con `dtype=str` y el merge de
    `build_dataframe()` compara contra columnas de texto. Con un DataFrame vacío pandas les
    pondría float64 y el merge no matchearía nada **sin dar error** — justo el día que la
    tabla esté vacía, que es el caso raro que nadie prueba.
    """
    filas = app2.leer_ingresos_protocolo()
    df = pd.DataFrame(
        {
            "LF": [f["lf"] for f in filas],
            "ProductoExcel": [f["producto"] for f in filas],
            "Protocolo": [f["protocolo"] for f in filas],
        }
    )
    return df[_COLS].astype(str).reset_index(drop=True)


def load_ingresos() -> pd.DataFrame:
    """Devuelve DataFrame con cols `LF`, `ProductoExcel`, `Protocolo`.

    Lee de `app_2_db` y cae al Excel si no puede. Ver la cabecera del módulo.
    """
    global ULTIMA_FUENTE_INGRESOS

    if app2.esta_configurado():
        try:
            df = _ingresos_desde_app2()
            ULTIMA_FUENTE_INGRESOS = "app_2_db"
            log.info("Ingresos (app_2_db): %d filas cargadas.", len(df))
            return df
        except Exception as e:
            log.error(
                "!! No se pudo leer los ingresos de app_2_db (%s). Se usa el Excel como "
                "RESPALDO: puede estar desactualizado y no incluir los certificados "
                "cargados desde la app.",
                e,
            )
    else:
        log.warning(
            "app_2_db no esta configurada (APP2_PG_*): se lee el Excel. Es el modo viejo; "
            "los certificados cargados desde /despacho/protocolos NO se van a ver."
        )

    ULTIMA_FUENTE_INGRESOS = "excel-fallback"
    return _load_ingresos_excel()


def _load_ingresos_excel() -> pd.DataFrame:
    """El camino viejo: leer `Protocolos x Ingreso OK.xlsx`. **Solo como respaldo.**"""
    s = get_settings()
    cols = _COLS
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
    log.warning("Ingresos xlsx (RESPALDO): %d filas cargadas.", len(out))
    return out.reset_index(drop=True)
