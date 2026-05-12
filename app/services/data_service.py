"""Acceso a datos: SQL Server (cabecera + items + cliente + producto) + Excel de clientes."""
from __future__ import annotations

import io
import threading
from datetime import date, datetime
from typing import Sequence

import pandas as pd

from ..config import get_settings
from ..db.connection import get_connection
from ..db import queries as q
from ..utils.format import format_serie
from ..utils.io import read_file_shared
from ..utils.logger import get_logger
from ..utils.validators import (
    REQUIRED_COLUMNS,
    validate_dataframe_columns,
    validate_date_range,
)
from .ingresos_service import load_ingresos
from .trazabilidad_service import lookup_trazabilidad

log = get_logger(__name__)

_CLIENT_COLUMNS_EXCEL = [
    "Código de cliente",
    "Region",
    "Protocolos",
    "Razon Social",
    "Mail Protocolos",
]

# Cache en memoria de los bytes de Excels descargados desde SharePoint.
# Vive mientras la app corre. Acelera "Buscar" del 2do click en adelante.
_SP_EXCEL_CACHE: dict[str, bytes] = {}
_SP_PATH_LOCKS: dict[str, threading.Lock] = {}
_SP_LOCKS_GUARD = threading.Lock()


def _sp_get_or_download(sp_path: str) -> bytes:
    """Devuelve los bytes del Excel.

    Estrategia 3 niveles:
      1) Cache en memoria (mismo run) → instantáneo.
      2) Cache en disco (%APPDATA%\\ProtocolosApp\\cache\\) → 1 GET de
         metadata a SharePoint para comparar `lastModifiedDateTime`.
         Si no cambió, leer del disco (~10ms).
      3) Descargar de SharePoint y guardar a disco.
    """
    cached = _SP_EXCEL_CACHE.get(sp_path)
    if cached is not None:
        log.info("Cache HIT (memoria): %s", sp_path)
        return cached

    with _SP_LOCKS_GUARD:
        lock = _SP_PATH_LOCKS.setdefault(sp_path, threading.Lock())

    with lock:
        cached = _SP_EXCEL_CACHE.get(sp_path)
        if cached is not None:
            log.info("Cache HIT (memoria, post-lock): %s", sp_path)
            return cached

        from ..sharepoint.client import get_client
        from ..utils import disk_cache

        client = get_client()

        # 1) Metadata del SP — rápido (~200ms).
        remote_modified = ""
        remote_etag = ""
        try:
            item = client.get_item(sp_path)
            if item:
                remote_modified = item.get("lastModifiedDateTime", "") or ""
                remote_etag = item.get("eTag", "") or ""
        except Exception as e:
            log.warning("get_item falló para %s: %s", sp_path, e)

        # 2) Disk cache (compara lastModifiedDateTime).
        if remote_modified:
            data = disk_cache.read_cached(sp_path, remote_modified)
            if data is not None:
                log.info("Cache HIT (disco, sin cambios en SP): %s (%.1f MB)",
                         sp_path, len(data) / (1024 * 1024))
                _SP_EXCEL_CACHE[sp_path] = data
                return data

        # 3) Descargar de SP.
        log.info("Cache MISS, descargando: %s", sp_path)
        data = client.download_file(sp_path)
        _SP_EXCEL_CACHE[sp_path] = data
        log.info("Descarga OK (%.1f MB): %s", len(data) / (1024 * 1024), sp_path)

        if remote_modified:
            disk_cache.write_cached(sp_path, data, remote_modified, remote_etag)

        return data


def clear_excel_cache() -> None:
    """Vacía el cache (forzá recarga la próxima búsqueda)."""
    _SP_EXCEL_CACHE.clear()


def warm_up_excels() -> None:
    """Pre-descarga los Excels al iniciar la app, en background.

    Llamarlo desde un thread daemon así el Buscar es instantáneo.
    Best-effort: si falla, log warning y se reintenta en la 1ra búsqueda.
    """
    s = get_settings()
    if s.storage_backend != "sharepoint":
        return
    log.info("Warm-up: pre-descargando Excels...")
    if s.sp_clients_file:
        try:
            _sp_get_or_download(s.sp_clients_file)
        except Exception as e:
            log.warning("Warm-up clients falló: %s", e)
    if s.sp_ingresos_file:
        try:
            _sp_get_or_download(s.sp_ingresos_file)
        except Exception as e:
            log.warning("Warm-up ingresos falló: %s", e)
    log.info("Warm-up completo.")


# ---------------------------------------------------------------------------
# Excel de clientes (filtro Protocolos = "S")
# ---------------------------------------------------------------------------
def load_clientes_protocolos() -> pd.DataFrame:
    s = get_settings()

    if s.storage_backend == "sharepoint":
        raw = _sp_get_or_download(s.sp_clients_file)
        df = pd.read_excel(
            io.BytesIO(raw),
            sheet_name=s.clients_sheet,
            engine="openpyxl",
            dtype=str,
        )
    else:
        if not s.clients_xlsx.exists():
            raise FileNotFoundError(f"Excel de clientes no encontrado: {s.clients_xlsx}")
        try:
            df = pd.read_excel(
                s.clients_xlsx,
                sheet_name=s.clients_sheet,
                engine="openpyxl",
                dtype=str,
            )
        except PermissionError:
            log.warning("Excel bloqueado (abierto en Excel/OneDrive) → lectura compartida.")
            raw = read_file_shared(s.clients_xlsx)
            df = pd.read_excel(
                io.BytesIO(raw),
                sheet_name=s.clients_sheet,
                engine="openpyxl",
                dtype=str,
            )

    df.columns = [str(c).replace("\xa0", " ").strip() for c in df.columns]

    missing = [c for c in _CLIENT_COLUMNS_EXCEL if c not in df.columns]
    if missing:
        raise ValueError(
            f"Columnas faltantes en hoja '{s.clients_sheet}' de {s.clients_xlsx.name}: {missing}. "
            f"Columnas presentes: {list(df.columns)}"
        )

    df["Protocolos"] = df["Protocolos"].fillna("").str.strip().str.upper()
    df = df[df["Protocolos"] == "S"].copy()
    df["Código de cliente"] = df["Código de cliente"].fillna("").astype(str).str.strip()
    df["Mail Protocolos"] = df["Mail Protocolos"].fillna("").astype(str).str.strip()
    df["Razon Social"] = df["Razon Social"].fillna("").astype(str).str.strip()

    log.info("Clientes con Protocolos='S': %d", len(df))
    return df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# SQL principal
# ---------------------------------------------------------------------------
def fetch_data(
    date_from: date,
    date_to: date,
    codigo_cliente: str | None = None,
    nrofor: str | None = None,
) -> pd.DataFrame:
    """Trae el detalle completo y devuelve un DataFrame con TODAS las columnas
    necesarias para tabla 1 (resumen) y tabla 2 (detalle SAFED).

    El filtrado fino por cliente / comprobante de la GUI se hace en pandas
    (cross-filter), por lo que `codigo_cliente`/`nrofor` aquí se usan solo
    para acotar la consulta cuando el usuario escribe valores explícitos.
    """
    validate_date_range(date_from, date_to)

    s = get_settings()
    if s.mock_data:
        log.warning("MOCK_DATA=true → usando dataset hardcoded")
        df = _mock_dataframe()
    else:
        df = _query_sql(date_from, date_to, codigo_cliente, nrofor)

    df = _normalize(df)
    df = _dedup_items(df)

    # Filtro Excel: solo clientes con Protocolos == "S"
    clientes = load_clientes_protocolos()
    if clientes.empty:
        log.warning("Excel de clientes sin filas con Protocolos='S' → resultado vacío.")
        empty = df.iloc[0:0].copy()
        if "Mail Protocolos" not in empty.columns:
            empty["Mail Protocolos"] = pd.Series(dtype=str)
        return _ensure_required_cols(empty)

    enriched = df.merge(
        clientes[["Código de cliente", "Mail Protocolos", "Razon Social"]].rename(
            columns={
                "Código de cliente": "Código de Cliente",
                "Mail Protocolos": "_Mail",
                "Razon Social": "_RazonExcel",
            }
        ),
        on="Código de Cliente",
        how="inner",
    )

    enriched["Mail Protocolos"] = enriched["_Mail"].fillna("")
    # Si SQL no devuelve nombre, completar con Excel
    enriched["Razón Social"] = enriched["Razón Social"].fillna("").astype(str)
    mask = enriched["Razón Social"].str.len() == 0
    enriched.loc[mask, "Razón Social"] = enriched.loc[mask, "_RazonExcel"]
    enriched = enriched.drop(columns=["_Mail", "_RazonExcel"])

    enriched = _ensure_required_cols(enriched)

    enriched = enrich_with_protocolos(enriched)

    validate_dataframe_columns(enriched)

    enriched = enriched.sort_values(["#Comprobante", "Producto"], na_position="last").reset_index(drop=True)
    log.info("Filas tras filtro de clientes con protocolos: %d", len(enriched))
    return enriched


# ---------------------------------------------------------------------------
# Enriquecimiento: Protocolo + Protocolo Serie LF (Iteración 2)
# ---------------------------------------------------------------------------
def enrich_with_protocolos(df: pd.DataFrame) -> pd.DataFrame:
    """Llena Protocolo, Protocolo Serie LF y #Protocolo PDF cruzando con
    FSBI.dbo.TrazabilidadPapel + Excel `Protocolos x Ingreso OK.xlsx`.

    Aplica solo a items con TipoProducto == 'SAFED'. Cualquier item sin
    match queda con valores vacíos (NO bloqueante).
    """
    if df.empty:
        return df

    df = df.copy()
    df["Serie"] = df["Serie"].fillna("").astype(str).str.strip()
    safed_mask = df["TipoProducto"].astype(str).str.upper() == "SAFED"
    df["SerieKey"] = ""
    df.loc[safed_mask, "SerieKey"] = df.loc[safed_mask, "Serie"].str[:7]

    serie_keys = sorted({s for s in df.loc[safed_mask, "SerieKey"] if s})
    if not serie_keys:
        log.info("No hay items SAFED con SerieKey → skip enriquecimiento.")
        df = df.drop(columns=["SerieKey"])
        return df

    # 1) Trazabilidad
    try:
        traza = lookup_trazabilidad(serie_keys)
    except Exception as e:
        log.exception("Error en lookup trazabilidad: %s", e)
        traza = pd.DataFrame(columns=["Serie", "FormularioCodigo", "FormularioNumero", "ProductoTraza"])

    if not traza.empty:
        df = df.merge(
            traza.rename(columns={"Serie": "SerieKey"}),
            on="SerieKey",
            how="left",
        )
    else:
        for c in ("FormularioCodigo", "FormularioNumero", "ProductoTraza"):
            df[c] = ""

    df["FormularioCodigo"] = df["FormularioCodigo"].fillna("").astype(str).str.strip()
    df["FormularioNumero"] = df["FormularioNumero"].fillna("").astype(str).str.strip()
    df["ProductoTraza"] = df["ProductoTraza"].fillna("").astype(str).str.strip()

    has_traza = df["FormularioCodigo"] != ""
    df["Protocolo Serie LF"] = ""
    df.loc[has_traza, "Protocolo Serie LF"] = (
        df.loc[has_traza, "FormularioCodigo"] + " - "
        + df.loc[has_traza, "FormularioNumero"] + " - "
        + df.loc[has_traza, "ProductoTraza"]
    )

    # 2) Excel ingresos
    try:
        ingresos = load_ingresos()
    except Exception as e:
        log.exception("Error leyendo ingresos xlsx: %s", e)
        ingresos = pd.DataFrame(columns=["LF", "ProductoExcel", "Protocolo"])

    df["Protocolo"] = ""
    if not ingresos.empty:
        df = df.merge(
            ingresos,
            left_on=["ProductoTraza", "FormularioNumero"],
            right_on=["ProductoExcel", "LF"],
            how="left",
            suffixes=("", "_excel"),
        )
        if "Protocolo_excel" in df.columns:
            df["Protocolo"] = df["Protocolo_excel"].fillna("").astype(str).str.strip()
            df = df.drop(columns=["Protocolo_excel"])
        else:
            df["Protocolo"] = df.get("Protocolo", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        for c in ("LF", "ProductoExcel"):
            if c in df.columns:
                df = df.drop(columns=[c])

    df["#Protocolo PDF"] = df["Protocolo"]

    df = df.drop(columns=[c for c in ("SerieKey",) if c in df.columns])

    n_total = int(safed_mask.sum())
    n_traza = int((df["FormularioCodigo"] != "").sum())
    n_proto = int((df["Protocolo"] != "").sum())
    log.info("Enriquecimiento: %d SAFED, %d con trazabilidad, %d con protocolo.",
             n_total, n_traza, n_proto)
    return df


def _query_sql(
    date_from: date,
    date_to: date,
    codigo_cliente: str | None,
    nrofor: str | None,
) -> pd.DataFrame:
    sql = q.build_detalle_query(by_cliente=bool(codigo_cliente), by_nrofor=bool(nrofor))
    params: list = [q.CODFOR_FILTER, q.CODFOR_FILTER, date_from, date_to]
    if codigo_cliente:
        params.append(codigo_cliente)
    if nrofor:
        params.append(nrofor)

    log.info("Ejecutando query SQL (rango %s → %s)", date_from, date_to)
    with get_connection() as conn:
        df = pd.read_sql(sql, conn, params=params)
    log.info("Filas devueltas por SQL: %d", len(df))
    return df


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza tipos y construye Descripcion2 / Producto / strings de claves."""
    if df.empty:
        return df

    for c in ("Código de Cliente", "#Comprobante", "Razón Social",
              "Producto", "Serie", "Número OC", "Observaciones",
              "TipoProducto", "_Descrp", "_Adhesi", "_Prolin",
              "CodFor", "NroFor"):
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    df["Descripcion2"] = (
        df.get("Producto", "").astype(str).str.strip() + " "
        + df.get("_Descrp", "").astype(str).str.strip() + " "
        + df.get("_Adhesi", "").astype(str).str.strip() + " "
        + df.get("_Prolin", "").astype(str).str.strip()
    ).str.replace(r"\s+", " ", regex=True).str.strip()

    df = df.drop(columns=[c for c in ("_Descrp", "_Adhesi", "_Prolin") if c in df.columns])

    # #Protocolo PDF queda vacío hasta confirmar de dónde sale
    if "#Protocolo PDF" not in df.columns:
        df["#Protocolo PDF"] = ""

    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")

    return df


def _dedup_items(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina duplicados generados por LEFT JOIN con VTMCLH/STMPDH cuando esas
    tablas tienen múltiples filas por la clave (sucursales, histórico, etc.).

    IMPORTANTE: cuando hay duplicados con distinto TipoProducto, prioriza
    `SAFED` para que esos items aparezcan en la tabla Detalle. Sin esta
    priorización, si STMPDH tiene el mismo ARTCOD con SAFED y otro tipo,
    el dedup podía quedarse con el tipo incorrecto y perder el item del Detalle.
    """
    if df.empty:
        return df
    key_cols = [c for c in ("CodFor", "NroFor", "Producto", "Serie", "Ancho", "Largo", "#m2")
                if c in df.columns]
    if not key_cols:
        return df
    before = len(df)

    df = df.copy()
    if "TipoProducto" in df.columns:
        # Marca filas con TipoProducto == "SAFED" → quedan primeras en el orden,
        # y `keep="first"` del drop_duplicates conserva esas.
        df["_safed_priority"] = (
            df["TipoProducto"].fillna("").astype(str).str.upper().str.strip() == "SAFED"
        ).astype(int)
        df = df.sort_values("_safed_priority", ascending=False, kind="stable")
        out = df.drop_duplicates(subset=key_cols, keep="first").reset_index(drop=True)
        out = out.drop(columns="_safed_priority")
    else:
        out = df.drop_duplicates(subset=key_cols, keep="first").reset_index(drop=True)

    removed = before - len(out)
    if removed:
        log.info("Deduplicación: %d fila(s) duplicadas eliminadas (de %d → %d).",
                 removed, before, len(out))
    return out


def _ensure_required_cols(df: pd.DataFrame) -> pd.DataFrame:
    for c in REQUIRED_COLUMNS:
        if c not in df.columns:
            df[c] = "" if c not in ("#m2", "Ancho", "Largo", "Fecha") else None
    return df


# ---------------------------------------------------------------------------
# Cross-filter helpers (pandas) usadas por la GUI
# ---------------------------------------------------------------------------
def distinct_clientes(df: pd.DataFrame) -> list[str]:
    if df.empty:
        return []
    return sorted({str(x) for x in df["Código de Cliente"].dropna().unique()})


def distinct_comprobantes(df: pd.DataFrame) -> list[str]:
    if df.empty:
        return []
    return sorted({str(x) for x in df["#Comprobante"].dropna().unique()})


def analyze_missing_protocols(df: pd.DataFrame) -> dict[str, dict]:
    """Detecta comprobantes con items SAFED sin #Protocolo PDF.

    Devuelve {comprobante: {cliente, razon_social, oc, mail, items: [ {producto,
    serie, m2, descripcion2}, ... ]}} solo para los comprobantes con AL MENOS
    un item SAFED faltante.
    """
    if df.empty:
        return {}

    safed = df[df["TipoProducto"].astype(str).str.upper() == "SAFED"].copy()
    if safed.empty:
        return {}

    safed["#Protocolo PDF"] = safed["#Protocolo PDF"].fillna("").astype(str).str.strip()
    missing = safed[safed["#Protocolo PDF"] == ""]
    if missing.empty:
        return {}

    out: dict[str, dict] = {}
    for comp, sub in missing.groupby("#Comprobante", sort=False):
        first = sub.iloc[0]
        items = []
        for _, row in sub.iterrows():
            m2 = row.get("#m2")
            try:
                m2_str = "" if pd.isna(m2) else f"{int(round(float(m2)))}"
            except (TypeError, ValueError):
                m2_str = str(m2 or "")
            items.append({
                "producto": str(row.get("Producto", "") or ""),
                "serie": str(row.get("Serie", "") or ""),
                "m2": m2_str,
                "descripcion2": str(row.get("Descripcion2", "") or ""),
            })
        out[str(comp)] = {
            "cliente": str(first.get("Código de Cliente", "") or ""),
            "razon_social": str(first.get("Razón Social", "") or ""),
            "oc": str(first.get("Número OC", "") or ""),
            "mail": str(first.get("Mail Protocolos", "") or ""),
            "items": items,
        }
    return out


def apply_filters(
    df: pd.DataFrame,
    clientes: Sequence[str] | None = None,
    comprobantes: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Aplica filtros de selección múltiple sobre el DataFrame ya cargado."""
    if df.empty:
        return df
    out = df
    if clientes:
        out = out[out["Código de Cliente"].astype(str).isin([str(c) for c in clientes])]
    if comprobantes:
        out = out[out["#Comprobante"].astype(str).isin([str(c) for c in comprobantes])]
    return out


# ---------------------------------------------------------------------------
# Vistas (tabla 1: resumen por comprobante / tabla 2: detalle SAFED)
# ---------------------------------------------------------------------------
def build_summary_view(df: pd.DataFrame) -> pd.DataFrame:
    """Tabla 1: una fila por #Comprobante.

    M2 = SUM(#m2) sin filtrar por TipoProducto (todos los items con CODFOR=RX0018,
    como pidió la spec).
    """
    if df.empty:
        return pd.DataFrame(columns=[
            "Cliente", "Comprobante", "Fecha", "M2", "Observaciones",
        ])

    df = df.copy()
    df["#m2"] = pd.to_numeric(df["#m2"], errors="coerce").fillna(0.0)

    grouped = df.groupby("#Comprobante", dropna=False, sort=False).agg(
        Cliente=("Código de Cliente", "first"),
        Fecha=("Fecha", "first"),
        M2=("#m2", "sum"),
        Observaciones=("Observaciones", "first"),
    ).reset_index().rename(columns={"#Comprobante": "Comprobante"})

    grouped["Fecha"] = pd.to_datetime(grouped["Fecha"], errors="coerce").dt.strftime("%d-%m-%Y").fillna("")
    grouped["M2"] = grouped["M2"].map(lambda x: f"{int(round(x))}" if pd.notna(x) else "")
    return grouped[["Cliente", "Comprobante", "Fecha", "M2", "Observaciones"]]


def build_detail_view(df: pd.DataFrame) -> pd.DataFrame:
    """Tabla 2: detalle por línea, filtrando STMPDH_TIPPRO == 'SAFED'."""
    if df.empty:
        return pd.DataFrame(columns=[
            "Cliente", "Protocolo Serie LF", "Protocolo",
            "Ancho", "Largo", "M2", "Serie",
            "Comprobante", "Descripción"
        ])
    sub = df[df["TipoProducto"].astype(str).str.upper() == "SAFED"].copy()
    if "Protocolo" not in sub.columns:
        sub["Protocolo"] = ""
    if "Protocolo Serie LF" not in sub.columns:
        sub["Protocolo Serie LF"] = ""
    sub["Protocolo"] = sub["Protocolo"].fillna("").astype(str)
    sub["Protocolo Serie LF"] = sub["Protocolo Serie LF"].fillna("").astype(str)
    sub["M2"] = pd.to_numeric(sub["#m2"], errors="coerce").map(
        lambda x: "" if pd.isna(x) else f"{int(round(x))}"
    )
    sub["Serie"] = sub["Serie"].map(format_serie)

    sub = sub.rename(columns={
        "Código de Cliente": "Cliente",
        "Descripcion2": "Descripción",
        "#Comprobante": "Comprobante",
    })
    return sub[[
        "Cliente", "Protocolo Serie LF", "Protocolo",
        "Ancho", "Largo", "M2", "Serie",
        "Comprobante", "Descripción",
    ]].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Mock para probar pipeline sin DB
# ---------------------------------------------------------------------------
def _mock_dataframe() -> pd.DataFrame:
    base = {
        "Código de Cliente": "001",
        "CodFor": "RX0018",
        "NroFor": "00012345",
        "#Comprobante": "RX0018-00012345",
        "Fecha": datetime.now(),
        "Observaciones": "Mock - sin observaciones",
        "Número OC": "OC-9999",
        "Razón Social": "Cliente Demo SA",
    }
    rows = [
        {**base, "Producto": "ART-A", "_Descrp": "OPP TC WHITE GLOSS",
         "_Adhesi": "AP903 PLUS", "_Prolin": "WG55",
         "TipoProducto": "SAFED",
         "Ancho": 1000, "Largo": 2000, "#m2": 2.0, "Serie": "S-001"},
        {**base, "Producto": "ART-B", "_Descrp": "OPP TC CLEAR GLOSS",
         "_Adhesi": "PF1", "_Prolin": "WG55",
         "TipoProducto": "SAFED",
         "Ancho": 1200, "Largo": 2500, "#m2": 3.0, "Serie": "S-002"},
        {**base, "NroFor": "00012346", "#Comprobante": "RX0018-00012346",
         "Producto": "ART-C", "_Descrp": "OPP TC OTHER",
         "_Adhesi": "PF2", "_Prolin": "WG55",
         "TipoProducto": "OTRO",
         "Ancho": 800, "Largo": 1500, "#m2": 1.2, "Serie": "S-003"},
    ]
    return pd.DataFrame(rows)
