"""Tracking de usuarios que usan la app.

Mantiene un Excel en SharePoint (`SP_USUARIOS_FILE`) con:
  Email | Nombre | Primer login | Último login | Sesiones | Ejecuciones

- "Sesiones" se incrementa al loguearse (1 vez por arranque de la app).
- "Ejecuciones" se incrementa al terminar una generación de PDFs.

Best-effort: si Graph falla, se loguea y se continúa (no rompe la app).
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Optional

import pandas as pd
import requests

from ..config import get_settings
from ..utils.logger import get_logger
from .auth import get_token
from .client import GRAPH, SharePointError, get_app_client

log = get_logger(__name__)

COLUMNS = ["Email", "Nombre", "Primer login", "Último login", "Sesiones", "Ejecuciones"]
_DT_FMT = "%Y-%m-%d %H:%M:%S"


# ---------------------------------------------------------------------------
# Graph /me
# ---------------------------------------------------------------------------
def get_current_user() -> tuple[str, str]:
    """Devuelve (email, displayName) del usuario logueado."""
    token = get_token(interactive=False) if False else get_token()
    resp = requests.get(
        f"{GRAPH}/me",
        headers={"Authorization": f"Bearer {token}"},
        timeout=15,
    )
    if not resp.ok:
        raise SharePointError(f"GET /me falló: {resp.status_code} {resp.text}")
    data = resp.json()
    email = data.get("mail") or data.get("userPrincipalName") or ""
    name = data.get("displayName") or ""
    return str(email), str(name)


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------
def _load_users_xlsx() -> pd.DataFrame:
    s = get_settings()
    if not s.sp_usuarios_file:
        return pd.DataFrame(columns=COLUMNS)
    try:
        client = get_app_client()
        if not client.file_exists(s.sp_usuarios_file):
            return pd.DataFrame(columns=COLUMNS)
        raw = client.download_file(s.sp_usuarios_file)
        df = pd.read_excel(io.BytesIO(raw), engine="openpyxl", dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        # Asegurar todas las columnas esperadas
        for c in COLUMNS:
            if c not in df.columns:
                df[c] = ""
        return df[COLUMNS].fillna("")
    except Exception as e:
        log.warning("No se pudo leer usuarios.xlsx (%s) → arranco vacío.", e)
        return pd.DataFrame(columns=COLUMNS)


def _save_users_xlsx(df: pd.DataFrame) -> None:
    s = get_settings()
    if not s.sp_usuarios_file:
        return
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    try:
        get_app_client().upload_file(s.sp_usuarios_file, buf.getvalue(), overwrite=True)
    except Exception as e:
        log.warning("No se pudo subir usuarios.xlsx: %s", e)


def _upsert(df: pd.DataFrame, email: str, name: str, *,
            inc_sesiones: int = 0, inc_ejecuciones: int = 0) -> pd.DataFrame:
    if not email:
        return df
    now_str = datetime.now().strftime(_DT_FMT)
    email_lower = email.strip().lower()
    df["_email_lower"] = df["Email"].astype(str).str.strip().str.lower()
    mask = df["_email_lower"] == email_lower

    if mask.any():
        idx = df.index[mask][0]
        df.at[idx, "Nombre"] = name or df.at[idx, "Nombre"]
        df.at[idx, "Último login"] = now_str
        prev_ses = pd.to_numeric(df.at[idx, "Sesiones"], errors="coerce")
        prev_ses = 0 if pd.isna(prev_ses) else int(prev_ses)
        prev_eje = pd.to_numeric(df.at[idx, "Ejecuciones"], errors="coerce")
        prev_eje = 0 if pd.isna(prev_eje) else int(prev_eje)
        df.at[idx, "Sesiones"] = str(prev_ses + inc_sesiones)
        df.at[idx, "Ejecuciones"] = str(prev_eje + inc_ejecuciones)
    else:
        df = pd.concat([df, pd.DataFrame([{
            "Email": email,
            "Nombre": name,
            "Primer login": now_str,
            "Último login": now_str,
            "Sesiones": str(inc_sesiones),
            "Ejecuciones": str(inc_ejecuciones),
            "_email_lower": email_lower,
        }])], ignore_index=True)

    df = df.drop(columns=["_email_lower"])
    return df


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
_CURRENT_USER: tuple[str, str] | None = None


def register_login() -> None:
    """Registra/actualiza el login del usuario actual. Best-effort."""
    global _CURRENT_USER
    try:
        email, name = get_current_user()
        _CURRENT_USER = (email, name)
        log.info("Usuario logueado: %s (%s)", email, name)
        df = _load_users_xlsx()
        df = _upsert(df, email, name, inc_sesiones=1)
        _save_users_xlsx(df)
    except Exception as e:
        log.warning("register_login falló: %s", e)


def register_run(executions: int = 1) -> None:
    """Incrementa el contador de ejecuciones del usuario actual."""
    if _CURRENT_USER is None:
        return
    email, name = _CURRENT_USER
    try:
        df = _load_users_xlsx()
        df = _upsert(df, email, name, inc_ejecuciones=executions)
        _save_users_xlsx(df)
    except Exception as e:
        log.warning("register_run falló: %s", e)


def current_user() -> tuple[str, str] | None:
    return _CURRENT_USER
