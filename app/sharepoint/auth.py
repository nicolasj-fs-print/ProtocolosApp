"""Auth contra Microsoft Graph (MSAL public client + token cache)."""
from __future__ import annotations

import os
import threading
from pathlib import Path
from typing import Optional

import msal

from ..config import get_settings
from ..utils.logger import get_logger

log = get_logger(__name__)

SCOPES = [
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read",
]

# Locks separados:
# - _BUILD_LOCK: solo para inicializar _APP/_CACHE (corto, una vez por sesión).
# - _CACHE_LOCK: solo para escribir el archivo de cache (corto).
# MSAL.PublicClientApplication.acquire_token_* es thread-safe internamente,
# así que NO necesitamos un lock global durante toda la operación de get_token.
_BUILD_LOCK = threading.Lock()
_CACHE_LOCK = threading.Lock()
_APP: msal.PublicClientApplication | None = None
_CACHE: msal.SerializableTokenCache | None = None


class AuthError(RuntimeError):
    pass


def _cache_path() -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    p = Path(base) / "ProtocolosApp"
    p.mkdir(parents=True, exist_ok=True)
    return p / "token.cache"


def _build_app() -> msal.PublicClientApplication:
    """Inicialización lazy del MSAL app + cache. Thread-safe vía _BUILD_LOCK
    (solo para la primera construcción)."""
    global _APP, _CACHE
    # Double-check: si ya está, sin lock.
    if _APP is not None:
        return _APP

    with _BUILD_LOCK:
        if _APP is not None:
            return _APP

        s = get_settings()
        if not s.ms_tenant_id or not s.ms_client_id:
            raise AuthError(
                "Faltan MS_TENANT_ID y/o MS_CLIENT_ID en .env (o no se cargó el .env)."
            )

        _CACHE = msal.SerializableTokenCache()
        cache_file = _cache_path()
        if cache_file.exists():
            try:
                _CACHE.deserialize(cache_file.read_text(encoding="utf-8"))
            except Exception as e:
                log.warning("Cache de token inválido, se ignora: %s", e)

        _APP = msal.PublicClientApplication(
            client_id=s.ms_client_id,
            authority=f"https://login.microsoftonline.com/{s.ms_tenant_id}",
            token_cache=_CACHE,
        )
        return _APP


def _persist_cache() -> None:
    if _CACHE is None or not _CACHE.has_state_changed:
        return
    with _CACHE_LOCK:
        try:
            _cache_path().write_text(_CACHE.serialize(), encoding="utf-8")
        except OSError as e:
            log.warning("No se pudo guardar cache de token: %s", e)


def get_token(interactive: bool = True) -> str:
    """Devuelve un access_token válido. Refresh silencioso, fallback interactive.

    NO toma lock global durante la operación: MSAL.acquire_token_silent y
    acquire_token_interactive son thread-safe internamente. Así múltiples
    threads pueden pedir token concurrentemente sin colgarse uno al otro.
    """
    app = _build_app()
    accounts = app.get_accounts()
    result: Optional[dict] = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        if not interactive:
            raise AuthError("No hay sesión cacheada (interactive=False).")
        log.info("Abriendo browser para login Microsoft...")
        try:
            result = app.acquire_token_interactive(
                scopes=SCOPES,
                prompt="select_account",
            )
        except Exception as e:
            raise AuthError(f"Login interactivo falló: {e}") from e

    if not result or "access_token" not in result:
        err = (result or {}).get("error_description") or (result or {}).get("error") or "desconocido"
        raise AuthError(f"No se pudo obtener token: {err}")

    _persist_cache()
    return result["access_token"]


def ensure_authenticated() -> str:
    """Garantiza que haya un token válido (puede abrir browser la primera vez)."""
    return get_token(interactive=True)


def sign_out() -> None:
    """Borra cache de token (próxima llamada a get_token vuelve a pedir login)."""
    with _BUILD_LOCK:
        global _APP, _CACHE
        try:
            _cache_path().unlink(missing_ok=True)
        except OSError:
            pass
        _APP = None
        _CACHE = None
        log.info("Sesión cerrada (cache de token eliminado).")
