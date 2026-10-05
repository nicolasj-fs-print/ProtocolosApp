"""Auth contra Microsoft Graph (MSAL public client + token cache).

Soporta múltiples token caches (uno por cuenta) para que el bot principal
(--auto) y el mailbot (--mailbot, logueado con AGENTE) NO se pisen entre sí.
Cada cache vive en %APPDATA%\\ProtocolosApp\\token_{name}.cache (o token.cache
para el default).
"""
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

# Scopes adicionales que se piden ON-DEMAND cuando un flujo los necesita.
# Mantener fuera del SCOPES default evita pedir admin consent en flujos que
# no los necesitan (ej. la GUI manual).
MAIL_SEND_SCOPE = "Mail.Send"
MAIL_READWRITE_SCOPE = "Mail.ReadWrite"

# Estado por nombre de cache. Default `"default"` es la cuenta del bot/usuario.
# El mailbot usa `"agente"` para no pisar tokens entre cuentas distintas.
_BUILD_LOCK = threading.Lock()
_CACHE_LOCK = threading.Lock()
_APPS: dict[str, msal.PublicClientApplication] = {}
_CACHES: dict[str, msal.SerializableTokenCache] = {}


class AuthError(RuntimeError):
    pass


def _cache_path(cache_name: str = "default") -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    p = Path(base) / "ProtocolosApp"
    p.mkdir(parents=True, exist_ok=True)
    fname = "token.cache" if cache_name == "default" else f"token_{cache_name}.cache"
    return p / fname


def _build_app(cache_name: str = "default") -> msal.PublicClientApplication:
    """Inicialización lazy del MSAL app + cache para el `cache_name` pedido."""
    # Double-check sin lock si ya está construido.
    app = _APPS.get(cache_name)
    if app is not None:
        return app

    with _BUILD_LOCK:
        app = _APPS.get(cache_name)
        if app is not None:
            return app

        s = get_settings()
        if not s.ms_tenant_id or not s.ms_client_id:
            raise AuthError(
                "Faltan MS_TENANT_ID y/o MS_CLIENT_ID en .env (o no se cargó el .env)."
            )

        cache = msal.SerializableTokenCache()
        cache_file = _cache_path(cache_name)
        if cache_file.exists():
            try:
                cache.deserialize(cache_file.read_text(encoding="utf-8"))
            except Exception as e:
                log.warning("Cache de token inválido (%s), se ignora: %s", cache_name, e)

        app = msal.PublicClientApplication(
            client_id=s.ms_client_id,
            authority=f"https://login.microsoftonline.com/{s.ms_tenant_id}",
            token_cache=cache,
        )
        _APPS[cache_name] = app
        _CACHES[cache_name] = cache
        return app


def _persist_cache(cache_name: str = "default") -> None:
    cache = _CACHES.get(cache_name)
    if cache is None or not cache.has_state_changed:
        return
    with _CACHE_LOCK:
        try:
            _cache_path(cache_name).write_text(cache.serialize(), encoding="utf-8")
        except OSError as e:
            log.warning("No se pudo guardar cache de token (%s): %s", cache_name, e)


def get_token(
    interactive: bool = True,
    extra_scopes: list[str] | None = None,
    cache_name: str = "default",
) -> str:
    """Devuelve un access_token válido. Refresh silencioso, fallback interactive.

    `extra_scopes` agrega permisos al set default (ej. `Mail.Send` para el modo auto).
    Si esos scopes requieren admin consent y el usuario no es admin, el login va a
    fallar — por eso solo se piden cuando realmente se necesitan.

    `cache_name` selecciona qué token cache usar. Por defecto `"default"` (la
    cuenta del bot/usuario). El mailbot pasa `"agente"` para no pisar tokens.

    NO toma lock global durante la operación: MSAL.acquire_token_silent y
    acquire_token_interactive son thread-safe internamente.
    """
    scopes = list(SCOPES)
    if extra_scopes:
        for s in extra_scopes:
            if s not in scopes:
                scopes.append(s)

    app = _build_app(cache_name)
    accounts = app.get_accounts()
    result: Optional[dict] = None

    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        if not interactive:
            raise AuthError(
                f"No hay sesión cacheada para '{cache_name}' (interactive=False)."
            )
        log.info("Abriendo browser para login Microsoft (cache=%s)...", cache_name)
        try:
            result = app.acquire_token_interactive(
                scopes=scopes,
                prompt="select_account",
            )
        except Exception as e:
            raise AuthError(f"Login interactivo falló: {e}") from e

    if not result or "access_token" not in result:
        err = (result or {}).get("error_description") or (result or {}).get("error") or "desconocido"
        raise AuthError(f"No se pudo obtener token: {err}")

    _persist_cache(cache_name)
    return result["access_token"]


def ensure_authenticated(
    extra_scopes: list[str] | None = None,
    cache_name: str = "default",
) -> str:
    """Garantiza que haya un token válido (puede abrir browser la primera vez)."""
    return get_token(interactive=True, extra_scopes=extra_scopes, cache_name=cache_name)


def sign_out(cache_name: str = "default") -> None:
    """Borra cache de token (próxima llamada a get_token vuelve a pedir login)."""
    with _BUILD_LOCK:
        try:
            _cache_path(cache_name).unlink(missing_ok=True)
        except OSError:
            pass
        _APPS.pop(cache_name, None)
        _CACHES.pop(cache_name, None)
        log.info("Sesión cerrada (cache='%s' eliminado).", cache_name)
