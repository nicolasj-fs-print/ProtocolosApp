"""Cliente Microsoft Graph para SharePoint (auth interactiva + REST)."""
from .auth import AuthError, ensure_authenticated, get_token, sign_out
from .client import SharePointClient, SharePointError
from .users import current_user, register_login, register_run

__all__ = [
    "AuthError",
    "SharePointClient",
    "SharePointError",
    "ensure_authenticated",
    "get_token",
    "sign_out",
    "current_user",
    "register_login",
    "register_run",
]
