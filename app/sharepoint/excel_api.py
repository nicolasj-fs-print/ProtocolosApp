"""Wrapper Microsoft Graph Excel API — edita celdas sin descargar el archivo.

Mucho más rápido que abrir + serializar un .xlsm grande con openpyxl
(~1-3s vs 60-90s para Stock Mendoza.xlsm de 20MB con macros).
"""
from __future__ import annotations

import urllib.parse

from ..utils.logger import get_logger
from .client import GRAPH, SharePointClient, SharePointError

log = get_logger(__name__)


def col_letter(col_num: int) -> str:
    """1→A, 26→Z, 27→AA, 52→AZ, 53→BA, etc."""
    if col_num <= 0:
        raise ValueError(f"col_num inválido: {col_num}")
    result = ""
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result
    return result


class ExcelSession:
    """Context manager para una sesión de Excel via Graph API.

    Crea una workbook session (más rápido que invocar cada llamada sin sesión,
    además permite agrupar cambios). `persist_changes=True` guarda los cambios
    en el archivo real; si es False son temporales.
    """

    def __init__(self, client: SharePointClient, item_id: str, *, persist_changes: bool = True):
        self.client = client
        self.item_id = item_id
        self.persist = persist_changes
        self.session_id: str | None = None

    # --- low-level ---------------------------------------------------------
    def _base(self) -> str:
        return f"{GRAPH}/sites/{self.client.site_id()}/drive/items/{self.item_id}/workbook"

    def _headers(self) -> dict[str, str]:
        return {"workbook-session-id": self.session_id} if self.session_id else {}

    # --- lifecycle ---------------------------------------------------------
    def __enter__(self) -> "ExcelSession":
        url = f"{self._base()}/createSession"
        resp = self.client._request(
            "POST", url, json_body={"persistChanges": self.persist}
        )
        if not resp.ok:
            raise SharePointError(
                f"createSession falló: {resp.status_code} {resp.text}"
            )
        self.session_id = resp.json().get("id")
        log.info("Excel session abierta (persist=%s)", self.persist)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.session_id is None:
            return
        try:
            url = f"{self._base()}/closeSession"
            self.client._request("POST", url, extra_headers=self._headers())
            log.info("Excel session cerrada.")
        except Exception as e:
            log.warning("closeSession falló (no crítico): %s", e)
        self.session_id = None

    # --- range I/O ---------------------------------------------------------
    @staticmethod
    def _ws_path(worksheet: str) -> str:
        # nombres con espacios deben quedar tipo `Clientes` o `'Mi Hoja'`.
        # Graph acepta directamente la forma con espacios URL-encoded.
        return urllib.parse.quote(worksheet, safe="")

    def get_range(self, worksheet: str, address: str) -> list[list]:
        """Devuelve `values` (matriz 2D) del rango pedido.

        `address` formato Excel: "A1:E1", "A:A", "1:1", "A2:A150", etc.
        """
        ws = self._ws_path(worksheet)
        addr = urllib.parse.quote(address, safe="")
        url = f"{self._base()}/worksheets/{ws}/range(address='{addr}')"
        resp = self.client._request("GET", url, extra_headers=self._headers())
        if not resp.ok:
            raise SharePointError(
                f"get_range({worksheet}!{address}) falló: {resp.status_code} {resp.text}"
            )
        return resp.json().get("values", []) or []

    def patch_range(self, worksheet: str, address: str, values: list[list]) -> None:
        """Sobrescribe el rango con `values` (matriz 2D)."""
        ws = self._ws_path(worksheet)
        addr = urllib.parse.quote(address, safe="")
        url = f"{self._base()}/worksheets/{ws}/range(address='{addr}')"
        resp = self.client._request(
            "PATCH", url,
            json_body={"values": values},
            extra_headers=self._headers(),
        )
        if not resp.ok:
            raise SharePointError(
                f"patch_range({worksheet}!{address}) falló: "
                f"{resp.status_code} {resp.text}"
            )

    def used_range_rows(self, worksheet: str) -> tuple[int, int]:
        """Devuelve (first_row, last_row) 1-based del usedRange de la hoja."""
        ws = self._ws_path(worksheet)
        url = (
            f"{self._base()}/worksheets/{ws}/usedRange(valuesOnly=true)"
            f"?$select=rowIndex,rowCount"
        )
        resp = self.client._request("GET", url, extra_headers=self._headers())
        if not resp.ok:
            raise SharePointError(
                f"usedRange({worksheet}) falló: {resp.status_code} {resp.text}"
            )
        data = resp.json()
        row_index = int(data.get("rowIndex", 0) or 0)  # 0-based
        row_count = int(data.get("rowCount", 0) or 0)
        return (row_index + 1, row_index + row_count)
