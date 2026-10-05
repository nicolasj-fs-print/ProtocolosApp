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


# ===========================================================================
# Escritura del Excel Protocolos x Ingreso (MailBot Fedrigoni, Iteración 8)
# ===========================================================================
# Constantes para los valores especiales que escribe / detecta el mailbot.
LF_PENDIENTE = "No ingresado a softland"
LF_ESCALADO = "⚠ Sin LF - escalado"

# Headers esperados en la fila 1 del Excel (case-insensitive, trim).
_HEADER_PACKING_LIST = "packing list"
_HEADER_FC_FEDRIGONI = "fc fedrigoni"
_HEADER_FECHA = "fecha"
_HEADER_LF = "lf"
_HEADER_PRODUCTO = "producto"
_HEADER_PROTOCOLO = "protocolo"


def _normalize_header(s) -> str:
    return str(s or "").replace("\xa0", " ").strip().lower()


class IngresosExcelWriter:
    """Context manager para escribir al Excel Protocolos x Ingreso via Graph Excel API.

    Abre una `ExcelSession` con `persist_changes=True`, lee la fila 1 para mapear
    headers a columnas dinámicamente, y expone métodos de lectura/escritura
    puntual sin tener que descargar/subir el archivo entero.

    Uso:
        with IngresosExcelWriter() as w:
            if not w.packing_list_exists("93200"):
                pr = w.next_protocolo()       # ej. "PR0401"
                w.append_row({
                    "packing_list": "93200",
                    "fc_fedrigoni": "93200",
                    "fecha": "22/05/2026",
                    "lf": "LF - 12345 - PRODUCT",
                    "producto": "B1147",
                    "protocolo": pr,
                })
    """

    def __init__(self):
        from ..sharepoint.client import get_client
        from ..sharepoint.excel_api import ExcelSession

        s = get_settings()
        if s.storage_backend != "sharepoint":
            raise RuntimeError(
                "IngresosExcelWriter solo funciona con storage_backend=sharepoint."
            )
        if not s.sp_ingresos_file:
            raise RuntimeError("SP_INGRESOS_FILE vacío en .env.")

        client = get_client()
        item = client.get_item(s.sp_ingresos_file)
        if not item:
            raise FileNotFoundError(
                f"No se encontró el Excel de ingresos en SharePoint: {s.sp_ingresos_file}"
            )

        self._client = client
        self._item_id = item["id"]
        # Si INGRESOS_SHEET está vacío en .env, se resuelve dinámicamente a la
        # primera hoja del workbook en __enter__ (NO asumimos "Sheet1").
        self._sheet = (s.ingresos_sheet or "").strip()
        self._session_cm = ExcelSession(client, self._item_id, persist_changes=True)
        self._session = None
        self._headers: dict[str, int] = {}  # header_normalized → col_index (1-based)
        # Offset que se incrementa cada vez que se "reserva" un PR sin escribir
        # al Excel (ej. modo TEST). En producción, append_row() escribe la fila
        # con ese PR y resetea el offset (el próximo PR se lee del Excel fresh).
        self._pr_reserved: int = 0

    def __enter__(self) -> "IngresosExcelWriter":
        self._session = self._session_cm.__enter__()
        # Si no se configuró sheet name, usar la primera hoja del workbook.
        if not self._sheet or self._sheet in ("0", 0):
            self._sheet = self._resolve_first_sheet_name()
            log.info("Excel sheet resuelta dinámicamente: %r", self._sheet)
        self._headers = self._read_headers()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._session_cm:
            self._session_cm.__exit__(exc_type, exc, tb)
        self._session = None

    # ----------- internals -----------
    def _resolve_first_sheet_name(self) -> str:
        """Cuando ingresos_sheet está vacío, pedimos el listado de worksheets
        y tomamos la primera."""
        from ..sharepoint.client import GRAPH, SharePointError
        url = f"{GRAPH}/sites/{self._client.site_id()}/drive/items/{self._item_id}/workbook/worksheets"
        resp = self._client._request(
            "GET", url,
            extra_headers={"workbook-session-id": self._session.session_id} if self._session and self._session.session_id else None,
        )
        if not resp.ok:
            raise SharePointError(
                f"No se pudieron listar hojas: {resp.status_code} {resp.text}"
            )
        names = [w.get("name") for w in resp.json().get("value", []) if w.get("name")]
        if not names:
            raise SharePointError("El workbook no tiene hojas.")
        return names[0]

    def _read_headers(self) -> dict[str, int]:
        rows = self._session.get_range(self._sheet, "1:1")
        if not rows or not rows[0]:
            raise RuntimeError(f"Hoja '{self._sheet}': fila 1 (headers) vacía.")
        headers: dict[str, int] = {}
        for idx, raw in enumerate(rows[0], start=1):
            norm = _normalize_header(raw)
            if norm and norm not in headers:
                headers[norm] = idx
        required = (
            _HEADER_PACKING_LIST, _HEADER_FC_FEDRIGONI, _HEADER_FECHA,
            _HEADER_LF, _HEADER_PRODUCTO, _HEADER_PROTOCOLO,
        )
        missing = [h for h in required if h not in headers]
        if missing:
            raise RuntimeError(
                f"Headers faltantes en Excel Ingresos: {missing}. "
                f"Headers encontrados: {sorted(headers.keys())}"
            )
        return headers

    def _col_letter(self, header_norm: str) -> str:
        from ..sharepoint.excel_api import col_letter
        return col_letter(self._headers[header_norm])

    def _used_rows(self) -> tuple[int, int]:
        return self._session.used_range_rows(self._sheet)

    # ----------- API pública -----------

    def next_protocolo(self) -> str:
        """Devuelve el siguiente PR correlativo (PR####).

        Lee la columna Protocolo del Excel y calcula max + 1. Además, suma
        `_pr_reserved` para que llamadas SUCESIVAS en la misma sesión sin
        haber escrito al Excel (ej. modo TEST) devuelvan PRs distintos.

        En producción `append_row()` resetea `_pr_reserved` después de escribir,
        así que el próximo `next_protocolo()` lee el Excel fresco.
        """
        L = self._col_letter(_HEADER_PROTOCOLO)
        first_row, last_row = self._used_rows()
        if last_row < 2:
            n = 1
        else:
            rows = self._session.get_range(self._sheet, f"{L}2:{L}{last_row}")
            nums: list[int] = []
            s = get_settings()
            prefix = (s.mailbot_protocol_prefix or "PR").upper()
            for r in rows:
                v = str((r[0] if r else "") or "").strip().upper()
                if not v.startswith(prefix):
                    continue
                try:
                    nums.append(int(v[len(prefix):]))
                except ValueError:
                    continue
            n = (max(nums) + 1) if nums else 1
        # Reservamos este PR — la próxima llamada sin escritura intermedia
        # devuelve el siguiente.
        n += self._pr_reserved
        self._pr_reserved += 1
        s = get_settings()
        return f"{s.mailbot_protocol_prefix}{n:0{s.mailbot_protocol_digits}d}"

    def packing_list_producto_exists(self, packing_list: str, producto: str) -> bool:
        """True si la combinación `Packing List` + `Producto` ya existe.

        Un mismo Packing List puede aparecer varias veces con productos
        distintos (varios certificados / pedidos para un mismo NF). La
        idempotencia se valida sobre AMBAS columnas para no rechazar mails
        legítimos por colisión de PL.
        """
        pl_target = str(packing_list or "").strip()
        prod_target = str(producto or "").strip()
        if not pl_target or not prod_target:
            return False
        L_pl = self._col_letter(_HEADER_PACKING_LIST)
        L_prod = self._col_letter(_HEADER_PRODUCTO)
        first_row, last_row = self._used_rows()
        if last_row < 2:
            return False
        rows_pl = self._session.get_range(self._sheet, f"{L_pl}2:{L_pl}{last_row}")
        rows_prod = self._session.get_range(self._sheet, f"{L_prod}2:{L_prod}{last_row}")
        n = min(len(rows_pl), len(rows_prod))
        for i in range(n):
            pl_v = str((rows_pl[i][0] if rows_pl[i] else "") or "").strip()
            prod_v = str((rows_prod[i][0] if rows_prod[i] else "") or "").strip()
            if pl_v == pl_target and prod_v == prod_target:
                return True
        return False

    def append_row(self, fields: dict) -> int:
        """Agrega una fila al final del usedRange. Devuelve el row_index (1-based).

        `fields` lleva los valores con keys normalizados:
          packing_list, fc_fedrigoni, fecha, lf, producto, protocolo.
        Las celdas se escriben solo en sus columnas correspondientes — las
        otras columnas de la fila quedan vacías (caller debe asegurar que
        las columnas necesarias estén entre las 6 cubiertas, sino editar
        este método).
        """
        from ..sharepoint.excel_api import col_letter

        first_row, last_row = self._used_rows()
        new_row = max(last_row + 1, 2)

        # Mapeo key → header normalizado.
        key_to_header = {
            "packing_list": _HEADER_PACKING_LIST,
            "fc_fedrigoni": _HEADER_FC_FEDRIGONI,
            "fecha": _HEADER_FECHA,
            "lf": _HEADER_LF,
            "producto": _HEADER_PRODUCTO,
            "protocolo": _HEADER_PROTOCOLO,
        }
        # PATCH celda a celda. Es 6 calls pero la performance es aceptable
        # (decenas de ms cada una) — se puede optimizar con un único PATCH
        # del rango contiguo si las 6 columnas están adyacentes, pero como
        # leemos headers dinámicamente no lo garantizamos.
        for key, header_norm in key_to_header.items():
            value = fields.get(key, "")
            if value is None:
                value = ""
            col_idx = self._headers[header_norm]
            L = col_letter(col_idx)
            addr = f"{L}{new_row}"
            self._session.patch_range(self._sheet, addr, [[str(value)]])
        # Después de escribir, reseteamos el offset reservado: el próximo
        # next_protocolo() leerá del Excel fresco (que ya incluye esta fila).
        self._pr_reserved = 0
        return new_row

    def find_pending_lf_rows(self) -> list[dict]:
        """Devuelve la lista de filas con LF == LF_PENDIENTE.

        Cada dict trae: `row_index`, `packing_list`, `fc_fedrigoni`, `fecha`,
        `producto`, `protocolo`. Las usa el mailbot en la Fase A para
        re-chequear contra Trazabilidad.
        """
        first_row, last_row = self._used_rows()
        if last_row < 2:
            return []

        # Leer las 4 columnas de interés en bloque (packing list, fc fedrigoni,
        # fecha, lf, producto, protocolo) — pedimos cada una por separado.
        cols = {
            "packing_list": self._col_letter(_HEADER_PACKING_LIST),
            "fc_fedrigoni": self._col_letter(_HEADER_FC_FEDRIGONI),
            "fecha": self._col_letter(_HEADER_FECHA),
            "lf": self._col_letter(_HEADER_LF),
            "producto": self._col_letter(_HEADER_PRODUCTO),
            "protocolo": self._col_letter(_HEADER_PROTOCOLO),
        }
        data: dict[str, list[str]] = {}
        for key, L in cols.items():
            rows = self._session.get_range(self._sheet, f"{L}2:{L}{last_row}")
            data[key] = [str((r[0] if r else "") or "").strip() for r in rows]

        out: list[dict] = []
        n = len(data["lf"])
        for i in range(n):
            if data["lf"][i] == LF_PENDIENTE:
                out.append({
                    "row_index": i + 2,  # +2 porque arrancamos en fila 2 (1-based)
                    "packing_list": data["packing_list"][i],
                    "fc_fedrigoni": data["fc_fedrigoni"][i],
                    "fecha": data["fecha"][i],
                    "producto": data["producto"][i],
                    "protocolo": data["protocolo"][i],
                })
        return out

    def update_lf(self, row_index: int, new_value: str) -> None:
        """Actualiza la celda LF de la fila indicada (1-based)."""
        L = self._col_letter(_HEADER_LF)
        addr = f"{L}{row_index}"
        self._session.patch_range(self._sheet, addr, [[str(new_value)]])
