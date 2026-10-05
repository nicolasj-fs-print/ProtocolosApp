"""Wrapper Microsoft Graph para SharePoint Lists.

Usado por el modo automático (Iteración 7) para persistir el tracking de
remitos enviados, así el cron no duplica envíos.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Iterable

from ..config import get_settings
from ..utils.logger import get_logger
from .client import GRAPH, SharePointError, get_app_client

log = get_logger(__name__)


# Estados terminales del tracking.
ESTADO_ENVIADO_CLIENTE = "enviado_cliente"
ESTADO_PENDING_CONTROL = "pending_control"
ESTADO_RESUELTO_MANUAL = "resuelto_manual"
# Cuando el bot detecta que el PDF ya existía en SharePoint (alguien lo generó
# manualmente desde la GUI antes de la corrida) → no remanda mail, solo registra.
ESTADO_ENVIADO_MANUAL = "enviado_manual"
ESTADO_ERROR = "error"

ALL_ESTADOS = (
    ESTADO_ENVIADO_CLIENTE,
    ESTADO_PENDING_CONTROL,
    ESTADO_RESUELTO_MANUAL,
    ESTADO_ENVIADO_MANUAL,
    ESTADO_ERROR,
)


@dataclass
class TrackingEntry:
    item_id: str
    comprobante: str
    cliente: str = ""
    razon_social: str = ""
    fecha_remito: str = ""
    fecha_procesado: str = ""
    estado: str = ""
    mail_destino: str = ""
    pdf_url: str = ""
    items_faltantes: str = ""
    run_id: str = ""


class SharePointLists:
    """Wrapper CRUD sobre SharePoint Lists vía Graph API.

    Reusa el cliente "app" (sitio Desarrollo / ProtocolosApp) — ahí ya viven
    Logs, Ejecuciones, Usuarios, así que el tracking encaja naturalmente.
    """

    def __init__(self):
        self.client = get_app_client()
        self._list_id_cache: dict[str, str] = {}

    # ----------- Helpers internos -----------

    def _site_id(self) -> str:
        return self.client.site_id()

    def _list_url(self, list_id: str) -> str:
        return f"{GRAPH}/sites/{self._site_id()}/lists/{list_id}"

    # ----------- ensure_list_exists -----------

    def list_exists(self, list_name: str) -> bool:
        """True si la SP List ya existe. No la crea.

        Útil para hooks readonly (ej. el de la GUI manual que cierra el ciclo
        pending_control → resuelto_manual): si el bot nunca corrió, la list
        no existe y NO debe crearse desde la GUI.
        """
        if list_name in self._list_id_cache:
            return True
        try:
            url = f"{GRAPH}/sites/{self._site_id()}/lists?$select=id,name,displayName"
            resp = self.client._request("GET", url)
            if not resp.ok:
                return False
            for item in resp.json().get("value", []):
                if item.get("displayName") == list_name or item.get("name") == list_name:
                    self._list_id_cache[list_name] = item["id"]
                    return True
            return False
        except Exception as e:
            log.debug("list_exists(%s) error: %s", list_name, e)
            return False

    def ensure_list_exists(self, list_name: str) -> str:
        """Devuelve el list_id, creando la lista si no existe.

        La estructura de columnas está hardcoded para el tracking de envíos.
        Si la lista ya existe pero le faltan columnas, NO las agrega (asume
        que un admin las definió o que fue creada por una versión previa).
        """
        if list_name in self._list_id_cache:
            return self._list_id_cache[list_name]

        # 1) Buscar la lista existente por displayName.
        url = f"{GRAPH}/sites/{self._site_id()}/lists?$select=id,name,displayName"
        resp = self.client._request("GET", url)
        if not resp.ok:
            raise SharePointError(
                f"No se pudieron listar las SP Lists: {resp.status_code} {resp.text}"
            )
        for item in resp.json().get("value", []):
            if item.get("displayName") == list_name or item.get("name") == list_name:
                lid = item["id"]
                self._list_id_cache[list_name] = lid
                log.info("SharePoint List existente: '%s' (id=%s)", list_name, lid)
                return lid

        # 2) Crear nueva.
        log.info("Creando SharePoint List nueva: '%s'", list_name)
        body = {
            "displayName": list_name,
            "list": {"template": "genericList"},
            "columns": [
                # Title (built-in) → guardamos el #Comprobante.
                {"name": "Cliente", "text": {}},
                {"name": "RazonSocial", "text": {}},
                {"name": "FechaRemito", "text": {}},
                {"name": "FechaProcesado", "text": {}},
                {
                    "name": "Estado",
                    "choice": {"choices": list(ALL_ESTADOS), "displayAs": "dropDownMenu"},
                },
                {"name": "MailDestino", "text": {}},
                {"name": "PdfUrl", "text": {"allowMultipleLines": True}},
                {"name": "ItemsFaltantes", "text": {"allowMultipleLines": True}},
                {"name": "RunId", "text": {}},
            ],
        }
        create_url = f"{GRAPH}/sites/{self._site_id()}/lists"
        cr = self.client._request("POST", create_url, json_body=body)
        if cr.status_code == 403:
            raise SharePointError(
                f"\n\n"
                f"❌ La SharePoint List '{list_name}' NO existe y no hay permisos para "
                f"crearla desde código (403 Access Denied).\n\n"
                f"SOLUCIÓN — creala manualmente UNA VEZ:\n"
                f"  1. Ir al sitio SharePoint de la app (SP_APP_SITE).\n"
                f"  2. + Nueva → Lista → Lista en blanco.\n"
                f"  3. Nombre EXACTO: {list_name}\n"
                f"  4. Agregar estas columnas (todas 'Una sola línea de texto' salvo "
                f"donde se indique):\n"
                f"       - Cliente\n"
                f"       - RazonSocial\n"
                f"       - FechaRemito\n"
                f"       - FechaProcesado\n"
                f"       - Estado  (Opción: enviado_cliente | pending_control | "
                f"resuelto_manual | error)\n"
                f"       - MailDestino\n"
                f"       - PdfUrl\n"
                f"       - ItemsFaltantes  (Varias líneas de texto)\n"
                f"       - RunId\n"
                f"     (La columna 'Título' viene built-in; el bot guarda ahí el "
                f"#Comprobante.)\n"
                f"  5. Volvé a correr el bot.\n"
            )
        if not cr.ok:
            raise SharePointError(
                f"No se pudo crear la SP List '{list_name}': "
                f"{cr.status_code} {cr.text}"
            )
        lid = cr.json()["id"]
        self._list_id_cache[list_name] = lid
        log.info("SharePoint List creada: '%s' (id=%s)", list_name, lid)
        return lid

    # ----------- Query helpers -----------

    @staticmethod
    def _fields_to_entry(item_id: str, fields: dict) -> TrackingEntry:
        return TrackingEntry(
            item_id=item_id,
            comprobante=str(fields.get("Title") or ""),
            cliente=str(fields.get("Cliente") or ""),
            razon_social=str(fields.get("RazonSocial") or ""),
            fecha_remito=str(fields.get("FechaRemito") or ""),
            fecha_procesado=str(fields.get("FechaProcesado") or ""),
            estado=str(fields.get("Estado") or ""),
            mail_destino=str(fields.get("MailDestino") or ""),
            pdf_url=str(fields.get("PdfUrl") or ""),
            items_faltantes=str(fields.get("ItemsFaltantes") or ""),
            run_id=str(fields.get("RunId") or ""),
        )

    def fetch_in_range(
        self,
        list_name: str,
        date_from: date,
        date_to: date,
    ) -> dict[str, TrackingEntry]:
        """Devuelve {comprobante: TrackingEntry} para items con
        FechaRemito entre `date_from` y `date_to` (inclusive).

        Como FechaRemito es text (no Date), filtramos client-side. Para volúmenes
        chicos de tracking (decenas de items por día) es OK.
        """
        list_id = self.ensure_list_exists(list_name)
        url = f"{self._list_url(list_id)}/items?$expand=fields&$top=500"
        out: dict[str, TrackingEntry] = {}

        date_from_s = date_from.isoformat()
        date_to_s = date_to.isoformat()

        next_url = url
        while next_url:
            resp = self.client._request("GET", next_url)
            if not resp.ok:
                raise SharePointError(
                    f"fetch_in_range falló: {resp.status_code} {resp.text}"
                )
            data = resp.json()
            for it in data.get("value", []):
                fields = it.get("fields") or {}
                fecha = str(fields.get("FechaRemito") or "")
                if not fecha:
                    continue
                if date_from_s <= fecha <= date_to_s:
                    entry = self._fields_to_entry(it.get("id", ""), fields)
                    if entry.comprobante:
                        out[entry.comprobante] = entry
            next_url = data.get("@odata.nextLink")

        return out

    def find_by_comprobante(
        self,
        list_name: str,
        comprobante: str,
    ) -> TrackingEntry | None:
        """Busca el item por Title (= #Comprobante). Devuelve None si no existe."""
        list_id = self.ensure_list_exists(list_name)
        # Filtro server-side por Title.
        comp_q = (comprobante or "").replace("'", "''")
        url = (
            f"{self._list_url(list_id)}/items"
            f"?$expand=fields&$filter=fields/Title eq '{comp_q}'&$top=5"
        )
        resp = self.client._request(
            "GET", url,
            extra_headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"},
        )
        if not resp.ok:
            raise SharePointError(
                f"find_by_comprobante falló: {resp.status_code} {resp.text}"
            )
        for it in resp.json().get("value", []):
            fields = it.get("fields") or {}
            if str(fields.get("Title") or "") == comprobante:
                return self._fields_to_entry(it.get("id", ""), fields)
        return None

    # ----------- Espejo en app_2_db -----------

    def _espejo_app2(
        self,
        comprobante: str,
        cliente: str,
        razon_social: str,
        fecha_remito: str,
        estado: str,
        mail_destino: str,
        pdf_url: str,
        items_faltantes: str,
        run_id: str,
    ) -> None:
        """Copia el tracking a app_2_db, la base de la app de Seguimiento de Órdenes.

        Se llama DESPUÉS de que la escritura a SharePoint salió bien, y NUNCA tira: si el
        VPS está caído el bot tiene que seguir mandando los mails igual. Antes de esto el
        bot no dependía de esa app y no puede empezar a depender ahora.

        SharePoint sigue siendo la fuente de la que el bot LEE para saber qué ya procesó
        (`fetch_in_range`). Esto es, por ahora, solo un espejo.

        ⚠️ El usuario `agente_dsk` SOLO puede tocar la tabla `protocolos_envios`: no puede
        leer ninguna otra de la app, ni borrar. Verificado conectándose con él. No usar acá
        `app_2_user`, que es dueño de las 55 tablas.

        Config (si falta cualquiera, no hace nada y no molesta):
            APP2_PG_HOST      10.250.2.5
            APP2_PG_PORT      5432
            APP2_PG_DB        app_2_db
            APP2_PG_USER      agente_dsk
            APP2_PG_PASSWORD  la que genero server/db/preparar-usuario-protocolos.mjs
        """
        import os

        cfg = {
            "host": os.getenv("APP2_PG_HOST"),
            "port": os.getenv("APP2_PG_PORT", "5432"),
            "dbname": os.getenv("APP2_PG_DB", "app_2_db"),
            "user": os.getenv("APP2_PG_USER"),
            "password": os.getenv("APP2_PG_PASSWORD"),
        }
        if not cfg["host"] or not cfg["user"] or not cfg["password"]:
            return

        # Los vacíos van como NULL para que el COALESCE de abajo conserve lo que ya había.
        def v(x):
            x = (x or "").strip()
            return x or None

        faltantes = self._faltantes_a_json(items_faltantes)

        try:
            import psycopg2

            conn = psycopg2.connect(
                host=cfg["host"],
                port=int(cfg["port"]),
                dbname=cfg["dbname"],
                user=cfg["user"],
                password=cfg["password"],
                # Sin timeout, un server que acepta la conexión y no contesta cuelga el bot.
                connect_timeout=8,
                # El tráfico cruza la LAN: se cifra. `require` no valida el certificado,
                # que es lo correcto acá porque el cert es para apps.fs-print.com y nos
                # conectamos por IP.
                sslmode="require",
                application_name="ProtocolosBot",
            )
            try:
                with conn:
                    with conn.cursor() as cur:
                        # El COALESCE replica la misma regla que este upsert usa contra
                        # SharePoint: los campos que vienen vacíos NO pisan lo que ya
                        # estaba. Sin eso, el segundo escribe de un remito (cuando pasa de
                        # pending_control a resuelto_manual, que manda solo el estado)
                        # borraría cliente, razón social, fecha e items_faltantes.
                        #
                        # `estado` y `fecha_procesado` se actualizan SIEMPRE, igual que allá.
                        cur.execute(
                            """
                            insert into public.protocolos_envios
                                (comprobante, cliente, razon_social, fecha_remito,
                                 fecha_procesado, estado, mail_destino, pdf_url,
                                 items_faltantes, run_id, origen, updated_at)
                            values (%s, %s, %s, %s, %s, %s, %s, %s, %s::jsonb, %s, 'bot', now())
                            on conflict (comprobante) do update set
                                cliente         = coalesce(excluded.cliente,         protocolos_envios.cliente),
                                razon_social    = coalesce(excluded.razon_social,    protocolos_envios.razon_social),
                                fecha_remito    = coalesce(excluded.fecha_remito,    protocolos_envios.fecha_remito),
                                mail_destino    = coalesce(excluded.mail_destino,    protocolos_envios.mail_destino),
                                pdf_url         = coalesce(excluded.pdf_url,         protocolos_envios.pdf_url),
                                items_faltantes = coalesce(excluded.items_faltantes, protocolos_envios.items_faltantes),
                                run_id          = coalesce(excluded.run_id,          protocolos_envios.run_id),
                                estado          = excluded.estado,
                                fecha_procesado = excluded.fecha_procesado,
                                origen          = 'bot',
                                updated_at      = now()
                            """,
                            (
                                comprobante, v(cliente), v(razon_social), v(fecha_remito),
                                datetime.now().isoformat(timespec="seconds"), v(estado),
                                v(mail_destino), v(pdf_url), faltantes, v(run_id),
                            ),
                        )
                log.info("Espejo app_2_db OK: %s", comprobante)
            finally:
                conn.close()
        except Exception as e:
            log.warning("Espejo app_2_db falló para %s (no es fatal): %s", comprobante, e)

    @staticmethod
    def _faltantes_a_json(valor: str) -> str | None:
        """`items_faltantes` a algo que entre en una columna jsonb, o None.

        El bot lo manda como JSON serializado, pero en algún caso puede ser texto suelto.
        Si no parsea, se envuelve en vez de perderlo (y de romper el insert entero).
        """
        import json

        s = (valor or "").strip()
        if not s:
            return None
        try:
            json.loads(s)
            return s
        except (ValueError, TypeError):
            return json.dumps({"texto": s})

    # ----------- Upsert -----------

    def upsert(
        self,
        list_name: str,
        comprobante: str,
        *,
        cliente: str = "",
        razon_social: str = "",
        fecha_remito: str = "",
        estado: str = "",
        mail_destino: str = "",
        pdf_url: str = "",
        items_faltantes: str = "",
        run_id: str = "",
    ) -> TrackingEntry:
        """Crea o actualiza una entrada de tracking por #Comprobante.

        Idempotente: si ya existe, hace PATCH a los fields no vacíos. Si no
        existe, hace POST con todos los fields. Devuelve la TrackingEntry
        resultante.
        """
        if estado and estado not in ALL_ESTADOS:
            log.warning("Estado desconocido '%s' (se guarda igual)", estado)

        existing = self.find_by_comprobante(list_name, comprobante)
        list_id = self.ensure_list_exists(list_name)

        fields: dict = {
            "Title": comprobante,
            "Cliente": cliente,
            "RazonSocial": razon_social,
            "FechaRemito": fecha_remito,
            "FechaProcesado": datetime.now().isoformat(timespec="seconds"),
            "Estado": estado,
            "MailDestino": mail_destino,
            "PdfUrl": pdf_url,
            "ItemsFaltantes": items_faltantes,
            "RunId": run_id,
        }
        # En PATCH solo mandamos lo no vacío para no pisar datos previos con strings vacíos.
        if existing is not None:
            patch_fields = {
                k: v for k, v in fields.items()
                if v not in ("", None) or k in ("Estado", "FechaProcesado")
            }
            url = f"{self._list_url(list_id)}/items/{existing.item_id}/fields"
            resp = self.client._request("PATCH", url, json_body=patch_fields)
            if not resp.ok:
                raise SharePointError(
                    f"upsert PATCH falló: {resp.status_code} {resp.text}"
                )
            log.info("Tracking UPDATED: %s → %s", comprobante, estado or "(sin estado)")
            self._espejo_app2(comprobante, cliente, razon_social, fecha_remito,
                              estado, mail_destino, pdf_url, items_faltantes, run_id)
            return self._fields_to_entry(existing.item_id, resp.json())

        # POST nuevo.
        url = f"{self._list_url(list_id)}/items"
        resp = self.client._request("POST", url, json_body={"fields": fields})
        if not resp.ok:
            raise SharePointError(
                f"upsert POST falló: {resp.status_code} {resp.text}"
            )
        item = resp.json()
        item_id = item.get("id", "")
        log.info("Tracking CREATED: %s → %s", comprobante, estado or "(sin estado)")
        self._espejo_app2(comprobante, cliente, razon_social, fecha_remito,
                          estado, mail_destino, pdf_url, items_faltantes, run_id)
        return self._fields_to_entry(item_id, item.get("fields") or fields)


# ----------- Singleton -----------

_LISTS: SharePointLists | None = None


def get_lists() -> SharePointLists:
    global _LISTS
    if _LISTS is None:
        _LISTS = SharePointLists()
    return _LISTS


# ----------- Helpers de alto nivel para el bot -----------

def is_terminal(estado: str) -> bool:
    """True si el estado significa 'no reprocesar más'."""
    return estado in (
        ESTADO_ENVIADO_CLIENTE,
        ESTADO_PENDING_CONTROL,
        ESTADO_RESUELTO_MANUAL,
        ESTADO_ENVIADO_MANUAL,
    )


def filter_unprocessed(
    comprobantes: Iterable[str],
    tracking: dict[str, TrackingEntry],
) -> list[str]:
    """Devuelve los comprobantes que NO tienen un estado terminal en el tracking.

    Los que solo tienen `error` previo se incluyen para reintentar.
    """
    out = []
    for c in comprobantes:
        entry = tracking.get(c)
        if entry is None:
            out.append(c)
        elif not is_terminal(entry.estado):
            out.append(c)
    return out
