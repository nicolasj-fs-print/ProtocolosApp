"""Lectura de `app_2_db`, la base de la app de Seguimiento de Órdenes.

De acá salen las DOS cosas que este bot leía de un Excel:

    quién recibe protocolo y a qué mails   `v_clientes_protocolo`    (2026-09-04)
    qué PR#### le toca a cada material     `v_protocolos_ingresos`   (2026-09-11)

La segunda es la que jubila `Protocolos x Ingreso OK.xlsx`, que era la última planilla de la
que dependía una corrida.

── Por qué se mudó ──────────────────────────────────────────────────────

Ese dato ya vivía en las columnas `protocolos` / `mail_protocolos` de la tabla `clientes` de
la app, sin que nadie las leyera: estaba en dos lados y había divergido. Medido el
2026-09-04: **13 marcados en el Excel contra 11 en la base**, dos clientes sin marcar y a uno
le faltaba un destinatario. Ahora la app es el único lugar donde se edita —desde
`/despacho/protocolos`, con el botón "Clientes con protocolo"— y este módulo es el que hace
que el bot vea esas ediciones.

⚠️ **Solo se LEE, y solo una VISTA.** El rol `agente_dsk` tiene SELECT sobre
`public.v_clientes_protocolo`, que expone tres columnas de los clientes marcados. **No puede
leer `clientes`** — verificado conectándose con él: `permission denied for table clientes`.
Eso importa porque esa tabla tiene CUIT, direcciones, condición de venta y las seis columnas
de listas de precios de los tres negocios, y esta credencial vive dentro de un bundle de
PyInstaller en otra máquina.

── Config ───────────────────────────────────────────────────────────────

Las mismas variables que ya usa el espejo del tracking (`app/sharepoint/lists.py`). Si falta
cualquiera, `esta_configurado()` da False y el llamador cae al Excel.

    APP2_PG_HOST      10.250.2.5
    APP2_PG_PORT      5432
    APP2_PG_DB        app_2_db
    APP2_PG_USER      agente_dsk
    APP2_PG_PASSWORD  la que generó server/db/preparar-usuario-protocolos.mjs
"""
from __future__ import annotations

import os

from ..utils.logger import get_logger

log = get_logger(__name__)

#: Los clientes marcados. La vista ya filtra `protocolos = true`.
_SQL_CLIENTES = """
    select codigo, nombre, mail_protocolos
      from public.v_clientes_protocolo
     order by codigo
"""


def _config() -> dict | None:
    cfg = {
        "host": os.getenv("APP2_PG_HOST"),
        "port": os.getenv("APP2_PG_PORT", "5432"),
        "dbname": os.getenv("APP2_PG_DB", "app_2_db"),
        "user": os.getenv("APP2_PG_USER"),
        "password": os.getenv("APP2_PG_PASSWORD"),
    }
    if not cfg["host"] or not cfg["user"] or not cfg["password"]:
        return None
    return cfg


def esta_configurado() -> bool:
    """True si hay con qué conectarse. No prueba la conexión."""
    return _config() is not None


def leer_clientes_protocolo() -> list[dict]:
    """Los clientes que reciben protocolo, como lista de dicts.

    ⚠️ **TIRA** si no puede leer, y es a propósito: devolver una lista vacía se vería como
    "ningún cliente recibe protocolo" —o sea, el bot no manda nada y no avisa—, que es
    exactamente el modo de falla que este proyecto ya pagó con las tablas espejo del sync.
    Quien llama decide qué hacer con la excepción; hoy `data_service` cae al Excel y grita.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError(
            "app_2_db no está configurada (faltan APP2_PG_HOST / APP2_PG_USER / "
            "APP2_PG_PASSWORD en el .env)."
        )

    import psycopg2

    conn = psycopg2.connect(
        host=cfg["host"],
        port=int(cfg["port"]),
        dbname=cfg["dbname"],
        user=cfg["user"],
        password=cfg["password"],
        # Sin timeout, un server que acepta la conexión y no contesta cuelga el bot entero.
        connect_timeout=8,
        # El tráfico cruza la LAN: se cifra. `require` no valida el certificado, que es lo
        # correcto acá porque el cert es para apps.fs-print.com y nos conectamos por IP.
        sslmode="require",
        application_name="ProtocolosBot/clientes",
    )
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(_SQL_CLIENTES)
                filas = [
                    {
                        "codigo": (r[0] or "").strip(),
                        "nombre": (r[1] or "").strip(),
                        "mails": (r[2] or "").strip(),
                    }
                    for r in cur.fetchall()
                ]
    finally:
        conn.close()

    log.info("app_2_db: %d cliente(s) con protocolo", len(filas))
    return filas


#: El mapeo (producto de trazabilidad, LF) -> PR####.
#:
#: ⚠️⚠️ El `order by orden` NO es cosmético y NO se puede sacar. Una clave `(producto, LF)`
#: puede tener VARIOS certificados —25 medidas el 2026-09-11, con protocolos distintos: un
#: ingreso puede traer el mismo producto llegado en varias Notas Fiscales y cada una tiene el
#: suyo—. `orden` es el número de fila del Excel, así que ordenar por él deja las filas
#: exactamente como venían de `pd.read_excel`. Sin eso, el orden lo decide el planificador de
#: Postgres y el PDF podría listar los certificados de un rollo distinto en cada corrida.
_SQL_INGRESOS = """
    select producto, lf, protocolo
      from public.v_protocolos_ingresos
     order by orden, protocolo
"""


def leer_ingresos_protocolo() -> list[dict]:
    """El mapeo material -> protocolo, como lista de dicts.

    ⚠️ **TIRA** si no puede leer, por el mismo motivo que `leer_clientes_protocolo()`:
    devolver una lista vacía se vería como "ningún rollo tiene protocolo" —el bot mandaría
    los PDF sin un solo certificado adjunto y sin avisar—, que es el modo de falla que este
    proyecto ya pagó. Quien llama decide; hoy `ingresos_service` cae al Excel y grita.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError(
            "app_2_db no está configurada (faltan APP2_PG_HOST / APP2_PG_USER / "
            "APP2_PG_PASSWORD en el .env)."
        )

    import psycopg2

    conn = psycopg2.connect(
        host=cfg["host"],
        port=int(cfg["port"]),
        dbname=cfg["dbname"],
        user=cfg["user"],
        password=cfg["password"],
        connect_timeout=8,
        sslmode="require",
        application_name="ProtocolosBot/ingresos",
    )
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(_SQL_INGRESOS)
                filas = [
                    {
                        "producto": (r[0] or "").strip(),
                        "lf": (r[1] or "").strip(),
                        "protocolo": (r[2] or "").strip(),
                    }
                    for r in cur.fetchall()
                ]
    finally:
        conn.close()

    log.info("app_2_db: %d fila(s) de ingresos (producto, LF -> protocolo)", len(filas))
    return filas
