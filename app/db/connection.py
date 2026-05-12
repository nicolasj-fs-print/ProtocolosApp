from __future__ import annotations

import time
from contextlib import contextmanager
from typing import Iterator

import pyodbc

from ..config import get_settings
from ..utils.logger import get_logger

log = get_logger(__name__)

_MAX_ATTEMPTS = 3
_BACKOFF_SECONDS = 1.5


@contextmanager
def get_connection(database: str | None = None) -> Iterator[pyodbc.Connection]:
    """Conexión pyodbc con reintentos. `database` permite override del default."""
    settings = get_settings()
    conn_str = settings.odbc_connection_string(database=database)
    last_err: Exception | None = None

    for attempt in range(1, _MAX_ATTEMPTS + 1):
        try:
            conn = pyodbc.connect(conn_str, timeout=15)
            try:
                yield conn
            finally:
                conn.close()
            return
        except pyodbc.Error as e:
            last_err = e
            log.warning("Conexión SQL fallida (intento %d/%d): %s", attempt, _MAX_ATTEMPTS, e)
            if attempt < _MAX_ATTEMPTS:
                time.sleep(_BACKOFF_SECONDS * attempt)

    raise ConnectionError(f"No se pudo conectar a SQL Server tras {_MAX_ATTEMPTS} intentos: {last_err}")


def test_connection() -> tuple[bool, str]:
    try:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT 1")
            cur.fetchone()
        return True, "OK"
    except Exception as e:
        return False, str(e)


def quick_test_connection(timeout: int = 5) -> tuple[bool, str]:
    """Test rápido: 1 solo intento con timeout corto (sin reintentos).
    Pensado para el splash al arrancar la app."""
    settings = get_settings()
    conn_str = settings.odbc_connection_string()
    try:
        conn = pyodbc.connect(conn_str, timeout=timeout)
        try:
            cur = conn.cursor()
            cur.execute("SELECT 1")
            cur.fetchone()
            return True, "OK"
        finally:
            conn.close()
    except Exception as e:
        return False, str(e)
