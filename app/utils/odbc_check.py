"""Verificación de ODBC Driver 18 for SQL Server.

Si no está instalado en la máquina, ofrece instalarlo desde el .msi embebido
en el bundle (tools/msodbcsql18.msi).
"""
from __future__ import annotations

import subprocess
from tkinter import messagebox

from .logger import get_logger
from .paths import resource_path

log = get_logger(__name__)

REQUIRED_DRIVER = "ODBC Driver 18 for SQL Server"
MSI_RELATIVE = "tools/msodbcsql18.msi"


def is_odbc_driver_installed() -> bool:
    """True si el driver ODBC requerido está en la lista del sistema."""
    try:
        import pyodbc
        drivers = [d.strip() for d in pyodbc.drivers()]
        return REQUIRED_DRIVER in drivers
    except Exception as e:
        log.warning("No se pudo listar drivers ODBC: %s", e)
        return False


def install_odbc_driver() -> bool:
    """Lanza el msiexec con el .msi embebido. Bloquea hasta que termine.
    Devuelve True si la instalación fue exitosa (driver detectable después)."""
    msi_path = resource_path(MSI_RELATIVE)
    if not msi_path.exists():
        log.error("MSI no encontrado en el bundle: %s", msi_path)
        return False

    log.info("Instalando ODBC Driver 18: %s", msi_path)
    try:
        # /passive: progress bar sin prompts. /norestart: no reinicia Windows.
        # IACCEPTMSODBCSQLLICENSETERMS=YES: requerido por Microsoft.
        proc = subprocess.run(
            [
                "msiexec",
                "/i", str(msi_path),
                "/passive",
                "/norestart",
                "IACCEPTMSODBCSQLLICENSETERMS=YES",
            ],
            check=False,
        )
        log.info("msiexec terminó con código %d", proc.returncode)
    except Exception as e:
        log.exception("Error ejecutando msiexec: %s", e)
        return False

    return is_odbc_driver_installed()


def ensure_odbc_driver() -> bool:
    """Verifica el driver. Si falta, pregunta al user si quiere instalarlo.
    Devuelve True si el driver está disponible para continuar."""
    if is_odbc_driver_installed():
        return True

    answer = messagebox.askyesno(
        "Falta ODBC Driver 18 for SQL Server",
        f"Esta aplicación necesita 'ODBC Driver 18 for SQL Server' y "
        f"no está instalado en este equipo.\n\n"
        f"¿Querés instalarlo ahora?\n\n"
        f"(Va a aparecer la barra de progreso de Windows. Puede pedirte permisos de administrador.)",
    )
    if not answer:
        messagebox.showwarning(
            "Driver no instalado",
            "La app no puede continuar sin el driver ODBC.\n"
            "Podés instalarlo manualmente con tools\\msodbcsql18.msi y reiniciar.",
        )
        return False

    ok = install_odbc_driver()
    if ok:
        messagebox.showinfo(
            "Instalación OK",
            "Driver ODBC instalado correctamente. La app va a continuar.",
        )
        return True

    messagebox.showerror(
        "Instalación falló",
        "No se pudo instalar el driver ODBC automáticamente.\n"
        "Probá instalarlo manualmente con tools\\msodbcsql18.msi y reiniciar la app.",
    )
    return False
