from __future__ import annotations

import os
from datetime import date
from pathlib import Path


REQUIRED_COLUMNS = [
    "Código de Cliente",
    "Razón Social",
    "#Comprobante",
    "Número OC",
    "Producto",
    "Descripcion2",
    "Ancho",
    "Largo",
    "#m2",
    "Serie",
    "#Protocolo PDF",
    "Mail Protocolos",
    "Fecha",
    "Observaciones",
    "TipoProducto",
    "NroFor",
]


def validate_date_range(date_from: date, date_to: date) -> None:
    if date_from is None or date_to is None:
        raise ValueError("Las fechas desde/hasta son obligatorias.")
    if date_from > date_to:
        raise ValueError("La fecha 'desde' no puede ser posterior a 'hasta'.")


def ensure_readable(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"{label} no encontrado: {path}. "
            f"Verificá que tengas acceso al sitio SharePoint y la carpeta esté sincronizada."
        )
    if not os.access(path, os.R_OK):
        raise PermissionError(f"Sin permisos de lectura sobre {label}: {path}")


def ensure_writable(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"{label} no encontrada: {path}. "
            f"Pedile acceso al sitio SharePoint VentasPowerBI o sincronizalo localmente."
        )
    if not os.access(path, os.W_OK):
        raise PermissionError(f"Sin permisos de escritura sobre {label}: {path}")


def validate_dataframe_columns(df, required: list[str] = REQUIRED_COLUMNS) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Faltan columnas en el resultado: " + ", ".join(missing) +
            f". Columnas presentes: {list(df.columns)}"
        )
