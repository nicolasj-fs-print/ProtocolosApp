"""Entrypoint para PyInstaller — importa el paquete `app` con imports absolutos.

Necesario porque `app/main.py` usa relative imports (`from .config import ...`)
que solo funcionan cuando se corre como módulo (`python -m app.main`), no como
script directo. PyInstaller arranca el archivo como script, así que necesita
este wrapper.
"""
from __future__ import annotations

import sys

from app.main import main


if __name__ == "__main__":
    sys.exit(main())
