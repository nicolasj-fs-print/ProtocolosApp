"""Entrypoint dedicado del bot (modo automático).

Hardcodea el flag `--auto`. Doble click sobre el .exe resultante (o Task
Scheduler) → corre el modo automático sin GUI.

Pasar args extras sigue funcionando, ej:
  ProtocolosBot.exe --setup-auth   (login interactivo una vez)
  ProtocolosBot.exe --dry-run      (simula sin enviar nada)
"""
from __future__ import annotations

import sys

# Inyectamos --auto al principio para forzar el modo automático.
# Si el user pasa flags extras (--setup-auth, --dry-run), se respetan.
if "--auto" not in sys.argv:
    sys.argv.insert(1, "--auto")

from app.main import main

if __name__ == "__main__":
    raise SystemExit(main())
