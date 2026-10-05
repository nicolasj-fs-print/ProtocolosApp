"""Entrypoint dedicado del MailBot Fedrigoni (sin GUI).

Hardcodea el flag `--mailbot`. Doble click sobre el .exe resultante (o Task
Scheduler) → corre el mailbot que procesa los mails de Fedrigoni.

Pasar args extras sigue funcionando, ej:
  MailBot.exe --setup-auth   (login interactivo AGENTE, una vez)
  MailBot.exe --dry-run      (simula sin escribir Excel ni mover mails)
"""
from __future__ import annotations

import sys

# Inyectamos --mailbot al principio para forzar el modo mailbot.
# Si el user pasa flags extras (--setup-auth, --dry-run), se respetan.
if "--mailbot" not in sys.argv:
    sys.argv.insert(1, "--mailbot")

from app.main import main

if __name__ == "__main__":
    raise SystemExit(main())
