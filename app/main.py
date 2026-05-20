"""Entrypoint de la aplicación."""
from __future__ import annotations

import sys
import traceback
import warnings
from tkinter import messagebox

# Silencia el aviso de pandas cuando `read_sql` recibe un pyodbc connection
# en lugar de SQLAlchemy. Es informativo, NO un error: nuestras queries son
# parametrizadas y funcionan bien. Migrar a SQLAlchemy sería un cambio mayor
# sin beneficio real para este caso de uso (queries de lectura puntuales).
warnings.filterwarnings(
    "ignore",
    message="pandas only supports SQLAlchemy connectable.*",
    category=UserWarning,
)

from .config import get_settings
from .utils.logger import new_run_id, setup_logger


def _run_auto(setup_auth: bool, dry_run: bool) -> int:
    """Modo automático: cron diario sin GUI. Ver app/auto/runner.py."""
    try:
        from .auto.runner import run as auto_run
        return auto_run(setup_auth=setup_auth, dry_run=dry_run)
    except Exception:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        return 1


def _run_gui() -> int:
    """Modo manual (interactivo): la GUI customtkinter de siempre."""
    settings = get_settings()
    try:
        # Verificación de ODBC Driver 18 antes de cualquier ventana grande.
        # Si falta, ofrecemos instalarlo desde el .msi embebido.
        from .utils.odbc_check import ensure_odbc_driver
        # Necesitamos un root tk efímero para que aparezcan los messagebox.
        import tkinter as _tk
        _root = _tk.Tk()
        _root.withdraw()
        try:
            if not ensure_odbc_driver():
                _root.destroy()
                return 1
        finally:
            _root.destroy()

        # En backend SharePoint, primero mostramos splash con auth + warm-up
        # de Excels. La GUI principal arranca con todo cacheado → primer Buscar
        # instantáneo.
        if settings.storage_backend == "sharepoint":
            from .gui.splash import SplashWindow
            splash = SplashWindow()
            ok = splash.run()
            if not ok:
                return 1

        from .gui.main_window import MainWindow
        app = MainWindow()
        app.mainloop()
        return 0
    except Exception:
        tb = traceback.format_exc()
        try:
            messagebox.showerror("Error fatal", tb)
        except Exception:
            print(tb, file=sys.stderr)
        return 1


def main() -> int:
    new_run_id()
    settings = get_settings()
    setup_logger(level=settings.log_level)

    args = sys.argv[1:]
    auto_mode = "--auto" in args
    setup_auth = "--setup-auth" in args
    dry_run = "--dry-run" in args

    # --setup-auth implica modo auto (login interactivo único, sin GUI).
    if setup_auth and not auto_mode:
        auto_mode = True

    if auto_mode:
        return _run_auto(setup_auth=setup_auth, dry_run=dry_run)
    return _run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
