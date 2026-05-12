"""Entrypoint de la aplicación."""
from __future__ import annotations

import sys
import traceback
from tkinter import messagebox

from .config import get_settings
from .utils.logger import new_run_id, setup_logger


def main() -> int:
    new_run_id()
    settings = get_settings()
    setup_logger(level=settings.log_level)

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


if __name__ == "__main__":
    raise SystemExit(main())
