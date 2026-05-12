"""Splash screen al iniciar la app — ejecuta auth + warm-up de Excels.

Bloquea hasta que termina la carga (o falla). Si todo OK, abre la MainWindow.
"""
from __future__ import annotations

import queue
import threading
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image

from ..config import get_settings
from ..utils.logger import get_logger
from ..utils.paths import resource_path

log = get_logger(__name__)


class SplashWindow(ctk.CTk):
    """Mini ventana de carga inicial. Mainloop propio."""

    def __init__(self):
        super().__init__()
        self.title("Generador de Protocolos de Calidad")
        self.geometry("460x220")
        self.resizable(False, False)
        try:
            ico = resource_path("assets/Icono.ico")
            if ico.exists():
                self.iconbitmap(default=str(ico))
        except Exception as e:
            log.warning("No se pudo aplicar Icono.ico: %s", e)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.failed = False
        self._events: queue.Queue = queue.Queue()

        # Centrar
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (sw - 460) // 2
        y = (sh - 220) // 2
        self.geometry(f"460x220+{x}+{y}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Logo FS
        try:
            img = Image.open(resource_path("assets/logo_fs.png"))
            ratio = img.width / max(img.height, 1)
            h = 56
            self._logo = ctk.CTkImage(
                light_image=img, dark_image=img,
                size=(int(h * ratio), h),
            )
            ctk.CTkLabel(self, image=self._logo, text="").grid(
                row=0, column=0, pady=(20, 4)
            )
        except Exception as e:
            log.warning("No se pudo cargar logo splash: %s", e)

        ctk.CTkLabel(
            self, text="Generador de Protocolos de Calidad",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=1, column=0, pady=(0, 4))

        self._lbl_status = ctk.CTkLabel(
            self, text="Iniciando...",
            font=ctk.CTkFont(size=12),
        )
        self._lbl_status.grid(row=2, column=0, pady=(8, 8), padx=20)

        self._progress = ctk.CTkProgressBar(self, mode="indeterminate", width=380)
        self._progress.grid(row=3, column=0, pady=(0, 20))
        self._progress.start()

        # Lanzar el worker en thread daemon
        threading.Thread(target=self._load_worker, daemon=True).start()
        self.after(100, self._poll)

    @staticmethod
    def _ensure_folder_with_timeout(client, path: str, timeout: int = 20) -> bool:
        """Llama client.ensure_folder(path) con timeout duro: si tarda más,
        loggea y sigue. Evita que el splash se cuelgue indefinidamente."""
        result = {"ok": False, "err": None}
        def runner():
            try:
                client.ensure_folder(path)
                result["ok"] = True
            except Exception as e:
                result["err"] = str(e)
        t = threading.Thread(target=runner, daemon=True)
        t.start()
        t.join(timeout=timeout)
        if t.is_alive():
            log.warning("ensure_folder(%s): timeout %ds — sigue corriendo en bg", path, timeout)
            return False
        if result["err"]:
            log.warning("ensure_folder(%s): error %s", path, result["err"])
            return False
        return True

    def _set_status(self, text: str) -> None:
        try:
            self._lbl_status.configure(text=text)
        except Exception:
            pass

    def _post(self, kind: str, msg: str = "") -> None:
        self._events.put((kind, msg))

    def _load_worker(self) -> None:
        """Test SQL + Auth M365 + resolver site/drive + pre-crear Reportes Finales + warm-up Excels.

        Pre-creamos SOLO la carpeta crítica (Reportes Finales). Las demás
        (Logs / Ejecuciones / Usuarios) se crean lazy en sus uploads daemon —
        si tardan o fallan no bloquean al user.
        """
        try:
            # 1) Test rápido de conexión SQL — 1 intento, 5s timeout.
            self._post("status", "Validando conexión SQL...")
            from ..db.connection import quick_test_connection
            ok, msg = quick_test_connection(timeout=5)
            if not ok:
                self._post(
                    "error",
                    "No se pudo establecer conexión a SQL.\n\n"
                    "Verificá:\n"
                    "  • Que tu IP esté autorizada en el servidor SQL.\n"
                    "  • Que el ODBC Driver 18 esté instalado.\n"
                    "  • Que las credenciales en .env sean correctas.\n\n"
                    f"Detalle técnico: {msg}",
                )
                return

            self._post("status", "Conectando con Microsoft 365...")
            from ..sharepoint.auth import ensure_authenticated
            from ..sharepoint.client import get_client
            ensure_authenticated()
            client = get_client()
            client.site_id()
            client.drive_id()

            settings = get_settings()

            self._post("status", "Verificando carpeta de salida...")
            if settings.sp_output_folder:
                try:
                    self._ensure_folder_with_timeout(client, settings.sp_output_folder, timeout=20)
                except Exception as e:
                    log.warning("ensure_folder(output) falló: %s", e)

            from ..services.data_service import _sp_get_or_download

            if settings.sp_clients_file:
                self._post("status", "Cargando lista de clientes...")
                try:
                    _sp_get_or_download(settings.sp_clients_file)
                except Exception as e:
                    log.warning("warm-up clients falló: %s", e)

            if settings.sp_ingresos_file:
                self._post("status", "Cargando ingresos...")
                try:
                    _sp_get_or_download(settings.sp_ingresos_file)
                except Exception as e:
                    log.warning("warm-up ingresos falló: %s", e)

            # Pre-popular el index de protocolos para que cualquier "Generar"
            # sea instantáneo (sin re-listar la carpeta).
            self._post("status", "Indexando protocolos...")
            try:
                from ..services.protocol_service import ProtocolFinder
                ProtocolFinder().index()
            except Exception as e:
                log.warning("warm-up protocol index falló: %s", e)

            # Pre-crear carpetas de telemetría (Logs / Ejecuciones / Usuarios)
            # con timeout, así los uploads daemon posteriores son rápidos.
            self._post("status", "Preparando carpetas de telemetría...")
            for folder in (
                settings.sp_logs_folder,
                settings.sp_ejecuciones_folder,
            ):
                if folder:
                    self._ensure_folder_with_timeout(client, folder, timeout=15)
            if settings.sp_usuarios_file:
                parent = "/".join(settings.sp_usuarios_file.strip("/").split("/")[:-1])
                if parent:
                    self._ensure_folder_with_timeout(client, parent, timeout=15)

            self._post("done", "Listo.")
        except Exception as e:
            log.exception("Error en splash worker")
            self._post("error", str(e))

    def _poll(self) -> None:
        try:
            while True:
                kind, msg = self._events.get_nowait()
                if kind == "status":
                    self._set_status(msg)
                elif kind == "done":
                    self._progress.stop()
                    self.after(150, self.destroy)
                    return
                elif kind == "error":
                    self._progress.stop()
                    self.failed = True
                    messagebox.showerror(
                        "Error de conexión",
                        f"No se pudo conectar a SharePoint:\n\n{msg}\n\n"
                        f"Cerrá la app y verificá tu conexión / credenciales.",
                        parent=self,
                    )
                    self.destroy()
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def run(self) -> bool:
        """Bloquea hasta que termine la carga. Devuelve True si OK, False si falló."""
        self.mainloop()
        return not self.failed
