"""Ventana principal de la app."""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
from PIL import Image
from tkcalendar import DateEntry

from ..config import Settings, get_settings
from ..services import data_service
from ..utils.logger import get_logger
from ..utils.paths import resource_path
from ..utils.validators import ensure_readable, ensure_writable
from ..services import mail_service
from ..services.pdf_service import ComprobanteResult
from ..workers.generation_worker import GenerationWorker, ProgressEvent
from .dialogs import (
    AddClientesDialog,
    AlreadyGeneratedDialog,
    AskMailsDialog,
    MissingProtocolsDialog,
)
from .widgets import DataTable, MultiSelectListbox

log = get_logger(__name__)

APP_TITLE = "Generador de Protocolos de Calidad"
APP_VERSION = "1.0.0"


SUMMARY_COLS = ["Cliente", "Remito", "Fecha", "M2", "Observaciones"]
SUMMARY_WIDTHS = [110, 160, 110, 90, 400]

DETAIL_COLS = [
    "Cliente", "LF - Producto", "Protocolo",
    "Ancho", "Largo", "M2", "Serie",
    "Remito", "Producto"
]
DETAIL_WIDTHS = [120, 200, 160, 60, 60, 60, 150, 120, 280]


def _fix_date_entry_focus_bug(de: DateEntry) -> None:
    """Monkey-patch del `_on_focus_out_cal` de tkcalendar.DateEntry.

    Bug original (ver tkcalendar/dateentry.py línea 248): cuando hacés click
    en las flechas de cambio de mes ◀ ▶, el Calendar interno pierde el focus
    y `focus_get()` puede devolver None. El método original entra al else
    final y hace `withdraw()` → popup cerrado.

    Este patch chequea SIEMPRE primero si el mouse está sobre el popup. Si
    está, re-focusea el calendar y NO cierra. Si está afuera, cierra normal.
    """
    import types

    def _patched(self, _event):
        try:
            x, y = self._top_cal.winfo_pointerxy()
            xc = self._top_cal.winfo_rootx()
            yc = self._top_cal.winfo_rooty()
            w = self._top_cal.winfo_width()
            h = self._top_cal.winfo_height()
            if xc <= x <= xc + w and yc <= y <= yc + h:
                # Mouse adentro del popup → mantener abierto.
                self._calendar.focus_force()
                return
        except Exception:
            pass
        # Mouse afuera → cerrar (comportamiento default).
        try:
            self._top_cal.withdraw()
            self.state(['!pressed'])
        except Exception:
            pass

    try:
        de._on_focus_out_cal = types.MethodType(_patched, de)
        # Re-bindear el evento para usar el método patcheado.
        de._calendar.unbind('<FocusOut>')
        de._calendar.bind('<FocusOut>', de._on_focus_out_cal)
    except Exception as e:
        log.debug("_fix_date_entry_focus_bug: %s", e)


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1480x880")
        self.minsize(1240, 720)
        # Icono de la ventana (title bar + taskbar de Windows).
        try:
            ico = resource_path("assets/Icono.ico")
            if ico.exists():
                self.iconbitmap(default=str(ico))
        except Exception as e:
            log.warning("No se pudo aplicar Icono.ico: %s", e)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.settings: Settings = get_settings()

        self.df: pd.DataFrame | None = None
        self.events: "queue.Queue[ProgressEvent]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: GenerationWorker | None = None

        self._logo_fs_img: ctk.CTkImage | None = None
        self._logo_fed_img: ctk.CTkImage | None = None

        self._build_layout()
        self._refresh_logos()
        if self.settings.storage_backend == "sharepoint":
            self._authenticate_sharepoint()
        else:
            self._validate_paths_on_start()
        self.after(120, self._poll_events)

    # ----------------------- Auth SharePoint -----------------------

    def _authenticate_sharepoint(self) -> None:
        from ..sharepoint.auth import AuthError, ensure_authenticated
        from ..sharepoint.client import SharePointError, get_client
        try:
            self.update_idletasks()
            self._log_console("Autenticando contra Microsoft 365...")
            ensure_authenticated()
            client = get_client()
            client.site_id()
            client.drive_id()
            self._log_console("Sesión SharePoint OK.")
            # NOTA: NO llamamos register_login() ni warm_up_excels() acá.
            # Esos hacían I/O contra Graph en paralelo con el primer Buscar y
            # estaban colgando la descarga de Stock Mendoza. Ahora la telemetría
            # corre al FINAL del worker (cuando ya no hay competencia).
        except (AuthError, SharePointError) as e:
            log.exception("Auth/SP falló")
            messagebox.showerror(
                "Sin conexión a SharePoint",
                f"No se pudo autenticar / conectar:\n\n{e}\n\nLa app va a quedar deshabilitada hasta reintentar.",
            )
            self.btn_buscar.configure(state="disabled")
            self.btn_generar.configure(state="disabled")
        except Exception as e:
            log.exception("Error inesperado en auth")
            messagebox.showerror("Error inesperado", str(e))

    # ----------------------- Layout -----------------------

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Top bar
        top = ctk.CTkFrame(self, corner_radius=0, height=70)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        self.lbl_logo_fs = ctk.CTkLabel(top, text="")
        self.lbl_logo_fs.grid(row=0, column=0, padx=(16, 8), pady=10, sticky="w")
        self.lbl_logo_fed = ctk.CTkLabel(top, text="")
        self.lbl_logo_fed.grid(row=0, column=2, padx=(8, 16), pady=10, sticky="e")

        ctk.CTkLabel(top, text=APP_TITLE,
                     font=ctk.CTkFont(size=18, weight="bold")
                     ).grid(row=0, column=1, padx=8, pady=10)

        self.theme_switch = ctk.CTkSegmentedButton(
            top, values=["Light", "Dark", "System"],
            command=self._on_theme_change,
        )
        self.theme_switch.set("System")
        self.theme_switch.grid(row=0, column=3, padx=12, pady=10, sticky="e")

        # Filtros
        filters = ctk.CTkFrame(self)
        filters.grid(row=1, column=0, sticky="ew", padx=12, pady=(12, 6))
        for i in range(10):
            filters.grid_columnconfigure(i, weight=0)
        filters.grid_columnconfigure(8, weight=1)  # spacer expansible

        date_font = ("Segoe UI", 16, "bold")
        lbl_font = ctk.CTkFont(size=15, weight="bold")

        ctk.CTkLabel(filters, text="Desde:", font=lbl_font).grid(
            row=0, column=0, padx=(12, 6), pady=12, sticky="e"
        )
        self.date_from = DateEntry(
            filters, date_pattern="dd-mm-yyyy", width=14,
            font=date_font, justify="center",
        )
        self.date_from.set_date(date.today() - timedelta(days=7))
        self.date_from.grid(row=0, column=1, padx=6, pady=12, ipady=4)
        _fix_date_entry_focus_bug(self.date_from)

        ctk.CTkLabel(filters, text="Hasta:", font=lbl_font).grid(
            row=0, column=2, padx=(16, 6), pady=12, sticky="e"
        )
        self.date_to = DateEntry(
            filters, date_pattern="dd-mm-yyyy", width=14,
            font=date_font, justify="center",
        )
        self.date_to.set_date(date.today())
        self.date_to.grid(row=0, column=3, padx=6, pady=12, ipady=4)
        _fix_date_entry_focus_bug(self.date_to)

        self.btn_buscar = ctk.CTkButton(filters, text="Buscar", width=110, command=self._on_buscar)
        self.btn_buscar.grid(row=0, column=4, padx=(16, 4), pady=10)

        self.btn_open_folder = ctk.CTkButton(
            filters, text="Abrir carpeta", width=120,
            command=self._on_open_output,
        )
        self.btn_open_folder.grid(row=0, column=5, padx=4, pady=10)

        self.btn_add_clientes = ctk.CTkButton(
            filters, text="Añadir clientes", width=140,
            command=self._on_add_clientes,
            fg_color="#2c5282", hover_color="#1f3c66",
        )
        self.btn_add_clientes.grid(row=0, column=6, padx=4, pady=10)

        self.btn_modify_clientes = ctk.CTkButton(
            filters, text="Modificar clientes", width=150,
            command=self._on_modify_clientes,
            fg_color="#2c5282", hover_color="#1f3c66",
        )
        self.btn_modify_clientes.grid(row=0, column=7, padx=4, pady=10)

        # col 8 = spacer expansible (weight=1) → empuja "Generar" al borde derecho

        self.btn_generar = ctk.CTkButton(
            filters, text="Generar protocolos", width=170,
            state="disabled", command=self._on_generar,
        )
        self.btn_generar.grid(row=0, column=9, padx=(20, 10), pady=10, sticky="e")

        # Entry de carpeta de salida — siempre creado por compatibilidad,
        # pero solo se muestra en backend "local". En "sharepoint" la ruta
        # se informa en el mensaje final de "Generación finalizada".
        is_sp = self.settings.storage_backend == "sharepoint"
        self.entry_output = ctk.CTkEntry(filters)
        if is_sp:
            self.entry_output.insert(0, self.settings.sp_output_folder)
            self.entry_output.configure(state="readonly")
        else:
            self.entry_output.insert(0, str(self.settings.default_output_folder))
        self.btn_browse = ctk.CTkButton(filters, text="...", width=36, command=self._on_browse_output)

        if not is_sp:
            ctk.CTkLabel(filters, text="Carpeta de salida:").grid(
                row=1, column=0, padx=(10, 4), pady=(0, 10), sticky="e"
            )
            self.entry_output.grid(row=1, column=1, columnspan=4, padx=4, pady=(0, 10), sticky="ew")
            self.btn_browse.grid(row=1, column=5, padx=4, pady=(0, 10), sticky="w")

        # Body: paneles principales
        body = ctk.CTkFrame(self)
        body.grid(row=2, column=0, sticky="nsew", padx=12, pady=6)
        body.grid_columnconfigure(0, weight=0)  # clientes list
        body.grid_columnconfigure(1, weight=0)  # comprobantes list
        body.grid_columnconfigure(2, weight=1)  # tablas
        body.grid_rowconfigure(0, weight=1)

        # Lista clientes
        self.list_clientes = MultiSelectListbox(
            body, title="Clientes",
            on_select=self._on_clientes_changed,
            height=12, width=200,
        )
        self.list_clientes.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=8)

        # Lista comprobantes
        self.list_comprobantes = MultiSelectListbox(
            body, title="Remitos",
            on_select=self._on_comprobantes_changed,
            height=12, width=240,
        )
        self.list_comprobantes.grid(row=0, column=1, sticky="nsew", padx=4, pady=8)

        # Panel tablas (resumen + detalle apilados)
        right = ctk.CTkFrame(body)
        right.grid(row=0, column=2, sticky="nsew", padx=(4, 8), pady=8)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)
        right.grid_rowconfigure(3, weight=2)

        ctk.CTkLabel(right, text="Resumen por remito",
                     anchor="w", font=ctk.CTkFont(weight="bold")
                     ).grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 2))
        self.tbl_summary = DataTable(right, columns=SUMMARY_COLS, widths=SUMMARY_WIDTHS, height=6)
        self.tbl_summary.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 6))

        ctk.CTkLabel(right, text="Detalle Remito",
                     anchor="w", font=ctk.CTkFont(weight="bold")
                     ).grid(row=2, column=0, sticky="ew", padx=8, pady=(2, 2))
        self.tbl_detail = DataTable(right, columns=DETAIL_COLS, widths=DETAIL_WIDTHS, height=10)
        self.tbl_detail.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))

        # Bottom: status + progreso (sin consola de logs, todo va al archivo .log)
        bottom = ctk.CTkFrame(self)
        bottom.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 12))
        bottom.grid_columnconfigure(1, weight=1)

        self.lbl_summary_text = ctk.CTkLabel(
            bottom, text="Sin datos.", anchor="w",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.lbl_summary_text.grid(row=0, column=0, sticky="w", padx=12, pady=(10, 0))

        self.lbl_status = ctk.CTkLabel(
            bottom, text="Listo.", anchor="w",
            font=ctk.CTkFont(size=13),
        )
        self.lbl_status.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))

        self.progress = ctk.CTkProgressBar(bottom)
        self.progress.grid(row=0, column=1, sticky="ew", padx=12, pady=(14, 4), rowspan=1)
        self.progress.set(0)

        actions = ctk.CTkFrame(bottom, fg_color="transparent")
        actions.grid(row=0, column=2, rowspan=2, sticky="e", padx=8, pady=8)

        self.btn_cancelar = ctk.CTkButton(
            actions, text="Cancelar", state="disabled",
            command=self._on_cancelar,
            fg_color="#9c2a2a", hover_color="#7a2020", width=110,
        )
        self.btn_cancelar.grid(row=0, column=0, padx=4, pady=2)

        self.btn_open_logs = ctk.CTkButton(
            actions, text="Abrir logs", width=110,
            command=self._on_open_logs,
        )
        self.btn_open_logs.grid(row=0, column=1, padx=4, pady=2)

    # ----------------------- Logos / Tema -----------------------

    def _refresh_logos(self) -> None:
        try:
            fs_light = Image.open(resource_path("assets/logo_fs.png"))
            fs_dark = Image.open(resource_path("assets/logo_fs_dark.png"))
            fed_light = Image.open(resource_path("assets/logo_fedrigoni.png"))
            fed_dark = Image.open(resource_path("assets/logo_fedrigoni_dark.png"))
        except FileNotFoundError as e:
            log.warning("Logo faltante: %s", e)
            return

        h = 42
        fs_size = self._fit(fs_light, h)
        fed_size = self._fit(fed_light, h)

        fs_light_r = fs_light.resize(fs_size, Image.LANCZOS)
        fs_dark_r = fs_dark.resize(fs_size, Image.LANCZOS)
        fed_light_r = fed_light.resize(fed_size, Image.LANCZOS)
        fed_dark_r = fed_dark.resize(fed_size, Image.LANCZOS)

        self._logo_fs_img = ctk.CTkImage(
            light_image=fs_light_r, dark_image=fs_dark_r, size=fs_size,
        )
        self._logo_fed_img = ctk.CTkImage(
            light_image=fed_light_r, dark_image=fed_dark_r, size=fed_size,
        )
        self.lbl_logo_fs.configure(image=self._logo_fs_img, text="")
        self.lbl_logo_fed.configure(image=self._logo_fed_img, text="")

    @staticmethod
    def _fit(img: Image.Image, target_h: int) -> tuple[int, int]:
        ratio = img.width / max(img.height, 1)
        return (max(1, int(target_h * ratio)), target_h)

    def _on_theme_change(self, value: str) -> None:
        ctk.set_appearance_mode(value)
        self._refresh_logos()

    # ----------------------- Validaciones iniciales -----------------------

    def _validate_paths_on_start(self) -> None:
        s = self.settings
        try:
            ensure_readable(s.protocols_folder, "Carpeta de protocolos")
            self._log_console(f"OK Carpeta protocolos: {s.protocols_folder}")
        except Exception as e:
            messagebox.showwarning("Acceso a SharePoint", str(e))
            self._log_console(f"⚠ {e}")

        try:
            ensure_writable(s.default_output_folder, "Carpeta de salida (Reportes Finales)")
            self._log_console(f"OK Carpeta salida: {s.default_output_folder}")
        except Exception as e:
            messagebox.showwarning("Acceso a SharePoint", str(e))
            self._log_console(f"⚠ {e}")

    # ----------------------- Acciones top -----------------------

    def _on_browse_output(self) -> None:
        initial = self.entry_output.get() or str(Path.home())
        folder = filedialog.askdirectory(initialdir=initial, title="Seleccionar carpeta de salida")
        if folder:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, folder)

    def _on_buscar(self) -> None:
        """Lanza fetch_data en un thread daemon — la GUI sigue responsiva."""
        try:
            d_from = self.date_from.get_date()
            d_to = self.date_to.get_date()
        except Exception as e:
            messagebox.showerror("Fechas inválidas", str(e))
            return

        self._log_console(f"Consultando datos {d_from} → {d_to}...")
        self.lbl_status.configure(text="Buscando datos...")
        self.btn_buscar.configure(state="disabled")
        self.btn_generar.configure(state="disabled")
        self.update_idletasks()

        threading.Thread(
            target=self._buscar_worker,
            args=(d_from, d_to),
            daemon=True,
        ).start()

    def _buscar_worker(self, d_from, d_to) -> None:
        try:
            df = data_service.fetch_data(d_from, d_to)
            self.events.put(ProgressEvent("fetch_done", payload=df))
        except Exception as e:
            log.exception("Error en buscar worker")
            self.events.put(ProgressEvent("fetch_error", message=str(e)))
            return

        # Después del fetch principal, levantamos en background los remitos ya
        # registrados en la SP List del bot (para pintarlos en verde en la GUI).
        # Si SP no responde o la list no existe, no rompe nada (best-effort).
        threading.Thread(
            target=self._load_tracking_async,
            args=(d_from, d_to),
            daemon=True,
            name="GUI-LoadTracking",
        ).start()

    def _load_tracking_async(self, d_from, d_to) -> None:
        """Consulta la SP List del bot y emite los #Comprobante con estado
        terminal 'positivo' (enviado_cliente / enviado_manual / resuelto_manual).
        Los pending_control y error NO se pintan."""
        if self.settings.storage_backend != "sharepoint":
            return
        try:
            from ..sharepoint.lists import (
                ESTADO_ENVIADO_CLIENTE,
                ESTADO_ENVIADO_MANUAL,
                ESTADO_RESUELTO_MANUAL,
                get_lists,
            )
            lists = get_lists()
            list_name = self.settings.sp_tracking_list_name
            if not lists.list_exists(list_name):
                # Bot nunca corrió → no hay nada para pintar.
                return
            tracking = lists.fetch_in_range(list_name, d_from, d_to)
            terminales_ok = {
                ESTADO_ENVIADO_CLIENTE,
                ESTADO_ENVIADO_MANUAL,
                ESTADO_RESUELTO_MANUAL,
            }
            comps = {c for c, e in tracking.items() if e.estado in terminales_ok}
            self.events.put(ProgressEvent("tracking_loaded", payload=comps))
        except Exception as e:
            log.warning("Tracking async load falló (no crítico): %s", e)

    def _on_add_clientes(self) -> None:
        """Flow para añadir clientes al listado (set Protocolos='S' + mails).

        Sync para simplificar (download + dialogs + upload). La GUI puede
        verse trabada brevemente durante la descarga/subida del Excel
        (~5-10s en total con red estable).
        """
        self.btn_add_clientes.configure(state="disabled")
        self.lbl_status.configure(text="Descargando lista de clientes...")
        self._log_console("Descargando lista completa de clientes desde SharePoint...")
        self.update_idletasks()

        try:
            df_all = data_service.load_clientes_all()
        except Exception as e:
            log.exception("Error al cargar clientes para añadir")
            messagebox.showerror("Error", f"No se pudo cargar el listado de clientes:\n{e}")
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        if "Protocolos" not in df_all.columns:
            messagebox.showerror(
                "Excel inválido",
                "No se encontró la columna 'Protocolos' en la hoja Clientes.",
            )
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        mask_no_s = (
            df_all["Protocolos"].fillna("").astype(str).str.strip().str.upper() != "S"
        )
        df_no_s = df_all[mask_no_s].copy()
        if df_no_s.empty:
            messagebox.showinfo(
                "Sin clientes pendientes",
                "Todos los clientes del Excel ya tienen el flag de protocolos en 'S'.",
            )
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self.lbl_status.configure(text=f"{len(df_no_s)} cliente(s) sin marca de protocolos.")
        self.update_idletasks()

        selected = AddClientesDialog(self, df_no_s).show()
        if not selected:
            self._log_console("Añadir clientes: cancelado.")
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        updates: dict[str, str] = {}
        for cli in selected:
            skipped, mails = AskMailsDialog(
                self,
                razon_social=cli["razon_social"],
                mail_actual=cli["mail_actual"],
            ).show()
            if skipped:
                self._log_console(f"  · {cli['razon_social']}: saltado.")
                continue
            updates[cli["codigo"]] = mails or ""

        if not updates:
            messagebox.showinfo("Sin cambios", "No se confirmó ningún cliente.")
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self.lbl_status.configure(text=f"Guardando {len(updates)} cliente(s) en SharePoint...")
        self._log_console(f"Aplicando cambios a {len(updates)} cliente(s) en Stock Mendoza.xlsm...")
        self.update_idletasks()

        try:
            n = data_service.update_clientes_in_excel(updates)
        except Exception as e:
            log.exception("Error al actualizar clientes en Excel")
            err_str = str(e)
            if "resourceLocked" in err_str or " 423 " in err_str or err_str.startswith("chunk upload falló: 423"):
                messagebox.showerror(
                    "Excel bloqueado",
                    "El archivo Stock Mendoza.xlsm está bloqueado en SharePoint.\n\n"
                    "Causas más comunes:\n"
                    "  • Lo tenés abierto en Excel (cerralo).\n"
                    "  • Otra persona lo está editando online.\n"
                    "  • OneDrive sigue sincronizando.\n\n"
                    "Cerralo en Excel y volvé a intentar.",
                )
            else:
                messagebox.showerror(
                    "Error al guardar",
                    f"No se pudieron guardar los cambios:\n{e}",
                )
            self.btn_add_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self._log_console(f"OK · {n} cliente(s) añadidos al listado.")
        messagebox.showinfo(
            "Clientes añadidos",
            f"{n} cliente(s) añadidos al listado.\n\n"
            "Volvé a apretar Buscar para verlos en las listas.",
        )
        self.btn_add_clientes.configure(state="normal")
        self.lbl_status.configure(text="Listo.")

    def _on_modify_clientes(self) -> None:
        """Flow para modificar mails de clientes que YA tienen Protocolos='S'.

        Muestra clientes con S + sus mails actuales. Permite editar y guardar.
        El flag 'S' se mantiene (no se quita). Si querés desactivar un cliente
        del listado, hay que tocar el Excel a mano (decisión consciente para
        evitar errores destructivos).
        """
        self.btn_modify_clientes.configure(state="disabled")
        self.lbl_status.configure(text="Descargando lista de clientes...")
        self._log_console("Descargando lista completa de clientes desde SharePoint...")
        self.update_idletasks()

        try:
            df_all = data_service.load_clientes_all()
        except Exception as e:
            log.exception("Error al cargar clientes para modificar")
            messagebox.showerror("Error", f"No se pudo cargar el listado de clientes:\n{e}")
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        if "Protocolos" not in df_all.columns:
            messagebox.showerror(
                "Excel inválido",
                "No se encontró la columna 'Protocolos' en la hoja Clientes.",
            )
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        mask_s = (
            df_all["Protocolos"].fillna("").astype(str).str.strip().str.upper() == "S"
        )
        df_s = df_all[mask_s].copy()
        if df_s.empty:
            messagebox.showinfo(
                "Sin clientes para modificar",
                "No hay clientes con flag 'S' en el Excel. Usá 'Añadir clientes' primero.",
            )
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self.lbl_status.configure(text=f"{len(df_s)} cliente(s) con marca de protocolos.")
        self.update_idletasks()

        selected = AddClientesDialog(
            self, df_s,
            window_title="Modificar clientes",
            header_text=f"Clientes con protocolos activos ({len(df_s)})",
            subtitle="Buscá por nombre o mail, marcá los que quieras modificar.",
            show_mail_in_label=True,
        ).show()
        if not selected:
            self._log_console("Modificar clientes: cancelado.")
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        updates: dict[str, str] = {}
        for cli in selected:
            skipped, mails = AskMailsDialog(
                self,
                razon_social=cli["razon_social"],
                mail_actual=cli["mail_actual"],
            ).show()
            if skipped:
                self._log_console(f"  · {cli['razon_social']}: saltado.")
                continue
            updates[cli["codigo"]] = mails or ""

        if not updates:
            messagebox.showinfo("Sin cambios", "No se confirmó ningún cliente.")
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self.lbl_status.configure(text=f"Guardando {len(updates)} cliente(s) en SharePoint...")
        self._log_console(f"Modificando {len(updates)} cliente(s) en Stock Mendoza.xlsm...")
        self.update_idletasks()

        try:
            n = data_service.update_clientes_in_excel(updates)
        except Exception as e:
            log.exception("Error al modificar clientes en Excel")
            err_str = str(e)
            if "resourceLocked" in err_str or " 423 " in err_str or err_str.startswith("chunk upload falló: 423"):
                messagebox.showerror(
                    "Excel bloqueado",
                    "El archivo Stock Mendoza.xlsm está bloqueado en SharePoint.\n\n"
                    "Causas más comunes:\n"
                    "  • Lo tenés abierto en Excel (cerralo).\n"
                    "  • Otra persona lo está editando online.\n"
                    "  • OneDrive sigue sincronizando.\n\n"
                    "Cerralo en Excel y volvé a intentar.",
                )
            else:
                messagebox.showerror(
                    "Error al guardar",
                    f"No se pudieron guardar los cambios:\n{e}",
                )
            self.btn_modify_clientes.configure(state="normal")
            self.lbl_status.configure(text="Listo.")
            return

        self._log_console(f"OK · {n} cliente(s) modificados.")
        messagebox.showinfo(
            "Clientes modificados",
            f"{n} cliente(s) modificados.\n\nLos nuevos mails se usan a partir de la próxima ejecución.",
        )
        self.btn_modify_clientes.configure(state="normal")
        self.lbl_status.configure(text="Listo.")

    def _on_generar(self) -> None:
        log.info("_on_generar invocado")
        if self.df is None or self.df.empty:
            log.info("→ self.df vacío, messagebox y return")
            messagebox.showinfo("Sin datos", "Primero buscá datos con un rango válido.")
            return

        df_filtered = self._current_filtered_df()
        log.info("→ df_filtered: %d filas", len(df_filtered))
        if df_filtered.empty:
            messagebox.showinfo("Sin datos", "El filtro actual no devuelve filas.")
            return

        # Si el worker anterior sigue vivo, no permitimos arrancar otro.
        if self.worker is not None and self.worker.is_alive():
            log.info("→ worker anterior sigue vivo, abort")
            messagebox.showinfo(
                "Procesando",
                "Hay una generación en curso. Esperá a que termine o cancelala primero.",
            )
            return

        if self.settings.storage_backend == "sharepoint":
            out_dir = self.settings.sp_output_folder
            if not out_dir:
                messagebox.showerror("Configuración", "SP_OUTPUT_FOLDER vacío en .env.")
                return
        else:
            out_dir = Path(self.entry_output.get().strip())
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
                ensure_writable(out_dir, "Carpeta de salida")
            except Exception as e:
                messagebox.showerror("Carpeta inválida", str(e))
                return

        # Pre-análisis: comprobantes con items SAFED sin protocolo
        missing = data_service.analyze_missing_protocols(df_filtered)
        log.info("→ missing: %d remito(s) con items sin protocolo", len(missing))
        pre_results: list[ComprobanteResult] = []

        if missing:
            try:
                detail_full = data_service.build_detail_view(df_filtered).copy()
                problematicos = list(missing.keys())
                export_df = detail_full[detail_full["Remito"].astype(str).isin(problematicos)].copy()
                if not export_df.empty:
                    mask_missing = export_df["Protocolo"].fillna("").astype(str).str.strip() == ""
                    # Solo las filas que realmente NO tienen protocolo.
                    export_df = export_df[mask_missing].copy()
                    export_df["Protocolo"] = "⚠ FALTA PROTOCOLO"
                    export_df = export_df.reindex(columns=DETAIL_COLS, fill_value="")
            except Exception as e:
                log.exception("No se pudo armar export_df")
                export_df = None

            log.info("Abriendo MissingProtocolsDialog (modal)")
            dialog = MissingProtocolsDialog(self, missing, export_df=export_df)
            decision = dialog.show()
            log.info("MissingProtocolsDialog cerrado con: %s", decision)
            if decision == "cancel":
                self._log_console("Generación cancelada por el usuario.")
                return
            if decision == "skip":
                problematicos = list(missing.keys())
                # OJO: df_filtered es el DataFrame INTERNO (de fetch_data),
                # tiene columna "#Comprobante", NO "Remito" (display name).
                df_filtered = df_filtered[~df_filtered["#Comprobante"].astype(str).isin(problematicos)]
                for comp, info in missing.items():
                    pre_results.append(ComprobanteResult(
                        comprobante=comp,
                        cliente=info["cliente"],
                        razon_social=info["razon_social"],
                        numero_oc=info["oc"],
                        mail=info["mail"],
                        estado="omitido_por_usuario",
                    ))
                self._log_console(f"Omitidos por el usuario: {len(problematicos)} remito(s).")

        log.info("→ Lanzando GenerationWorker (df=%d filas, pre_results=%d)",
                 len(df_filtered), len(pre_results))
        self.cancel_event.clear()
        self.btn_generar.configure(state="disabled")
        self.btn_buscar.configure(state="disabled")
        self.btn_cancelar.configure(state="normal")
        self.progress.set(0)
        self.lbl_status.configure(text="Generando...")

        self.worker = GenerationWorker(
            df=df_filtered,
            output_dir=out_dir,
            protocols_folder=self.settings.protocols_folder,
            events=self.events,
            cancel_event=self.cancel_event,
            pre_results=pre_results,
        )
        self.worker.start()
        log.info("→ Worker lanzado, alive=%s", self.worker.is_alive())

    def _on_cancelar(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self._log_console("Cancelación solicitada... esperando a que termine el remito actual.")

    def _on_open_output(self) -> None:
        if self.settings.storage_backend == "sharepoint":
            try:
                import webbrowser
                from ..sharepoint.client import get_client
                url = get_client().folder_web_url(self.settings.sp_output_folder)
                webbrowser.open(url)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir la carpeta SharePoint:\n{e}")
            return
        path = self.entry_output.get().strip()
        if path and Path(path).exists():
            os.startfile(path)
        else:
            messagebox.showinfo("Carpeta", "La carpeta no existe todavía.")

    def _on_open_logs(self) -> None:
        from ..utils.paths import runtime_dir
        logs_dir = runtime_dir() / "logs"
        if logs_dir.exists():
            os.startfile(logs_dir)
        else:
            messagebox.showinfo("Logs", f"Aún no hay logs ({logs_dir}).")

    # ----------------------- Cross-filter -----------------------

    def _on_clientes_changed(self, selected: list[str]) -> None:
        if self.df is None:
            return
        if selected:
            allowed = self.df[self.df["Código de Cliente"].astype(str).isin(selected)]
            self.list_comprobantes.set_items_filtered(data_service.distinct_comprobantes(allowed))
        else:
            self.list_comprobantes.set_items_filtered(data_service.distinct_comprobantes(self.df))
        self._refresh_tables()

    def _on_comprobantes_changed(self, selected: list[str]) -> None:
        if self.df is None:
            return
        if selected:
            # OJO: la columna en self.df es "#Comprobante" (nombre interno),
            # NO "Remito" (que es solo el display name de la tabla).
            allowed = self.df[self.df["#Comprobante"].astype(str).isin(selected)]
            self.list_clientes.set_items_filtered(data_service.distinct_clientes(allowed))
        else:
            self.list_clientes.set_items_filtered(data_service.distinct_clientes(self.df))
        self._refresh_tables()

    def _current_filtered_df(self) -> pd.DataFrame:
        if self.df is None:
            return pd.DataFrame()
        return data_service.apply_filters(
            self.df,
            clientes=self.list_clientes.get_selected() or None,
            comprobantes=self.list_comprobantes.get_selected() or None,
        )

    def _refresh_tables(self) -> None:
        df = self._current_filtered_df()
        summary = data_service.build_summary_view(df)
        detail = data_service.build_detail_view(df).copy()

        # Marca visual: items SAFED sin Protocolo → "⚠ FALTA PROTOCOLO" + tag rojo bold
        if not detail.empty:
            mask_missing = detail["Protocolo"].fillna("").astype(str).str.strip() == ""
            detail.loc[mask_missing, "Protocolo"] = "⚠ FALTA PROTOCOLO"

        self.tbl_summary.set_rows(summary.to_dict(orient="records"))
        self.tbl_detail.set_rows(
            detail.to_dict(orient="records"),
            flag_predicate=lambda r: r.get("Protocolo") == "⚠ FALTA PROTOCOLO",
        )
        n_comp = len(summary)
        n_filas = len(df)
        n_safed = len(detail)
        self.lbl_summary_text.configure(
            text=f"Remitos: {n_comp}  |  Filas: {n_filas}  |  Items SAFED: {n_safed}"
        )

    # ----------------------- Polling de eventos worker -----------------------

    def _poll_events(self) -> None:
        try:
            while True:
                ev = self.events.get_nowait()
                self._handle_event(ev)
        except queue.Empty:
            pass
        self.after(120, self._poll_events)

    def _handle_event(self, ev: ProgressEvent) -> None:
        if ev.kind == "log":
            self._log_console(ev.message)
        elif ev.kind == "progress":
            if ev.total > 0:
                self.progress.set(ev.current / ev.total)
                self.lbl_status.configure(text=f"{ev.current}/{ev.total}")
        elif ev.kind == "result":
            if ev.message:
                self._log_console(ev.message)
        elif ev.kind == "fetch_done":
            df = ev.payload
            self.df = df
            self.list_clientes.set_items(data_service.distinct_clientes(df))
            self.list_comprobantes.set_items(data_service.distinct_comprobantes(df))
            self._refresh_tables()
            n_filas = len(df) if df is not None else 0
            n_cli = len(self.list_clientes._all_items)
            n_comp = len(self.list_comprobantes._all_items)
            self._log_console(
                f"Datos cargados: {n_filas} filas, {n_cli} cliente(s), {n_comp} remito(s)."
            )
            self.lbl_status.configure(text="Datos cargados.")
            self.btn_buscar.configure(state="normal")
            self.btn_generar.configure(state="normal" if n_filas > 0 else "disabled")
        elif ev.kind == "tracking_loaded":
            comps = ev.payload or set()
            try:
                self.list_comprobantes.set_highlighted(comps)
            except Exception as e:
                log.warning("set_highlighted falló: %s", e)
            if comps:
                self._log_console(
                    f"Tracking: {len(comps)} remito(s) ya procesado(s) por el bot (resaltados en verde)."
                )
        elif ev.kind == "fetch_error":
            self._log_console(f"✗ Error: {ev.message}")
            self.lbl_status.configure(text="Error.")
            self.btn_buscar.configure(state="normal")
            self.btn_generar.configure(state="disabled")
            messagebox.showerror("Error al buscar datos", ev.message)
        elif ev.kind == "done":
            self._log_console(ev.message)
            self.progress.set(1.0)
            self.lbl_status.configure(text="Listo.")
            self.btn_generar.configure(state="normal")
            self.btn_buscar.configure(state="normal")
            self.btn_cancelar.configure(state="disabled")
            self._maybe_open_drafts(ev.payload or [])
        elif ev.kind == "error":
            self._log_console(ev.message)
            self.lbl_status.configure(text="Error.")
            self.btn_generar.configure(state="normal")
            self.btn_buscar.configure(state="normal")
            self.btn_cancelar.configure(state="disabled")
            messagebox.showerror("Error", ev.message)

    def _output_location_text(self) -> str:
        """Texto amigable de la carpeta de salida para mostrar al usuario."""
        if self.settings.storage_backend == "sharepoint":
            site = self.settings.sharepoint_site or "VentasPowerBI"
            sub = (self.settings.sp_output_folder or "").replace("\\", "/")
            crumbs = " > ".join([p for p in sub.split("/") if p])
            return f"{site} > Documentos compartidos > {crumbs}"
        return str(self.settings.default_output_folder)

    def _maybe_open_drafts(self, results: list[ComprobanteResult]) -> None:
        log.info("_maybe_open_drafts: %d results", len(results))
        # 1) Si hubo remitos que ya estaban generados, mostrar dialog informativo.
        ya_existen = [r.comprobante for r in results if r.estado == "ya_existe"]
        if ya_existen:
            try:
                AlreadyGeneratedDialog(
                    self,
                    comprobantes=ya_existen,
                    open_folder_callback=self._on_open_output,
                    location_text=self._output_location_text(),
                ).show()
            except Exception as e:
                log.warning("No se pudo mostrar dialog ya_generados: %s", e)

        # 2) Borradores Outlook para los que SÍ se generaron (ok / solo_caratula).
        eligibles = [
            r for r in results
            if r.estado in ("ok", "solo_caratula") and r.mail and r.pdf_path
        ]
        n = len(eligibles)
        n_generados = sum(1 for r in results if r.estado in ("ok", "solo_caratula"))
        done_msg = (
            f"Protocolos generados: {n_generados} remito(s).\n"
            f"Guardado en: {self._output_location_text()}"
        )

        if n == 0:
            messagebox.showinfo("Finalizado", done_msg + "\n\nNo hay borradores para abrir.")
            return

        ask = messagebox.askyesno(
            "Borradores Outlook",
            f"{done_msg}\n\n¿Querés abrir {n} borrador(es) en Outlook ahora?",
        )
        if not ask:
            self._log_console("El usuario eligió no abrir borradores.")
            return

        self._log_console(f"Abriendo {n} borrador(es) en Outlook...")
        ok = 0
        sent_manually: list[ComprobanteResult] = []
        for r in eligibles:
            try:
                # Outlook necesita un path LOCAL para adjuntar.
                attach_path = Path(r.pdf_local_path) if r.pdf_local_path else Path(r.pdf_path)
                mail_service.open_outlook_draft(
                    to=r.mail,
                    razon_social=r.razon_social,
                    comprobante=r.comprobante,
                    pdf_path=attach_path,
                    protocolos_encontrados=r.protocolos_encontrados,
                    protocolos_no_encontrados=r.protocolos_no_encontrados,
                )
                ok += 1
                sent_manually.append(r)
            except Exception as e:
                log.exception("Error abriendo borrador Outlook para %s", r.comprobante)
                self._log_console(f"⚠ No se pudo abrir Outlook para {r.comprobante}: {e}")

        skipped = [r for r in results if r.estado in ("ok", "solo_caratula") and not r.mail]
        if skipped:
            self._log_console(
                f"{len(skipped)} remito(s) sin Mail Protocolos → no se abrió borrador."
            )
        self._log_console(f"Borradores abiertos: {ok}/{n}.")

        # Registrar el envío manual en el tracking del bot.
        # Best-effort en thread daemon: si falla (sin red, sin permisos), el ciclo
        # manual sigue funcionando igual, solo queda el row del bot sin actualizar.
        # Asumimos que si el draft se abrió OK, el operador va a apretar Enviar.
        if sent_manually and self.settings.storage_backend == "sharepoint":
            threading.Thread(
                target=self._register_manual_sends,
                args=(sent_manually,),
                daemon=True,
            ).start()

    def _register_manual_sends(self, results: list[ComprobanteResult]) -> None:
        """Best-effort: registra cada envío manual en el tracking del bot.

        Lógica por remito:
          - Si NO está en tracking → CREA fila con estado `enviado_manual`.
          - Si está como `pending_control` (el bot lo flagueó) → pasa a `resuelto_manual`.
          - Si está en otro estado terminal (enviado_cliente, enviado_manual,
            resuelto_manual) → no toca (ya estaba cerrado).
          - Si está como `error` → lo pisa con `enviado_manual` (consideramos que
            el operador resolvió lo que el bot no pudo).

        Solo se activa si la SharePoint List existe (no la crea desde la GUI).
        """
        try:
            from ..sharepoint.lists import (
                ESTADO_ENVIADO_CLIENTE,
                ESTADO_ENVIADO_MANUAL,
                ESTADO_ERROR,
                ESTADO_PENDING_CONTROL,
                ESTADO_RESUELTO_MANUAL,
                get_lists,
            )
            lists = get_lists()
            list_name = self.settings.sp_tracking_list_name
            if not lists.list_exists(list_name):
                # Bot nunca corrió → no hay tracking todavía. NO la creamos desde
                # la GUI (esa es responsabilidad del modo --auto).
                return

            terminales_cerrados = {
                ESTADO_ENVIADO_CLIENTE,
                ESTADO_ENVIADO_MANUAL,
                ESTADO_RESUELTO_MANUAL,
            }

            from datetime import datetime
            run_id = "gui-" + datetime.now().strftime("%Y%m%d_%H%M%S")
            for r in results:
                try:
                    comp = r.comprobante
                    existing = lists.find_by_comprobante(list_name, comp)

                    if existing is None:
                        # Nunca tocado por el bot → crear como envío manual.
                        try:
                            fecha_str = ""
                            if self.df is not None and not self.df.empty:
                                sub = self.df[self.df["#Comprobante"].astype(str) == comp]
                                if not sub.empty:
                                    f = sub.iloc[0].get("Fecha")
                                    try:
                                        fecha_str = pd.to_datetime(f).date().isoformat()
                                    except Exception:
                                        pass
                        except Exception:
                            fecha_str = ""

                        lists.upsert(
                            list_name,
                            comprobante=comp,
                            cliente=r.cliente,
                            razon_social=r.razon_social,
                            fecha_remito=fecha_str,
                            estado=ESTADO_ENVIADO_MANUAL,
                            mail_destino=r.mail or "(desde GUI)",
                            pdf_url=r.pdf_path or "",
                            run_id=run_id,
                        )
                        log.info("Tracking: %s → enviado_manual (nuevo, desde GUI)", comp)
                        continue

                    if existing.estado in terminales_cerrados:
                        # Ya estaba cerrado → no tocar.
                        continue

                    if existing.estado == ESTADO_PENDING_CONTROL:
                        lists.upsert(
                            list_name,
                            comprobante=comp,
                            estado=ESTADO_RESUELTO_MANUAL,
                        )
                        log.info("Tracking: %s → resuelto_manual", comp)
                        continue

                    if existing.estado == ESTADO_ERROR:
                        # El bot había fallado; el operador lo resolvió.
                        lists.upsert(
                            list_name,
                            comprobante=comp,
                            estado=ESTADO_ENVIADO_MANUAL,
                            mail_destino=r.mail or "(desde GUI)",
                            pdf_url=r.pdf_path or "",
                            run_id=run_id,
                        )
                        log.info("Tracking: %s (estaba en error) → enviado_manual", comp)
                        continue

                    # Cualquier otro estado desconocido → log warning, no tocar.
                    log.warning(
                        "Tracking: %s tiene estado desconocido '%s', no se modifica",
                        comp, existing.estado,
                    )
                except Exception as e:
                    log.warning("No se pudo registrar %s en tracking: %s", comp, e)
        except Exception as e:
            log.warning("Hook _register_manual_sends falló: %s", e)

    def _log_console(self, text: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {text}"
        log.info(text)
        try:
            self.lbl_status.configure(text=line)
        except Exception:
            pass
