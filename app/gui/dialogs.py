"""Dialogs custom para la GUI."""
from __future__ import annotations

import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Callable, Literal

import customtkinter as ctk
import pandas as pd

Decision = Literal["all", "skip", "cancel"]


class MissingProtocolsDialog(ctk.CTkToplevel):
    """Dialog modal que lista comprobantes con items SAFED sin protocolo
    y deja al user decidir: generar todos / omitir problemáticos / cancelar.
    Permite además exportar la lista a CSV o Excel.
    """

    def __init__(
        self,
        parent,
        missing: dict[str, dict],
        export_df: pd.DataFrame | None = None,
    ):
        super().__init__(parent)
        self.title("Items sin protocolo detectados")
        self.geometry("780x560")
        self.minsize(640, 420)

        self._decision: Decision = "cancel"
        self.missing = missing
        self.export_df = export_df

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        n = len(missing)
        total_items = sum(len(d["items"]) for d in missing.values())

        header = ctk.CTkLabel(
            self,
            text=(
                f"Hay {n} comprobante(s) con {total_items} item(s) SAFED sin protocolo.\n"
                "¿Querés generar igual o omitir los problemáticos?"
            ),
            font=ctk.CTkFont(size=14, weight="bold"),
            justify="left", anchor="w",
        )
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))

        body = ctk.CTkScrollableFrame(self)
        body.grid(row=1, column=0, sticky="nsew", padx=16, pady=4)
        body.grid_columnconfigure(0, weight=1)

        title_font = ctk.CTkFont(size=13, weight="bold")
        item_font = ctk.CTkFont(size=12)

        row_idx = 0
        for comp, info in missing.items():
            block = ctk.CTkFrame(body)
            block.grid(row=row_idx, column=0, sticky="ew", padx=4, pady=6)
            block.grid_columnconfigure(0, weight=1)

            n_items = len(info["items"])
            heading = (
                f"Cliente {info['cliente']} — {info['razon_social']}\n"
                f"Remito {comp} — OC {info['oc']}   ({n_items} item(s) sin protocolo)"
            )
            ctk.CTkLabel(
                block, text=heading, font=title_font,
                justify="left", anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))

            for j, it in enumerate(info["items"], 1):
                line = f"   • {it['producto']}  |  Serie: {it['serie'] or '-'}  |  m²: {it['m2'] or '-'}"
                if it["descripcion2"]:
                    line += f"\n     {it['descripcion2']}"
                ctk.CTkLabel(
                    block, text=line, font=item_font,
                    justify="left", anchor="w",
                ).grid(row=j, column=0, sticky="ew", padx=12, pady=(0, 2))

            row_idx += 1

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=12, pady=12)
        footer.grid_columnconfigure(1, weight=1)

        # Botón Exportar a la izquierda
        self.btn_export = ctk.CTkButton(
            footer, text="Exportar...", width=130,
            fg_color="#2c5282", hover_color="#1f3c66",
            command=self._on_export,
        )
        self.btn_export.grid(row=0, column=0, sticky="w", padx=4)
        if self.export_df is None or self.export_df.empty:
            self.btn_export.configure(state="disabled")

        btns = ctk.CTkFrame(footer, fg_color="transparent")
        btns.grid(row=0, column=2, sticky="e")

        self.btn_cancel = ctk.CTkButton(
            btns, text="Cancelar", width=120,
            fg_color="#555555", hover_color="#3d3d3d",
            command=self._on_cancel,
        )
        self.btn_cancel.grid(row=0, column=0, padx=4)

        self.btn_skip = ctk.CTkButton(
            btns, text="Omitir problemáticos", width=180,
            command=self._on_skip,
        )
        self.btn_skip.grid(row=0, column=1, padx=4)

        self.btn_all = ctk.CTkButton(
            btns, text="Generar todos", width=160,
            fg_color="#1f6f3f", hover_color="#155226",
            command=self._on_all,
        )
        self.btn_all.grid(row=0, column=2, padx=4)

        self.after(50, self._center_on_parent)

    def _on_export(self) -> None:
        if self.export_df is None or self.export_df.empty:
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Exportar items sin protocolo",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
            ],
            initialfile="items_sin_protocolo.xlsx",
        )
        if not path:
            return
        try:
            ext = Path(path).suffix.lower()
            if ext == ".csv":
                self.export_df.to_csv(path, index=False, encoding="utf-8-sig", sep=";")
            else:
                # Default: Excel (.xlsx) — incluye casos donde el user escribe sin extensión
                if ext != ".xlsx":
                    path = str(Path(path).with_suffix(".xlsx"))
                self.export_df.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo(
                "Exportación OK",
                f"Archivo guardado en:\n{path}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Error al exportar",
                f"No se pudo guardar el archivo:\n{e}",
                parent=self,
            )

    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            parent = self.master
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _on_cancel(self) -> None:
        self._decision = "cancel"
        self.destroy()

    def _on_skip(self) -> None:
        self._decision = "skip"
        self.destroy()

    def _on_all(self) -> None:
        self._decision = "all"
        self.destroy()

    def show(self) -> Decision:
        """Bloquea hasta cerrar y devuelve la decisión."""
        self.wait_window()
        return self._decision


class AlreadyGeneratedDialog(ctk.CTkToplevel):
    """Dialog informativo que muestra la lista de remitos cuyo PDF ya existía
    (no se regeneró) y ofrece un botón para abrir la carpeta de salida."""

    def __init__(
        self,
        parent,
        comprobantes: list[str],
        open_folder_callback: Callable[[], None] | None = None,
        location_text: str = "",
    ):
        super().__init__(parent)
        self.title("Remitos ya generados")
        self.geometry("640x420")
        self.minsize(520, 320)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        n = len(comprobantes)
        header_text = (
            f"{n} remito(s) ya estaban generados — no se regeneraron.\n"
            f"Si querés volver a generarlos, eliminá el PDF existente primero."
        )
        if location_text:
            header_text += f"\n\nUbicación: {location_text}"

        ctk.CTkLabel(
            self, text=header_text,
            font=ctk.CTkFont(size=13, weight="bold"),
            justify="left", anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))

        body = ctk.CTkScrollableFrame(self)
        body.grid(row=1, column=0, sticky="nsew", padx=16, pady=4)
        body.grid_columnconfigure(0, weight=1)

        item_font = ctk.CTkFont(size=12)
        for i, comp in enumerate(comprobantes):
            ctk.CTkLabel(
                body, text=f"   • {comp}", font=item_font,
                justify="left", anchor="w",
            ).grid(row=i, column=0, sticky="ew", padx=6, pady=2)

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=12, pady=12)
        footer.grid_columnconfigure(0, weight=1)

        btns = ctk.CTkFrame(footer, fg_color="transparent")
        btns.grid(row=0, column=0, sticky="e")

        if open_folder_callback is not None:
            self.btn_open = ctk.CTkButton(
                btns, text="Abrir carpeta", width=140,
                fg_color="#2c5282", hover_color="#1f3c66",
                command=lambda: (open_folder_callback(),),
            )
            self.btn_open.grid(row=0, column=0, padx=4)

        self.btn_ok = ctk.CTkButton(
            btns, text="Continuar", width=120,
            command=self._on_close,
        )
        self.btn_ok.grid(row=0, column=1, padx=4)

        self.after(50, self._center_on_parent)

    def _center_on_parent(self) -> None:
        try:
            self.update_idletasks()
            parent = self.master
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _on_close(self) -> None:
        self.destroy()

    def show(self) -> None:
        self.wait_window()
