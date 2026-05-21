"""Widgets reutilizables para la GUI."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Iterable, Sequence

import customtkinter as ctk


class LogConsole(ctk.CTkTextbox):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("state", "disabled")
        super().__init__(master, **kwargs)

    def append(self, text: str) -> None:
        self.configure(state="normal")
        self.insert("end", text + "\n")
        self.see("end")
        self.configure(state="disabled")

    def clear(self) -> None:
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


class DataTable(ctk.CTkFrame):
    """Tabla genérica (ttk.Treeview embebido) con columnas dinámicas."""

    def __init__(
        self,
        master,
        columns: Sequence[str],
        widths: Sequence[int] | None = None,
        height: int = 8,
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._columns = list(columns)
        widths = list(widths or [120] * len(columns))

        self.tree = ttk.Treeview(self, columns=self._columns, show="headings", height=height)
        for col, w in zip(self._columns, widths):
            self.tree.heading(col, text=col)
            anchor = "e" if col.lower() in {"m2", "ancho", "largo", "#m2"} else "w"
            self.tree.column(col, width=w, anchor=anchor, stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Tag para resaltar filas con problema (ej. items sin protocolo)
        self.tree.tag_configure(
            "missing",
            foreground="#c0392b",
            font=("Segoe UI", 10, "bold"),
        )

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    def set_rows(
        self,
        rows: Sequence[dict],
        *,
        flag_predicate=None,
        flag_tag: str = "missing",
    ) -> None:
        for it in self.tree.get_children():
            self.tree.delete(it)
        for r in rows:
            values = [r.get(c, "") for c in self._columns]
            tags = (flag_tag,) if (flag_predicate and flag_predicate(r)) else ()
            self.tree.insert("", "end", values=values, tags=tags)

    def clear(self) -> None:
        for it in self.tree.get_children():
            self.tree.delete(it)


class MultiSelectListbox(ctk.CTkFrame):
    """Lista con multi-selección + buscador. Notifica vía callback al cambiar selección.

    Mantiene el conjunto FULL de items (universo) y filtra cuando se buscan o
    cuando se setean items mediante `set_items_with_filter`.
    """

    def __init__(
        self,
        master,
        title: str,
        on_select: Callable[[list[str]], None] | None = None,
        height: int = 12,
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._on_select = on_select
        self._all_items: list[str] = []
        self._visible_items: list[str] = []
        self._selected: set[str] = set()
        # Items que se pintan en verde (ej. remitos ya procesados por el bot).
        self._highlighted: set[str] = set()
        self._highlight_bg = "#dcfce7"  # verde claro
        self._highlight_fg = "#15803d"  # verde oscuro
        self._suppress_event = False

        self._lbl = ctk.CTkLabel(
            self, text=title, anchor="w",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        self._lbl.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))

        self._search_var = tk.StringVar()
        self._search = ctk.CTkEntry(
            self,
            placeholder_text="Buscar...",
            textvariable=self._search_var,
            font=ctk.CTkFont(size=14),
            height=36,
        )
        self._search.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 4))
        self._search_var.trace_add("write", lambda *_: self._refresh_visible())

        frame = tk.Frame(self, bd=0, highlightthickness=0)
        frame.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 6))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self._listbox = tk.Listbox(
            frame,
            selectmode="extended",
            exportselection=False,
            height=height,
            activestyle="dotbox",
            font=("Segoe UI", 13),
        )
        self._listbox.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._listbox.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self._listbox.configure(yscrollcommand=vsb.set)

        self._listbox.bind("<<ListboxSelect>>", self._on_listbox_select)

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 6))
        bottom.grid_columnconfigure(2, weight=1)

        self._btn_clear = ctk.CTkButton(bottom, text="Limpiar", width=80,
                                        command=self.clear_selection)
        self._btn_clear.grid(row=0, column=0, padx=(0, 4))

        self._lbl_count = ctk.CTkLabel(bottom, text="0/0 seleccionados", anchor="e")
        self._lbl_count.grid(row=0, column=2, sticky="e")

    # -------- API pública --------
    def set_items(self, items: Sequence[str]) -> None:
        self._all_items = list(items)
        self._selected.clear()
        # Nuevo "Buscar" → reset de highlights. El tracking se vuelve a cargar
        # asincrónicamente y repinta lo que corresponda.
        self._highlighted.clear()
        self._search_var.set("")
        self._refresh_visible()

    def set_highlighted(self, items: Iterable[str]) -> None:
        """Marca los items recibidos en verde (ej. remitos ya procesados por el bot).

        Items que no estén en `_all_items` se ignoran. No afecta la selección
        ni el filtro de búsqueda. Llamar varias veces sobrescribe.
        """
        self._highlighted = {str(x) for x in items}
        self._refresh_visible()

    def set_items_filtered(self, items: Sequence[str]) -> None:
        """Restringe el universo visible (cross-filter) sin perder selección
        si los items siguen presentes."""
        self._all_items = list(items)
        keep = self._selected & set(items)
        self._selected = set(keep)
        self._refresh_visible()

    def get_selected(self) -> list[str]:
        return sorted(self._selected)

    def clear_selection(self) -> None:
        if not self._selected:
            return
        self._selected.clear()
        self._refresh_visible()
        self._fire_event()

    # -------- internos --------
    def _refresh_visible(self) -> None:
        text = self._search_var.get().strip().lower()
        if text:
            self._visible_items = [x for x in self._all_items if text in str(x).lower()]
        else:
            self._visible_items = list(self._all_items)

        self._suppress_event = True
        self._listbox.delete(0, "end")
        for it in self._visible_items:
            self._listbox.insert("end", it)
        for i, it in enumerate(self._visible_items):
            if it in self._selected:
                self._listbox.selection_set(i)
            if it in self._highlighted:
                # Pinta el item en verde (fondo claro + texto oscuro).
                try:
                    self._listbox.itemconfig(
                        i, background=self._highlight_bg, foreground=self._highlight_fg,
                        selectbackground="#86efac", selectforeground="#14532d",
                    )
                except tk.TclError:
                    pass
        self._suppress_event = False
        self._lbl_count.configure(
            text=f"{len(self._selected)}/{len(self._all_items)} seleccionados"
        )

    def _on_listbox_select(self, _event=None) -> None:
        if self._suppress_event:
            return
        idxs = self._listbox.curselection()
        visible_selected = {self._visible_items[i] for i in idxs}
        text = self._search_var.get().strip().lower()
        if text:
            preserved = {s for s in self._selected if text not in str(s).lower()}
            self._selected = preserved | visible_selected
        else:
            self._selected = set(visible_selected)
        self._lbl_count.configure(
            text=f"{len(self._selected)}/{len(self._all_items)} seleccionados"
        )
        self._fire_event()

    def _fire_event(self) -> None:
        if self._on_select is not None:
            self._on_select(self.get_selected())
