"""
Stundenrechner – Hauptanwendung mit grafischer Benutzeroberfläche.

Eine benutzerfreundliche App zum Erfassen, Verwalten und Exportieren
von Arbeitsstunden mit Aufgabenzuordnung.
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from database import Database
from exporter import ExcelExporter

# ── Konstanten ────────────────────────────────────────────────

GERMAN_MONTHS = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

APP_TITLE = "Stundenrechner"
WINDOW_SIZE = (760, 840)
MIN_SIZE = (660, 720)
THEME = "cosmo"
FONT_FAMILY = "Segoe UI"


class StundenrechnerApp:
    """Hauptklasse für die Stundenrechner-Anwendung."""

    def __init__(self):
        self.root = ttk.Window(
            title=APP_TITLE,
            themename=THEME,
            size=WINDOW_SIZE,
            resizable=(True, True),
            minsize=MIN_SIZE,
        )
        # Native Titelleiste mit Schließen/Minimieren/Maximieren erzwingen
        self.root.overrideredirect(False)
        self.root.attributes("-toolwindow", False)
        self.root.place_window_center()

        self.db = Database()
        self._month_map: dict[str, tuple[int, int]] = {}
        self._default_export_path = str(Path.home() / "Documents")
        self._user_name = self.db.get_setting("user_name") or ""

        # Beim ersten Start nach dem Namen fragen
        if not self._user_name:
            self._ask_user_name()

        self._build_ui()
        self._refresh_all()
        self._start_date_polling()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ══════════════════════════════════════════════════════════
    # UI-Aufbau
    # ══════════════════════════════════════════════════════════

    def _build_ui(self):
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=BOTH, expand=YES)

        self._build_header()
        self._build_input_section()
        self._build_entries_section()
        self._build_monthly_section()

    def _build_header(self):
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=X, pady=(0, 5))

        ttk.Label(
            header_frame,
            text="\u23F1  Stundenrechner",
            font=(FONT_FAMILY, 22, "bold"),
            bootstyle="primary",
        ).pack(side=LEFT)

        ttk.Separator(self.main_frame).pack(fill=X, pady=(0, 15))

    # ── Eingabebereich ────────────────────────────────────────

    def _build_input_section(self):
        frame = ttk.Labelframe(
            self.main_frame,
            text="  Neuer Eintrag  ",
            padding=15,
            bootstyle="primary",
        )
        frame.pack(fill=X, pady=(0, 15))
        frame.columnconfigure(1, weight=1)

        # Validierungsfunktionen
        _vcmd_int = (frame.register(lambda s: s == "" or s.isdigit()), "%P")
        _vcmd_num = (frame.register(
            lambda s: s == "" or all(c in "0123456789.," for c in s)
            and s.count(".") + s.count(",") <= 1
        ), "%P")

        # Datum
        ttk.Label(frame, text="Datum:", font=(FONT_FAMILY, 10), width=14).grid(
            row=0, column=0, sticky=W, padx=(0, 10), pady=6
        )
        self.date_entry = ttk.DateEntry(
            frame, bootstyle="primary", dateformat="%d.%m.%Y", firstweekday=0
        )
        self.date_entry.grid(row=0, column=1, sticky=EW, pady=6)

        # Kunde
        ttk.Label(frame, text="Kunde:", font=(FONT_FAMILY, 10), width=14).grid(
            row=1, column=0, sticky=W, padx=(0, 10), pady=6
        )
        self.customer_var = ttk.StringVar()
        self.customer_entry = ttk.Entry(
            frame, textvariable=self.customer_var, font=(FONT_FAMILY, 10)
        )
        self.customer_entry.grid(row=1, column=1, sticky=EW, pady=6)

        # Komissionsnummer
        ttk.Label(frame, text="Komissions-Nr.:", font=(FONT_FAMILY, 10), width=14).grid(
            row=2, column=0, sticky=W, padx=(0, 10), pady=6
        )
        self.commission_var = ttk.StringVar()
        self.commission_entry = ttk.Entry(
            frame, textvariable=self.commission_var, font=(FONT_FAMILY, 10),
            validate="key", validatecommand=_vcmd_int,
        )
        self.commission_entry.grid(row=2, column=1, sticky=EW, pady=6)

        # Aufgabe (Combobox mit Autovervollständigung)
        ttk.Label(frame, text="Aufgabe:", font=(FONT_FAMILY, 10), width=14).grid(
            row=3, column=0, sticky=W, padx=(0, 10), pady=6
        )
        self.task_var = ttk.StringVar()
        self.task_combo = ttk.Combobox(
            frame, textvariable=self.task_var, font=(FONT_FAMILY, 10)
        )
        self.task_combo.grid(row=3, column=1, sticky=EW, pady=6)

        # Stunden
        ttk.Label(frame, text="Stunden:", font=(FONT_FAMILY, 10), width=14).grid(
            row=4, column=0, sticky=W, padx=(0, 10), pady=6
        )
        self.hours_var = ttk.StringVar()
        self.hours_entry = ttk.Entry(
            frame, textvariable=self.hours_var, font=(FONT_FAMILY, 10),
            validate="key", validatecommand=_vcmd_num,
        )
        self.hours_entry.grid(row=4, column=1, sticky=EW, pady=6)
        self.hours_entry.bind("<Return>", lambda _: self._add_entry())

        # Hinzufügen-Button
        ttk.Button(
            frame,
            text="\u2795  Eintrag hinzufügen",
            command=self._add_entry,
            bootstyle="success",
            width=25,
        ).grid(row=5, column=0, columnspan=2, pady=(12, 2))

    # ── Tagesübersicht ────────────────────────────────────────

    def _build_entries_section(self):
        self.entries_lf = ttk.Labelframe(
            self.main_frame,
            text="  Tagesübersicht  ",
            padding=15,
            bootstyle="info",
        )
        self.entries_lf.pack(fill=BOTH, expand=YES, pady=(0, 15))

        tree_frame = ttk.Frame(self.entries_lf)
        tree_frame.pack(fill=BOTH, expand=YES)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("customer", "commission", "task", "hours"),
            show="headings",
            height=8,
            bootstyle="info",
        )
        self.tree.heading("customer", text="Kunde", anchor=W)
        self.tree.heading("commission", text="Komissions-Nr.", anchor=W)
        self.tree.heading("task", text="Aufgabe", anchor=W)
        self.tree.heading("hours", text="Stunden", anchor=CENTER)
        self.tree.column("customer", width=140, minwidth=80, anchor=W)
        self.tree.column("commission", width=120, minwidth=80, anchor=W)
        self.tree.column("task", width=260, minwidth=120, anchor=W)
        self.tree.column("hours", width=80, minwidth=60, anchor=CENTER)

        tree_scroll = ttk.Scrollbar(
            tree_frame, orient=VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)
        tree_scroll.pack(side=RIGHT, fill=Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Delete>", lambda _: self._delete_entry())

        # Unterer Bereich
        bottom = ttk.Frame(self.entries_lf)
        bottom.pack(fill=X, pady=(10, 0))

        self.delete_btn = ttk.Button(
            bottom,
            text="\U0001F5D1  Eintrag löschen",
            command=self._delete_entry,
            bootstyle="danger-outline",
            state=DISABLED,
        )
        self.delete_btn.pack(side=LEFT)

        self.daily_total_label = ttk.Label(
            bottom,
            text="Tagesgesamt: 0,00 Std",
            font=(FONT_FAMILY, 11, "bold"),
            bootstyle="info",
        )
        self.daily_total_label.pack(side=RIGHT)

    # ── Monatsübersicht & Export ──────────────────────────────

    def _build_monthly_section(self):
        frame = ttk.Labelframe(
            self.main_frame,
            text="  Monatsübersicht & Export  ",
            padding=15,
            bootstyle="success",
        )
        frame.pack(fill=X)

        self.monthly_total_label = ttk.Label(
            frame,
            text="Monatsstunden: 0,00 Std",
            font=(FONT_FAMILY, 14, "bold"),
            bootstyle="success",
        )
        self.monthly_total_label.pack(anchor=W, pady=(0, 15))

        # Export-Pfad
        path_frame = ttk.Frame(frame)
        path_frame.pack(fill=X, pady=(0, 10))

        ttk.Label(
            path_frame, text="Export-Pfad:", font=(FONT_FAMILY, 10)
        ).pack(side=LEFT, padx=(0, 10))

        self.export_path_var = ttk.StringVar(value=self._default_export_path)
        self.export_path_entry = ttk.Entry(
            path_frame, textvariable=self.export_path_var, font=(FONT_FAMILY, 10)
        )
        self.export_path_entry.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))

        ttk.Button(
            path_frame,
            text="\U0001F4C1  Durchsuchen",
            command=self._browse_export_path,
            bootstyle="secondary-outline",
            width=16,
        ).pack(side=LEFT)

        # Export-Zeile
        export_frame = ttk.Frame(frame)
        export_frame.pack(fill=X)

        ttk.Label(
            export_frame, text="Monat exportieren:", font=(FONT_FAMILY, 10)
        ).pack(side=LEFT, padx=(0, 10))

        self.export_month_var = ttk.StringVar()
        self.export_month_combo = ttk.Combobox(
            export_frame,
            textvariable=self.export_month_var,
            state="readonly",
            font=(FONT_FAMILY, 10),
            width=18,
        )
        self.export_month_combo.pack(side=LEFT, padx=(0, 15))

        ttk.Button(
            export_frame,
            text="\U0001F4CA  Als Excel exportieren",
            command=self._export_month,
            bootstyle="success",
            width=25,
        ).pack(side=LEFT)

    # ══════════════════════════════════════════════════════════
    # Benutzername
    # ══════════════════════════════════════════════════════════

    def _ask_user_name(self):
        """Fragt beim ersten Start nach dem vollständigen Namen."""
        import tkinter.simpledialog as simpledialog
        while True:
            name = simpledialog.askstring(
                "Willkommen",
                "Bitte geben Sie Ihren vollständigen Namen ein:",
                parent=self.root,
            )
            if name and name.strip():
                self._user_name = name.strip()
                self.db.set_setting("user_name", self._user_name)
                return
            messagebox.showwarning(
                "Name erforderlich",
                "Bitte geben Sie einen Namen ein, damit die\n"
                "exportierten Dateien korrekt benannt werden.",
            )

    # ══════════════════════════════════════════════════════════
    # Hilfsmethoden – Datum
    # ══════════════════════════════════════════════════════════

    def _get_date_iso(self) -> str | None:
        try:
            raw = self.date_entry.entry.get()
            dt = datetime.strptime(raw, "%d.%m.%Y")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            return None

    def _get_date_display(self) -> str:
        return self.date_entry.entry.get()

    def _get_year_month(self) -> tuple[int | None, int | None]:
        try:
            raw = self.date_entry.entry.get()
            dt = datetime.strptime(raw, "%d.%m.%Y")
            return dt.year, dt.month
        except ValueError:
            return None, None

    # ══════════════════════════════════════════════════════════
    # Geschäftslogik
    # ══════════════════════════════════════════════════════════

    def _add_entry(self):
        date_iso = self._get_date_iso()
        if not date_iso:
            messagebox.showerror("Fehler", "Bitte ein gültiges Datum auswählen.")
            return

        task = self.task_var.get().strip()
        if not task:
            messagebox.showerror("Fehler", "Bitte eine Aufgabe eingeben.")
            self.task_combo.focus_set()
            return

        hours_str = self.hours_var.get().strip().replace(",", ".")
        try:
            hours = float(hours_str)
            if hours <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Fehler",
                "Bitte gültige Stunden eingeben (z.\u202fB. 1,5 oder 2.25).",
            )
            self.hours_entry.focus_set()
            return

        customer = self.customer_var.get().strip()
        commission = self.commission_var.get().strip()

        self.db.add_entry(date_iso, task, hours, customer, commission)
        self.hours_var.set("")
        self.task_var.set("")
        self.customer_var.set("")
        self.commission_var.set("")
        self._refresh_all()
        self.customer_entry.focus_set()

    def _delete_entry(self):
        selected = self.tree.selection()
        if not selected:
            return
        entry_id = int(selected[0])
        if messagebox.askyesno("Eintrag löschen", "Möchten Sie diesen Eintrag wirklich löschen?"):
            self.db.delete_entry(entry_id)
            self._refresh_all()

    def _browse_export_path(self):
        """Öffnet einen Ordner-Dialog zur Auswahl des Export-Pfads."""
        current = self.export_path_var.get()
        initial = current if os.path.isdir(current) else self._default_export_path
        folder = filedialog.askdirectory(
            title="Export-Ordner auswählen",
            initialdir=initial,
        )
        if folder:
            self.export_path_var.set(folder)

    def _export_month(self):
        month_display = self.export_month_var.get()
        if not month_display or month_display not in self._month_map:
            messagebox.showerror("Fehler", "Bitte einen Monat auswählen.")
            return

        year, month = self._month_map[month_display]
        entries = self.db.get_entries_by_month(year, month)

        if not entries:
            messagebox.showinfo("Keine Daten", f"Keine Einträge für {month_display} vorhanden.")
            return

        # Export-Pfad prüfen
        export_dir = self.export_path_var.get().strip()
        if not export_dir or not os.path.isdir(export_dir):
            messagebox.showerror("Fehler", "Bitte einen gültigen Export-Ordner auswählen.")
            return

        safe_name = self._user_name.replace(" ", "_")
        filename = f"Stundenzettel_{GERMAN_MONTHS[month - 1]}_{year}_{safe_name}.xlsx"
        filepath = os.path.join(export_dir, filename)

        # Bei bestehender Datei nachfragen
        if os.path.exists(filepath):
            if not messagebox.askyesno(
                "Datei existiert",
                f"Die Datei '{filename}' existiert bereits.\nÜberschreiben?",
            ):
                return

        try:
            ExcelExporter.export(entries, year, month, filepath)
            messagebox.showinfo("Export erfolgreich", f"Stundenübersicht wurde exportiert:\n\n{filepath}")
        except Exception as exc:
            messagebox.showerror("Export-Fehler", f"Fehler beim Exportieren:\n{exc}")

    # ══════════════════════════════════════════════════════════
    # UI-Aktualisierung
    # ══════════════════════════════════════════════════════════

    def _refresh_all(self):
        self._load_entries()
        self._update_monthly_info()
        self._update_task_list()
        self._update_export_months()

    def _load_entries(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        date_iso = self._get_date_iso()
        if not date_iso:
            return

        entries = self.db.get_entries_by_date(date_iso)
        for entry_id, task, hours, customer, commission in entries:
            self.tree.insert("", END, iid=str(entry_id),
                            values=(customer, commission, task, f"{hours:.2f}"))

        total = self.db.get_daily_total(date_iso)
        formatted = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.daily_total_label.configure(text=f"Tagesgesamt: {formatted} Std")

        date_display = self._get_date_display()
        self.entries_lf.configure(text=f"  Tagesübersicht – {date_display}  ")
        self.delete_btn.configure(state=DISABLED)

    def _update_monthly_info(self):
        year, month = self._get_year_month()
        if year is None or month is None:
            return
        total = self.db.get_monthly_total(year, month)
        month_name = GERMAN_MONTHS[month - 1]
        formatted = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.monthly_total_label.configure(
            text=f"Monatsstunden ({month_name} {year}): {formatted} Std",
        )

    def _update_task_list(self):
        tasks = self.db.get_all_tasks()
        self.task_combo["values"] = tasks

    def _update_export_months(self):
        available = self.db.get_available_months()

        now = datetime.now()
        current_ym = f"{now.year:04d}-{now.month:02d}"
        if current_ym not in available:
            available.append(current_ym)

        year, month = self._get_year_month()
        if year and month:
            selected_ym = f"{year:04d}-{month:02d}"
            if selected_ym not in available:
                available.append(selected_ym)

        available.sort(reverse=True)

        self._month_map.clear()
        display_values = []
        for ym in available:
            y, m = int(ym[:4]), int(ym[5:7])
            display = f"{GERMAN_MONTHS[m - 1]} {y}"
            display_values.append(display)
            self._month_map[display] = (y, m)

        self.export_month_combo["values"] = display_values

        current_selection = self.export_month_var.get()
        if current_selection not in display_values and display_values:
            self.export_month_combo.set(display_values[0])

    # ── Events ────────────────────────────────────────────────

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        self.delete_btn.configure(state=NORMAL if selected else DISABLED)

    def _start_date_polling(self):
        self._last_date_str = self.date_entry.entry.get()
        self._poll_date()

    def _poll_date(self):
        current = self.date_entry.entry.get()
        if current != self._last_date_str:
            self._last_date_str = current
            self._load_entries()
            self._update_monthly_info()
            self._update_export_months()
        self.root.after(300, self._poll_date)

    def _on_close(self):
        self.db.close()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = StundenrechnerApp()
    app.run()
