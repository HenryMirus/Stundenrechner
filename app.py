"""
Stundenrechner Hauptanwendung mit grafischer Benutzeroberfläche.

Eine benutzerfreundliche App zum Erfassen, Verwalten und Exportieren
von Arbeitsstunden mit Aufgabenzuordnung.
Unterstützt mehrere Microsoft-Konten mit OneDrive-Export.
"""

import os
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from auth import MicrosoftAuth
from database import Database
from exporter import ExcelExporter
from onedrive import OneDriveClient

# Konstanten 

GERMAN_MONTHS = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

APP_TITLE = "Stundenrechner"
WINDOW_SIZE = (760, 920)
MIN_SIZE = (660, 920)
THEME = "cosmo"
FONT_FAMILY = "Segoe UI"


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# OneDrive Ordner-Auswahl Dialog
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

class OneDriveFolderDialog(ttk.Toplevel):
    """Dialog zum Auswählen eines OneDrive-Ordners."""

    def __init__(self, parent, onedrive: OneDriveClient):
        super().__init__(parent)
        self.title("OneDrive-Ordner auswählen")
        self.geometry("600x450")
        self.resizable(True, True)
        self.grab_set()
        self.transient(parent)

        self._onedrive = onedrive
        self._selected_id: str | None = None
        self._selected_name: str | None = None
        self._selected_path: str | None = None
        # Stack für Navigation: Liste von (id, name)
        self._nav_stack: list[tuple[str, str]] = [("root", "OneDrive")]

        self._build()
        self._load_children("root")
        self.place_window_center()

    def _build(self):
        frame = ttk.Frame(self, padding=15)
        frame.pack(fill=BOTH, expand=YES)

        # Pfad-Anzeige / Navigation
        top = ttk.Frame(frame)
        top.pack(fill=X, pady=(0, 8))

        self._back_btn = ttk.Button(
            top, text="Zurück", command=self._go_back,
            bootstyle="secondary-outline", width=10, state=DISABLED,
        )
        self._back_btn.pack(side=LEFT, padx=(0, 10))

        self._path_label = ttk.Label(
            top, text="OneDrive", font=(FONT_FAMILY, 10, "bold"),
            bootstyle="primary",
        )
        self._path_label.pack(side=LEFT, fill=X, expand=YES)

        # Ordnerliste
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=BOTH, expand=YES, pady=(0, 12))

        self._listbox = ttk.Treeview(
            list_frame, columns=("name", "items"), show="headings",
            height=12, bootstyle="primary",
        )
        self._listbox.heading("name", text="Ordnername", anchor=W)
        self._listbox.heading("items", text="Unterordner", anchor=CENTER)
        self._listbox.column("name", anchor=W, stretch=YES)
        self._listbox.column("items", width=100, anchor=CENTER, stretch=NO)

        sb = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=sb.set)
        self._listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        sb.pack(side=RIGHT, fill=Y)

        self._listbox.bind("<Double-1>", self._on_double_click)
        self._listbox.bind("<<TreeviewSelect>>", self._on_select)

        # Status
        self._status_label = ttk.Label(
            frame, text="", font=(FONT_FAMILY, 9), bootstyle="secondary",
        )
        self._status_label.pack(anchor=W, pady=(0, 10))

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=X)

        ttk.Button(
            btn_frame, text="Abbrechen", command=self.destroy,
            bootstyle="secondary-outline", width=12,
        ).pack(side=RIGHT, padx=(8, 0))

        self._select_btn = ttk.Button(
            btn_frame, text="Hier speichern", command=self._confirm_selection,
            bootstyle="success", width=20, state=DISABLED,
        )
        self._select_btn.pack(side=RIGHT)

    def _load_children(self, folder_id: str):
        """Lädt Unterordner asynchron."""
        self._status_label.configure(text="Lade Ordner")
        self._listbox.delete(*self._listbox.get_children())
        self._select_btn.configure(state=DISABLED)

        def _fetch():
            try:
                items = self._onedrive.list_folder_children(folder_id)
                self.after(0, lambda: self._populate(items))
            except Exception as exc:
                self.after(0, lambda: self._status_label.configure(
                    text=f"Fehler: {exc}"
                ))
        threading.Thread(target=_fetch, daemon=True).start()

    def _populate(self, items: list[dict]):
        self._listbox.delete(*self._listbox.get_children())
        for item in items:
            cnt = item["child_count"]
            cnt_str = str(cnt) if cnt else "â€“"
            self._listbox.insert("", END, iid=item["id"],
                                 values=(item["name"], cnt_str))
        current_name = self._nav_stack[-1][1]
        self._path_label.configure(
            text=" / ".join(n for _, n in self._nav_stack)
        )
        self._status_label.configure(
            text=f"{len(items)} Unterordner in '{current_name}'"
        )
        self._select_btn.configure(state=NORMAL)

    def _on_select(self, _=None):
        self._select_btn.configure(state=NORMAL)

    def _on_double_click(self, _=None):
        sel = self._listbox.selection()
        if not sel:
            return
        item_id = sel[0]
        name = self._listbox.item(item_id, "values")[0]
        self._nav_stack.append((item_id, name))
        self._back_btn.configure(state=NORMAL)
        self._load_children(item_id)

    def _go_back(self):
        if len(self._nav_stack) <= 1:
            return
        self._nav_stack.pop()
        if len(self._nav_stack) == 1:
            self._back_btn.configure(state=DISABLED)
        self._load_children(self._nav_stack[-1][0])

    def _confirm_selection(self):
        sel = self._listbox.selection()
        if sel:
            item_id = sel[0]
            folder_name = self._listbox.item(item_id, "values")[0]
            self._selected_id = item_id
            self._selected_name = folder_name
            self._selected_path = (
                " / ".join(n for _, n in self._nav_stack) + f" / {folder_name}"
            )
        else:
            current_id, current_name = self._nav_stack[-1]
            self._selected_id = current_id
            self._selected_name = current_name
            self._selected_path = " / ".join(n for _, n in self._nav_stack)
        self.destroy()

    @property
    def result(self) -> tuple[str | None, str | None, str | None]:
        """Gibt (folder_id, folder_name, folder_path) zurück."""
        return self._selected_id, self._selected_name, self._selected_path


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# Haupt-Anwendungsklasse
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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
        self.root.overrideredirect(False)
        self.root.attributes("-toolwindow", False)
        self.root.place_window_center()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._auth = MicrosoftAuth()
        self.db: Database | None = None
        self._onedrive: OneDriveClient | None = None
        self._user_info: dict | None = None
        self._month_map: dict[str, tuple[int, int]] = {}
        self._main_frame: ttk.Frame | None = None
        self._polling_active: bool = False

        self._show_login_screen()

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # Login-Screen
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

    def _clear_window(self):
        """Entfernt alle Widgets aus dem Hauptfenster."""
        for widget in self.root.winfo_children():
            widget.destroy()

    def _show_login_screen(self):
        """Zeigt den Anmelde-Bildschirm mit gespeicherten Konten."""
        self._clear_window()
        self.root.title(APP_TITLE)

        outer = ttk.Frame(self.root, padding=40)
        outer.pack(fill=BOTH, expand=YES)

        ttk.Label(
            outer,
            text="\u23F1  Stundenrechner",
            font=(FONT_FAMILY, 24, "bold"),
            bootstyle="primary",
        ).pack(pady=(0, 5))

        ttk.Label(
            outer,
            text="Bitte mit einem Microsoft-Konto anmelden",
            font=(FONT_FAMILY, 11),
            bootstyle="secondary",
        ).pack(pady=(0, 25))

        ttk.Separator(outer).pack(fill=X, pady=(0, 20))

        accounts = self._auth.get_accounts()

        if accounts:
            ttk.Label(
                outer,
                text="Gespeicherte Konten:",
                font=(FONT_FAMILY, 10, "bold"),
            ).pack(anchor=W, pady=(0, 8))

            accounts_frame = ttk.Frame(outer)
            accounts_frame.pack(fill=X, pady=(0, 20))

            for acc in accounts:
                self._build_account_row(accounts_frame, acc)

            ttk.Separator(outer).pack(fill=X, pady=(0, 20))

        ttk.Button(
            outer,
            text="\u2795  Neues Konto hinzufügen",
            command=self._login_new_account,
            bootstyle="primary",
            width=30,
        ).pack(pady=(0, 10))

        self._login_status = ttk.Label(
            outer, text="", font=(FONT_FAMILY, 9), bootstyle="danger",
        )
        self._login_status.pack(pady=(5, 0))

    def _build_account_row(self, parent: ttk.Frame, acc: dict):
        """Erzeugt eine Zeile für ein gespeichertes Konto."""
        row = ttk.Frame(parent, padding=(0, 4))
        row.pack(fill=X)

        username = acc.get("username", "Unbekanntes Konto")
        name = acc.get("name", "")
        display = f"{name}  ({username})" if name and name != username else username

        ttk.Button(
            row,
            text=f"\U0001F464  {display}",
            command=lambda a=acc: self._login_existing_account(a),
            bootstyle="outline",
            width=42,
        ).pack(side=LEFT, fill=X, expand=YES, padx=(0, 8))

        ttk.Button(
            row,
            text="\U0001F5D1",
            command=lambda a=acc: self._remove_account(a),
            bootstyle="danger-outline",
            width=4,
        ).pack(side=LEFT)

    def _login_existing_account(self, account: dict):
        """Meldet einen gespeicherten Benutzer still an."""
        self._login_status.configure(text="Anmeldung wird durchgeführt")
        self.root.update()
        if self._auth.switch_account(account):
            self._on_login_success()
        else:
            self._login_status.configure(
                text="Token abgelaufen, bitte erneut anmelden."
            )
            if self._auth.login_interactive():
                self._on_login_success()
            else:
                self._login_status.configure(
                    text="Anmeldung fehlgeschlagen. Bitte erneut versuchen."
                )

    def _login_new_account(self):
        """Öffnet den Browser für eine neue Microsoft-Anmeldung."""
        self._login_status.configure(text="Browser wird geöffnet")
        self.root.update()
        if self._auth.login_interactive():
            self._on_login_success()
        else:
            self._login_status.configure(
                text="Anmeldung abgebrochen oder fehlgeschlagen."
            )

    def _remove_account(self, account: dict):
        """Entfernt ein Konto aus dem lokalen Cache."""
        username = account.get("username", "dieses Konto")
        if messagebox.askyesno(
            "Konto entfernen",
            f"Möchten Sie '{username}' aus der Liste entfernen?\n\n"
            "Die gespeicherten Stunden bleiben erhalten.",
            parent=self.root,
        ):
            self._auth.logout(account)
            self._show_login_screen()

    def _on_login_success(self):
        """Wird nach erfolgreicher Anmeldung aufgerufen und initialisiert die App."""
        self._user_info = self._auth.get_user_info()
        uid_short = self._auth.current_user_id_short

        self.db = Database(user_id_short=uid_short)
        self._onedrive = OneDriveClient(self._auth)

        export_path = self.db.get_setting("export_path") or str(Path.home() / "Documents")
        self._default_export_path = export_path

        self._build_main_ui()

    # Haupt-UI

    def _build_main_ui(self):
        """Baut die komplette Haupt-Oberfläche nach dem Login auf."""
        self._clear_window()
        user_name = (self._user_info or {}).get("name", "")
        self.root.title(f"{APP_TITLE} {user_name}" if user_name else APP_TITLE)

        self._main_frame = ttk.Frame(self.root, padding=20)
        self._main_frame.pack(fill=BOTH, expand=YES)

        self._build_header()
        self._build_input_section()
        self._build_entries_section()
        self._build_monthly_section()

        self._refresh_all()
        self._start_date_polling()

    def _build_header(self):
        header_frame = ttk.Frame(self._main_frame)
        header_frame.pack(fill=X, pady=(0, 5))

        ttk.Label(
            header_frame,
            text="\u23F1  Stundenrechner",
            font=(FONT_FAMILY, 22, "bold"),
            bootstyle="primary",
        ).pack(side=LEFT)

        # Benutzer-Info & Abmelden (rechts)
        user_frame = ttk.Frame(header_frame)
        user_frame.pack(side=RIGHT)

        if self._user_info:
            name = self._user_info.get("name", "")
            email = self._user_info.get("email", "")
            if name:
                ttk.Label(
                    user_frame,
                    text=f"\U0001F464  {name}",
                    font=(FONT_FAMILY, 9, "bold"),
                    bootstyle="secondary",
                ).pack(anchor=E)
            if email:
                ttk.Label(
                    user_frame,
                    text=email,
                    font=(FONT_FAMILY, 8),
                    bootstyle="secondary",
                ).pack(anchor=E)

        ttk.Button(
            user_frame,
            text="Abmelden",
            command=self._logout,
            bootstyle="secondary-outline",
            width=10,
        ).pack(anchor=E, pady=(2, 0))

        ttk.Separator(self._main_frame).pack(fill=X, pady=(5, 15))

    def _logout(self):
        """Meldet den aktuellen Benutzer ab und zeigt den Login-Screen."""
        self._polling_active = False  # Datum-Polling stoppen
        if self.db:
            self.db.close()
            self.db = None
        self._onedrive = None
        self._user_info = None
        # Konto bleibt im Cache, nur aktive Sitzung beenden
        self._auth._current_account = None
        self._auth._current_token = None
        self._show_login_screen()

    # Eingabebereich

    def _build_input_section(self):
        frame = ttk.Labelframe(
            self._main_frame,
            text="  Neuer Eintrag  ",
            padding=15,
            bootstyle="primary",
        )
        frame.pack(fill=X, pady=(0, 15))
        frame.columnconfigure(1, weight=1)

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

        # Aufgabe
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

        ttk.Button(
            frame,
            text="\u2795  Eintrag hinzufügen",
            command=self._add_entry,
            bootstyle="success",
            width=25,
        ).grid(row=5, column=0, columnspan=2, pady=(12, 2))

    # Tagesübersicht

    def _build_entries_section(self):
        self.entries_lf = ttk.Labelframe(
            self._main_frame,
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

    # Monatsübersicht & Export

    def _build_monthly_section(self):
        frame = ttk.Labelframe(
            self._main_frame,
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
        self.monthly_total_label.pack(anchor=W, pady=(0, 12))

        # Export-Modus (Lokal / OneDrive)
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=X, pady=(0, 8))

        ttk.Label(
            mode_frame, text="Speicherort:", font=(FONT_FAMILY, 10)
        ).pack(side=LEFT, padx=(0, 12))

        self._export_mode = ttk.StringVar(
            value=self.db.get_setting("export_mode") or "local"
        )
        ttk.Radiobutton(
            mode_frame, text="Lokal", variable=self._export_mode,
            value="local", bootstyle="success",
            command=self._on_export_mode_change,
        ).pack(side=LEFT, padx=(0, 16))
        ttk.Radiobutton(
            mode_frame, text="OneDrive", variable=self._export_mode,
            value="onedrive", bootstyle="success",
            command=self._on_export_mode_change,
        ).pack(side=LEFT)

        # Lokaler Pfad
        self._local_path_frame = ttk.Frame(frame)
        self._local_path_frame.pack(fill=X, pady=(0, 6))

        ttk.Label(
            self._local_path_frame, text="Export-Ordner:", font=(FONT_FAMILY, 10)
        ).pack(side=LEFT, padx=(0, 10))

        self.export_path_var = ttk.StringVar(value=self._default_export_path)
        self.export_path_entry = ttk.Entry(
            self._local_path_frame, textvariable=self.export_path_var,
            font=(FONT_FAMILY, 10),
        )
        self.export_path_entry.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))

        ttk.Button(
            self._local_path_frame,
            text="\U0001F4C1  Durchsuchen",
            command=self._browse_export_path,
            bootstyle="secondary-outline",
            width=16,
        ).pack(side=LEFT)

        # OneDrive-Ordner
        self._onedrive_frame = ttk.Frame(frame)
        self._onedrive_frame.pack(fill=X, pady=(0, 6))

        ttk.Label(
            self._onedrive_frame, text="OneDrive-Ordner:", font=(FONT_FAMILY, 10)
        ).pack(side=LEFT, padx=(0, 10))

        saved_od_name = self.db.get_setting("onedrive_folder_name") or ""
        self._onedrive_folder_label = ttk.Label(
            self._onedrive_frame,
            text=saved_od_name if saved_od_name else "(noch kein Ordner gewählt)",
            font=(FONT_FAMILY, 10),
            bootstyle="success" if saved_od_name else "secondary",
        )
        self._onedrive_folder_label.pack(side=LEFT, fill=X, expand=YES, padx=(0, 10))

        ttk.Button(
            self._onedrive_frame,
            text="\U0001F4C1  Ordner wählen",
            command=self._pick_onedrive_folder,
            bootstyle="secondary-outline",
            width=16,
        ).pack(side=LEFT)

        # Quota-Warnung (wird asynchron befüllt)
        self._quota_label = ttk.Label(
            frame, text="", font=(FONT_FAMILY, 9),
        )
        self._quota_label.pack(anchor=W, pady=(0, 2))

        # Export-Aktion
        export_frame = ttk.Frame(frame)
        export_frame.pack(fill=X, pady=(8, 0))

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

        # Initiale Sichtbarkeit
        self._on_export_mode_change()

    def _on_export_mode_change(self):
        """Zeigt lokalen oder OneDrive-Pfad je nach gewähltem Modus."""
        mode = self._export_mode.get()
        if mode == "local":
            self._onedrive_frame.pack_forget()
            self._quota_label.pack_forget()
            self._local_path_frame.pack(fill=X, pady=(0, 6))
        else:
            self._local_path_frame.pack_forget()
            self._onedrive_frame.pack(fill=X, pady=(0, 6))
            self._quota_label.pack(anchor=W, pady=(0, 2))
            self._check_quota_async()
        if self.db:
            self.db.set_setting("export_mode", mode)

    def _check_quota_async(self):
        """Prüft die OneDrive-Quota im Hintergrund und zeigt eine Warnung."""
        self._quota_label.configure(
            text="OneDrive-Speicher wird geprüft…", bootstyle="secondary"
        )
        def _fetch():
            try:
                quota = self._onedrive.get_quota_info() if self._onedrive else None
            except Exception:
                quota = None
            self.root.after(0, lambda: self._update_quota_label(quota))
        threading.Thread(target=_fetch, daemon=True).start()

    def _update_quota_label(self, quota: dict | None):
        """Aktualisiert das Quota-Label mit den abgerufenen Informationen."""
        # Widget könnte durch Logout/Neuaufbau bereits zerstört sein
        try:
            if not self._quota_label.winfo_exists():
                return
        except Exception:
            return
        if not quota:
            self._quota_label.configure(text="", bootstyle="secondary")
            return
        used_gb = quota["used"] / 1_073_741_824
        total_gb = quota["total"] / 1_073_741_824
        state = quota["state"]
        if state == "exceeded":
            self._quota_label.configure(
                text=f"⚠  OneDrive VOLL: {used_gb:.1f} GB von {total_gb:.1f} GB genutzt – Upload nicht möglich!",
                bootstyle="danger",
            )
        elif state == "nearing":
            self._quota_label.configure(
                text=f"⚠  OneDrive fast voll: {used_gb:.1f} GB von {total_gb:.1f} GB genutzt",
                bootstyle="warning",
            )
        else:
            rem_gb = quota["remaining"] / 1_073_741_824
            self._quota_label.configure(
                text=f"✓  OneDrive: {used_gb:.1f} GB von {total_gb:.1f} GB genutzt  ({rem_gb:.1f} GB frei)",
                bootstyle="success",
            )

    # Hilfsmethoden Datum

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

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # Geschäftslogik
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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
        if messagebox.askyesno(
            "Eintrag löschen", "Möchten Sie diesen Eintrag wirklich löschen?"
        ):
            self.db.delete_entry(entry_id)
            self._refresh_all()

    def _browse_export_path(self):
        """Öffnet einen Ordner-Dialog zur Auswahl des lokalen Export-Pfads."""
        current = self.export_path_var.get()
        initial = current if os.path.isdir(current) else self._default_export_path
        folder = filedialog.askdirectory(
            title="Export-Ordner auswählen",
            initialdir=initial,
        )
        if folder:
            self.export_path_var.set(folder)
            self.db.set_setting("export_path", folder)

    def _pick_onedrive_folder(self):
        """Öffnet den OneDrive-Ordner-Browser-Dialog."""
        dialog = OneDriveFolderDialog(self.root, self._onedrive)
        self.root.wait_window(dialog)
        folder_id, folder_name, folder_path = dialog.result
        if folder_id and folder_name:
            self.db.set_setting("onedrive_folder_id", folder_id)
            self.db.set_setting("onedrive_folder_name", folder_path or folder_name)
            self._onedrive_folder_label.configure(
                text=folder_path or folder_name,
                bootstyle="success",
            )

    def _export_month(self):
        month_display = self.export_month_var.get()
        if not month_display or month_display not in self._month_map:
            messagebox.showerror("Fehler", "Bitte einen Monat auswählen.")
            return

        year, month = self._month_map[month_display]
        entries = self.db.get_entries_by_month(year, month)

        if not entries:
            messagebox.showinfo(
                "Keine Daten", f"Keine Einträge für {month_display} vorhanden."
            )
            return

        user_name = (self._user_info or {}).get("name", "")
        safe_name = user_name.replace(" ", "_") if user_name else "Export"
        filename = f"Stundenzettel_{GERMAN_MONTHS[month - 1]}_{year}_{safe_name}.xlsx"

        if self._export_mode.get() == "onedrive":
            self._export_to_onedrive(entries, year, month, filename)
        else:
            self._export_locally(entries, year, month, filename)

    def _export_locally(self, entries, year: int, month: int, filename: str):
        """Exportiert die Datei lokal ins Dateisystem."""
        export_dir = self.export_path_var.get().strip()
        if not export_dir or not os.path.isdir(export_dir):
            messagebox.showerror(
                "Fehler", "Bitte einen gültigen Export-Ordner auswählen."
            )
            return

        filepath = os.path.join(export_dir, filename)
        if os.path.exists(filepath):
            if not messagebox.askyesno(
                "Datei existiert",
                f"Die Datei '{filename}' existiert bereits.\nÜberschreiben?",
            ):
                return

        try:
            ExcelExporter.export(entries, year, month, filepath)
            messagebox.showinfo(
                "Export erfolgreich",
                f"Stundenübersicht wurde exportiert:\n\n{filepath}",
            )
        except Exception as exc:
            messagebox.showerror("Export-Fehler", f"Fehler beim Exportieren:\n{exc}")

    def _export_to_onedrive(self, entries, year: int, month: int, filename: str):
        """Exportiert die Datei nach OneDrive."""
        folder_id = self.db.get_setting("onedrive_folder_id")
        folder_name = self.db.get_setting("onedrive_folder_name")

        if not folder_id:
            messagebox.showerror(
                "Fehler", "Bitte zuerst einen OneDrive-Ordner auswählen."
            )
            return

        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp_path = tmp.name
            ExcelExporter.export(entries, year, month, tmp_path)
            self._onedrive.upload_file(tmp_path, folder_id, filename)
        except Exception as exc:
            messagebox.showerror("Export-Fehler", f"Fehler beim Exportieren:\n{exc}")
            return
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

        web_url = self._onedrive.get_file_web_url(folder_id, filename)
        msg = (
            f"Stundenübersicht wurde erfolgreich nach OneDrive hochgeladen:\n\n"
            f"\U0001F4C1 {folder_name or 'OneDrive'} / {filename}"
        )
        if web_url:
            msg += f"\n\n{web_url}"
        messagebox.showinfo("Export erfolgreich", msg)

    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # UI-Aktualisierung
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

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
            self.tree.insert(
                "", END, iid=str(entry_id),
                values=(customer, commission, task, f"{hours:.2f}"),
            )

        total = self.db.get_daily_total(date_iso)
        formatted = f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.daily_total_label.configure(text=f"Tagesgesamt: {formatted} Std")

        date_display = self._get_date_display()
        self.entries_lf.configure(text=f"  Tagesübersicht {date_display}  ")
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

    # Events 

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        self.delete_btn.configure(state=NORMAL if selected else DISABLED)

    def _start_date_polling(self):
        self._polling_active = True
        self._last_date_str = self.date_entry.entry.get()
        self._poll_date()

    def _poll_date(self):
        if not self._polling_active:
            return
        try:
            current = self.date_entry.entry.get()
        except Exception:
            self._polling_active = False
            return
        if current != self._last_date_str:
            self._last_date_str = current
            self._load_entries()
            self._update_monthly_info()
            self._update_export_months()
        self.root.after(300, self._poll_date)

    def _on_close(self):
        if self.db:
            self.db.close()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = StundenrechnerApp()
    app.run()

