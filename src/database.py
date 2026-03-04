"""
Datenbank-Modul für den Stundenrechner.
Verwaltet alle SQLite-Operationen für Einträge und Aufgaben.
Unterstützt benutzer-spezifische Datenbanken (eine DB pro Microsoft-Konto).
"""

import sqlite3
import os


class Database:
    """SQLite-Datenbank für Stundeneinträge und gespeicherte Aufgaben."""

    def __init__(self, user_id_short: str | None = None):
        """
        Initialisiert die Datenbankverbindung.

        Args:
            user_id_short: Kurz-Hash der Microsoft-User-ID (12 Zeichen).
                           Wenn angegeben, wird eine benutzer-spezifische DB verwendet.
                           Wenn None, wird die Legacy-Datenbank 'stundenrechner.db' genutzt.
        """
        app_dir = os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")),
            "Stundenrechner"
        )
        os.makedirs(app_dir, exist_ok=True)
        if user_id_short:
            db_filename = f"stundenrechner_{user_id_short}.db"
        else:
            db_filename = "stundenrechner.db"
        self.db_path = os.path.join(app_dir, db_filename)
        self.conn = sqlite3.connect(self.db_path)
        self._create_tables()

    def _create_tables(self):
        """Erstellt die Datenbanktabellen, falls sie nicht existieren."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                task TEXT NOT NULL,
                hours REAL NOT NULL,
                customer TEXT NOT NULL DEFAULT '',
                commission TEXT NOT NULL DEFAULT ''
            )
        """)
        # Migration: Spalten hinzufügen falls sie fehlen
        self._migrate_entries(cursor)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        self.conn.commit()

    def _migrate_entries(self, cursor):
        """Fügt neue Spalten hinzu, falls sie in einer älteren DB fehlen."""
        cursor.execute("PRAGMA table_info(entries)")
        columns = {row[1] for row in cursor.fetchall()}
        if "customer" not in columns:
            cursor.execute("ALTER TABLE entries ADD COLUMN customer TEXT NOT NULL DEFAULT ''")
        if "commission" not in columns:
            cursor.execute("ALTER TABLE entries ADD COLUMN commission TEXT NOT NULL DEFAULT ''")

    # ── Einstellungen ─────────────────────────────────────────

    def get_setting(self, key: str) -> str | None:
        """Gibt den Wert einer Einstellung zurück oder None."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row[0] if row else None

    def set_setting(self, key: str, value: str):
        """Speichert eine Einstellung (überschreibt vorhandene)."""
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, value),
        )
        self.conn.commit()

    # ── Einträge ──────────────────────────────────────────────

    def add_entry(self, date_iso: str, task: str, hours: float,
                  customer: str = "", commission: str = "") -> int:
        """Fügt einen neuen Stundeneintrag hinzu und speichert die Aufgabe."""
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO entries (date, task, hours, customer, commission) "
            "VALUES (?, ?, ?, ?, ?)",
            (date_iso, task, hours, customer, commission),
        )
        # Aufgabe automatisch für spätere Wiederverwendung speichern
        cursor.execute(
            "INSERT OR IGNORE INTO tasks (name) VALUES (?)",
            (task,),
        )
        self.conn.commit()
        return cursor.lastrowid

    def delete_entry(self, entry_id: int):
        """Löscht einen Eintrag anhand seiner ID."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM entries WHERE id = ?", (entry_id,))
        self.conn.commit()

    def get_entries_by_date(self, date_iso: str) -> list:
        """Gibt alle Einträge für ein bestimmtes Datum zurück."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT id, task, hours, customer, commission FROM entries WHERE date = ? ORDER BY id",
            (date_iso,),
        )
        return cursor.fetchall()

    def get_entries_by_month(self, year: int, month: int) -> list:
        """Gibt alle Einträge für einen bestimmten Monat zurück."""
        cursor = self.conn.cursor()
        prefix = f"{year:04d}-{month:02d}"
        cursor.execute(
            "SELECT id, date, task, hours, customer, commission FROM entries "
            "WHERE date LIKE ? ORDER BY date, id",
            (prefix + "%",),
        )
        return cursor.fetchall()

    # ── Summen ────────────────────────────────────────────────

    def get_daily_total(self, date_iso: str) -> float:
        """Berechnet die Gesamtstunden für einen Tag."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT COALESCE(SUM(hours), 0) FROM entries WHERE date = ?",
            (date_iso,),
        )
        return cursor.fetchone()[0]

    def get_monthly_total(self, year: int, month: int) -> float:
        """Berechnet die Gesamtstunden für einen Monat."""
        cursor = self.conn.cursor()
        prefix = f"{year:04d}-{month:02d}"
        cursor.execute(
            "SELECT COALESCE(SUM(hours), 0) FROM entries WHERE date LIKE ?",
            (prefix + "%",),
        )
        return cursor.fetchone()[0]

    # ── Aufgaben ──────────────────────────────────────────────

    def get_all_tasks(self) -> list:
        """Gibt alle gespeicherten Aufgabennamen zurück."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM tasks ORDER BY name")
        return [row[0] for row in cursor.fetchall()]

    def get_all_customers(self) -> list:
        """Gibt alle einzigartigen Kundennamen zurück."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT DISTINCT customer FROM entries "
            "WHERE customer != '' ORDER BY customer"
        )
        return [row[0] for row in cursor.fetchall()]

    def get_all_commissions(self) -> list:
        """Gibt alle einzigartigen Komissionsnummern zurück."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT DISTINCT commission FROM entries "
            "WHERE commission != '' ORDER BY commission"
        )
        return [row[0] for row in cursor.fetchall()]

    # ── Verfügbare Monate ─────────────────────────────────────

    def get_available_months(self) -> list:
        """Gibt alle Monate zurück, für die Einträge existieren."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT DISTINCT substr(date, 1, 7) AS ym "
            "FROM entries ORDER BY ym DESC"
        )
        return [row[0] for row in cursor.fetchall()]

    # ── Verbindung ────────────────────────────────────────────

    def close(self):
        """Schließt die Datenbankverbindung."""
        self.conn.close()
