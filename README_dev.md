# Stundenrechner – Entwicklerdokumentation

---

## Projektstruktur

```
Stundenrechner/
├── app.py              # Hauptanwendung (GUI)
├── database.py         # Datenbankzugriff (SQLite)
├── exporter.py         # Excel-Export (openpyxl)
├── requirements.txt    # Python-Abhängigkeiten
├── build.bat           # Build-Skript (PyInstaller)
├── README.md           # Benutzer-Dokumentation
├── README_dev.md       # Diese Datei
├── .venv/              # Virtuelle Umgebung (nicht versioniert)
├── build/              # PyInstaller Zwischendateien (nicht versioniert)
└── dist/               # Fertige EXE (nicht versioniert)
```

---

## Anforderungen

- Python 3.11+
- Abhängigkeiten: siehe `requirements.txt`

```
ttkbootstrap>=1.10.0
openpyxl>=3.1.0
```

---

## Entwicklungsumgebung einrichten

```bash
# Virtuelle Umgebung erstellen
python -m venv .venv

# Aktivieren (Windows PowerShell)
.venv\Scripts\Activate.ps1

# Aktivieren (Windows CMD)
.venv\Scripts\activate.bat

# Abhängigkeiten installieren
pip install -r requirements.txt
```

### App starten

```bash
python app.py
# oder mit explizitem venv-Python:
.venv\Scripts\python.exe app.py
```

---

## Architektur

### `database.py` – Datenbankschicht

Kapselt alle SQLite-Operationen. Die Datenbankdatei liegt unter:
`%APPDATA%\Stundenrechner\stundenrechner.db`

**Tabellen:**

| Tabelle | Beschreibung |
|---|---|
| `entries` | Stundeneinträge (date, task, hours, customer, commission) |
| `tasks` | Gespeicherte Aufgaben zur Wiederverwendung |
| `settings` | App-Einstellungen (aktuell: `user_name`) |

**Wichtig:** Die Methode `_migrate_entries()` sorgt für automatische Schemamigrationen bei bestehenden Datenbanken (z. B. neue Spalten). Neue Spalten immer dort eintragen.

**Zentrale Methoden:**

```python
db.add_entry(date_iso, task, hours, customer, commission) -> int
db.delete_entry(entry_id)
db.get_entries_by_date(date_iso)        # -> [(id, task, hours, customer, commission)]
db.get_entries_by_month(year, month)    # -> [(id, date, task, hours, customer, commission)]
db.get_daily_total(date_iso)            # -> float
db.get_monthly_total(year, month)       # -> float
db.get_all_tasks()                      # -> [str]
db.get_available_months()               # -> ["YYYY-MM"]
db.get_setting(key)                     # -> str | None
db.set_setting(key, value)
```

---

### `app.py` – GUI-Schicht

Basiert auf **ttkbootstrap** (Theme: `cosmo`), einem modernen Wrapper um tkinter.

**Klasse `StundenrechnerApp`:**

| Methode | Beschreibung |
|---|---|
| `_build_ui()` | Erstellt das gesamte UI |
| `_build_input_section()` | Eingabebereich (Datum, Kunde, Komissions-Nr., Aufgabe, Stunden) |
| `_build_entries_section()` | Tagesübersicht-Treeview |
| `_build_monthly_section()` | Monatsübersicht, Export-Pfad und Export-Button |
| `_ask_user_name()` | Namens-Dialog beim ersten Start |
| `_add_entry()` | Validiert und speichert neuen Eintrag |
| `_delete_entry()` | Löscht ausgewählten Eintrag |
| `_export_month()` | Startet den Excel-Export |
| `_refresh_all()` | Aktualisiert alle UI-Elemente |
| `_poll_date()` | Überwacht Datumsänderungen (alle 300 ms) |

**Eingabevalidierung:**  
Erfolgt direkt über tkinter `validatecommand`:
- Komissions-Nr.: nur Ganzzahlen (`str.isdigit()`)
- Stunden: Dezimalzahlen mit `.` oder `,`

---

### `exporter.py` – Export-Schicht

Erstellt formatierte `.xlsx`-Dateien mit **openpyxl**.

**Spalten in der exportierten Datei:** Datum | Kunde | Komissions-Nr. | Aufgabe | Stunden

**Besonderheiten:**
- Alternierende Zeilenfärbung für bessere Lesbarkeit
- Tagesgesamt-Zeile (grün hinterlegt) nach jedem Tag
- Monatsgesamt-Zeile (dunkelblau) am Ende
- Datum wird nur in der ersten Zeile eines Tages angezeigt (inkl. Wochentag)
- Druckbereich wird automatisch gesetzt

---

## Neues Feld hinzufügen (Anleitung)

1. **`database.py`**: Neue Spalte in `CREATE TABLE` + Migration in `_migrate_entries()` + Parameter in `add_entry()` + Spalte in `get_entries_by_date()` und `get_entries_by_month()` ergänzen
2. **`app.py`**: Eingabefeld in `_build_input_section()` + Treeview-Spalte in `_build_entries_section()` + Variable in `_add_entry()` lesen und an `db.add_entry()` übergeben + Treeview-Befüllung in `_load_entries()` anpassen
3. **`exporter.py`**: `NUM_COLS` erhöhen, Spaltenbreite, Header, Entpacken des Tupels und Zellenformatierung ergänzen

---

## Build (EXE erstellen)

```bash
# PyInstaller installieren
pip install pyinstaller

# Build ausführen
pyinstaller --onefile --windowed --name Stundenrechner --collect-all ttkbootstrap --clean app.py
```

Die fertige EXE liegt unter `dist\Stundenrechner.exe`.

**Oder:** `build.bat` ausführen – erledigt alle Schritte automatisch.

> **Hinweis:** Immer als normaler Benutzer (nicht als Administrator) builden, da PyInstaller 7.0 Admin-Builds blockiert.

---

## Datenbankpfad für Tests zurücksetzen

Um den Erststart-Dialog (Namenseingabe) erneut zu testen, entweder:

- Die Datenbank löschen: `%APPDATA%\Stundenrechner\stundenrechner.db`
- Oder den `user_name`-Eintrag direkt entfernen:

```bash
# PowerShell
& ".venv\Scripts\python.exe" -c "from database import Database; db = Database(); db.conn.execute('DELETE FROM settings WHERE key=''user_name'''); db.conn.commit()"
```
