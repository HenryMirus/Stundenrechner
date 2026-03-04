# Stundenrechner – Entwicklerdokumentation

---

## Projektstruktur

```
stundenrechner/
├── main.py             # Einstiegspunkt (fügt src/ zum Suchpfad hinzu)
├── src/
│   ├── app.py          # Hauptanwendung (GUI + Login)
│   ├── auth.py         # Microsoft-Authentifizierung (MSAL)
│   ├── onedrive.py     # OneDrive-Integration (Graph API)
│   ├── database.py     # Datenbankzugriff (SQLite, pro Benutzer)
│   └── exporter.py     # Excel-Export (openpyxl)
├── requirements.txt    # Python-Abhängigkeiten
├── build.bat           # Build-Skript (PyInstaller)
├── Stundenrechner.spec # PyInstaller-Spec-Datei
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
msal>=1.24.0
requests>=2.31.0
```

### Azure App Registration (Voraussetzung)

Für die Microsoft-Anmeldung ist eine Azure App Registration erforderlich:

1. [Azure Portal](https://portal.azure.com) → **App-Registrierungen → Neue Registrierung**
2. Kontotyp: **Konten in beliebigen Organisationsverzeichnissen + persönliche Konten**
3. Umleitungs-URI: **Öffentlicher Client/nativ** → `http://localhost`
4. API-Berechtigungen (delegiert): **User.Read**, **Files.ReadWrite**
5. Client-ID in `src/auth.py` → `MS_CLIENT_ID` eintragen

---

## Entwicklungsumgebung einrichten

```bash
# Virtuelle Umgebung erstellen
python -m venv .venv

# Aktivieren (Windows PowerShell)
.venv\Scripts\Activate.ps1

# Abhängigkeiten installieren
pip install -r requirements.txt
```

### App starten

```bash
python main.py
# oder:
.venv\Scripts\python.exe main.py
```

---

## Architektur

### `auth.py` – Authentifizierungsschicht

Verwaltet Microsoft OAuth2-Anmeldungen via **MSAL** mit persistentem Token-Cache.

**Klasse `MicrosoftAuth`:**

| Methode / Property | Beschreibung |
|---|---|
| `get_accounts()` | Alle gespeicherten Konten aus dem Cache |
| `login_interactive()` | Browser-Anmeldung (`acquire_token_interactive`) |
| `login_silent(account)` | Stille Token-Erneuerung |
| `switch_account(account)` | Wechselt zu einem anderen Konto (setzt `_current_account` explizit) |
| `logout(account)` | Entfernt Konto aus dem lokalen Cache |
| `get_token()` | Gibt gültiges Access-Token zurück (auto-refresh) |
| `get_user_info()` | Ruft Name + E-Mail via `/me` ab |
| `current_user_id_short` | 12-stelliger SHA256-Hash der `home_account_id` → Dateiname-Suffix |

**Token-Cache:** `%APPDATA%\Stundenrechner\auth\token_cache.bin` (serialisiert nach jeder Änderung)

---

### `onedrive.py` – OneDrive-Schicht

Kapselt alle Microsoft Graph API-Aufrufe für OneDrive.

**Klasse `OneDriveClient`:**

| Methode | Beschreibung |
|---|---|
| `list_folder_children(folder_id)` | Listet Unterordner eines Ordners |
| `get_folder_info(folder_id)` | Name und Pfad eines Ordners |
| `get_quota_info()` | Speicherplatz: `{total, used, remaining, state}` |
| `upload_file(local_path, folder_id, filename)` | Datei hochladen (prüft Quota vorher, ersetzt vorhandene Datei) |
| `get_file_web_url()` | Web-URL der zuletzt hochgeladenen Datei |

**Wichtig:**
- Dateinamen werden via `urllib.parse.quote(filename, safe="")` kodiert (Umlaute!)
- `@microsoft.graph.conflictBehavior=replace` steht **literal im URL-String** (nicht als `params`-Dict, da `requests` das `@` sonst zu `%40` enkodiert)
- Upload prüft Quota vorab; bei `state == "exceeded"` wird `RuntimeError` mit Klartextmeldung geworfen

---

### `database.py` – Datenbankschicht

Kapselt alle SQLite-Operationen. Pro Benutzer gibt es eine eigene Datenbankdatei:

```
%APPDATA%\Stundenrechner\stundenrechner_{user_id_short}.db
```

Fällt `user_id_short` weg (Legacy), wird `stundenrechner.db` verwendet.

**Tabellen:**

| Tabelle | Beschreibung |
|---|---|
| `entries` | Stundeneinträge (date, task, hours, customer, commission) |
| `tasks` | Gespeicherte Aufgaben zur Wiederverwendung |
| `settings` | App-Einstellungen (user_name, export_path, export_mode, onedrive_folder_id, onedrive_folder_name) |

**Wichtig:** `_migrate_entries()` sorgt für automatische Schemamigrationen. Neue Spalten dort eintragen.

**Zentrale Methoden:**

```python
Database(user_id_short="abc123")        # Pro-Benutzer-Instanz
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

Basiert auf **ttkbootstrap** (Theme: `cosmo`). Enthält Login-Screen und Haupt-UI.

**Klasse `StundenrechnerApp`:**

| Methode | Beschreibung |
|---|---|
| `_show_login_screen()` | Anmeldebildschirm mit gespeicherten Konten |
| `_build_account_row()` | Zeile pro Konto (Login-Button + 🗑-Button) |
| `_login_existing_account(account)` | Stille Anmeldung, Fallback auf interaktiv |
| `_login_new_account()` | Browser-Anmeldung für neues Konto |
| `_on_login_success()` | Erstellt `Database` + `OneDriveClient`, baut Haupt-UI |
| `_build_main_ui()` | Komplette Haupt-Oberfläche nach Login |
| `_build_header()` | Titelzeile mit Benutzername + „Abmelden"-Button |
| `_logout()` | Beendet Sitzung, stoppt Polling, zeigt Login-Screen |
| `_build_input_section()` | Eingabebereich |
| `_build_entries_section()` | Tagesübersicht-Treeview |
| `_build_monthly_section()` | Monatsübersicht, Export-Modus-Wahl, Export-Button |
| `_on_export_mode_change()` | Wechsel lokal/OneDrive; triggert asynchrone Quota-Prüfung |
| `_check_quota_async()` | Quota-Abfrage im Hintergrund-Thread |
| `_update_quota_label(quota)` | Quota-Status-Label aktualisieren (mit `winfo_exists`-Guard) |
| `_export_to_onedrive()` | Temp-Datei → Excel-Export → Upload → Temp löschen |
| `_export_locally()` | Lokaler Datei-Export |
| `_poll_date()` | Datumsänderungen überwachen (300 ms); stoppt bei `_polling_active = False` |
| `_refresh_all()` | Alle UI-Elemente aktualisieren |

**Klasse `OneDriveFolderDialog`:**  
Ordner-Browser für OneDrive mit Breadcrumb-Navigation (`_nav_stack`), asynchronem Laden via `threading.Thread` und Doppelklick zum Navigieren.

---

### `exporter.py` – Export-Schicht

Erstellt formatierte `.xlsx`-Dateien mit **openpyxl**.

**Spalten:** Datum | Kunde | Komissions-Nr. | Aufgabe | Stunden

**Besonderheiten:**
- Alternierende Zeilenfärbung
- Tagesgesamt-Zeile (grün) nach jedem Tag
- Monatsgesamt-Zeile (dunkelblau) am Ende
- Datum nur in der ersten Zeile eines Tages (inkl. Wochentag)
- Druckbereich automatisch gesetzt

---

## Neues Feld hinzufügen (Anleitung)

1. **`database.py`**: Neue Spalte in `CREATE TABLE` + Migration in `_migrate_entries()` + Parameter in `add_entry()` + Spalte in `get_entries_by_date()` und `get_entries_by_month()`
2. **`app.py`**: Eingabefeld in `_build_input_section()` + Treeview-Spalte in `_build_entries_section()` + Variable in `_add_entry()` + Treeview-Befüllung in `_load_entries()`
3. **`exporter.py`**: `NUM_COLS` erhöhen, Spaltenbreite, Header, Tupel-Entpackung und Zellenformatierung ergänzen

---

## Build (EXE erstellen)

```bat
pip install pyinstaller

pyinstaller --onefile --windowed --name Stundenrechner ^
  --collect-all ttkbootstrap ^
  --collect-all msal ^
  --hidden-import=requests ^
  --hidden-import=msal ^
  --paths src ^
  --clean main.py
```

Die fertige EXE liegt unter `dist\Stundenrechner.exe`.

**Oder:** `build.bat` ausführen – erledigt alle Schritte automatisch.

> **Hinweis:** Immer als normaler Benutzer (nicht als Administrator) builden, da PyInstaller 7.0 Admin-Builds blockiert.

---

## Debugging / Tests

### Module-Check

```powershell
& ".venv\Scripts\python.exe" -c "
from auth import MicrosoftAuth, MS_CLIENT_ID
from database import Database
db = Database(user_id_short='testuser12')
db.close()
print('OK | Client-ID:', MS_CLIENT_ID[:12])
"
```

### Quota prüfen

```powershell
& ".venv\Scripts\python.exe" -c "
from auth import MicrosoftAuth
from onedrive import OneDriveClient
auth = MicrosoftAuth()
accounts = auth.get_accounts()
if accounts:
    auth.switch_account(accounts[0])
    od = OneDriveClient(auth)
    q = od.get_quota_info()
    print(f'Quota: {q[chr(34)+chr(117)+chr(115)+chr(101)+chr(100)+chr(34)]/1e9:.2f} GB | Status: {q[chr(34)+chr(115)+chr(116)+chr(97)+chr(116)+chr(101)+chr(34)]}')
"
```

### Token-Cache zurücksetzen

```powershell
Remove-Item "$env:APPDATA\Stundenrechner\auth\token_cache.bin"
```

### Benutzerdatenbank zurücksetzen

```powershell
Remove-Item "$env:APPDATA\Stundenrechner\stundenrechner_*.db"
```
