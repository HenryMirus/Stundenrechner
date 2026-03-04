# Stundenrechner

Eine benutzerfreundliche Windows-App zur Erfassung, Verwaltung und dem Export von Arbeitsstunden.

---

## ⬇️ Download

**[➡ Stundenrechner.exe herunterladen (v1.0)](https://github.com/HenryMirus/Stundenrechner/releases/download/v1.0/Stundenrechner.exe)**

> Keine Installation notwendig – einfach herunterladen und starten.

---

## Voraussetzungen

- Windows 10 oder neuer
- Ein Microsoft-Konto (Outlook, Hotmail, Microsoft 365 oder Arbeits-/Schulkonto)

---

## Installation & erster Start

1. Die `Stundenrechner.exe` über den Link oben herunterladen
2. Die Datei an einen beliebigen Ort speichern (z. B. Desktop oder `C:\Programme`)
3. Doppelklick auf `Stundenrechner.exe` – fertig, keine Installation nötig

### ⚠️ Windows SmartScreen-Warnung

Beim ersten Öffnen kann Windows SmartScreen eine Sicherheitswarnung anzeigen, da die Datei nicht signiert ist. So umgehen Sie die Warnung:

1. Klicken Sie auf **„Weitere Informationen"**
2. Klicken Sie auf **„Trotzdem ausführen"**

---

## Erster Start & Anmeldung

Beim ersten Start erscheint ein **Anmeldebildschirm**. Klicken Sie auf **„➕ Neues Konto hinzufügen"** – es öffnet sich ein Browser-Fenster, in dem Sie sich mit Ihrem Microsoft-Konto anmelden.

Nach erfolgreicher Anmeldung wird das Konto gespeichert. Beim nächsten Start genügt ein Klick auf Ihren Namen – eine erneute Browser-Anmeldung ist in der Regel nicht notwendig.

### Mehrere Konten

Mehrere Microsoft-Konten können auf demselben PC verwendet werden. Im Anmeldebildschirm erscheinen alle gespeicherten Konten zur Auswahl. Jedes Konto hat eine **eigene, getrennte Datenbankdatei** – die Stunden eines Benutzers sind für andere nicht sichtbar.

### Abmelden / Konto wechseln

- **Abmelden:** Schaltfläche **„Abmelden"** oben rechts – Sie gelangen zurück zum Anmeldebildschirm. Das Konto bleibt gespeichert.
- **Konto entfernen:** Auf das 🗑-Symbol neben dem Konto klicken – entfernt die gespeicherten Anmeldedaten. Die erfassten Stunden bleiben erhalten.

---

## Bedienung

### Neuen Eintrag hinzufügen

1. **Datum** auswählen (standardmäßig das heutige Datum)
2. **Kunde** eingeben
3. **Komissions-Nr.** eingeben (nur Zahlen)
4. **Aufgabe** auswählen oder neu eingeben  
   → Einmal eingegebene Aufgaben werden automatisch gespeichert und stehen beim nächsten Mal zur Auswahl
5. **Stunden** eingeben (z. B. `1,5` oder `2.25`)
6. Auf **„Eintrag hinzufügen"** klicken oder Enter drücken

### Eintrag löschen

- Eintrag in der Tagesübersicht anklicken → **„Eintrag löschen"** klicken
- Alternativ: Eintrag auswählen und die **Entf-Taste** drücken

### Datum wechseln

Über das Datumsfeld oben können Sie zwischen Tagen wechseln – die Tagesübersicht aktualisiert sich automatisch.

---

## Monatsübersicht

Unterhalb der Tagesübersicht werden die **Gesamtstunden des aktuellen Monats** angezeigt. Diese aktualisieren sich automatisch bei jedem neuen Eintrag.

---

## Excel-Export

Die App unterstützt zwei Export-Modi, die pro Konto gespeichert werden:

### Lokal speichern

1. Exportmodus **„Lokal"** auswählen
2. Über **„Durchsuchen"** einen Zielordner wählen (Standard: Dokumente-Ordner)
3. **Monat** in der Dropdown-Liste auswählen
4. Auf **„Als Excel exportieren"** klicken

### In OneDrive speichern

1. Exportmodus **„OneDrive"** auswählen
2. Über **„📁 Ordner wählen"** den gewünschten OneDrive-Ordner auswählen
3. **Monat** in der Dropdown-Liste auswählen
4. Auf **„Als Excel exportieren"** klicken

Die App zeigt beim Umschalten auf OneDrive automatisch den aktuellen Speicherplatz an (frei / fast voll / voll).

> **Hinweis:** Wenn der OneDrive-Speicher voll ist, schlägt der Upload fehl. In diesem Fall bitte Dateien in OneDrive löschen oder den Speicherplan upgraden.

### Aufbau der exportierten Datei

Die exportierte Datei wird automatisch benannt:  
`Stundenzettel_Monat_Jahr_Vorname_Nachname.xlsx`

| Datum | Kunde | Komissions-Nr. | Aufgabe | Stunden |
|---|---|---|---|---|
| Mo, 04.03.2026 | Musterfirma | 12345 | Projektplanung | 2,50 |
| | Musterfirma | 12345 | Implementierung | 4,00 |
| | | **Tagesgesamt** | | **6,50** |
| ... | | | | |
| | | **MONATSGESAMT** | | **42,00** |

---

## Datenspeicherung

Alle Daten werden lokal auf Ihrem Computer gespeichert. Jedes Konto erhält eine eigene Datenbankdatei:

```
C:\Benutzer\[IhrName]\AppData\Roaming\Stundenrechner\stundenrechner_[KontoID].db
```

Die Anmeldedaten (Token-Cache) liegen unter:

```
C:\Benutzer\[IhrName]\AppData\Roaming\Stundenrechner\auth\token_cache.bin
```

> **Hinweis:** Die Datenbankdateien sind gerätespezifisch. Möchten Sie Daten auf ein anderes Gerät übertragen, kopieren Sie die entsprechende `.db`-Datei an denselben Pfad auf dem Zielgerät.
