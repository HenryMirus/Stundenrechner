# Stundenrechner

Eine benutzerfreundliche Windows-App zur Erfassung, Verwaltung und dem Export von Arbeitsstunden.

---

## Voraussetzungen

- Windows 10 oder neuer
- Keine Installation notwendig – einfach die `Stundenrechner.exe` starten

---

## Erster Start

Beim ersten Start wird nach Ihrem **vollständigen Namen** gefragt. Dieser Name wird dauerhaft gespeichert und automatisch an den Dateinamen der exportierten Excel-Dateien angehangen.

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

1. **Export-Pfad** auswählen (Standardmäßig: Dokumente-Ordner)  
   → Über „Durchsuchen" einen anderen Ordner wählen
2. **Monat** in der Dropdown-Liste auswählen
3. Auf **„Als Excel exportieren"** klicken

### Aufbau der exportierten Datei

Die exportierte Datei wird automatisch benannt:  
`Stundenzettel_Monat_Jahr_Vorname_Nachname.xlsx`

Die Datei enthält:

| Datum | Kunde | Komissions-Nr. | Aufgabe | Stunden |
|---|---|---|---|---|
| Mo, 04.03.2026 | Musterfirma | 12345 | Projektplanung | 2,50 |
| | Musterfirma | 12345 | Implementierung | 4,00 |
| | | **Tagesgesamt** | | **6,50** |
| ... | | | | |
| | | **MONATSGESAMT** | | **42,00** |

---

## Datenspeicherung

Alle Daten werden lokal auf Ihrem Computer gespeichert unter:

```
C:\Benutzer\[IhrName]\AppData\Roaming\Stundenrechner\stundenrechner.db
```

> **Hinweis:** Diese Datenbank-Datei ist gerätespezifisch. Möchten Sie die Daten auf ein anderes Gerät übertragen, kopieren Sie diese Datei an denselben Speicherort auf dem Zielgerät.

---

## Hinweis beim ersten Start auf einem neuen PC

Windows SmartScreen kann beim ersten Öffnen der `Stundenrechner.exe` eine Warnung anzeigen, da die Datei nicht signiert ist. So umgehen Sie die Warnung:

1. Klicken Sie auf **„Weitere Informationen"**
2. Klicken Sie auf **„Trotzdem ausführen"**
