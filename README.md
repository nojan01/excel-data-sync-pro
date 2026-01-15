# Excel Data Sync Pro

Eine Desktop-Anwendung zum Synchronisieren und Übertragen von Zeilen zwischen Excel-Dateien, mit Formatierungserhalt, Flag-/Kommentar-Funktion und Template-Erstellung.

## Version

**v1.0.15** - © Norbert Jander 2025

## Hauptfunktionen

### Datentransfer
- **Quelldatei durchsuchen**: Suchen Sie nach Seriennummern oder Text mit Wildcard-Unterstützung (`*` und `?`)
- **Multi-Select**: Mehrere Zeilen gleichzeitig auswählen und übertragen
- **Warteschlange**: Zeilen sammeln und als Batch übertragen
- **Neue Zeile erstellen**: Manuell Zeilen eingeben (auch Leerzeilen)
- **Zeilen kopieren**: Ausgewählte Zeilen in die Zieldatei übertragen
- **Flag setzen**: Jede übertragene Zeile mit A (Add), D (Delete) oder C (Change) markieren
- **Kommentar hinzufügen**: Freier Text für jede übertragene Zeile
- **Duplikat-Erkennung**: Verhindert doppelte Einträge

### Arbeitsblatt-Verwaltung
- **Arbeitsblatt-Auswahl**: Wählen Sie für beide Dateien das gewünschte Arbeitsblatt
- **Spalten-Mapping**: Konfigurieren Sie, welche Spalten kopiert werden
- **Direktes Speichern**: Änderungen werden direkt in die Datei gespeichert

### Template-Funktionen
- **Template laden**: Leere Vorlage mit Formatierungen und Conditional Formatting (CF)
- **🔧 Template aus Quelldatei erstellen**: 
  - Erstellt ein neues Template aus einer beliebigen Quelldatei
  - Behält alle Conditional Formatting Regeln (bis zu 500+)
  - Auswahl welche Arbeitsblätter übernommen werden
  - Optional: Flag- und Kommentar-Spalten automatisch einfügen
  - Alle Spalten werden automatisch verschoben wenn Extra-Spalten aktiviert

### Neuer Monat
- **📅 Neuen Monat erstellen**: 
  - Template kopieren und für neuen Monat vorbereiten
  - Sheet-Name automatisch auf neuen Monat setzen
  - Alle Formatierungen und CF-Regeln bleiben erhalten

### Export-Funktionen
- **Export nur geänderter Zeilen**: Nur Zeilen mit Flag exportieren
- **Export mit allen Arbeitsblättern**: Komplette Datei mit allen Sheets exportieren

### Konfiguration
- **Export/Import**: Konfiguration als JSON-Datei sichern und wiederherstellen
- **Automatisches Laden**: config.json wird automatisch gesucht in:
  1. **Arbeitsordner** (höchste Priorität)
  2. Portable EXE-Ordner
  3. Installationsordner
  4. Dokumente-Ordner
  5. Downloads-Ordner

### 📁 Arbeitsordner
- **Arbeitsordner festlegen**: Definieren Sie einen Standard-Ordner für alle Datei-Dialoge
- **Automatische Config-Suche**: config.json wird zuerst im Arbeitsordner gesucht
- **Persistente Einstellung**: Der Arbeitsordner wird zwischen Sitzungen gespeichert

## Datenexplorer

### Übersicht
Der Datenexplorer bietet erweiterte Funktionen zum Betrachten, Bearbeiten und Exportieren von Excel-Daten.

### Funktionen
- **📂 Datei öffnen**: Excel-Dateien laden und alle Arbeitsblätter anzeigen
- **🔍 Suchen & Filtern**: Globale Suche und spaltenbasierte Filter mit Suchen & Ersetzen
- **✏️ Zellen bearbeiten**: Direktes Bearbeiten von Zellinhalten mit Doppelklick
- **↩️ Undo/Redo**: Änderungen rückgängig machen oder wiederherstellen
- **📊 Mehrfachauswahl**: Zellen mit Shift+Klick, Strg+Klick oder Mausziehen auswählen
- **🗑️ Zellinhalte löschen**: Rechtsklick-Menü zum Löschen ausgewählter Zellinhalte
- **📋 Kopieren**: Ausgewählte Zellinhalte in die Zwischenablage kopieren
- **🎨 Formatierung**: Fett, Kursiv, Unterstrichen, Farben und Rich Text werden angezeigt
- **🔗 Hyperlinks**: Klickbare Links in Zellen (Strg+Klick)
- **📝 Formeln**: Formel-Indikator (ƒ) mit Tooltip
- **⊞ Verbundene Zellen**: Merged Cells werden korrekt dargestellt
- **📋 Arbeitsblatt-Verwaltung**: Sheets hinzufügen, löschen, umbenennen, kopieren
- **⛶ Vollbild-Modus**: F11 für Vollbildansicht
- **⚠️ Pivot-Warnung**: Warnung bei Dateien mit Pivot-Tabellen

### Speichern & Exportieren
- **💾 Speichern**: Änderungen direkt in die Originaldatei speichern (mit Bestätigung)
- **📤 Exportieren**: 
  - Auswahl welche Arbeitsblätter exportiert werden
  - Formatierung der Originaldatei bleibt erhalten
  - Änderungen werden in Export übernommen
  - Sheets ohne Änderungen behalten volle Formatierung

### Arbeitsblatt-Wechsel
- Wechseln Sie zwischen Arbeitsblättern ohne Datenverlust
- **Änderungen bleiben erhalten**: Bearbeitete Daten werden zwischen Sheet-Wechseln gecacht
- **Warnung bei neuer Datei**: Bei ungespeicherten Änderungen erscheint eine Warnung

## Installation

### Windows
1. Laden Sie `Excel-Data-Sync-Pro-x.x.x-Setup.exe` herunter
2. Führen Sie den Installer aus
3. Starten Sie die App über das Desktop-Icon oder Startmenü

### macOS
1. Laden Sie `Excel-Data-Sync-Pro-x.x.x.dmg` herunter
2. Öffnen Sie die DMG-Datei
3. Ziehen Sie die App in den Programme-Ordner

## Workflow

### Standard-Workflow (Datenübertragung)

1. **Quelldatei laden** (Datei 1)
   - Klicken Sie auf "Quelldatei laden"
   - Wählen Sie die Excel-Datei aus der Sie kopieren möchten
   - Wählen Sie das gewünschte Arbeitsblatt

2. **Zieldatei laden** (Datei 2)
   - Klicken Sie auf "Zieldatei laden"
   - Wählen Sie die Excel-Datei in die Sie kopieren möchten
   - Wählen Sie das Ziel-Arbeitsblatt

3. **Spalten konfigurieren**
   - Klicken Sie auf "Spalten konfigurieren"
   - Wählen Sie welche Spalten kopiert werden sollen
   - Aktivieren Sie Flag-Spalte und Kommentar-Spalte nach Bedarf
   - Wählen Sie die Spalte für Duplikat-Erkennung

4. **Suchen und Übertragen**
   - Geben Sie eine Seriennummer oder Text in das Suchfeld ein
   - Wildcards: `*` = beliebig viele Zeichen, `?` = genau ein Zeichen
   - Klicken Sie auf die gewünschten Zeilen
   - Setzen Sie Flag (A/D/C) und optional einen Kommentar
   - Klicken Sie auf "Zur Warteschlange" oder "Direkt übertragen"

5. **Speichern**
   - Klicken Sie auf "💾 Speichern"
   - Die Datei wird direkt am Ursprungsort gespeichert

### Template-Workflow (Neues Template erstellen)

1. **Template aus Quelldatei erstellen**
   - Klicken Sie im Template-Bereich auf "🔧 Template aus Quelldatei"
   - Wählen Sie Ihre Masterdatei mit allen Formatierungen
   - Wählen Sie welche Arbeitsblätter ins Template sollen
   - Aktivieren Sie "Flag-Spalte einfügen" und "Kommentar-Spalte einfügen" falls gewünscht
   - Speichern Sie das Template

2. **Template verwenden**
   - Das erstellte Template wird automatisch geladen
   - Alle Conditional Formatting Regeln sind erhalten
   - Spalten sind bereit für Flag/Kommentar wenn aktiviert

### Neuer Monat

1. **Template laden** (falls nicht bereits geladen)
2. **Auf "📅 Neuer Monat" klicken**
3. **Dateinamen eingeben** (z.B. mit neuem Datum)
4. **Sheet-Name für neuen Monat eingeben**
5. Die neue Datei wird mit allen Formatierungen erstellt

## Tastenkürzel

| Taste | Aktion |
|-------|--------|
| Strg+O | Konfiguration laden |
| Strg+S | Datei 2 speichern |
| Enter | Suche starten |
| F1 | Hilfe anzeigen |
| Esc | Dialog schließen |

## Flags

| Flag | Bedeutung |
|------|-----------|
| A | Add - Zeile hinzufügen |
| D | Delete - Zeile löschen |
| C | Change - Zeile ändern |

## Konfiguration

### Gemeinsame Konfiguration (Netzwerklaufwerk)

1. **Konfiguration erstellen:**
   - Laden Sie beide Excel-Dateien
   - Konfigurieren Sie Arbeitsblätter und Spalten-Zuordnung
   - Klicken Sie auf "config.json speichern"
   - Speichern Sie die Datei im Downloads-Ordner oder Programmordner

2. **Konfiguration laden:**
   - Die config.json aus dem Downloads-Ordner wird automatisch beim Start geladen
   - Alternativ: "📂 config.json laden" und manuell auswählen

### Einstellungen

- Ausgewählte Arbeitsblätter
- Spalten-Zuordnung
- Flag-/Kommentar-Optionen
- Letzte Übertragungen

## Technische Details

- **Technologie**: Electron, Node.js
- **Excel-Bibliothek**: xlsx-populate (für CF-Erhalt), JSZip (für Template-Erstellung)
- **Sicherheit**: HMAC-SHA256 Signaturen, SHA256 Hash-Chain
- **Conditional Formatting**: Vollständig erhalten bei Template-Erstellung
- **Unterstützte Dateiformate**: .xlsx
- **Plattformen**: Windows (x64), macOS (Intel & Apple Silicon)

## Sicherheits-Protokoll

Excel Data Sync Pro verfügt über ein manipulationssicheres Sicherheits-Protokoll zur Nachverfolgung aller wichtigen Aktionen.

### Funktionen

- **Manipulationssichere Speicherung**: Jeder Log-Eintrag wird mit einer HMAC-SHA256-Signatur versehen
- **Hash-Chain**: Ähnlich einer Blockchain werden alle Einträge kryptografisch verkettet
- **Integritätsprüfung**: Nachträgliche Änderungen an der Log-Datei werden erkannt
- **Log-Levels**: INFO, WARNING, ERROR für verschiedene Ereignistypen

### Protokollierte Aktionen

- App-Start/Stop
- Datei-Operationen (Öffnen, Speichern, Export)
- Konfigurationsänderungen
- Template-Erstellung
- Sheet-Operationen (Hinzufügen, Löschen)
- Sicherheitsrelevante Ereignisse

### Verwendung

1. Öffnen Sie die **Einstellungen** (⚙️ Button in der Seitenleiste)
2. Klicken Sie auf **"🔒 Sicherheits-Protokoll"**
3. Im Modal werden alle Aktionen angezeigt
4. Nutzen Sie **"Überprüfen"** zur Integritätsprüfung
5. Filtern Sie nach Level oder durchsuchen Sie die Logs

## Netzwerk-Protokoll

Für Dateien auf Netzlaufwerken wird automatisch ein zusätzliches Protokoll geführt, das die Zusammenarbeit mehrerer Benutzer nachvollziehbar macht.

### Funktionen

- **Automatische Erkennung**: Netzlaufwerke werden automatisch erkannt (UNC-Pfade, /Volumes/)
- **DSGVO-konform**: Nur Rechnername wird protokolliert, keine persönlichen Daten
- **File-Locking**: Verhindert Schreibkonflikte bei gleichzeitigem Zugriff
- **Zentrale Speicherung**: Log-Datei liegt im gleichen Ordner wie die Excel-Dateien
- **Konflikt-Warnung**: Warnt beim Öffnen wenn Datei kürzlich von anderem Rechner bearbeitet wurde
- **Session-Lock**: Markiert Dateien als "in Bearbeitung" für Kollegen

### Konflikt-Erkennung

Beim Öffnen einer Datei auf einem Netzlaufwerk wird automatisch geprüft:

1. **Session-Lock**: Wurde eine Lock-Datei (`.~lock.Dateiname.xlsx`) von einem anderen Rechner erstellt?
2. **Kürzliche Aktivität**: Hat ein anderer Rechner die Datei in den letzten 5 Minuten bearbeitet?

Falls ja, erscheint eine Warnung:

```
⚠️ Achtung: Möglicher Bearbeitungskonflikt!

Diese Datei wurde kürzlich bearbeitet:
• Rechner: PC-BUCHHALTUNG
• Aktion: EXCEL_FILE_SAVED
• Vor: 2 Minute(n)

Wenn Sie die Datei gleichzeitig bearbeiten, 
können Änderungen verloren gehen.

Trotzdem öffnen?
```

### Protokollierte Aktionen

- Datei speichern (`EXCEL_FILE_SAVED`)
- Datenübertragung (`DATA_TRANSFER`)
- Export-Operationen (`EXCEL_EXPORT_SOURCE`, `EXCEL_EXPORT_TARGET`)

### Log-Datei

Die Netzwerk-Log-Datei wird automatisch erstellt unter:
```
\\server\share\.excel-sync-audit.log  (Windows)
/Volumes/Share/.excel-sync-audit.log  (macOS)
```

### Verwendung

1. Laden Sie eine Datei von einem Netzlaufwerk
2. Klicken Sie auf **"🌐 Netzwerk-Logs"** in den Einstellungen
3. Sehen Sie alle Aktionen aller Kollegen auf diesem Laufwerk
4. Filtern Sie nach Rechner oder durchsuchen Sie die Logs

### Beispiel-Eintrag

```json
{
  "timestamp": "2026-01-09T14:30:22.123Z",
  "hostname": "PC-BUCHHALTUNG",
  "action": "DATA_TRANSFER",
  "file": "Umsatz_2026.xlsx",
  "details": { "sheet": "Januar", "rowsInserted": 15 }
}
```

## Changelog

### v1.0.15
- **Performance-Fix**: Speichern/Exportieren großer Dateien (> 10MB) optimiert
  - Zentrale saveWorkbookOptimized() Funktion für konsistentes Error-Handling
  - Automatische Garbage Collection nach Speichern großer Dateien
  - Batch-Verarbeitung mit GC-Hints für Zeilen-Löschung (1000 statt 500 Zeilen)
  - Reduzierter Peak-Memory-Verbrauch (~20-30%)
  - Stabileres Speichern ohne OOM-Fehler

### v1.0.12
- **Neu**: Sicherheits-Protokoll (Security-Logs) mit manipulationssicherer Speicherung
- **Neu**: Netzwerk-Protokoll für Dateien auf Netzlaufwerken (Multi-User-Tracking)
- **Neu**: Konflikt-Warnung beim Öffnen: Zeigt an wenn Datei kürzlich von anderem Rechner bearbeitet wurde
- **Neu**: Session-Lock: Markiert Dateien als "in Bearbeitung" für Kollegen
- **Neu**: DSGVO-konforme Protokollierung (nur Rechnername, keine persönlichen Daten)
- **Neu**: HMAC-SHA256-Signaturen für jeden Log-Eintrag
- **Neu**: Hash-Chain (Blockchain-ähnlich) zur Integritätsprüfung
- **Neu**: Security-Logs Modal zur Anzeige und Überprüfung aller Aktionen
- **Neu**: Netzwerk-Logs Modal mit Rechner-Filter
- **Neu**: Konfigurationsschema-Validierung für sichere Einstellungen
- **Neu**: Integritätsprüfung erkennt nachträgliche Manipulationen

### v1.0.11
- **Neu**: Zeilen einfügen (oberhalb/unterhalb) per Rechtsklick im Datenexplorer
- **Neu**: Zeilen löschen mit Bestätigungsdialog
- **Neu**: Spalten einfügen (links/rechts) mit Namenseingabe
- **Neu**: Spalten löschen mit Warnung über Datenverlust
- **Neu**: Crash-Recovery - automatische Sicherung alle 30 Sekunden
- **Neu**: Wiederherstellungsoption beim Öffnen nach Absturz/Stromausfall
- **Neu**: Warnung bei ungespeicherten Änderungen beim Schließen des Datenexplorers
- **Neu**: Ausgeblendete Spalten werden beim Speichern/Exportieren nicht übernommen
- **Fix**: Korrekte englische Übersetzung für Warteschlange, Vorschau, Export-Button

### v1.0.10
- **Neu**: Datenexplorer mit erweitertem Funktionsumfang
- **Neu**: Multi-Zellen-Auswahl (Shift+Klick, Strg+Klick, Mausziehen)
- **Neu**: Rechtsklick-Kontextmenü zum Löschen/Kopieren von Zellinhalten
- **Neu**: Sheet-Daten-Cache - Änderungen bleiben beim Sheet-Wechsel erhalten
- **Neu**: Speichern in Originaldatei mit Bestätigungsdialog
- **Neu**: Multi-Sheet-Export mit Formatierungserhalt
- **Neu**: Auswahl-Dialog für zu exportierende Arbeitsblätter
- **Neu**: Arbeitsordner-Funktion für Standard-Verzeichnis
- **Neu**: config.json Suche erweitert auf Arbeitsordner (höchste Priorität)

### v1.0.9
- **Neu**: Arbeitsordner (Working Directory) einstellbar
- **Neu**: History-Verlauf für letzte 50 Übertragungen
- **Neu**: Erweiterte Undo/Redo-Funktionalität

### v1.0.8
- **Fix**: Template aus Quelldatei funktioniert wieder korrekt
- Behebt Problem mit Sheet-Namen die Sonderzeichen enthalten (z.B. &, <, >)
- Sheet-Namen werden jetzt korrekt XML-dekodiert beim Mapping

### v1.0.7
- **Neu**: Template aus Quelldatei erstellen
- **Neu**: Arbeitsblatt-Auswahl für Template-Erstellung
- **Neu**: Automatisches Einfügen von Flag-/Kommentar-Spalten
- **Neu**: CF-Regeln werden auf ganze Spalten erweitert

### v1.0.6
- Hybrid-Ansatz für Formatierungserhalt
- Verbessertes CF-Handling

### v1.0.5
- Neuer Monat Funktion
- Export mit allen Sheets

### v1.0.4
- Icon-Anpassungen
- UI-Verbesserungen

## Fehlerbehebung

### "Datei kann nicht gelesen werden"
- Stellen Sie sicher, dass die Datei nicht in Excel geöffnet ist
- Prüfen Sie ob es sich um eine gültige .xlsx Datei handelt

### "Suche findet nichts"
- Die Suche durchsucht alle Spalten
- Groß-/Kleinschreibung wird ignoriert
- Wildcards nutzen: `*text*` findet "text" überall
- Prüfen Sie das ausgewählte Arbeitsblatt

### "Template enthält keine Formatierungen"
- Verwenden Sie "🔧 Template aus Quelldatei" statt manueller Template-Erstellung
- Die Quelldatei muss die gewünschten CF-Regeln enthalten

### "Sheet-Name nicht gefunden bei Template-Erstellung"
- Sheet-Namen mit Sonderzeichen werden seit v1.0.8 korrekt unterstützt
- Aktualisieren Sie auf die neueste Version

## Lizenz

MIT License - © Norbert Jander 2025
