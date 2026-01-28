# Excel Live Session - Architektur-Dokumentation

## Übersicht

Die **Live-Session-Architektur** löst das fundamentale Index-Problem des bisherigen Batch-Modus:

### Das Problem (Batch-Modus)

```
Benutzer löscht Zeile 3
Benutzer verschiebt Zeile 5 nach oben
→ Problem: Nach dem Löschen von Zeile 3 ist "Zeile 5" jetzt eigentlich "Zeile 4"!
→ Der Batch-Modus kennt diese Änderung nicht → Falsche Zeile wird verschoben
```

### Die Lösung (Live-Session)

```
App startet → Excel öffnet Datei im Hintergrund (unsichtbar)
Benutzer löscht Zeile 3 → SOFORT in Excel ausgeführt
Benutzer verschiebt Zeile 5 → SOFORT in Excel ausgeführt (Zeile 5 ist tatsächlich Zeile 5!)
Benutzer klickt "Export" → Excel speichert unter neuem Namen → Fertig
```

## Komponenten

### 1. Python: `excel_live_session.py`

Langlebiger Python-Prozess der Excel im Hintergrund offen hält.

**Kommunikation:** JSON über stdin/stdout

**Unterstützte Befehle:**

| Kategorie | Aktion | Parameter |
|-----------|--------|-----------|
| Session | `open` | `filePath`, `sheetName` |
| Session | `save` | `outputPath` (optional) |
| Session | `close` | - |
| Session | `getData` | - |
| Zeilen | `deleteRow` | `rowIndex` |
| Zeilen | `insertRow` | `rowIndex`, `count` |
| Zeilen | `moveRow` | `fromIndex`, `toIndex` |
| Zeilen | `hideRow` | `rowIndex`, `hidden` |
| Zeilen | `highlightRow` | `rowIndex`, `color` |
| Spalten | `deleteColumn` | `colIndex` |
| Spalten | `insertColumn` | `colIndex`, `count`, `headers` |
| Spalten | `moveColumn` | `fromIndex`, `toIndex` |
| Spalten | `hideColumn` | `colIndex`, `hidden` |
| Zellen | `setCellValue` | `rowIndex`, `colIndex`, `value` |

### 2. JavaScript: `excel_live_bridge.js`

Node.js-Klasse die den Python-Prozess verwaltet.

```javascript
const { getLiveSession } = require('./python/excel_live_bridge');

// Session holen (Singleton)
const session = getLiveSession();

// Starten
await session.start();

// Datei öffnen
await session.openFile('/path/to/file.xlsx', 'Tabelle1');

// Operationen ausführen (SOFORT in Excel!)
await session.deleteRow(2);       // Zeile 3 löschen
await session.moveRow(4, 1);      // Zeile 5 nach Position 2
await session.hideRow(3);         // Zeile 4 verstecken
await session.highlightRow(5, 'green');

// Speichern
await session.saveFile('/path/to/output.xlsx');

// Beenden
await session.close();
```

### 3. Test-Skript: `test-live-session.js`

```bash
node test-live-session.js /path/to/test.xlsx SheetName
```

## Vorteile

| Aspekt | Batch-Modus | Live-Session |
|--------|-------------|--------------|
| Index-Tracking | ❌ Kompliziert, fehleranfällig | ✅ Automatisch korrekt |
| Formatierung | ❌ Kann verloren gehen | ✅ Immer erhalten |
| Feedback | ❌ Erst am Ende sichtbar | ✅ Sofort nach jeder Operation |
| Performance | ❌ Alles am Ende | ✅ Inkrementell |
| Undo | ❌ Nicht möglich | ✅ Excel-Undo verfügbar |

## Integration in Frontend

### Option A: Reaktiv bei jeder Aktion

```javascript
// Wenn Benutzer auf "Zeile löschen" klickt:
async function onDeleteRowClicked(rowIndex) {
    const result = await session.deleteRow(rowIndex);
    if (result.success) {
        // Tabelle im Frontend aktualisieren
        await refreshTableFromExcel();
    }
}
```

### Option B: Hybrid (später implementieren)

- Schnelle Preview-Änderungen im Frontend
- Sync mit Excel im Hintergrund
- Bei Export: Excel-Version ist bereits aktuell

## Nächste Schritte

1. ✅ Python Live-Session implementiert
2. ✅ JavaScript Bridge implementiert
3. ⏳ Integration in `main.js` (IPC-Handlers)
4. ⏳ Integration in Frontend (`preload.js`)
5. ⏳ Test mit echten Daten
6. ⏳ Fehlerbehandlung verbessern (Verbindungsverlust, etc.)

## Fallback

Falls Excel nicht verfügbar ist oder abstürzt:
- Fallback auf bisherigen Batch-Modus mit openpyxl
- Warnung an Benutzer

## Plattform-Unterschiede

| Feature | macOS | Windows |
|---------|-------|---------|
| Excel verstecken | AppleScript | `app.visible = False` |
| Kopieren | `copy_range()` | `Copy(Destination=)` |
| Python | `/usr/bin/python3` oder Homebrew | Embedded oder System |
