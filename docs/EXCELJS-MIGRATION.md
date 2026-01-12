# ExcelJS Migration

## Überblick

Dieser Branch (`exceljs-migration`) testet die Migration von **xlsx-populate** zu **exceljs**.

## Warum ExcelJS?

### xlsx-populate (aktuelle Version im master)
- ✅ Erhält Formatierung perfekt
- ❌ Sehr langsam (11.5 Sekunden für 7MB Datei)
- ❌ 500x Memory Bloat (4.2MB → 2.3GB)
- ❌ Seit 6 Jahren nicht mehr gewartet
- ❌ OOM-Crashes bei großen Dateien

### exceljs (diese Migration)
- ✅ Aktiv gewartet (3.3M Downloads/Woche)
- ✅ Schneller beim Parsen
- ✅ Unterstützt Formatierung (Styles, Formeln, RichText)
- ✅ Weniger Memory-Verbrauch
- ⚠️ Zu testen: Formatierungs-Erhaltung bei Row-Moves

## Implementierung

### Dateien
- `exceljs-reader.js` - Neue Read-Funktion mit ExcelJS
- `test-exceljs.js` - Standalone Performance-Test
- `main.js` - IPC-Handler für A/B-Test (`excel:readSheetTest`)

### Performance testen (Kommandozeile)

```bash
# Test mit deiner Excel-Datei
node test-exceljs.js "/pfad/zu/datei.xlsx" "SheetName"

# Beispiel
node test-exceljs.js test.xlsx "DEFENCE&SPACE Aug-2025"
```

Das Skript zeigt:
- ⏱️ Ladezeit xlsx-populate vs ExcelJS
- 📊 Anzahl Zeilen/Spalten/Zellen
- 🚀 Geschwindigkeits-Vergleich in %
- 📋 Qualität: Styles, Formeln, Hyperlinks, RichText

### In der App testen

Die App hat einen Test-Handler `excel:readSheetTest` der beide Methoden vergleicht und die Performance loggt.

## Branches

- **master**: Stabile Version mit xlsx-populate
- **exceljs-migration**: Test-Version mit exceljs

## Branch wechseln

```bash
# Zurück zum master (xlsx-populate)
git checkout master

# Zur exceljs-Version wechseln
git checkout exceljs-migration
```

## Test-Checkliste

- [ ] Performance: ExcelJS schneller als xlsx-populate?
- [ ] Datei öffnen und Sheet laden
- [ ] Einfache Zell-Änderungen
- [ ] Zeilen verschieben (Row-Moves)
- [ ] Formatierung bleibt erhalten (Styles)
- [ ] RichText-Zellen werden korrekt gelesen
- [ ] Formeln werden extrahiert
- [ ] Hyperlinks funktionieren
- [ ] Versteckte Zeilen/Spalten
- [ ] Conditional Formatting (CF)
- [ ] Große Dateien (> 5MB)
- [ ] Memory-Verbrauch akzeptabel

## Nächste Schritte

1. **Performance testen**: `node test-exceljs.js <datei> <sheet>`
2. **Formatierung prüfen**: Styles, RichText, Farben vergleichen
3. **Export implementieren**: ExcelJS-Write-Funktion erstellen
4. **Row-Moves testen**: Formatierung nach Verschieben prüfen
5. **Entscheidung**: Bei Erfolg → merge in master, sonst → xlsx-populate behalten

## Status

🚧 **Phase 1: READ-PERFORMANCE** - ExcelJS Reader implementiert, Performance-Tests möglich

Nächste Phase: Write-Funktion für Export/Save

---

**WICHTIG**: DO NOT MERGE ohne vollständige Tests!
