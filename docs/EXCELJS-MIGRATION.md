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

## Branches

- **master**: Stabile Version mit xlsx-populate
- **exceljs-migration**: Test-Version mit exceljs

## Vergleich zurückwechseln

```bash
# Zurück zum master (xlsx-populate)
git checkout master

# Zur exceljs-Version wechseln
git checkout exceljs-migration
```

## Test-Checkliste

- [ ] Datei öffnen und Sheet laden
- [ ] Einfache Zell-Änderungen
- [ ] Zeilen verschieben (Row-Moves)
- [ ] Formatierung bleibt erhalten
- [ ] RichText-Zellen
- [ ] Formeln
- [ ] Hyperlinks
- [ ] Conditional Formatting
- [ ] Große Dateien (> 5MB)
- [ ] Performance-Messung

## Status

🚧 In Entwicklung - DO NOT MERGE ohne vollständige Tests!
