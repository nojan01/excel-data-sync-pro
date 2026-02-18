# ToDo Liste

## Offen

(keine offenen Punkte)

## Erledigt

- [x] **Schwebende Bilder** — Klassische Drawing-Bilder (twoCellAnchor/oneCellAnchor) werden über `restore_external_links_from_original()` vollständig aus der Originaldatei wiederhergestellt (xl/media/, xl/drawings/, Drawing-Rels, Sheet-Rels, Content_Types). Funktioniert in allen FALL-Pfaden (3a direct-XML, 3b openpyxl, 2 structural, Save-As, Multi-Sheet). JSZip (pendingSheetOperations) und xlsx-populate (Passwortschutz) erhalten Bilder ebenfalls.
- [x] Zell-Formatierung beim Kopieren/Einfügen erhalten (cellStyles vollständig übergeben)
- [x] Excel 365 Zellbilder (richData/vm-System) beim Speichern erhalten
- [x] XML-Element-Reihenfolge (drawing/legacyDrawing vor tableParts/extLst)
- [x] Speichern-in-gleicher-Datei: Backup für Restore
- [x] RichText-Export: `name 'max_row' is not defined` behoben
- [x] Merged Cells Export
