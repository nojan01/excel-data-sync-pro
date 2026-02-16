# ToDo Liste

## Offen

- [ ] **Schwebende Bilder** — Unterstützung für klassische Drawing-Bilder (über Zellen schwebend, nicht richData/vm-Zellbilder). openpyxl nutzt `ws.add_image(Image(...), 'A1')` für diese Art von Bildern.

## Erledigt

- [x] Zell-Formatierung beim Kopieren/Einfügen erhalten (cellStyles vollständig übergeben)
- [x] Excel 365 Zellbilder (richData/vm-System) beim Speichern erhalten
- [x] XML-Element-Reihenfolge (drawing/legacyDrawing vor tableParts/extLst)
- [x] Speichern-in-gleicher-Datei: Backup für Restore
- [x] RichText-Export: `name 'max_row' is not defined` behoben
- [x] Merged Cells Export
