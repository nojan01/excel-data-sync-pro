# MVMS-Tool - Verbesserungen

## 🔴 Hohe Priorität (Produktivität)

- [ ] **1. Undo/Redo für Bearbeitungen**
  - Strg+Z / Strg+Y um Änderungen rückgängig zu machen
  - Wichtig für versehentliche Edits im Datenexplorer und Suchergebnissen
  - Undo-Stack mit max. 50 Aktionen

- [ ] **2. Tastenkombinationen (Shortcuts)**
  - Strg+S → Warteschlange speichern/exportieren
  - Strg+F → Fokus auf Suchfeld
  - Strg+Enter → Direkt übertragen
  - F5 → Datei neu laden
  - Escape → Modal schließen

- [ ] **3. Auto-Save der Bearbeitungen**
  - Bearbeitete Zellen periodisch sichern (LocalStorage)
  - Bei Absturz/Neustart wiederherstellbar
  - Hinweis beim Start wenn ungespeicherte Änderungen vorhanden

---

## 🟡 Mittlere Priorität (UX)

- [ ] **4. Such-Historie**
  - Letzte 10 Suchbegriffe merken
  - Dropdown mit Vorschlägen

- [ ] **5. Mehrfach-Suche (AND/OR)**
  - z.B. `Eurofighter AND 2025` oder `A400M OR C-130`
  - Erweiterte Suchsyntax

- [ ] **6. Spalten-Sortierung im Datenexplorer**
  - Klick auf Header → aufsteigend/absteigend sortieren
  - Sehr nützlich bei großen Datensätzen

- [ ] **7. Zeilen-Markierung/Highlighting**
  - Wichtige Zeilen farblich markieren
  - z.B. Rechtsklick → "Als wichtig markieren"

---

## 🟢 Niedrige Priorität (Nice-to-have)

- [ ] **8. Statistiken/Dashboard**
  - Anzahl Zeilen pro Monat
  - Letzte Übertragungen
  - Grafische Auswertung

- [ ] **9. Vorlagen für häufige Transfers**
  - Oft verwendete Spalten-Mappings speichern
  - Schnellauswahl

- [ ] **10. Diff-Ansicht vor Transfer**
  - Zeige was sich ändern wird bevor übertragen wird
  - "Vorschau"-Button

---

## ✅ Erledigt

- [x] Editierbare Zellen im Datenexplorer (2026-01-03)
- [x] Sicherheitsaudit + Fixes (2026-01-03)
- [x] Eurofighter Icon (2026-01-03)
- [x] Version 1.0.5 Release (2026-01-03)
