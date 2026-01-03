# MVMS-Tool - Verbesserungen

## 🔴 Hohe Priorität (Produktivität)

- [x] **1. Undo/Redo für Bearbeitungen** ✅ 2026-01-03
  - Strg+Z / Strg+Y um Änderungen rückgängig zu machen
  - Wichtig für versehentliche Edits im Datenexplorer und Suchergebnissen
  - Undo-Stack mit max. 50 Aktionen

- [x] **2. Tastenkombinationen (Shortcuts)** ✅ 2026-01-03
  - Strg+S → Warteschlange speichern/exportieren
  - Strg+F → Fokus auf Suchfeld
  - Strg+Enter → Direkt übertragen
  - F5 → Datei neu laden
  - Escape → Modal schließen

- [x] **3. Auto-Save der Bearbeitungen** ✅ 2026-01-03
  - Bearbeitete Zellen periodisch sichern (LocalStorage)
  - Bei Absturz/Neustart wiederherstellbar
  - Hinweis beim Start wenn ungespeicherte Änderungen vorhanden

---

## 🟡 Mittlere Priorität (UX)

- [x] **4. Such-Historie** ✅ 2026-01-03
  - Letzte 15 Suchbegriffe merken
  - Dropdown mit Vorschlägen (Pfeiltasten navigieren)
  - Gespeichert in LocalStorage

- [x] **5. Mehrfach-Suche (AND/OR)** ✅ 2026-01-03
  - z.B. `Eurofighter AND 2025` oder `A400M OR C-130`
  - Erweiterte Suchsyntax mit AND/OR Operatoren
  - Kombinierbar mit Platzhaltern (* ?)

- [x] **6. Spalten-Sortierung im Datenexplorer** ✅ 2026-01-03
  - Klick auf Header → aufsteigend sortieren
  - Zweiter Klick → absteigend sortieren
  - Dritter Klick → Sortierung aufheben

- [x] **7. Zeilen-Markierung/Highlighting** ✅ 2026-01-03
  - Rechtsklick auf Zeile → Kontextmenü mit 6 Farben
  - Grün, Gelb, Orange, Rot, Blau, Lila
  - Markierung entfernen möglich

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
