# Excel Data Sync Pro - Features Roadmap

## Übersicht

**NEU: xlwings-Integration (Januar 2026)**

Die App wurde auf xlwings umgestellt für perfekte Excel-Kompatibilität:
- ✅ Conditional Formatting (CF) wird bei strukturellen Änderungen automatisch angepasst
- ✅ Spalten löschen/einfügen mit CF-Erhalt
- ✅ Zeilen löschen mit CF-Erhalt  
- ✅ Alle Excel-Features werden von Excel selbst verarbeitet

### Architektur
- **xlwings**: Native Excel-Integration (Excel macht alle strukturellen Änderungen)
- **openpyxl**: Fallback für Lesen von Metadaten (merged cells, hidden columns)
- **Plattformen**: macOS (mit Excel) und Windows

---

## ✅ Bereits implementiert
- [x] Lesen/Schreiben von Zellwerten
- [x] Spalten ausblenden (`column.hidden()`)
- [x] Hidden-Status beim Laden/Speichern erhalten
- [x] Zeilen löschen (einzeln und mehrfach)
- [x] Arbeitsblätter lesen und wechseln
- [x] **Conditional Formatting Erhalt bei Spalten löschen** (xlwings)
- [x] **Conditional Formatting Erhalt bei Spalten einfügen** (xlwings)

---

## 🔴 Priorität HOCH

### 1. Suchen & Ersetzen
- **Status:** ✅ Implementiert
- **API:** `sheet.find(pattern, replacement)`, `workbook.find(pattern, replacement)`
- **Nutzen:** Schnelle Massenänderungen im DatenExplorer
- **UI:** Suchfeld + Ersetzen-Feld in Toolbar, Rückgängig-Funktion

### 2. Data Validation (Dropdown-Listen)
- **Status:** ✅ Implementiert
- **API:** `cell.dataValidation()` 
- **Nutzen:** Spalten mit vordefinierten Werten als Dropdown anzeigen
- **UI:** Dropdown in Zellen mit Validierung, unterstützt Listen und Bereichsreferenzen

### 3. Styles lesen/anzeigen
- **Status:** ✅ Implementiert
- **API:** `cell.style("bold")`, `cell.style("fill")`, `cell.style("fontColor")`
- **Nutzen:** Formatierungen visuell darstellen
- **UI:** Zellen entsprechend formatiert anzeigen (Fett, Kursiv, Unterstrichen, Durchgestrichen, Schriftfarbe, Hintergrundfarbe, Schriftgröße, Ausrichtung)

---

## 🟡 Priorität MITTEL

### 4. Zeilen ausblenden
- **Status:** ✅ Implementiert
- **API:** `row.hidden(true/false)`
- **Nutzen:** Analog zu Spalten auch Zeilen temporär ausblenden
- **UI:** Kontextmenü mit "Zeile ausblenden", Indikator-Button zum Einblenden

### 5. Formeln anzeigen
- **Status:** ✅ Implementiert
- **API:** `cell.formula()`
- **Nutzen:** Transparenz - Benutzer sieht ob Zelle Formel oder Wert enthält
- **UI:** Formel-Icon (ƒ) in der Ecke von Formelzellen, Tooltip mit vollständiger Formel

### 6. AutoFilter erhalten
- **Status:** ✅ Implementiert
- **API:** `sheet.autoFilter()`, `range.autoFilter()`
- **Nutzen:** Excel-AutoFilter beim Speichern nicht verlieren
- **UI:** Automatisch beim Speichern erhalten (xlsx-populate erhält AutoFilter im XML)

---

## 🟢 Priorität NIEDRIG

### 7. Passwortschutz
- **Status:** ✅ Implementiert
- **API:** `fromFileAsync(path, { password })`, `toFileAsync(path, { password })`
- **Nutzen:** Passwortgeschützte Dateien öffnen/speichern/exportieren
- **UI:** Passwort-Dialog beim Speichern und Exportieren mit Optionen (kein Schutz / beibehalten / neues Passwort)

### 8. Hyperlinks
- **Status:** ✅ Implementiert
- **API:** `cell.hyperlink()`
- **Nutzen:** Links in Zellen klickbar machen
- **UI:** Klickbare Links im DatenExplorer (Ctrl+Klick oder Doppelklick öffnet den Link)

### 9. Zellen verbinden (Merged Cells)
- **Status:** ✅ Implementiert
- **API:** `range.merged()`
- **Nutzen:** Verbundene Zellen korrekt darstellen
- **UI:** Visuell verbundene Zellen mit ⊞-Icon, colspan für horizontale Merges

### 10. Rich Text
- **Status:** ✅ Implementiert
- **API:** `RichText` Klasse
- **Nutzen:** Gemischte Formatierung in einer Zelle
- **UI:** Formatierter Text mit unterschiedlichen Styles pro Fragment (Fett, Kursiv, Unterstrichen, Farben, Schriftgrößen)

### 11. Freeze Panes
- **Status:** ✅ Verifiziert
- **API:** `sheet.freezePanes(x, y)`
- **Nutzen:** Fixierung erhalten beim Speichern
- **UI:** Automatisch erhalten (xlsx-populate behält sheetViews/pane-Struktur)

### 12. Arbeitsblätter verwalten
- **Status:** ✅ Implementiert
- **API:** `addSheet()`, `deleteSheet()`, `cloneSheet()`, `moveSheet()`, `sheet.name()`
- **Nutzen:** Blätter hinzufügen/löschen/kopieren/umbenennen, Reihenfolge ändern
- **UI:** Sheet-Verwaltung Modal (📋 Button neben Dropdown)

---

## 🔧 Strukturelle Operationen

### Spalten-Operationen (Column Operations)
- **Status:** ✅ Abgeschlossen
- **Funktionen:**
  - [x] Spalte löschen - Mit Style-Shifting, CF-Anpassung, Formel-Referenzen
  - [x] Spalte verschieben (Drag & Drop) - Reihenfolge ändern
  - [x] Spalte ausblenden/einblenden - Hidden-Status persistent
  - [x] Spalte einfügen - Neue leere Spalte an Position
  - [x] AutoFilter-Bereich anpassen bei Strukturänderungen
  - [x] Bedingte Formatierungen (CF) bei Spaltenänderungen anpassen
  - [x] Formeln mit Spaltenreferenzen aktualisieren
  - [x] Export mit fullRewrite bei strukturellen Änderungen

### Zeilen-Operationen (Row Operations)
- **Status:** 🟡 In Arbeit
- **Geplante Funktionen:**
  - [x] Zeile löschen (einzeln und mehrfach)
  - [x] Zeile ausblenden/einblenden
  - [x] Gefilterte Zeilen exportieren
  - [ ] Zeile einfügen - Neue leere Zeile an Position
  - [ ] Zeile duplizieren - Kopie mit Styles und Formeln
  - [ ] Zeilen verschieben (Drag & Drop)
  - [ ] Mehrfachauswahl für Zeilen-Operationen
- **UI:** Kontextmenü für Zeilen-Operationen, Zeilenauswahl mit Shift/Ctrl

---

## Änderungshistorie

| Datum | Version | Änderung |
|-------|---------|----------|
| 2026-01-14 | 1.0.16 | **Spalten-Operationen abgeschlossen**: Löschen, Verschieben, Ausblenden, Einfügen mit vollständiger Style/CF/Formel-Unterstützung |
| 2026-01-14 | 1.0.16 | **Filter-Export**: Nur gefilterte Zeilen exportieren, überschüssige Zeilen löschen, kompakte Filter-UI |
| 2026-01-12 | 1.0.15 | **Performance-Fix**: Speichern/Exportieren großer Dateien optimiert - Buffer+Streaming für Dateien > 10MB, Garbage Collection, 2x schnelleres Schreiben |
| 2026-01-10 | 1.0.13 | Computer-spezifische Konfiguration - config.json mit Abschnitten pro Computer für unterschiedliche Netzwerkpfade |
| 2026-01-10 | 1.0.12 | DatenExplorer: Excel-Spaltenbuchstaben (A, B, C...) als zusätzliche Header-Zeile |
| 2026-01-10 | 1.0.12 | DatenExplorer: Kopieren/Einfügen mit Formatierung (Styles, Formeln, Hyperlinks, Rich Text) |
| 2026-01-08 | 1.0.12 | Passwortschutz implementiert (Prio NIEDRIG #7) - Speichern und Exportieren mit Excel-kompatibler Verschlüsselung |
| 2026-01-08 | 1.0.12 | Datum-Filter für DatenExplorer - Fällig in X Tagen / Überfällig seit X Tagen |
| 2026-01-08 | 1.0.12 | Pivot-Tabellen Warnung implementiert - Erkennung beim Laden, Warnung vor Datenverlust |
| 2026-01-08 | 1.0.12 | DatenExplorer Vollbild-Modus (F11, ⛶ Button) und sichtbarer Resize-Handle |
| 2026-01-08 | 1.0.12 | Version auf 1.0.12 angehoben - alle 12 geplanten Features implementiert |
| 2026-01-08 | 1.0.11 | Arbeitsblätter verwalten implementiert (Prio NIEDRIG #12) - Hinzufügen, Löschen, Umbenennen, Kopieren, Reihenfolge ändern |
| 2026-01-08 | 1.0.11 | Freeze Panes verifiziert (Prio NIEDRIG #11) - xlsx-populate erhält Freeze Panes automatisch |
| 2026-01-08 | 1.0.11 | Rich Text implementiert (Prio NIEDRIG #10) - Gemischte Formatierung in Zellen dargestellt |
| 2026-01-08 | 1.0.11 | Merged Cells implementiert (Prio NIEDRIG #9) - Verbundene Zellen visuell dargestellt |
| 2026-01-08 | 1.0.11 | Hyperlinks implementiert (Prio NIEDRIG #8) - Links in Zellen klickbar (Ctrl+Klick/Doppelklick) |
| 2026-01-08 | 1.0.11 | AutoFilter erhalten verifiziert (Prio MITTEL #6) - xlsx-populate erhält AutoFilter automatisch |
| 2026-01-08 | 1.0.11 | Formeln anzeigen implementiert (Prio MITTEL #5) |
| 2026-01-08 | 1.0.11 | Zeilen ausblenden implementiert (Prio MITTEL #4) |
| 2026-01-08 | 1.0.11 | Dokument erstellt, Prio HOCH begonnen |

