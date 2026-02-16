# Code-Review: Excel Data Sync Pro
**Datum:** 10. Februar 2026  
**Version:** 1.1.3  
**Geprüfte Dateien:** main.js, preload.js, python_bridge.js, excel_live_bridge.js, exceljs-reader.js, exceljs-writer.js, package.json

---

## 🔴 P0 — KRITISCHE PROBLEME

### 1. Endlose Rekursion in `safeLog`/`safeError`
- **Datei:** `python/python_bridge.js` Zeile 14–31
- **Problem:** Beide Funktionen rufen sich selbst auf statt `console.log`/`console.error`. Jeder Aufruf führt zu Stack Overflow.
- **Fix:** `safeLog(...args)` → `console.log(...args)`, `safeError(...args)` → `console.error(...args)`
- **Status:** ✅ Behoben

### 2. Timeout-Race-Condition in `_sendCommand`
- **Datei:** `python/excel_live_bridge.js` Zeile 260–288
- **Problem:** `setTimeout`-Timer wird nie mit `clearTimeout` bereinigt. Feuert auch nach erfolgreicher Antwort und rejectet ggf. den nächsten Command.
- **Fix:** Timer-ID speichern, in resolve/reject mit `clearTimeout` aufräumen. Command-Queue implementiert.
- **Status:** ✅ Behoben

### 3. Globaler `uncaughtException`-Handler crasht Electron
- **Datei:** `python/excel_live_bridge.js` Zeile 15–22
- **Problem:** Handler fängt nur EPIPE ab, re-throwt alle anderen Exceptions → crasht Electron Main Process.
- **Fix:** `throw err` entfernt, loggt stattdessen den Fehler.
- **Status:** ✅ Behoben

---

## 🟠 P1 — HOHE SICHERHEITSRISIKEN

### 4. `callPython` ohne Pfad-Validierung
- **Datei:** `python/python_bridge.js` Zeile 254
- **Problem:** `scriptName` wird direkt in `path.join(basePath, scriptName)` verwendet. Path-Traversal möglich.
- **Fix:** `path.resolve` + `startsWith`-Guard implementiert.
- **Status:** ✅ Behoben

### 5. Command-Queue fehlt in Live Bridge
- **Datei:** `python/excel_live_bridge.js` Zeile 250–256
- **Problem:** `currentResolve`/`currentReject` wird bei parallelen Commands überschrieben. Erste Antwort geht verloren.
- **Fix:** Vollständige Command-Queue mit `_processNextCommand()` implementiert.
- **Status:** ✅ Behoben

### 6. `deleteRecoveryFile` nimmt beliebigen Pfad
- **Datei:** `main.js` (liveSession:deleteRecoveryFile Handler)
- **Problem:** Keine Prüfung ob der Pfad im Recovery-Verzeichnis liegt. Beliebige Dateien löschbar.
- **Fix:** Plattform-spezifische Recovery-Verzeichnis-Validierung implementiert.
- **Status:** ✅ Behoben

### 7. stderr-Inhalte in Fehlermeldungen
- **Datei:** `python/python_bridge.js` Zeile 276
- **Problem:** Python-Stacktraces mit internen Pfaden könnten ans Frontend gelangen.
- **Fix:** Generische Fehlermeldung an Frontend implementiert, Details nur im Log.
- **Status:** ✅ Behoben

---

## 🟡 P2 — MITTLERE RISIKEN & EFFIZIENZ

### 8. Vorhersagbare Temp-Dateinamen
- **Datei:** `exceljs-reader.js` Zeile 281
- **Problem:** `mvms_decrypt_${Date.now()}` ist erratbar.
- **Fix:** `crypto.randomUUID()` implementiert.
- **Status:** ✅ Behoben

### 9. DevTools-Menü in Production sichtbar
- **Datei:** `main.js` Zeile 1300–1320
- **Problem:** `View → Toggle DevTools`, `Reload`, `Force Reload` sind auch in Production verfügbar.
- **Fix:** Menüeinträge mit `isDevMode`-Flag konditionalisiert.
- **Status:** ✅ Behoben

### 10. Doppelte Funktionsdefinitionen
- **Dateien:** `main.js`, `exceljs-reader.js`, `exceljs-writer.js`
- **Problem:** `numberToColumnLetter()` zweimal in main.js (Zeile 188 + 3878), `colLetterToNumber` in reader + writer dupliziert.
- **Fix:** Duplikate in main.js entfernt (zweite Definition von `numberToColumnLetter`/`columnLetterToNumber`).
- **Status:** ✅ Behoben

### 11. ~800 Zeilen toter/auskommentierter Code
- **Datei:** `main.js` Zeile ~2000–2800
- **Problem:** Alte xlsx-populate-Version ist komplett auskommentiert.
- **Fix:** 652 Zeilen toter Code entfernt.
- **Status:** ✅ Behoben

### 12. Redundante `require`-Aufrufe
- **Datei:** `main.js` Zeile 1618, 1635, 1636
- **Problem:** `fs` und `path` werden innerhalb von Handlern erneut importiert, obwohl sie bereits oben importiert sind.
- **Fix:** Inline-require entfernt (nutzt Top-Level-Imports).
- **Status:** ✅ Behoben

### 13. `getPythonPath()` ohne Cache
- **Dateien:** `python/python_bridge.js` Zeile 39–120, `python/excel_live_bridge.js`
- **Problem:** ~10× `fs.existsSync` bei jedem Aufruf. Wird bei jeder Python-Operation aufgerufen.
- **Fix:** Ergebnis wird nach erstem Aufruf gecacht (in `python/python_env.js`).
- **Status:** ✅ Behoben

### 14. ~100 Zeilen duplizierter Python-Pfad-Code
- **Dateien:** `python/python_bridge.js` + `python/excel_live_bridge.js`
- **Problem:** `getPythonPath`, `getPythonBasePath`, `getPythonEnv` sind fast identisch kopiert.
- **Fix:** Gemeinsames Modul `python/python_env.js` extrahiert. Beide Dateien nutzen es.
- **Status:** ✅ Behoben

### 15. V8 Heap auf 16GB gesetzt
- **Datei:** `main.js` Zeile 9, `package.json` Scripts
- **Problem:** `--max-old-space-size=16384` ist überdimensioniert für die meisten Maschinen.
- **Fix:** Auf 4 GB (`--max-old-space-size=4096`) reduziert.
- **Status:** ✅ Behoben

---

## 🔵 P3 — VERBESSERUNGSVORSCHLÄGE

### 16. App-Cleanup bei Quit
- **Problem:** Live Session wird beim Beenden nicht sauber geschlossen.
- **Fix:** `app.on('before-quit')` Handler mit `session.close()` implementiert.
- **Status:** ✅ Behoben

### 17. main.js aufteilen (4.500 Zeilen)
- **Empfehlung:** In Module aufteilen:
  - `ipc/excel-handlers.js` (Excel-Operationen)
  - `ipc/config-handlers.js` (Konfiguration)
  - `ipc/live-session-handlers.js` (Live Session IPC)
  - `security/logger.js` (Security + Network Logging)

### 18. Python-Timeout fehlt
- **Datei:** `python/python_bridge.js` Zeile 254
- **Problem:** `callPython` hat keinen Timeout. Hängendes Script blockiert die App ewig.
- **Fix:** `setTimeout` + `proc.kill()` mit konfigurierbarem Timeout (Standard: 120s) implementiert.
- **Status:** ✅ Behoben

### 19. Encoding-Probleme
- **Datei:** `main.js` (z.B. Zeile 2272)
- **Problem:** Mehrere Stellen mit `�` statt Umlauten (z.B. `Pr�fen`, `l�schen`).
- **Fix:** 16 kaputte Umlaute (ü, ö, ß) manuell repariert. Null offene übrig.
- **Status:** ✅ Behoben

### 20. xlsx-populate ohne Maintainer
- **Dependency:** `xlsx-populate ^1.21.0`
- **Problem:** Paket hat keinen aktiven Maintainer mehr.
- **Empfehlung:** Langfristig Ablösung durch rein ExcelJS-basierte Lösung prüfen.

---

## ✅ Was gut funktioniert

- **preload.js:** Saubere `contextBridge`-Implementierung mit `contextIsolation: true`
- **Security Logger:** HMAC-signierte Blockchain-artige Log-Kette
- **Network Logging:** DSGVO-konform (nur Hostname), File-Locking für Netzlaufwerke  
- **Path-Validierung:** `isValidFilePath()` mit Null-Byte und Traversal-Checks
- **Config-Schema:** Validierung + Sanitize-Pattern
- **URL-Whitelist:** `shell:openExternal` erlaubt nur `http://`, `https://`, `mailto:`

---

## Fortschritt

| # | Priorität | Thema | Status |
|---|-----------|-------|--------|
| 1 | P0 | safeLog/safeError Rekursion | ✅ |
| 2 | P0 | Timeout-Race _sendCommand | ✅ |
| 3 | P0 | uncaughtException Handler | ✅ |
| 4 | P1 | callPython Pfad-Validierung | ✅ |
| 5 | P1 | Command-Queue Live Bridge | ✅ |
| 6 | P1 | deleteRecoveryFile Pfad-Check | ✅ |
| 7 | P1 | stderr in Fehlermeldungen | ✅ |
| 8 | P2 | Temp-Datei randomUUID | ✅ |
| 9 | P2 | DevTools nur Dev-Modus | ✅ |
| 10 | P2 | Doppelte Funktionen | ✅ |
| 11 | P2 | Toten Code entfernen | ✅ |
| 12 | P2 | Redundante require | ✅ |
| 13 | P2 | getPythonPath Cache | ✅ |
| 14 | P2 | Python-Pfad-Code dupliziert | ✅ |
| 15 | P2 | V8 Heap 16GB → 4GB | ✅ |
| 16 | P3 | App-Cleanup bei Quit | ✅ |
| 17 | P3 | main.js aufteilen | ⬜ (Architektur-Refactoring, separat) |
| 18 | P3 | Python-Timeout | ✅ |
| 19 | P3 | Encoding-Probleme | ✅ |
| 20 | P3 | xlsx-populate Ablösung | ⬜ (langfristig) |

---

## 📋 To-Do: Bild-in-Zelle (richData) Wiederherstellung

### Erledigt ✅

| Commit | Beschreibung |
|--------|-------------|
| `ae40d28` | rId-Mismatch Fix: tableParts/pageSetup/Namespaces vom Original wiederherstellen |
| `d5e4f2a` | Diagnostik-Dump für Image-Debugging hinzugefügt |
| `592f756` | Windows-Pfad-Robustheit: `os.path.relpath` + `os.path.normpath` × 3 |
| `978fba0` | Content_Types/Rels Konsistenz + FALL 1 restore (siehe unten) |

### Fix-Details (978fba0)

1. **Content_Types Konsistenz** — `[Content_Types].xml` wird vom Original kopiert, referenziert aber Dateien die openpyxl nicht erzeugt (z.B. `calcChain.xml`). Excel findet die fehlende Datei → Reparaturmodus → richData/Bilder werden entfernt. **Fix:** Fehlende Dateien aus Original nachkopieren, inexistente Einträge entfernen.

2. **workbook.xml.rels Konsistenz** — Gleiche Problematik für `xl/_rels/workbook.xml.rels`.

3. **FALL 1 (fromFile) restore fehlte** — Wenn ein Sheet als `fromFile: true` verarbeitet wird, fehlten `restore_table_xml_from_original()` und `restore_external_links_from_original()` Aufrufe komplett. Ergänzt.

### Offen ⬜

- [ ] Windows-Test: Bestätigen dass Bild nach Export sichtbar ist (kein Placeholder-Icon)
- [ ] Test mit mehreren Bildern / mehreren Sheets
- [ ] Test: "Speichern" (gleiche Datei) vs "Speichern unter" (neue Datei)
- [ ] Test: Zelle mit Bild editieren → Speichern → Bild erhalten?
- [ ] XlsxPopulate Passwortschutz prüfen (Zeile 997 python_bridge.js) — könnte richData nach restore zerstören
