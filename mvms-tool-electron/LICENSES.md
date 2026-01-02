# Lizenzübersicht - MVMS-Tool v1.0.4

Dieses Dokument enthält eine Übersicht aller verwendeten Abhängigkeiten und deren Lizenzen.

## Projektlizenz

**MVMS-Tool** ist lizenziert unter der **MIT License**.

---

## Produktionsabhängigkeiten (dependencies)

| Modul | Version | Lizenz | Beschreibung |
|-------|---------|--------|--------------|
| [xlsx-populate](https://www.npmjs.com/package/xlsx-populate) | ^1.21.0 | MIT | Excel-Bibliothek zum Lesen/Schreiben von XLSX-Dateien mit Formatierungserhalt |

---

## Entwicklungsabhängigkeiten (devDependencies)

| Modul | Version | Lizenz | Beschreibung |
|-------|---------|--------|--------------|
| [electron](https://www.electronjs.org/) | ^39.2.7 | MIT | Framework für Desktop-Anwendungen mit Web-Technologien |
| [electron-builder](https://www.electron.build/) | ^26.0.12 | MIT | Build-Tool für Electron-Anwendungen (Installer, Portable) |
| [electron-reload](https://www.npmjs.com/package/electron-reload) | ^2.0.0-alpha.1 | MIT | Hot-Reload für Electron während der Entwicklung |
| [eslint](https://eslint.org/) | ^9.39.2 | MIT | JavaScript/TypeScript Linter |

---

## Transitive Abhängigkeiten

Die oben genannten Module bringen weitere Abhängigkeiten mit. Alle transitiven Abhängigkeiten verwenden ebenfalls Open-Source-Lizenzen (hauptsächlich MIT, ISC, Apache-2.0, BSD).

Eine vollständige Liste aller transitiven Abhängigkeiten kann mit folgendem Befehl generiert werden:
```bash
npm ls --all
```

---

## Lizenztext (MIT License)

```
MIT License

Copyright (c) 2025 Norbert Jander

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Hinweise

- **xlsx-populate**: Diese Bibliothek wird verwendet, um Excel-Dateien (.xlsx) zu lesen und zu schreiben, wobei Formatierungen (Farben, Schriftarten, etc.) erhalten bleiben.
- **Electron**: Ermöglicht die Erstellung einer nativen Desktop-Anwendung mit HTML, CSS und JavaScript.
- Alle verwendeten Bibliotheken sind mit der MIT-Lizenz des Projekts kompatibel.

---

*Erstellt am: Januar 2025*
