/**
 * Excel Live Session Bridge
 * 
 * Kommuniziert mit dem Python-Prozess (excel_live_session.py)
 * der Excel im Hintergrund offen hält.
 * 
 * Jede Operation wird SOFORT in Excel ausgeführt!
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// Globaler Handler für EPIPE-Fehler (verhindert Crash beim Beenden)
process.on('uncaughtException', (err) => {
    if (err.code === 'EPIPE' || err.message?.includes('EPIPE')) {
        // EPIPE während Shutdown ignorieren
        console.error('[LiveSession] EPIPE ignoriert (Shutdown)');
        return;
    }
    // Andere Fehler loggen statt re-throw (verhindert Electron Main Process Crash)
    console.error('[LiveSession] Unbehandelter Fehler:', err.message || err);
});

// Python-Pfad ermitteln (übernommen von python_bridge.js)
function getPythonBasePath() {
    const isPackaged = process.mainModule 
        ? process.mainModule.filename.includes('app.asar')
        : (require.main && require.main.filename.includes('app.asar'));
    
    const hasAsar = process.resourcesPath && fs.existsSync(path.join(process.resourcesPath, 'app.asar'));
    
    if (isPackaged || hasAsar) {
        return path.join(process.resourcesPath, 'app.asar.unpacked', 'python');
    }
    return __dirname;
}

// Globale Variable für PYTHONPATH (wird von getPythonPath gesetzt)
let _pythonEnvPath = null;

function getPythonPath() {
    const basePath = getPythonBasePath();
    const isPackaged = basePath.includes('app.asar.unpacked');
    
    if (isPackaged) {
        const resourcesPath = process.resourcesPath;
        
        if (process.platform === 'darwin') {
            // macOS: System-Python verwenden da venv nur Symlinks enthält
            // Das venv kann auf anderen Macs nicht funktionieren
            // Aber wir brauchen das venv für die Python-Pakete (xlwings, openpyxl)
            const venvPath = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'mac-arm64', 'python-venv');
            const sitePackages = path.join(venvPath, 'lib', 'python3.14', 'site-packages');
            
            // System-Python mit venv site-packages verwenden
            const macPythonPaths = [
                '/opt/homebrew/bin/python3',        // Homebrew Apple Silicon
                '/usr/local/bin/python3',           // Homebrew Intel
                '/usr/bin/python3',                 // System Python
                '/Library/Frameworks/Python.framework/Versions/Current/bin/python3'
            ];
            
            for (const pyPath of macPythonPaths) {
                if (fs.existsSync(pyPath)) {
                    console.log(`[LiveSession] macOS: System-Python gefunden: ${pyPath}`);
                    // Site-packages für späteren Gebrauch speichern
                    if (fs.existsSync(sitePackages)) {
                        _pythonEnvPath = sitePackages;
                        console.log(`[LiveSession] macOS: PYTHONPATH wird gesetzt auf: ${sitePackages}`);
                    }
                    return pyPath;
                }
            }
            
            console.log('[LiveSession] WARNUNG: Kein Python auf macOS gefunden');
        } else if (process.platform === 'win32') {
            const embeddedPython = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'win-x64', 'python.exe');
            if (fs.existsSync(embeddedPython)) return embeddedPython;
        }
    }
    
    // Dev-Modus: venv
    if (!isPackaged) {
        const venvPath = path.join(basePath, '..', '.venv');
        if (fs.existsSync(venvPath)) {
            if (process.platform === 'win32') {
                return path.join(venvPath, 'Scripts', 'python.exe');
            } else {
                return path.join(venvPath, 'bin', 'python3');
            }
        }
    }
    
    // macOS System-Python
    if (process.platform === 'darwin') {
        const macPythonPaths = [
            '/opt/homebrew/bin/python3',
            '/usr/local/bin/python3',
            '/usr/bin/python3'
        ];
        for (const pyPath of macPythonPaths) {
            if (fs.existsSync(pyPath)) return pyPath;
        }
    }
    
    return process.platform === 'win32' ? 'python' : 'python3';
}

/**
 * Gibt die Umgebungsvariablen für Python-Prozesse zurück
 */
function getPythonEnv() {
    const env = { ...process.env };
    if (_pythonEnvPath) {
        // Für macOS: aeosa.pth wird nicht verarbeitet, daher aeosa-Verzeichnis explizit hinzufügen
        // damit 'from appscript import ...' funktioniert
        const aeosaPath = path.join(_pythonEnvPath, 'aeosa');
        let pythonPath = _pythonEnvPath;
        if (fs.existsSync(aeosaPath)) {
            pythonPath = _pythonEnvPath + path.delimiter + aeosaPath;
        }
        env.PYTHONPATH = pythonPath + (env.PYTHONPATH ? path.delimiter + env.PYTHONPATH : '');
    }
    return env;
}

class ExcelLiveSession {
    constructor() {
        this.pythonProcess = null;
        this.commandQueue = [];
        this.currentResolve = null;
        this.currentReject = null;
        this.isReady = false;
        this.responseBuffer = '';
        this.isBusy = false;  // true wenn gerade ein Befehl verarbeitet wird
        this._currentTimeoutId = null;
    }

    /**
     * Startet die Python Live-Session
     */
    async start() {
        if (this.pythonProcess) {
            console.log('[LiveSession] Bereits gestartet');
            return { success: true };
        }

        return new Promise((resolve, reject) => {
            const pythonScript = path.join(getPythonBasePath(), 'excel_live_session.py');
            const pythonPath = getPythonPath();
            
            // cwd muss auf unpacked-Verzeichnis zeigen, nicht auf __dirname (das wäre im asar)
            const cwd = getPythonBasePath();
            
            // Umgebungsvariablen für Python (inkl. PYTHONPATH)
            const env = getPythonEnv();

            console.log('[LiveSession] Starte Python-Prozess:', pythonPath, pythonScript);
            console.log('[LiveSession] CWD:', cwd);
            console.log('[LiveSession] PYTHONPATH:', env.PYTHONPATH || '(nicht gesetzt)');
            
            this.pythonProcess = spawn(pythonPath, [pythonScript], {
                stdio: ['pipe', 'pipe', 'pipe'],
                cwd: cwd,
                env: env
            });

            // Flag um zu verhindern, dass nach close noch geloggt wird
            let processEnded = false;

            // stderr = Log-Output
            this.pythonProcess.stderr.on('data', (data) => {
                if (processEnded) return;
                try {
                    console.log('[Python]', data.toString().trim());
                } catch (e) {
                    // Ignore EPIPE errors during shutdown
                }
            });

            // stdout = JSON-Responses
            this.pythonProcess.stdout.on('data', (data) => {
                if (processEnded) return;
                try {
                    this.responseBuffer += data.toString();
                    
                    // Verarbeite vollständige JSON-Zeilen
                    const lines = this.responseBuffer.split('\n');
                    this.responseBuffer = lines.pop(); // Letzte (möglicherweise unvollständige) Zeile behalten
                    
                    for (const line of lines) {
                        if (line.trim()) {
                            try {
                                const response = JSON.parse(line);
                                if (this.currentResolve) {
                                    this.currentResolve(response);
                                    this.currentResolve = null;
                                    this.currentReject = null;
                                }
                            } catch (e) {
                                console.error('[LiveSession] JSON Parse Error:', e, 'Line:', line);
                            }
                        }
                    }
                } catch (e) {
                    // Ignore EPIPE errors during shutdown
                }
            });

            this.pythonProcess.on('error', (err) => {
                processEnded = true;
                try {
                    console.error('[LiveSession] Prozess-Fehler:', err);
                } catch (e) {
                    // Ignore logging errors
                }
                reject(err);
            });

            this.pythonProcess.on('close', (code) => {
                processEnded = true;
                try {
                    console.log('[LiveSession] Prozess beendet mit Code:', code);
                } catch (e) {
                    // Ignore logging errors during shutdown
                }
                this.pythonProcess = null;
                this.isReady = false;
                if (this.currentReject) {
                    this.currentReject(new Error('Python process closed'));
                }
            });

            // Ping um sicherzugehen dass der Prozess läuft
            setTimeout(async () => {
                try {
                    const result = await this._sendCommand({ action: 'ping' });
                    if (result.success) {
                        this.isReady = true;
                        resolve({ success: true });
                    } else {
                        reject(new Error('Ping failed'));
                    }
                } catch (e) {
                    reject(e);
                }
            }, 500);
        });
    }

    /**
     * Sendet einen Befehl an Python und wartet auf Antwort.
     * Bei parallelen Aufrufen werden Befehle in einer Queue gehalten
     * und sequenziell abgearbeitet.
     * @param {Object} command - Der Befehl
     * @param {number} timeout - Timeout in ms (default: 30000)
     */
    _sendCommand(command, timeout = 30000) {
        return new Promise((resolve, reject) => {
            if (!this.pythonProcess) {
                reject(new Error('Python-Prozess nicht gestartet'));
                return;
            }

            const entry = { command, resolve, reject, timeout };

            if (this.isBusy) {
                // Befehl in Queue einreihen — wird nach aktuellem Befehl ausgeführt
                this.commandQueue.push(entry);
                return;
            }

            this._executeCommand(entry);
        });
    }

    /**
     * Führt einen einzelnen Befehl aus (intern)
     */
    _executeCommand(entry) {
        const { command, resolve, reject, timeout } = entry;

        this.isBusy = true;

        this._currentTimeoutId = setTimeout(() => {
            if (this.isBusy) {
                this.isBusy = false;
                this.currentResolve = null;
                this.currentReject = null;
                this._currentTimeoutId = null;
                reject(new Error('Timeout waiting for response'));
                this._processNextCommand();
            }
        }, timeout);

        this.currentResolve = (result) => {
            if (this._currentTimeoutId) clearTimeout(this._currentTimeoutId);
            this._currentTimeoutId = null;
            this.isBusy = false;
            this.currentResolve = null;
            this.currentReject = null;
            resolve(result);
            this._processNextCommand();
        };
        this.currentReject = (error) => {
            if (this._currentTimeoutId) clearTimeout(this._currentTimeoutId);
            this._currentTimeoutId = null;
            this.isBusy = false;
            this.currentResolve = null;
            this.currentReject = null;
            reject(error);
            this._processNextCommand();
        };

        const cmdJson = JSON.stringify(command) + '\n';
        try {
            this.pythonProcess.stdin.write(cmdJson);
        } catch (e) {
            if (this._currentTimeoutId) clearTimeout(this._currentTimeoutId);
            this._currentTimeoutId = null;
            this.isBusy = false;
            this.currentResolve = null;
            this.currentReject = null;
            reject(e);
            this._processNextCommand();
        }
    }

    /**
     * Verarbeitet den nächsten Befehl aus der Queue
     */
    _processNextCommand() {
        if (this.commandQueue.length > 0 && !this.isBusy && this.pythonProcess) {
            const next = this.commandQueue.shift();
            this._executeCommand(next);
        }
    }

    /**
     * Bricht alle wartenden und laufenden Befehle ab.
     * Wird von close() aufgerufen um sofort beenden zu können.
     */
    _cancelAllPending() {
        // Queue leeren — alle wartenden Befehle ablehnen
        const queue = this.commandQueue;
        this.commandQueue = [];
        for (const entry of queue) {
            try { entry.reject(new Error('Session closing')); } catch(e) {}
        }

        // Laufenden Befehl abbrechen
        if (this._currentTimeoutId) {
            clearTimeout(this._currentTimeoutId);
            this._currentTimeoutId = null;
        }
        if (this.currentReject) {
            try { this.currentReject(new Error('Session closing')); } catch(e) {}
        }
        this.currentResolve = null;
        this.currentReject = null;
        this.isBusy = false;
    }

    /**
     * Öffnet eine Excel-Datei
     * @param {string} filePath - Pfad zur Datei
     * @param {string} sheetName - Name des Sheets
     * @param {string|null} password - Optionales Passwort
     */
    async openFile(filePath, sheetName, password = null) {
        console.log('[LiveSession] Öffne:', filePath, sheetName, password ? '(mit Passwort)' : '');
        this._openedFilePath = filePath;  // Merken für Cleanup bei force-close
        return this._sendCommand({
            action: 'open',
            filePath: filePath,
            sheetName: sheetName,
            password: password
        });
    }

    /**
     * Speichert die Datei
     * @param {string|null} outputPath - Optionaler neuer Pfad
     * @param {string|null} password - Optionales Passwort (null=beibehalten, ''=entfernen, 'xxx'=neu)
     */
    async saveFile(outputPath = null, password = null) {
        return this._sendCommand({
            action: 'save',
            outputPath: outputPath,
            password: password
        });
    }
    
    /**
     * Setzt oder entfernt das Passwort
     * @param {string|null} password - Neues Passwort (null oder '' zum Entfernen)
     */
    async setPassword(password) {
        return this._sendCommand({
            action: 'setPassword',
            password: password
        });
    }
    
    /**
     * Gibt den Passwort-Status zurück
     */
    async getPasswordStatus() {
        return this._sendCommand({
            action: 'getPasswordStatus'
        });
    }

    /**
     * Schließt die Session
     */
    async close() {
        if (!this.pythonProcess) {
            return { success: true };
        }
        
        // Wenn Befehle laufen oder in der Queue sind: alles abbrechen und force-kill
        if (this.isBusy || this.commandQueue.length > 0) {
            console.log('[LiveSession] Close: Offene Commands (' + 
                (this.isBusy ? 'busy' : '') + 
                (this.commandQueue.length > 0 ? ' queue:' + this.commandQueue.length : '') + 
                '), force-close');
            this._cancelAllPending();
            this._forceKillProcess();
            this._forceCloseExcel();
            return { success: true };
        }
        
        try {
            await this._sendCommand({ action: 'close' }, 10000);
        } catch (e) {
            console.error('[LiveSession] Fehler beim Schließen:', e);
            this._forceCloseExcel();
        }
        
        this._forceKillProcess();
        return { success: true };
    }
    
    /**
     * Beendet den Python-Prozess sofort
     */
    _forceKillProcess() {
        if (this.pythonProcess) {
            try {
                this.pythonProcess.kill();
            } catch(e) {}
            this.pythonProcess = null;
        }
        this.isReady = false;
        this.isBusy = false;
        this.commandQueue = [];
        if (this._currentTimeoutId) {
            clearTimeout(this._currentTimeoutId);
            this._currentTimeoutId = null;
        }
    }
    
    /**
     * Beendet Excel separat über OS-Mechanismen
     */
    _forceCloseExcel() {
        if (process.platform === 'darwin') {
            const { exec } = require('child_process');
            const fileName = this._openedFilePath 
                ? path.basename(this._openedFilePath) 
                : null;
            
            if (fileName) {
                const script = `tell application "Microsoft Excel"
                    try
                        close workbook "${fileName}" saving no
                    end try
                    if (count of workbooks) is 0 then quit
                end tell`;
                exec(`osascript -e '${script}'`, { timeout: 5000 }, (err) => {
                    if (err) console.log('[LiveSession] macOS Excel-Cleanup Fehler:', err.message);
                    else console.log('[LiveSession] macOS: Excel Workbook geschlossen');
                });
            } else {
                exec(`osascript -e 'tell application "Microsoft Excel" to quit saving no'`,
                    { timeout: 5000 });
            }
        }
    }

    /**
     * Beendet den Python-Prozess komplett
     */
    async quit() {
        if (!this.pythonProcess) {
            return { success: true };
        }
        
        try {
            await this._sendCommand({ action: 'quit' });
        } catch (e) {
            // Ignorieren, Prozess wird eh beendet
        }
        
        if (this.pythonProcess) {
            this.pythonProcess.kill();
            this.pythonProcess = null;
        }
        
        this.isReady = false;
        return { success: true };
    }

    /**
     * Liest alle Daten aus dem aktuellen Sheet
     */
    async getData() {
        return this._sendCommand({ action: 'getData' });
    }

    // =========================================================================
    // ZEILEN-OPERATIONEN
    // =========================================================================

    /**
     * Löscht eine Zeile
     * @param {number} rowIndex - 0-basierter Index (ohne Header)
     */
    async deleteRow(rowIndex) {
        console.log('[LiveSession] deleteRow:', rowIndex);
        return this._sendCommand({
            action: 'deleteRow',
            rowIndex: rowIndex
        });
    }

    /**
     * Fügt leere Zeilen ein
     * @param {number} rowIndex - Position für die neuen Zeilen
     * @param {number} count - Anzahl der Zeilen
     */
    async insertRow(rowIndex, count = 1) {
        console.log('[LiveSession] insertRow:', rowIndex, 'count:', count);
        return this._sendCommand({
            action: 'insertRow',
            rowIndex: rowIndex,
            count: count
        });
    }

    /**
     * Verschiebt eine Zeile
     * @param {number} fromIndex - Quell-Index
     * @param {number} toIndex - Ziel-Index
     */
    async moveRow(fromIndex, toIndex) {
        console.log('[LiveSession] moveRow:', fromIndex, '->', toIndex);
        return this._sendCommand({
            action: 'moveRow',
            fromIndex: fromIndex,
            toIndex: toIndex
        });
    }

    /**
     * Versteckt oder zeigt eine Zeile
     */
    async hideRow(rowIndex, hidden = true) {
        return this._sendCommand({
            action: 'hideRow',
            rowIndex: rowIndex,
            hidden: hidden
        });
    }

    /**
     * Versteckt oder zeigt mehrere Zeilen auf einmal (Performance-optimiert)
     * @param {number[]} rowIndices - Array von Zeilen-Indizes
     * @param {boolean} hidden - true zum Verstecken, false zum Anzeigen
     */
    async hideRowsBatch(rowIndices, hidden = true) {
        return this._sendCommand({
            action: 'hideRowsBatch',
            rowIndices: rowIndices,
            hidden: hidden
        });
    }

    /**
     * Markiert eine Zeile mit Farbe
     * @param {number} rowIndex
     * @param {string|null} color - 'green', 'yellow', 'red', etc. oder null zum Entfernen
     */
    async highlightRow(rowIndex, color) {
        return this._sendCommand({
            action: 'highlightRow',
            rowIndex: rowIndex,
            color: color
        });
    }

    // =========================================================================
    // SPALTEN-OPERATIONEN
    // =========================================================================

    /**
     * Löscht eine Spalte
     * @param {number} colIndex - 0-basierter Index
     */
    async deleteColumn(colIndex) {
        console.log('[LiveSession] deleteColumn:', colIndex);
        return this._sendCommand({
            action: 'deleteColumn',
            colIndex: colIndex
        });
    }

    /**
     * Fügt Spalten ein
     */
    async insertColumn(colIndex, count = 1, headers = null) {
        console.log('[LiveSession] insertColumn:', colIndex, 'count:', count);
        return this._sendCommand({
            action: 'insertColumn',
            colIndex: colIndex,
            count: count,
            headers: headers
        });
    }

    /**
     * Verschiebt eine Spalte
     */
    async moveColumn(fromIndex, toIndex) {
        console.log('[LiveSession] moveColumn:', fromIndex, '->', toIndex);
        return this._sendCommand({
            action: 'moveColumn',
            fromIndex: fromIndex,
            toIndex: toIndex
        });
    }

    /**
     * Versteckt oder zeigt eine Spalte
     */
    async hideColumn(colIndex, hidden = true) {
        return this._sendCommand({
            action: 'hideColumn',
            colIndex: colIndex,
            hidden: hidden
        });
    }

    // =========================================================================
    // ZELL-OPERATIONEN
    // =========================================================================

    /**
     * Setzt den Wert einer Zelle
     */
    async setCellValue(rowIndex, colIndex, value) {
        return this._sendCommand({
            action: 'setCellValue',
            rowIndex: rowIndex,
            colIndex: colIndex,
            value: value
        });
    }
    
    /**
     * Setzt alle Werte einer Spalte auf einmal (effizienter bei vielen Werten)
     * @param {number} colIndex - 0-basierter Spaltenindex
     * @param {Array} values - Array von Werten für jede Zeile
     * @param {number} startRow - 0-basierter Start-Zeilenindex (Default: 0)
     */
    async setColumnValues(colIndex, values, startRow = 0) {
        return this._sendCommand({
            action: 'setColumnValues',
            colIndex: colIndex,
            values: values,
            startRow: startRow
        });
    }
    
    /**
     * Setzt mehrere Zellen auf einmal (für einzelne Zellen)
     * @param {Array} cells - Array von {row, col, value}
     */
    async setCellsBatch(cells) {
        // Längerer Timeout für Bulk-Operationen (2 Minuten)
        const timeout = Math.max(60000, cells.length * 30); // Min 60s, +30ms pro Zelle
        return this._sendCommand({
            action: 'setCellsBatch',
            cells: cells
        }, timeout);
    }
    
    /**
     * Nutzt Excel's native Suchen & Ersetzen - extrem schnell!
     * @param {string} searchText - Suchtext
     * @param {string} replaceText - Ersetzungstext
     * @param {boolean} matchCase - Groß-/Kleinschreibung
     * @param {boolean} wholeWord - Nur ganze Wörter
     */
    async findReplace(searchText, replaceText, matchCase = false, wholeWord = false) {
        return this._sendCommand({
            action: 'findReplace',
            searchText: searchText,
            replaceText: replaceText,
            matchCase: matchCase,
            wholeWord: wholeWord
        });
    }
    
    /**
     * Setzt AutoFilter in Excel
     * @param {Array} filters - Array von {colIndex, criteria, operator}
     */
    async setAutoFilter(filters) {
        return this._sendCommand({
            action: 'setAutoFilter',
            filters: filters
        });
    }
    
    /**
     * Entfernt alle AutoFilter
     */
    async clearAutoFilter() {
        return this._sendCommand({
            action: 'clearAutoFilter'
        });
    }
    
    /**
     * Zeigt oder versteckt das Excel-Fenster
     */
    async setVisible(visible) {
        return this._sendCommand({
            action: 'setVisible',
            visible: visible
        });
    }
    
    /**
     * Prüft ob Excel und das Workbook noch aktiv sind
     * Wenn gerade ein Befehl läuft, wird angenommen dass alles OK ist
     */
    async checkAlive() {
        // Wenn gerade ein Befehl läuft, ist Excel definitiv noch aktiv
        if (this.isBusy) {
            return { success: true, alive: true, reason: 'busy' };
        }
        
        return this._sendCommand({
            action: 'checkAlive'
        });
    }
    
    /**
     * Gibt verfügbare Recovery-Dateien zurück
     */
    async getRecoveryFiles() {
        return this._sendCommand({
            action: 'getRecoveryFiles'
        });
    }
    
    /**
     * Löscht eine Recovery-Datei
     */
    async deleteRecoveryFile(filePath) {
        return this._sendCommand({
            action: 'deleteRecoveryFile',
            filePath: filePath
        });
    }
}

// Singleton-Instanz
let liveSession = null;

/**
 * Gibt die Live-Session-Instanz zurück (erstellt sie bei Bedarf)
 */
function getLiveSession() {
    if (!liveSession) {
        liveSession = new ExcelLiveSession();
    }
    return liveSession;
}

module.exports = {
    ExcelLiveSession,
    getLiveSession
};
