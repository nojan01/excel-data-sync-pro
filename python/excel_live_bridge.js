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

function getPythonPath() {
    const basePath = getPythonBasePath();
    const isPackaged = basePath.includes('app.asar.unpacked');
    
    if (isPackaged) {
        const resourcesPath = process.resourcesPath;
        
        if (process.platform === 'darwin') {
            const embeddedPython = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'mac-arm64', 'python-venv', 'bin', 'python3');
            if (fs.existsSync(embeddedPython)) return embeddedPython;
            
            const embeddedPythonX64 = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'mac-x64', 'python-venv', 'bin', 'python3');
            if (fs.existsSync(embeddedPythonX64)) return embeddedPythonX64;
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

class ExcelLiveSession {
    constructor() {
        this.pythonProcess = null;
        this.commandQueue = [];
        this.currentResolve = null;
        this.currentReject = null;
        this.isReady = false;
        this.responseBuffer = '';
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

            console.log('[LiveSession] Starte Python-Prozess:', pythonPath, pythonScript);
            
            this.pythonProcess = spawn(pythonPath, [pythonScript], {
                stdio: ['pipe', 'pipe', 'pipe'],
                cwd: __dirname
            });

            // stderr = Log-Output
            this.pythonProcess.stderr.on('data', (data) => {
                console.log('[Python]', data.toString().trim());
            });

            // stdout = JSON-Responses
            this.pythonProcess.stdout.on('data', (data) => {
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
            });

            this.pythonProcess.on('error', (err) => {
                console.error('[LiveSession] Prozess-Fehler:', err);
                reject(err);
            });

            this.pythonProcess.on('close', (code) => {
                console.log('[LiveSession] Prozess beendet mit Code:', code);
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
     * Sendet einen Befehl an Python und wartet auf Antwort
     */
    _sendCommand(command) {
        return new Promise((resolve, reject) => {
            if (!this.pythonProcess) {
                reject(new Error('Python-Prozess nicht gestartet'));
                return;
            }

            this.currentResolve = resolve;
            this.currentReject = reject;

            const cmdJson = JSON.stringify(command) + '\n';
            this.pythonProcess.stdin.write(cmdJson);

            // Timeout
            setTimeout(() => {
                if (this.currentReject === reject) {
                    reject(new Error('Timeout waiting for response'));
                    this.currentResolve = null;
                    this.currentReject = null;
                }
            }, 30000);
        });
    }

    /**
     * Öffnet eine Excel-Datei
     */
    async openFile(filePath, sheetName) {
        console.log('[LiveSession] Öffne:', filePath, sheetName);
        return this._sendCommand({
            action: 'open',
            filePath: filePath,
            sheetName: sheetName
        });
    }

    /**
     * Speichert die Datei
     */
    async saveFile(outputPath = null) {
        return this._sendCommand({
            action: 'save',
            outputPath: outputPath
        });
    }

    /**
     * Schließt die Session
     */
    async close() {
        if (!this.pythonProcess) {
            return { success: true };
        }
        
        try {
            await this._sendCommand({ action: 'close' });
        } catch (e) {
            console.error('[LiveSession] Fehler beim Schließen:', e);
        }
        
        if (this.pythonProcess) {
            this.pythonProcess.kill();
            this.pythonProcess = null;
        }
        
        this.isReady = false;
        return { success: true };
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
