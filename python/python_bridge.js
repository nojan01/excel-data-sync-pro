/**
 * Python Bridge für Excel Data Sync Pro
 * Ermöglicht die Kommunikation zwischen Node.js und Python
 * 
 * HINWEIS: Diese Datei verwendet jetzt primär xlwings für perfekte CF-Erhaltung.
 * Fallback auf openpyxl wenn xlwings nicht verfügbar.
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// Sichere Log-Funktion (verhindert EIO-Fehler wenn keine Konsole vorhanden)
function safeLog(...args) {
    try {
        if (process.stdout && process.stdout.writable) {
            safeLog(...args);
        }
    } catch (e) {
        // Ignoriere Konsolenfehler
    }
}

function safeError(...args) {
    try {
        if (process.stderr && process.stderr.writable) {
            safeError(...args);
        }
    } catch (e) {
        // Ignoriere Konsolenfehler
    }
}

// Python-Pfad ermitteln (embedded oder system)
function getPythonPath() {
    const basePath = getPythonBasePath();
    const isPackaged = basePath.includes('app.asar.unpacked');
    
    // Im gepackten Modus: Eingebettetes Python verwenden
    if (isPackaged) {
        const resourcesPath = process.resourcesPath;
        
        if (process.platform === 'darwin') {
            // macOS: Eingebettetes Python in python-embed/mac-arm64/python-venv
            const embeddedPython = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'mac-arm64', 'python-venv', 'bin', 'python3');
            if (fs.existsSync(embeddedPython)) {
                safeLog(`[Python] Eingebettetes Python gefunden: ${embeddedPython}`);
                return embeddedPython;
            }
            // Fallback Intel Mac
            const embeddedPythonX64 = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'mac-x64', 'python-venv', 'bin', 'python3');
            if (fs.existsSync(embeddedPythonX64)) {
                safeLog(`[Python] Eingebettetes Python (x64) gefunden: ${embeddedPythonX64}`);
                return embeddedPythonX64;
            }
        } else if (process.platform === 'win32') {
            // Windows: Eingebettetes Python in python-embed/win-x64
            const embeddedPython = path.join(resourcesPath, 'app.asar.unpacked', 'python-embed', 'win-x64', 'python.exe');
            if (fs.existsSync(embeddedPython)) {
                safeLog(`[Python] Eingebettetes Python gefunden: ${embeddedPython}`);
                return embeddedPython;
            }
        }
        
        safeLog('[Python] WARNUNG: Eingebettetes Python nicht gefunden, versuche System-Python');
    }
    
    // Dev-Modus: Prüfe ob venv existiert
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
    
    // Fallback: System-Python suchen
    if (process.platform === 'darwin') {
        const macPythonPaths = [
            '/opt/homebrew/bin/python3',        // Homebrew Apple Silicon
            '/usr/local/bin/python3',           // Homebrew Intel
            '/usr/bin/python3',                 // System Python
            '/Library/Frameworks/Python.framework/Versions/Current/bin/python3'
        ];
        
        for (const pyPath of macPythonPaths) {
            if (fs.existsSync(pyPath)) {
                safeLog(`[Python] System-Python gefunden: ${pyPath}`);
                return pyPath;
            }
        }
    }
    
    if (process.platform === 'win32') {
        const winPythonPaths = [
            'C:\\Python312\\python.exe',
            'C:\\Python311\\python.exe',
            'C:\\Python310\\python.exe',
            'C:\\Python39\\python.exe',
            (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python312\\python.exe',
            (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python311\\python.exe',
            (process.env.LOCALAPPDATA || '') + '\\Programs\\Python\\Python310\\python.exe'
        ];
        
        for (const pyPath of winPythonPaths) {
            if (fs.existsSync(pyPath)) {
                safeLog(`[Python] System-Python gefunden: ${pyPath}`);
                return pyPath;
            }
        }
    }
    
    // Letzter Fallback
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    safeLog(`[Python] Fallback auf PATH-Python: ${pythonCmd}`);
    return pythonCmd;
}

// Liefert den Basis-Pfad für Python-Skripte (entpackt im Produktionsmodus)
function getPythonBasePath() {
    // Prüfe ob wir in einer gepackten App laufen
    // process.mainModule ist deprecated, nutze require.main oder app.isPackaged
    const isPackaged = process.mainModule 
        ? process.mainModule.filename.includes('app.asar')
        : (require.main && require.main.filename.includes('app.asar'));
    
    // Alternative: Prüfe ob resourcesPath auf app.asar zeigt
    const hasAsar = process.resourcesPath && fs.existsSync(path.join(process.resourcesPath, 'app.asar'));
    
    if (isPackaged || hasAsar) {
        // In Electron production build: Python-Skripte liegen in app.asar.unpacked/python
        const unpackedPath = path.join(process.resourcesPath, 'app.asar.unpacked', 'python');
        safeLog(`[Python] Gepackter Modus - Python-Pfad: ${unpackedPath}`);
        safeLog(`[Python] Existiert: ${fs.existsSync(unpackedPath)}`);
        return unpackedPath;
    }
    // Im Dev-Modus: wie gehabt
    safeLog(`[Python] Dev-Modus - Python-Pfad: ${__dirname}`);
    return __dirname;
}

// Cache für Excel-Verfügbarkeit
let _excelAvailableCache = null;
let _excelCheckPromise = null;

// Konfigurierte Engine ('auto', 'xlwings', 'openpyxl')
let _configuredEngine = 'auto';

/**
 * Setzt die zu verwendende Excel-Engine
 * @param {string} engine - 'auto', 'xlwings' oder 'openpyxl'
 */
function setExcelEngine(engine) {
    const validEngines = ['auto', 'xlwings', 'openpyxl'];
    if (validEngines.includes(engine)) {
        _configuredEngine = engine;
        safeLog(`[Python] Excel-Engine gesetzt auf: ${engine}`);
        // Cache zurücksetzen wenn Engine geändert wird
        resetExcelCache();
    } else {
        console.warn(`[Python] Ungültige Engine '${engine}', verwende 'auto'`);
        _configuredEngine = 'auto';
    }
}

/**
 * Gibt die aktuell konfigurierte Engine zurück
 * @returns {string} 'auto', 'xlwings' oder 'openpyxl'
 */
function getExcelEngine() {
    return _configuredEngine;
}

// Prüfe ob xlwings-Scripts existieren
function hasXlwingsSupport() {
    const basePath = getPythonBasePath();
    const readerPath = path.join(basePath, 'excel_reader_xlwings.py');
    const writerPath = path.join(basePath, 'excel_writer_xlwings.py');
    const utilsPath = path.join(basePath, 'excel_utils.py');
    
    const readerExists = fs.existsSync(readerPath);
    const writerExists = fs.existsSync(writerPath);
    const utilsExists = fs.existsSync(utilsPath);
    
    safeLog(`[Python] hasXlwingsSupport Check:`);
    safeLog(`[Python]   basePath: ${basePath}`);
    safeLog(`[Python]   excel_reader_xlwings.py: ${readerExists}`);
    safeLog(`[Python]   excel_writer_xlwings.py: ${writerExists}`);
    safeLog(`[Python]   excel_utils.py: ${utilsExists}`);
    
    return readerExists && writerExists;
}

/**
 * Prüft asynchron ob Microsoft Excel installiert und verfügbar ist.
 * Berücksichtigt die konfigurierte Engine.
 * Das Ergebnis wird gecached für schnelle wiederholte Abfragen.
 * @returns {Promise<boolean>} true wenn Excel/xlwings verwendet werden soll
 */
async function isExcelAvailable() {
    // Wenn Engine auf 'openpyxl' gesetzt, immer false zurückgeben
    if (_configuredEngine === 'openpyxl') {
        safeLog('[Python] Engine auf openpyxl gesetzt - xlwings deaktiviert');
        return false;
    }
    
    // Wenn Engine auf 'xlwings' gesetzt, prüfen ob verfügbar
    if (_configuredEngine === 'xlwings') {
        // Cache zurückgeben wenn vorhanden
        if (_excelAvailableCache !== null) {
            return _excelAvailableCache;
        }
        
        if (!hasXlwingsSupport()) {
            console.warn('[Python] xlwings erzwungen aber Scripts nicht gefunden!');
            _excelAvailableCache = false;
            return false;
        }
        
        try {
            const result = await callPython('excel_utils.py', ['check_excel']);
            _excelAvailableCache = result.available === true;
            if (!_excelAvailableCache) {
                console.warn('[Python] xlwings erzwungen aber Excel nicht verfügbar!');
            }
            return _excelAvailableCache;
        } catch (error) {
            console.warn('[Python] xlwings erzwungen aber Check fehlgeschlagen:', error.message);
            _excelAvailableCache = false;
            return false;
        }
    }
    
    // Auto-Modus: Cache zurückgeben wenn vorhanden
    if (_excelAvailableCache !== null) {
        return _excelAvailableCache;
    }
    
    // Wenn bereits ein Check läuft, darauf warten
    if (_excelCheckPromise) {
        return _excelCheckPromise;
    }
    
    // Neuen Check starten
    _excelCheckPromise = (async () => {
        // Ohne xlwings-Scripts kein Excel-Support möglich
        if (!hasXlwingsSupport()) {
            _excelAvailableCache = false;
            return false;
        }
        
        try {
            const result = await callPython('excel_utils.py', ['check_excel']);
            _excelAvailableCache = result.available === true;
            safeLog(`[Python] Excel-Verfügbarkeit: ${_excelAvailableCache ? 'JA' : 'NEIN'}`);
            return _excelAvailableCache;
        } catch (error) {
            safeLog('[Python] Excel-Check fehlgeschlagen:', error.message);
            _excelAvailableCache = false;
            return false;
        }
    })();
    
    const result = await _excelCheckPromise;
    _excelCheckPromise = null;
    return result;
}

/**
 * Setzt den Excel-Cache zurück (für Tests oder nach Neuinstallation)
 */
function resetExcelCache() {
    _excelAvailableCache = null;
    _excelCheckPromise = null;
}

/**
 * Führt ein Python-Script aus und gibt das JSON-Ergebnis zurück
 */
async function callPython(scriptName, args = []) {
    const pythonPath = getPythonPath();
    const basePath = getPythonBasePath();
    const scriptPath = path.join(basePath, scriptName);
    
    safeLog(`[Python] callPython: ${scriptName} ${args.join(' ')}`);
    safeLog(`[Python]   pythonPath: ${pythonPath}`);
    safeLog(`[Python]   scriptPath: ${scriptPath}`);
    safeLog(`[Python]   scriptExists: ${fs.existsSync(scriptPath)}`);
    
    return new Promise((resolve, reject) => {
        const startTime = Date.now();
        const proc = spawn(pythonPath, [scriptPath, ...args]);
        
        let stdout = '';
        let stderr = '';
        
        proc.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        
        proc.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        proc.on('close', (code) => {
            const duration = Date.now() - startTime;
            safeLog(`[Python] Script beendet in ${duration}ms, code=${code}`);
            
            if (stderr) {
                safeLog(`[Python] stderr: ${stderr.substring(0, 500)}`);
            }
            
            if (code !== 0) {
                safeError(`[Python] Error:`, stderr);
                reject(new Error(stderr || `Python script exited with code ${code}`));
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                resolve(result);
            } catch (parseError) {
                safeError(`[Python] JSON parse error:`, parseError.message);
                safeError(`[Python] stdout:`, stdout.substring(0, 500));
                reject(new Error(`Failed to parse Python output: ${parseError.message}`));
            }
        });
        
        proc.on('error', (error) => {
            safeError(`[Python] Spawn error:`, error.message);
            reject(error);
        });
    });
}

/**
 * Liste alle Sheets in einer Excel-Datei
 * Verwendet openpyxl (schneller zum Lesen der Metadaten)
 */
async function listSheets(filePath) {
    return await callPython('excel_reader.py', ['list_sheets', filePath]);
}

/**
 * Liest ein Sheet mit allen Styles
 * Verwendet primär xlwings wenn Excel verfügbar, sonst openpyxl als Fallback
 * 
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des Sheets
 * @returns {Promise<Object>} Sheet-Daten im Format für die GUI
 */
async function readSheet(filePath, sheetName) {
    let result;
    let method = 'openpyxl';
    
    // Prüfe ob Excel verfügbar ist
    const excelAvailable = await isExcelAvailable();
    
    if (excelAvailable) {
        // Primär: xlwings verwenden (native Excel-Integration)
        try {
            result = await callPython('excel_reader_xlwings.py', ['read_sheet', filePath, sheetName]);
            method = 'xlwings';
        } catch (xlwingsError) {
            safeLog(`[Python] xlwings-Lesen fehlgeschlagen, Fallback auf openpyxl: ${xlwingsError.message}`);
            // Fallback auf openpyxl
            result = await callPython('excel_reader.py', ['read_sheet', filePath, sheetName]);
            method = 'openpyxl (fallback)';
        }
    } else {
        // Kein Excel: openpyxl verwenden
        result = await callPython('excel_reader.py', ['read_sheet', filePath, sheetName]);
    }
    
    if (!result.success) {
        return result;
    }
    
    result.method = method;
    
    // Konvertiere zum Frontend-Format (0-basierte Indizes, kompatibel mit ExcelJS Format)
    return {
        success: true,
        headers: result.headers || [],
        data: result.data || [],
        sheetName: result.sheetName,
        rowCount: result.rowCount,
        columnCount: result.columnCount,
        
        // Style-Daten
        cellStyles: result.cellStyles || {},
        cellFonts: result.cellFonts || {},
        defaultFont: result.defaultFont || { name: 'Calibri', size: 11 },
        
        // Struktur-Daten
        mergedCells: result.mergedCells || [],
        autoFilterRange: result.autoFilterRange || null,
        hiddenColumns: result.hiddenColumns || [],
        hiddenRows: result.hiddenRows || [],
        columnWidths: result.columnWidths || {},
        
        // Formeln
        cellFormulas: result.cellFormulas || {},
        
        // Rich Text (falls vorhanden)
        richTextCells: result.richTextCells || {},
        
        // Hyperlinks (falls vorhanden)
        cellHyperlinks: result.cellHyperlinks || {},
        
        // Methode die verwendet wurde
        method: result.method || 'openpyxl'
    };
}

/**
 * Schreibt Daten in eine Excel-Datei mit vollständiger Style-Erhaltung
 * Verwendet primär xlwings für perfekte CF-Erhaltung, Fallback auf openpyxl
 */
async function writeExcel(config) {
    const pythonPath = getPythonPath();
    
    // Prüfe ob Excel verfügbar ist
    const excelAvailable = await isExcelAvailable();
    
    let scriptPath;
    let useXlwings = false;
    
    if (excelAvailable) {
        scriptPath = path.join(getPythonBasePath(), 'excel_writer_xlwings.py');
        useXlwings = true;
        safeLog(`[Python] Verwende xlwings für Schreiboperation`);
        safeLog(`[Python] Script: ${scriptPath}`);
        safeLog(`[Python] Python: ${pythonPath}`);
    } else {
        scriptPath = path.join(getPythonBasePath(), 'excel_writer.py');
        safeLog('[Python] Verwende openpyxl für Schreiboperation (kein Excel verfügbar)');
    }
    
    // Prüfe ob Script existiert
    if (!fs.existsSync(scriptPath)) {
        safeError(`[Python] Script nicht gefunden: ${scriptPath}`);
        return { success: false, error: `Script nicht gefunden: ${scriptPath}`, method: 'error' };
    }
    
    return new Promise((resolve, reject) => {
        const startTime = Date.now();
        safeLog(`[Python] Starte: ${pythonPath} ${scriptPath} write_sheet`);
        const pythonProcess = spawn(pythonPath, [scriptPath, 'write_sheet']);
        
        let stdout = '';
        let stderr = '';
        
        pythonProcess.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        
        pythonProcess.stderr.on('data', (data) => {
            const chunk = data.toString();
            stderr += chunk;
            // LIVE output für Debugging
            process.stdout.write(chunk);
        });
        
        pythonProcess.on('close', async (code) => {
            const duration = Date.now() - startTime;
            
            if (code !== 0) {
                safeError(`[Python] Write Error (code ${code}):`, stderr);
                
                // Bei xlwings-Fehlern automatisch auf openpyxl wechseln
                if (useXlwings) {
                    const isMacPermission = stderr.includes('OSERROR: -1743');
                    const isEpipe = stderr.includes('EPIPE') || stderr.includes('Broken pipe');
                    const isWin32Error = stderr.includes('win32com') || stderr.includes('pywintypes') || stderr.includes('pythoncom');
                    
                    if (isMacPermission || isEpipe || isWin32Error || code !== 0) {
                        const reason = isMacPermission ? 'Berechtigung' : 
                                      isEpipe ? 'EPIPE' : 
                                      isWin32Error ? 'win32com' : 'Unbekannt';
                        safeLog(`[Python] xlwings fehlgeschlagen (${reason}) - wechsle zu openpyxl...`);
                        try {
                            const openpyxlResult = await writeExcelOpenpyxl(config);
                            openpyxlResult.method = 'openpyxl (fallback)';
                            openpyxlResult.warning = `xlwings nicht verfügbar (${reason}). openpyxl verwendet.`;
                            resolve(openpyxlResult);
                            return;
                        } catch (fallbackError) {
                            reject(new Error(`xlwings fehlgeschlagen (${reason}), openpyxl auch: ${fallbackError.message}`));
                            return;
                        }
                    }
                }
                
                reject(new Error(stderr || `Python writer exited with code ${code}`));
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                result.method = useXlwings ? 'xlwings' : 'openpyxl';
                resolve(result);
            } catch (parseError) {
                safeError(`[Python] JSON parse error:`, parseError.message);
                safeError(`[Python] stdout:`, stdout.substring(0, 500));
                reject(new Error(`Failed to parse Python output: ${parseError.message}`));
            }
        });

        pythonProcess.on('error', (error) => {
            safeError(`[Python] Spawn error:`, error.message);
            reject(error);
        });

        // Fehlerhandler für stdin (verhindert EPIPE crashes)
        pythonProcess.stdin.on('error', (error) => {
            safeError(`[Python] stdin error:`, error.message);
            // Nicht reject - warte auf close event für vollständige Fehlermeldung
        });

        // Sende Daten über stdin (für große Datenmengen)
        const jsonData = JSON.stringify(config);
        pythonProcess.stdin.write(jsonData);
        pythonProcess.stdin.end();
    });
}

/**
 * Fallback: Schreibt mit openpyxl (falls xlwings nicht verfügbar)
 */
async function writeExcelOpenpyxl(config) {
    const pythonPath = getPythonPath();
    const scriptPath = path.join(getPythonBasePath(), 'excel_writer.py');
    
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn(pythonPath, [scriptPath, 'write_sheet']);
        
        let stdout = '';
        let stderr = '';
        
        pythonProcess.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        
        pythonProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        pythonProcess.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(stderr || `Python writer exited with code ${code}`));
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                result.method = 'openpyxl';
                resolve(result);
            } catch (parseError) {
                reject(new Error(`Failed to parse Python output: ${parseError.message}`));
            }
        });
        
        pythonProcess.on('error', reject);
        
        pythonProcess.stdin.write(JSON.stringify(config));
        pythonProcess.stdin.end();
    });
}

/**
 * Exportiert mehrere Sheets mit xlwings/openpyxl
 * Öffnet Original-Datei, modifiziert Sheets und speichert unter neuem Pfad
 */
async function exportMultipleSheets(sourcePath, targetPath, sheets, options = {}) {
    const results = [];
    let hasError = false;
    let errorMessage = '';
    let actualMethod = null; // Track the ACTUAL method used, not what was planned
    
    // Original-Datei für Style-Wiederherstellung (falls Markierungen entfernt werden)
    const originalSourcePath = options.originalSourcePath || sourcePath;
    
    // Prüfe ob Quelldatei existiert
    if (!fs.existsSync(sourcePath)) {
        safeError(`[Python] Quelldatei nicht gefunden:`, sourcePath);
        return { 
            success: false, 
            error: `Quelldatei nicht gefunden: "${sourcePath}"\n\nDie Datei wurde möglicherweise verschoben, umbenannt oder gelöscht. Bitte öffnen Sie die Datei erneut.` 
        };
    }
    
    // Zuerst: Kopiere die Original-Datei zum Ziel (falls unterschiedlich)
    // So bleiben alle Sheets, Formatierungen, etc. erhalten
    if (sourcePath !== targetPath) {
        try {
            fs.copyFileSync(sourcePath, targetPath);
        } catch (copyError) {
            safeError(`[Python] Fehler beim Kopieren:`, copyError.message);
            return { success: false, error: `Fehler beim Kopieren: ${copyError.message}` };
        }
    }
    
    // Jetzt: Nur Sheets mit echten Änderungen modifizieren
    for (const sheet of sheets) {
        // Überspringe Sheets ohne Änderungen (fromFile: true und keine editedCells/data)
        if (sheet.fromFile && !sheet.changedCells && !sheet.data?.length && !sheet.fullRewrite) {
            results.push(sheet.sheetName);
            continue;
        }
        
        try {
            // Prüfe ob kombinierte Operationen (Zeilen UND Spalten)
            const hasRowOps = (sheet.rowOperationsQueue && sheet.rowOperationsQueue.length > 0) ||
                              (sheet.deletedRowIndices && sheet.deletedRowIndices.length > 0) ||
                              sheet.insertedRowInfo || sheet.rowOrder;
            const hasColOps = (sheet.columnOperationsQueue && sheet.columnOperationsQueue.length > 0) ||
                              (sheet.deletedColumnIndices && sheet.deletedColumnIndices.length > 0) ||
                              sheet.insertedColumnInfo || sheet.columnOrder;
            
            if (hasRowOps && hasColOps) {
                // KOMBINIERTE OPERATIONEN: Erst Zeilen, dann Spalten (zwei separate Aufrufe)
                safeLog(`[Python] Kombinierte Ops: Erst Zeilen, dann Spalten für "${sheet.sheetName}"`);
                
                // SCHRITT 1: Zeilen-Operationen (OHNE Spalten-Ops, OHNE fullRewrite)
                const rowConfig = {
                    filePath: targetPath,
                    outputPath: targetPath,
                    originalPath: originalSourcePath,
                    sheetName: sheet.sheetName,
                    changes: {
                        headers: sheet.headers || [],
                        data: sheet.data || [],
                        editedCells: {},
                        cellStyles: {},
                        rowHighlights: {},
                        deletedColumns: [],  // Keine Spalten-Ops im ersten Durchlauf
                        insertedColumns: null,
                        deletedRowIndices: sheet.deletedRowIndices || [],
                        insertedRowInfo: sheet.insertedRowInfo || null,
                        rowOrder: sheet.rowOrder || null,
                        hiddenColumns: [],
                        hiddenRows: [],
                        rowMapping: sheet.rowMapping || null,
                        fromFile: false,
                        fullRewrite: false,  // WICHTIG: Keine Daten schreiben, nur Zeilen-Ops
                        structuralChange: true,
                        clearedRowHighlights: [],
                        columnOrder: null,  // Keine Spalten-Reorder im ersten Durchlauf
                        affectedRows: sheet.affectedRows || [],
                        autoFilterRange: null
                    }
                };
                
                const rowResult = await writeExcel(rowConfig);
                if (!rowResult.success) {
                    hasError = true;
                    errorMessage = rowResult.error;
                    safeError(`[Python] Zeilen-Ops für "${sheet.sheetName}" fehlgeschlagen:`, rowResult.error);
                    continue;
                }
                // Track actual method used (might be fallback)
                if (rowResult.method) actualMethod = rowResult.method;
                safeLog(`[Python] Zeilen-Ops für "${sheet.sheetName}" erfolgreich (${rowResult.method})`);
                
                // SCHRITT 2: Spalten-Operationen (mit allen Daten, fullRewrite=true)
                const colConfig = {
                    filePath: targetPath,
                    outputPath: targetPath,
                    originalPath: originalSourcePath,
                    sheetName: sheet.sheetName,
                    changes: {
                        headers: sheet.headers || [],
                        data: sheet.data || [],
                        editedCells: sheet.changedCells || {},
                        cellStyles: sheet.cellStyles || {},
                        rowHighlights: sheet.rowHighlights || {},
                        deletedColumns: sheet.deletedColumnIndices || [],
                        insertedColumns: sheet.insertedColumnInfo || null,
                        deletedRowIndices: [],  // Keine Zeilen-Ops mehr (schon erledigt)
                        insertedRowInfo: null,
                        rowOrder: null,
                        hiddenColumns: sheet.hiddenColumns || [],
                        hiddenRows: sheet.hiddenRows || [],
                        rowMapping: null,  // Kein rowMapping mehr (Zeilen schon gelöscht)
                        fromFile: false,
                        fullRewrite: true,  // WICHTIG: Jetzt Daten schreiben
                        structuralChange: sheet.structuralChange || false,
                        clearedRowHighlights: sheet.clearedRowHighlights || [],
                        columnOrder: sheet.columnOrder || null,
                        affectedRows: [],
                        autoFilterRange: sheet.autoFilterRange || null
                    }
                };
                
                const colResult = await writeExcel(colConfig);
                if (!colResult.success) {
                    hasError = true;
                    errorMessage = colResult.error;
                    safeError(`[Python] Spalten-Ops für "${sheet.sheetName}" fehlgeschlagen:`, colResult.error);
                } else {
                    results.push(sheet.sheetName);
                    // Track actual method used
                    if (colResult.method) actualMethod = colResult.method;
                    safeLog(`[Python] Spalten-Ops für "${sheet.sheetName}" erfolgreich (${colResult.method})`);
                }
                
            } else {
                // EINZELNE OPERATIONEN: Normaler Aufruf (bestehender Code)
                const config = {
                    filePath: targetPath,
                    outputPath: targetPath,
                    originalPath: originalSourcePath,
                    sheetName: sheet.sheetName,
                    changes: {
                        headers: sheet.headers || [],
                        data: sheet.data || [],
                        editedCells: sheet.changedCells || {},
                        cellStyles: sheet.cellStyles || {},
                        rowHighlights: sheet.rowHighlights || {},
                        deletedColumns: sheet.deletedColumnIndices || [],
                        insertedColumns: sheet.insertedColumnInfo || null,
                        deletedRowIndices: sheet.deletedRowIndices || [],
                        insertedRowInfo: sheet.insertedRowInfo || null,
                        rowOrder: sheet.rowOrder || null,
                        hiddenColumns: sheet.hiddenColumns || [],
                        hiddenRows: sheet.hiddenRows || [],
                        rowMapping: sheet.rowMapping || null,
                        fromFile: sheet.fromFile || false,
                        fullRewrite: sheet.fullRewrite || false,
                        structuralChange: sheet.structuralChange || false,
                        clearedRowHighlights: sheet.clearedRowHighlights || [],
                        columnOrder: sheet.columnOrder || null,
                        affectedRows: sheet.affectedRows || [],
                        autoFilterRange: sheet.autoFilterRange || null
                    }
                };
                
                const result = await writeExcel(config);
            
                if (!result.success) {
                    hasError = true;
                    errorMessage = result.error;
                    safeError(`[Python] Sheet "${sheet.sheetName}" failed:`, result.error);
                } else {
                    results.push(sheet.sheetName);
                    // Track actual method used
                    if (result.method) actualMethod = result.method;
                    safeLog(`[Python] Sheet "${sheet.sheetName}" erfolgreich (${result.method})`);
                }
            }
            
        } catch (error) {
            hasError = true;
            errorMessage = error.message;
            safeError(`[Python] Sheet "${sheet.sheetName}" exception:`, error.message);
        }
    }
    
    if (hasError && results.length === 0) {
        return { success: false, error: errorMessage };
    }
    
    // Passwortschutz anwenden falls gewünscht
    // xlwings/openpyxl unterstützt keinen Passwortschutz, daher verwenden wir xlsx-populate
    if (options.password) {
        try {
            const XlsxPopulate = require('xlsx-populate');
            const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath);
            await pwWorkbook.toFileAsync(targetPath, { password: options.password });
        } catch (pwError) {
            safeError('[Python] Fehler beim Passwortschutz:', pwError.message);
            // Datei wurde bereits gespeichert, nur ohne Passwort
        }
    }
    
    // Ermittle verwendete Methode - nutze tatsächliche Methode falls verfügbar
    const finalMethod = actualMethod || (await isExcelAvailable() ? 'xlwings' : 'openpyxl');
    
    return {
        success: true,
        message: `${results.length} Sheet(s) exportiert`,
        sheetsExported: results,
        method: finalMethod
    };
}

/**
 * Prüft ob Microsoft Excel installiert und verfügbar ist
 * Verwendet den zentralen isExcelAvailable() Check mit Caching
 */
async function checkExcelAvailable() {
    const available = await isExcelAvailable();
    const engine = getExcelEngine();
    
    return {
        success: true,
        excelAvailable: available,
        configuredEngine: engine,
        method: available ? 'xlwings' : 'openpyxl',
        message: available 
            ? `Microsoft Excel verfügbar - xlwings wird verwendet (Engine: ${engine})`
            : `Microsoft Excel nicht verfügbar - openpyxl wird verwendet (Engine: ${engine})`
    };
}

module.exports = {
    getPythonPath,
    callPython,
    listSheets,
    readSheet,
    writeExcel,
    writeExcelOpenpyxl,
    exportMultipleSheets,
    checkExcelAvailable,
    hasXlwingsSupport,
    isExcelAvailable,
    resetExcelCache,
    setExcelEngine,
    getExcelEngine
};
