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

// Gemeinsames Python-Umgebungsmodul (gecachte Pfad-Ermittlung)
const { getPythonPath, getPythonBasePath, getPythonEnv, safeLog } = require('./python_env');

// Sichere Error-Log-Funktion
function safeError(...args) {
    try {
        if (process.stderr && process.stderr.writable) {
            console.error(...args);
        }
    } catch (e) {
        // Ignoriere Konsolenfehler
    }
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
 * @param {string} scriptName - Name des Python-Scripts
 * @param {string[]} args - Argumente
 * @param {number} timeoutMs - Timeout in Millisekunden (Standard: 120s, 0 = kein Timeout)
 */
async function callPython(scriptName, args = [], timeoutMs = 120000) {
    const pythonPath = getPythonPath();
    const basePath = getPythonBasePath();
    const scriptPath = path.resolve(basePath, scriptName);
    const env = getPythonEnv();
    // UTF-8 für Python erzwingen (Windows nutzt sonst CP1252 für Umlaute)
    env.PYTHONUTF8 = '1';
    env.PYTHONIOENCODING = 'utf-8';
    
    // Sicherheitsprüfung: Script muss innerhalb von basePath liegen
    if (!scriptPath.startsWith(path.resolve(basePath))) {
        throw new Error('Ungültiger Script-Pfad: Path Traversal erkannt');
    }
    
    safeLog(`[Python] callPython: ${scriptName} ${args.join(' ')}`);
    safeLog(`[Python]   pythonPath: ${pythonPath}`);
    safeLog(`[Python]   scriptPath: ${scriptPath}`);
    safeLog(`[Python]   scriptExists: ${fs.existsSync(scriptPath)}`);
    safeLog(`[Python]   PYTHONPATH: ${env.PYTHONPATH || '(nicht gesetzt)'}`);
    
    return new Promise((resolve, reject) => {
        const startTime = Date.now();
        const proc = spawn(pythonPath, [scriptPath, ...args], { env });
        let isSettled = false;
        
        let stdout = '';
        let stderr = '';
        
        // Timeout: Prozess nach Ablauf killen
        let timeoutHandle = null;
        if (timeoutMs > 0) {
            timeoutHandle = setTimeout(() => {
                if (!isSettled) {
                    isSettled = true;
                    safeError(`[Python] Timeout nach ${timeoutMs}ms - Prozess wird beendet`);
                    proc.kill('SIGTERM');
                    // Falls SIGTERM nicht wirkt, nach 5s SIGKILL
                    setTimeout(() => {
                        try { proc.kill('SIGKILL'); } catch (e) { /* already dead */ }
                    }, 5000);
                    reject(new Error(`Python-Skript Timeout nach ${timeoutMs / 1000}s`));
                }
            }, timeoutMs);
        }
        
        proc.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        
        proc.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        proc.on('close', (code) => {
            if (timeoutHandle) clearTimeout(timeoutHandle);
            if (isSettled) return; // Timeout hat bereits rejected
            isSettled = true;
            
            const duration = Date.now() - startTime;
            safeLog(`[Python] Script beendet in ${duration}ms, code=${code}`);
            
            if (stderr) {
                safeLog(`[Python] stderr: ${stderr.substring(0, 500)}`);
            }
            
            if (code !== 0) {
                safeError(`[Python] Error:`, stderr);
                // Generische Fehlermeldung ans Frontend, Details nur im Log
                reject(new Error(`Python-Skript fehlgeschlagen (Code ${code})`));
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
            if (timeoutHandle) clearTimeout(timeoutHandle);
            if (isSettled) return;
            isSettled = true;
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
    
    // Umgebungsvariablen für Python
    const env = getPythonEnv();
    env.PYTHONUTF8 = '1';
    env.PYTHONIOENCODING = 'utf-8';
    
    return new Promise((resolve, reject) => {
        const startTime = Date.now();
        safeLog(`[Python] Starte: ${pythonPath} ${scriptPath} write_sheet`);
        safeLog(`[Python] PYTHONPATH: ${env.PYTHONPATH || '(nicht gesetzt)'}`);
        const pythonProcess = spawn(pythonPath, [scriptPath, 'write_sheet'], { env });
        
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
    const env = getPythonEnv();
    env.PYTHONUTF8 = '1';
    env.PYTHONIOENCODING = 'utf-8';
    
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn(pythonPath, [scriptPath, 'write_sheet'], { env });
        
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
 * Wendet deferred Sheet-Operationen auf eine XLSX-Datei an (JSZip-basiert).
 * Wird beim Export aufgerufen, um Add/Delete/Rename/Clone/Move/Visibility-Ops anzuwenden,
 * die im Offline-Modus nur im Speicher gehalten wurden.
 */
async function applyPendingSheetOperations(filePath, operations) {
    if (!operations || operations.length === 0) return { success: true };

    const JSZip = require('jszip');
    const fileData = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(fileData);

    let workbookXml = await zip.file('xl/workbook.xml').async('string');
    let relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    let contentTypesXml = await zip.file('[Content_Types].xml').async('string');

    function xmlEncode(name) {
        return name
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
    function regexEscape(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    for (const op of operations) {
        try {
            switch (op.type) {
                case 'add': {
                    const enc = xmlEncode(op.sheetName);

                    const sheetIdMatches = [...workbookXml.matchAll(/sheetId="(\d+)"/g)];
                    const maxSheetId = sheetIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
                    const newSheetId = maxSheetId + 1;

                    const rIdMatches = [...relsXml.matchAll(/Id="rId(\d+)"/g)];
                    const maxRId = rIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
                    const newRId = `rId${maxRId + 1}`;

                    const wsFiles = Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
                    const nums = wsFiles.map(f => parseInt(f.match(/sheet(\d+)/)[1]));
                    const newNum = (nums.length > 0 ? Math.max(...nums) : 0) + 1;
                    const newFile = `worksheets/sheet${newNum}.xml`;

                    zip.file(`xl/${newFile}`,
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                        '<sheetData/></worksheet>');

                    workbookXml = workbookXml.replace('</sheets>',
                        `<sheet name="${enc}" sheetId="${newSheetId}" r:id="${newRId}"/></sheets>`);
                    relsXml = relsXml.replace('</Relationships>',
                        `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newFile}"/></Relationships>`);
                    contentTypesXml = contentTypesXml.replace('</Types>',
                        `<Override PartName="/xl/${newFile}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);

                    safeLog(`[PendingOps] Added sheet "${op.sheetName}"`);
                    break;
                }

                case 'delete': {
                    const enc = xmlEncode(op.sheetName);
                    const esc = regexEscape(enc);

                    // Find rId
                    const sheetRe = new RegExp(`<sheet[^>]*name="${esc}"[^>]*r:id="(rId\\d+)"[^>]*/>`);
                    const sheetM = workbookXml.match(sheetRe);
                    if (!sheetM) { safeLog(`[PendingOps] Sheet "${op.sheetName}" not found for delete`); break; }
                    const rId = sheetM[1];

                    // Find target file
                    const relRe = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"[^>]*/>`);
                    const relM = relsXml.match(relRe);
                    const target = relM ? relM[1] : null;

                    // Remove from workbook.xml
                    workbookXml = workbookXml.replace(new RegExp(`\\s*<sheet[^>]*name="${esc}"[^>]*/>`), '');

                    // Remove relationship
                    relsXml = relsXml.replace(new RegExp(`\\s*<Relationship[^>]*Id="${rId}"[^>]*/>`), '');

                    // Remove content type and file
                    if (target) {
                        const partName = target.startsWith('/') ? target : `/xl/${target}`;
                        contentTypesXml = contentTypesXml.replace(
                            new RegExp(`\\s*<Override[^>]*PartName="${regexEscape(partName)}"[^>]*/>`), '');
                        const zipPath = partName.startsWith('/') ? partName.slice(1) : `xl/${target}`;
                        zip.remove(zipPath);
                    }

                    safeLog(`[PendingOps] Deleted sheet "${op.sheetName}"`);
                    break;
                }

                case 'rename': {
                    const encOld = xmlEncode(op.oldName);
                    const encNew = xmlEncode(op.newName);
                    const escOld = regexEscape(encOld);
                    workbookXml = workbookXml.replace(
                        new RegExp(`(<sheet[^>]*name=")${escOld}(")`), `$1${encNew}$2`);
                    safeLog(`[PendingOps] Renamed "${op.oldName}" -> "${op.newName}"`);
                    break;
                }

                case 'clone': {
                    const encSrc = xmlEncode(op.sourceSheet);
                    const encNew = xmlEncode(op.newName);
                    const escSrc = regexEscape(encSrc);

                    // Find source rId
                    const srcRe = new RegExp(`<sheet[^>]*name="${escSrc}"[^>]*r:id="(rId\\d+)"[^>]*/>`);
                    const srcM = workbookXml.match(srcRe);
                    if (!srcM) { safeLog(`[PendingOps] Source "${op.sourceSheet}" not found for clone`); break; }
                    const srcRId = srcM[1];

                    // Find source target file
                    const srcRelRe = new RegExp(`<Relationship[^>]*Id="${srcRId}"[^>]*Target="([^"]+)"[^>]*/>`);
                    const srcRelM = relsXml.match(srcRelRe);
                    if (!srcRelM) break;
                    const srcTarget = srcRelM[1];
                    const srcZipPath = srcTarget.startsWith('/') ? srcTarget.slice(1) : `xl/${srcTarget}`;

                    const srcFile = zip.file(srcZipPath);
                    if (!srcFile) break;
                    const srcXml = await srcFile.async('string');

                    // New file
                    const wsFiles = Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
                    const nums = wsFiles.map(f => parseInt(f.match(/sheet(\d+)/)[1]));
                    const newNum = (nums.length > 0 ? Math.max(...nums) : 0) + 1;
                    const newFile = `worksheets/sheet${newNum}.xml`;
                    zip.file(`xl/${newFile}`, srcXml);

                    // Copy sheet-level rels if exist
                    const srcSheetNum = srcTarget.match(/sheet(\d+)/)?.[1];
                    if (srcSheetNum) {
                        const srcRelsPath = `xl/worksheets/_rels/sheet${srcSheetNum}.xml.rels`;
                        const srcSheetRelsFile = zip.file(srcRelsPath);
                        if (srcSheetRelsFile) {
                            zip.file(`xl/worksheets/_rels/sheet${newNum}.xml.rels`,
                                await srcSheetRelsFile.async('string'));
                        }
                    }

                    // Next IDs
                    const sheetIdMatches = [...workbookXml.matchAll(/sheetId="(\d+)"/g)];
                    const maxSId = sheetIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
                    const rIdMatches = [...relsXml.matchAll(/Id="rId(\d+)"/g)];
                    const maxRId = rIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1])), 0);
                    const newSheetId = maxSId + 1;
                    const newRId = `rId${maxRId + 1}`;

                    // Insert after source sheet tag
                    const newTag = `<sheet name="${encNew}" sheetId="${newSheetId}" r:id="${newRId}"/>`;
                    workbookXml = workbookXml.replace(srcM[0], `${srcM[0]}${newTag}`);

                    relsXml = relsXml.replace('</Relationships>',
                        `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newFile}"/></Relationships>`);
                    contentTypesXml = contentTypesXml.replace('</Types>',
                        `<Override PartName="/xl/${newFile}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);

                    safeLog(`[PendingOps] Cloned "${op.sourceSheet}" -> "${op.newName}"`);
                    break;
                }

                case 'move': {
                    const sheetTagRegex = /<sheet[^>]+\/>/g;
                    const sheetTags = workbookXml.match(sheetTagRegex);
                    if (!sheetTags || sheetTags.length < 2) break;

                    const enc = xmlEncode(op.sheetName);
                    const currentIdx = sheetTags.findIndex(t => t.includes(`name="${enc}"`));
                    if (currentIdx === -1) break;

                    const [movedTag] = sheetTags.splice(currentIdx, 1);
                    sheetTags.splice(op.newIndex, 0, movedTag);

                    workbookXml = workbookXml.replace(/<sheets>[\s\S]*?<\/sheets>/,
                        `<sheets>${sheetTags.join('')}</sheets>`);

                    safeLog(`[PendingOps] Moved "${op.sheetName}" to index ${op.newIndex}`);
                    break;
                }

                case 'visibility': {
                    const enc = xmlEncode(op.sheetName);
                    const esc = regexEscape(enc);

                    if (op.visible) {
                        workbookXml = workbookXml.replace(
                            new RegExp(`(<sheet[^>]*name="${esc}"[^>]*?)\\s+state="(?:hidden|veryHidden)"`, 'g'), '$1');
                    } else {
                        const hasState = new RegExp(`<sheet[^>]*name="${esc}"[^>]*state=`).test(workbookXml);
                        if (hasState) {
                            workbookXml = workbookXml.replace(
                                new RegExp(`(<sheet[^>]*name="${esc}"[^>]*)state="[^"]*"`, 'g'), '$1state="hidden"');
                        } else {
                            workbookXml = workbookXml.replace(
                                new RegExp(`(<sheet[^>]*name="${esc}"[^/]*)(/>)`, 'g'), '$1 state="hidden"$2');
                        }
                    }

                    safeLog(`[PendingOps] Set "${op.sheetName}" visibility=${op.visible}`);
                    break;
                }
            }
        } catch (opError) {
            safeError(`[PendingOps] Error processing ${op.type} "${op.sheetName}":`, opError.message);
        }
    }

    // Write back
    zip.file('xl/workbook.xml', workbookXml);
    zip.file('xl/_rels/workbook.xml.rels', relsXml);
    zip.file('[Content_Types].xml', contentTypesXml);

    const outputBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(filePath, outputBuffer);

    safeLog(`[PendingOps] Applied ${operations.length} pending operations to ${path.basename(filePath)}`);
    return { success: true };
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
    
    // Deferred Sheet-Operationen anwenden (Add/Delete/Rename/Clone/Move/Visibility)
    // Diese wurden im Offline-Modus nur im Speicher gehalten und müssen jetzt
    // auf die Zieldatei angewendet werden, BEVOR die Sheet-Daten geschrieben werden.
    if (options.pendingSheetOperations && options.pendingSheetOperations.length > 0) {
        const opsResult = await applyPendingSheetOperations(targetPath, options.pendingSheetOperations);
        if (!opsResult.success) {
            return { success: false, error: `Fehler bei Sheet-Operationen: ${opsResult.error}` };
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
