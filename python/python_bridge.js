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
                result.debugLog = stderr || '';  // Python stderr für Debugging
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
                result.debugLog = stderr || '';  // Python stderr für Debugging
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
                    let clonedXml = await srcFile.async('string');

                    // Strip elements that reference shared resources (tables, drawings, slicers)
                    // These can't be duplicated by simple copy and cause Excel repair errors

                    // 1. <tableParts> — references original tables → conflict
                    clonedXml = clonedXml.replace(/<tableParts[\s\S]*?<\/tableParts>/g, '');
                    clonedXml = clonedXml.replace(/<tableParts[^>]*\/>/g, '');

                    // 2. <drawing r:id="..."/> — references drawing via removed relationship
                    clonedXml = clonedXml.replace(/<drawing\s[^>]*\/>/g, '');
                    clonedXml = clonedXml.replace(/<drawing[\s\S]*?<\/drawing>/g, '');

                    // 3. <legacyDrawing r:id="..."/> — legacy drawing shapes (comments, form controls)
                    clonedXml = clonedXml.replace(/<legacyDrawing\s[^>]*\/>/g, '');

                    // 4. Slicer/timeline extensions in <extLst> that reference removed relationships
                    // Remove <ext> blocks containing slicer or timeline references
                    clonedXml = clonedXml.replace(/<ext\s[^>]*>[\s\S]*?<\/ext>/g, (match) => {
                        if (/slicer|timeline/i.test(match)) return '';
                        return match;
                    });
                    // Clean up empty <extLst> after removing extensions
                    clonedXml = clonedXml.replace(/<extLst>\s*<\/extLst>/g, '');

                    // 5. <oleObjects>, <controls> — embedded objects referencing removed rels
                    clonedXml = clonedXml.replace(/<oleObjects[\s\S]*?<\/oleObjects>/g, '');
                    clonedXml = clonedXml.replace(/<controls[\s\S]*?<\/controls>/g, '');

                    // New file
                    const wsFiles = Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));
                    const nums = wsFiles.map(f => parseInt(f.match(/sheet(\d+)/)[1]));
                    const newNum = (nums.length > 0 ? Math.max(...nums) : 0) + 1;
                    const newFile = `worksheets/sheet${newNum}.xml`;
                    zip.file(`xl/${newFile}`, clonedXml);

                    // Copy sheet-level rels if exist, but EXCLUDE table/slicer/pivotTable/drawing refs
                    // These reference shared resources that can't be duplicated by simple copy
                    const srcSheetNum = srcTarget.match(/sheet(\d+)/)?.[1];
                    if (srcSheetNum) {
                        const srcRelsPath = `xl/worksheets/_rels/sheet${srcSheetNum}.xml.rels`;
                        const srcSheetRelsFile = zip.file(srcRelsPath);
                        if (srcSheetRelsFile) {
                            let sheetRelsXml = await srcSheetRelsFile.async('string');
                            // Remove relationships to tables, slicers, pivotTables, drawings
                            // These share internal IDs and would cause Excel repair errors
                            const excludeTypes = [
                                'table', 'slicer', 'pivotTable', 'drawing',
                                'pivotCacheDefinition', 'slicerCache'
                            ];
                            for (const t of excludeTypes) {
                                sheetRelsXml = sheetRelsXml.replace(
                                    new RegExp(`\\s*<Relationship[^>]*Type="[^"]*${t}[^"]*"[^>]*/>`, 'gi'), '');
                            }
                            // Only write rels file if there are remaining relationships
                            if (/<Relationship\s/i.test(sheetRelsXml)) {
                                zip.file(`xl/worksheets/_rels/sheet${newNum}.xml.rels`, sheetRelsXml);
                            }
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
    let debugLogs = null; // Collect Python stderr debug output
    
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
    
    // Wenn die Quelldatei passwortgeschützt ist, muss die Kopie entschlüsselt werden
    // damit openpyxl die Datei bearbeiten kann.
    // Nach dem Schreiben wird die Datei je nach Benutzer-Wahl wieder verschlüsselt.
    if (options.sourcePassword) {
        try {
            const XlsxPopulate = require('xlsx-populate');
            const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath, { password: options.sourcePassword });
            await pwWorkbook.toFileAsync(targetPath); // Ohne Passwort speichern = entschlüsseln
            safeLog('[Python] Datei erfolgreich entschlüsselt für Bearbeitung');
        } catch (decryptError) {
            safeError('[Python] Fehler beim Entschlüsseln der Quelldatei:', decryptError.message);
            return { success: false, error: `Fehler beim Entschlüsseln: ${decryptError.message}` };
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
                        editedCells: sheet.changedCells || {},  // WICHTIG: editedCells mitsenden für Zeilen-Verschiebung ohne Block-Write
                        cellStyles: {},
                        rowHighlights: {},
                        deletedColumns: [],  // Keine Spalten-Ops im ersten Durchlauf
                        insertedColumns: null,
                        deletedRowIndices: sheet.deletedRowIndices || [],
                        insertedRowInfo: sheet.insertedRowInfo || null,
                        rowOrder: sheet.rowOrder || null,
                        hiddenColumns: [],
                        hiddenRows: sheet.hiddenRows || [],  // Hidden Rows im Zeilen-Pass anwenden (ZIP-ANSATZ unterstützt sie)
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
                // WICHTIG: originalPath = targetPath (NICHT originalSourcePath!)
                // Pass 1 hat die Zeilen-Ops bereits in targetPath gespeichert.
                // Wenn wir hier originalSourcePath verwenden, kopiert XML-DIREKT
                // von der unveränderten Datei und überschreibt die Zeilen-Änderungen!
                const colConfig = {
                    filePath: targetPath,
                    outputPath: targetPath,
                    originalPath: targetPath,
                    sheetName: sheet.sheetName,
                    changes: {
                        headers: sheet.headers || [],
                        data: sheet.data || [],
                        editedCells: sheet.changedCells || {},
                        cellStyles: sheet.cellStyles || {},
                        cellFonts: sheet.cellFonts || {},
                        richTextCells: sheet.richTextCells || {},
                        rowHighlights: sheet.rowHighlights || {},
                        mergedCells: sheet.mergedCells || [],
                        deletedColumns: sheet.deletedColumnIndices || [],
                        insertedColumns: sheet.insertedColumnInfo || null,
                        deletedRowIndices: [],  // Keine Zeilen-Ops mehr (schon erledigt)
                        insertedRowInfo: null,
                        rowOrder: null,
                        hiddenColumns: sheet.hiddenColumns || [],
                        hiddenRows: [],  // Schon in Pass 1 angewendet
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
                    // Collect debug logs from Python stderr
                    if (colResult.debugLog) {
                        if (!debugLogs) debugLogs = [];
                        debugLogs.push(colResult.debugLog);
                    }
                    safeLog(`[Python] Spalten-Ops für "${sheet.sheetName}" erfolgreich (${colResult.method})`);
                }
                
            } else if (hasColOps && !hasRowOps && 
                       sheet.changedCells && Object.keys(sheet.changedCells).some(k => !k.startsWith('_'))) {
                // SPALTEN + ZELL-EDITS: Erst Spalten via XML-DIREKT, dann Zell-Edits via FALL 3a
                // Ohne diese Trennung blockiert has_cell_edits den XML-DIREKT-Pfad,
                // und der openpyxl-Roundtrip zerstört AutoFilter/Tabelle.
                safeLog(`[Python] Spalten+Zell-Edits: Erst Spalten, dann Zell-Edits für "${sheet.sheetName}"`);
                
                // Separiere echte Zell-Edits von Marker-Einträgen (_columnDeleted etc.)
                const onlyCellEdits = {};
                if (sheet.changedCells) {
                    for (const [key, value] of Object.entries(sheet.changedCells)) {
                        if (!key.startsWith('_')) {
                            onlyCellEdits[key] = value;
                        }
                    }
                }
                
                // SCHRITT 1: Spalten-Operationen via XML-DIREKT (OHNE Zell-Edits)
                const colOnlyConfig = {
                    filePath: targetPath,
                    outputPath: targetPath,
                    originalPath: originalSourcePath,
                    sheetName: sheet.sheetName,
                    changes: {
                        headers: sheet.headers || [],
                        data: sheet.data || [],
                        editedCells: {},  // KEINE Zell-Edits → only_column_ops = True → XML-DIREKT
                        cellStyles: {},
                        cellFonts: {},
                        richTextCells: {},
                        rowHighlights: {},
                        mergedCells: [],
                        deletedColumns: sheet.deletedColumnIndices || [],
                        insertedColumns: sheet.insertedColumnInfo || null,
                        deletedRowIndices: [],
                        insertedRowInfo: null,
                        rowOrder: null,
                        hiddenColumns: [],
                        hiddenRows: [],
                        rowMapping: null,
                        fromFile: false,
                        fullRewrite: false,
                        structuralChange: true,  // Nötig um in FALL 1/2 Block zu kommen
                        clearedRowHighlights: [],
                        columnOrder: sheet.columnOrder || null,
                        affectedRows: [],
                        autoFilterRange: null,
                        hasFormatChanges: false
                    }
                };
                
                const colOnlyResult = await writeExcel(colOnlyConfig);
                if (!colOnlyResult.success) {
                    hasError = true;
                    errorMessage = colOnlyResult.error;
                    safeError(`[Python] Spalten-Ops für "${sheet.sheetName}" fehlgeschlagen:`, colOnlyResult.error);
                    continue;
                }
                if (colOnlyResult.method) actualMethod = colOnlyResult.method;
                safeLog(`[Python] Spalten-Ops für "${sheet.sheetName}" erfolgreich (${colOnlyResult.method})`);
                
                // SCHRITT 2: Zell-Edits + Highlights + Visibility via FALL 3a (OHNE Spalten-Ops)
                // originalPath = targetPath → Pass 1 hat Spalten-Ops bereits in targetPath
                const hasCellWork = Object.keys(onlyCellEdits).length > 0 ||
                    (sheet.rowHighlights && Object.keys(sheet.rowHighlights).length > 0) ||
                    (sheet.hiddenColumns && sheet.hiddenColumns.length > 0) ||
                    (sheet.hiddenRows && sheet.hiddenRows.length > 0) ||
                    (sheet.clearedRowHighlights && sheet.clearedRowHighlights.length > 0) ||
                    sheet.hasFormatChanges;
                
                if (hasCellWork) {
                    const cellConfig = {
                        filePath: targetPath,
                        outputPath: targetPath,
                        originalPath: targetPath,  // WICHTIG: Von der bereits modifizierten Datei lesen
                        sheetName: sheet.sheetName,
                        changes: {
                            headers: sheet.headers || [],
                            data: [],
                            editedCells: onlyCellEdits,
                            cellStyles: sheet.cellStyles || {},
                            cellFonts: sheet.cellFonts || {},
                            richTextCells: sheet.richTextCells || {},
                            rowHighlights: sheet.rowHighlights || {},
                            mergedCells: sheet.mergedCells || [],
                            deletedColumns: [],
                            insertedColumns: null,
                            deletedRowIndices: [],
                            insertedRowInfo: null,
                            rowOrder: null,
                            hiddenColumns: sheet.hiddenColumns || [],
                            hiddenRows: sheet.hiddenRows || [],
                            rowMapping: null,
                            fromFile: false,
                            fullRewrite: false,  // KEIN Full-Rewrite → FALL 3
                            structuralChange: false,  // KEIN Structural Change → FALL 3
                            clearedRowHighlights: sheet.clearedRowHighlights || [],
                            columnOrder: null,
                            affectedRows: [],
                            autoFilterRange: null,
                            hasFormatChanges: sheet.hasFormatChanges || false
                        }
                    };
                    
                    const cellResult = await writeExcel(cellConfig);
                    if (!cellResult.success) {
                        hasError = true;
                        errorMessage = cellResult.error;
                        safeError(`[Python] Zell-Edits für "${sheet.sheetName}" fehlgeschlagen:`, cellResult.error);
                    } else {
                        results.push(sheet.sheetName);
                        if (cellResult.method) actualMethod = cellResult.method;
                        if (cellResult.debugLog) {
                            if (!debugLogs) debugLogs = [];
                            debugLogs.push(cellResult.debugLog);
                        }
                        safeLog(`[Python] Zell-Edits für "${sheet.sheetName}" erfolgreich (${cellResult.method})`);
                    }
                } else {
                    // Nur Spalten-Ops, keine Zell-Edits → schon fertig
                    results.push(sheet.sheetName);
                    safeLog(`[Python] Nur Spalten-Ops für "${sheet.sheetName}" (keine Zell-Edits)`);
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
                        cellFonts: sheet.cellFonts || {},
                        richTextCells: sheet.richTextCells || {},
                        rowHighlights: sheet.rowHighlights || {},
                        mergedCells: sheet.mergedCells || [],
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
                        autoFilterRange: sheet.autoFilterRange || null,
                        hasFormatChanges: sheet.hasFormatChanges || false
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
                    // Collect debug logs from Python stderr
                    if (result.debugLog) {
                        if (!debugLogs) debugLogs = [];
                        debugLogs.push(result.debugLog);
                    }
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
    
    // Passwortschutz anwenden
    // xlwings/openpyxl unterstützt keinen Passwortschutz, daher verwenden wir xlsx-populate
    // options.password == null (null/undefined): Checkbox nicht aktiviert → Original-Passwort beibehalten
    // options.password === '':                   Checkbox aktiviert, leer → Passwort entfernen
    // options.password === 'xxx':                Checkbox aktiviert mit Wert → Neues Passwort setzen
    // WICHTIG: IPC-Layer konvertiert undefined → null (Default-Parameter), daher != null (loose equality)
    const finalPassword = (options.password != null) ? options.password : options.sourcePassword;
    if (finalPassword) {
        try {
            const XlsxPopulate = require('xlsx-populate');
            // Datei ist jetzt entschlüsselt (wurde oben entschlüsselt oder war nie verschlüsselt)
            const pwWorkbook = await XlsxPopulate.fromFileAsync(targetPath);
            await pwWorkbook.toFileAsync(targetPath, { password: finalPassword });
            safeLog(`[Python] Passwortschutz angewendet (${options.password !== undefined ? 'neues Passwort' : 'Original-Passwort beibehalten'})`);
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
        method: finalMethod,
        debugLog: debugLogs ? debugLogs.join('\n---\n') : ''
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
