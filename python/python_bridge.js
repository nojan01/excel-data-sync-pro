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

// Python-Pfad ermitteln (venv oder system)
function getPythonPath() {
    const venvPath = path.join(__dirname, '..', '.venv');
    
    // Prüfe ob venv existiert
    if (fs.existsSync(venvPath)) {
        if (process.platform === 'win32') {
            return path.join(venvPath, 'Scripts', 'python.exe');
        } else {
            return path.join(venvPath, 'bin', 'python3');
        }
    }
    
    // Fallback auf System-Python
    return process.platform === 'win32' ? 'python' : 'python3';
}

// Prüfe ob xlwings-Scripts existieren
function hasXlwingsSupport() {
    // TEMPORÄR DEAKTIVIERT FÜR FALLBACK-TEST - wieder aktivieren mit: return fs.existsSync(readerPath) && fs.existsSync(writerPath);
    return false;
    const readerPath = path.join(__dirname, 'excel_reader_xlwings.py');
    const writerPath = path.join(__dirname, 'excel_writer_xlwings.py');
    return fs.existsSync(readerPath) && fs.existsSync(writerPath);
}

/**
 * Führt ein Python-Script aus und gibt das JSON-Ergebnis zurück
 */
async function callPython(scriptName, args = []) {
    const pythonPath = getPythonPath();
    const scriptPath = path.join(__dirname, scriptName);
    
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
            
            if (code !== 0) {
                console.error(`[Python] Error:`, stderr);
                reject(new Error(stderr || `Python script exited with code ${code}`));
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                resolve(result);
            } catch (parseError) {
                console.error(`[Python] JSON parse error:`, parseError.message);
                console.error(`[Python] stdout:`, stdout.substring(0, 500));
                reject(new Error(`Failed to parse Python output: ${parseError.message}`));
            }
        });
        
        proc.on('error', (error) => {
            console.error(`[Python] Spawn error:`, error.message);
            reject(error);
        });
    });
}

/**
 * Liste alle Sheets in einer Excel-Datei
 */
async function listSheets(filePath) {
    // openpyxl ist schneller zum Lesen
    return await callPython('excel_reader.py', ['list_sheets', filePath]);
}

/**
 * Liest ein Sheet mit allen Styles
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des Sheets
 * @returns {Promise<Object>} Sheet-Daten im Format für die GUI
 */
async function readSheet(filePath, sheetName) {
    let result;
    
    // xlwings ist auf macOS zu langsam zum Lesen (AppleScript Overhead)
    // Verwende openpyxl zum Lesen - xlwings nur zum Schreiben (für CF-Erhaltung)
    result = await callPython('excel_reader.py', ['read_sheet', filePath, sheetName]);
    if (result.success) {
        result.method = 'openpyxl';
    }
    
    if (!result.success) {
        return result;
    }
    
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
 * Verwendet primär xlwings für perfekte CF-Erhaltung
 */
async function writeExcel(config) {
    const pythonPath = getPythonPath();
    
    // Versuche zuerst xlwings für CF-Erhalt
    let scriptPath;
    let useXlwings = false;
    
    if (hasXlwingsSupport()) {
        scriptPath = path.join(__dirname, 'excel_writer_xlwings.py');
        useXlwings = true;
    } else {
        scriptPath = path.join(__dirname, 'excel_writer.py');
    }
    
    return new Promise((resolve, reject) => {
        const startTime = Date.now();
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
        
        pythonProcess.on('close', (code) => {
            const duration = Date.now() - startTime;
            
            if (code !== 0) {
                console.error(`[Python] Write Error:`, stderr);
                
                // WICHTIG: Kein Fallback mehr - wir wollen den echten Fehler sehen!
                // Der openpyxl Fallback verursacht doppeltes Excel-Öffnen und hängt.
                reject(new Error(stderr || `Python writer exited with code ${code}`));
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                result.method = useXlwings ? 'xlwings' : 'openpyxl';
                resolve(result);
            } catch (parseError) {
                console.error(`[Python] JSON parse error:`, parseError.message);
                console.error(`[Python] stdout:`, stdout.substring(0, 500));
                reject(new Error(`Failed to parse Python output: ${parseError.message}`));
            }
        });

        pythonProcess.on('error', (error) => {
            console.error(`[Python] Spawn error:`, error.message);
            reject(error);
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
    const scriptPath = path.join(__dirname, 'excel_writer.py');
    
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
    
    // Zuerst: Kopiere die Original-Datei zum Ziel (falls unterschiedlich)
    // So bleiben alle Sheets, Formatierungen, etc. erhalten
    if (sourcePath !== targetPath) {
        try {
            fs.copyFileSync(sourcePath, targetPath);
        } catch (copyError) {
            console.error(`[Python] Fehler beim Kopieren:`, copyError.message);
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
            const config = {
                filePath: targetPath,  // Jetzt immer vom Ziel lesen (wir haben es kopiert)
                outputPath: targetPath,
                originalPath: sourcePath,  // Original-Pfad für restore_table_xml
                sheetName: sheet.sheetName,
                changes: {
                    headers: sheet.headers || [],
                    data: sheet.data || [],
                    editedCells: sheet.changedCells || {},
                    cellStyles: sheet.cellStyles || {},
                    rowHighlights: sheet.rowHighlights || {},
                    deletedColumns: sheet.deletedColumnIndices || [],
                    insertedColumns: sheet.insertedColumnInfo || null,
                    // Zeilen-Operationen (NEU - analog zu Spalten)
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
                console.error(`[Python] Sheet "${sheet.sheetName}" failed:`, result.error);
            } else {
                results.push(sheet.sheetName);
            }
            
        } catch (error) {
            hasError = true;
            errorMessage = error.message;
            console.error(`[Python] Sheet "${sheet.sheetName}" exception:`, error.message);
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
            console.error('[Python] Fehler beim Passwortschutz:', pwError.message);
            // Datei wurde bereits gespeichert, nur ohne Passwort
        }
    }
    
    return {
        success: true,
        message: `${results.length} Sheet(s) exportiert`,
        sheetsExported: results
    };
}

/**
 * Prüft ob Microsoft Excel installiert und verfügbar ist
 * Wenn ja, werden strukturelle Änderungen mit xlwings durchgeführt
 * für perfekten CF-Erhalt
 */
async function checkExcelAvailable() {
    // Versuche zuerst xlwings
    if (hasXlwingsSupport()) {
        try {
            return await callPython('excel_writer_xlwings.py', ['check_excel']);
        } catch (error) {
            console.log('[Python] xlwings check failed:', error.message);
        }
    }
    
    // Fallback auf openpyxl check
    const pythonPath = getPythonPath();
    const scriptPath = path.join(__dirname, 'excel_writer.py');
    
    return new Promise((resolve, reject) => {
        const pythonProcess = spawn(pythonPath, [scriptPath, 'check_excel']);
        
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
                console.log(`[Python] Excel check failed:`, stderr);
                resolve({ 
                    success: true, 
                    excelAvailable: false,
                    message: 'Excel-Prüfung fehlgeschlagen'
                });
                return;
            }
            
            try {
                const result = JSON.parse(stdout);
                console.log(`[Python] Excel status:`, result);
                resolve(result);
            } catch (parseError) {
                resolve({ 
                    success: true, 
                    excelAvailable: false,
                    message: 'Excel-Status konnte nicht ermittelt werden'
                });
            }
        });
        
        pythonProcess.on('error', (error) => {
            resolve({ 
                success: true, 
                excelAvailable: false,
                message: error.message
            });
        });
    });
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
    hasXlwingsSupport
};
