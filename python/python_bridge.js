/**
 * Python Bridge für Excel Data Sync Pro
 * Ermöglicht die Kommunikation zwischen Node.js und Python/openpyxl
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

/**
 * Führt ein Python-Script aus und gibt das JSON-Ergebnis zurück
 */
async function callPython(scriptName, args = []) {
    const pythonPath = getPythonPath();
    const scriptPath = path.join(__dirname, scriptName);
    
    return new Promise((resolve, reject) => {
        console.log(`[Python] Calling ${scriptName} with args:`, args.slice(0, 2)); // Log only first 2 args for brevity
        
        const startTime = Date.now();
        const process = spawn(pythonPath, [scriptPath, ...args]);
        
        let stdout = '';
        let stderr = '';
        
        process.stdout.on('data', (data) => {
            stdout += data.toString();
        });
        
        process.stderr.on('data', (data) => {
            stderr += data.toString();
        });
        
        process.on('close', (code) => {
            const duration = Date.now() - startTime;
            console.log(`[Python] Script completed in ${duration}ms with code ${code}`);
            
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
        
        process.on('error', (error) => {
            console.error(`[Python] Spawn error:`, error.message);
            reject(error);
        });
    });
}

/**
 * Liste alle Sheets in einer Excel-Datei
 */
async function listSheets(filePath) {
    return callPython('excel_reader.py', ['list_sheets', filePath]);
}

/**
 * Liest ein Sheet mit allen Styles
 * @param {string} filePath - Pfad zur Excel-Datei
 * @param {string} sheetName - Name des Sheets
 * @returns {Promise<Object>} Sheet-Daten im Format für die GUI
 */
async function readSheet(filePath, sheetName) {
    const result = await callPython('excel_reader.py', ['read_sheet', filePath, sheetName]);
    
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
        cellHyperlinks: result.cellHyperlinks || {}
    };
}

/**
 * Schreibt Daten in eine Excel-Datei mit vollständiger Style-Erhaltung
 * Verwendet stdin für große Datenmengen
 */
async function writeExcel(config) {
    const pythonPath = getPythonPath();
    const scriptPath = path.join(__dirname, 'excel_writer.py');
    
    return new Promise((resolve, reject) => {
        console.log(`[Python] Writing Excel: ${config.outputPath}`);
        
        const startTime = Date.now();
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
            const duration = Date.now() - startTime;
            console.log(`[Python] Write completed in ${duration}ms with code ${code}`);
            
            if (code !== 0) {
                console.error(`[Python] Write Error:`, stderr);
                reject(new Error(stderr || `Python writer exited with code ${code}`));
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
 * Exportiert mehrere Sheets mit Python/openpyxl
 * Öffnet Original-Datei, modifiziert Sheets und speichert unter neuem Pfad
 */
async function exportMultipleSheets(sourcePath, targetPath, sheets, options = {}) {
    console.log(`[Python] Export: ${sheets.length} Sheets von ${sourcePath} nach ${targetPath}`);
    
    const results = [];
    let hasError = false;
    let errorMessage = '';
    
    // Zuerst: Kopiere die Original-Datei zum Ziel (falls unterschiedlich)
    // So bleiben alle Sheets, Formatierungen, etc. erhalten
    if (sourcePath !== targetPath) {
        const fsSync = require('fs');
        try {
            fsSync.copyFileSync(sourcePath, targetPath);
            console.log(`[Python] Datei kopiert: ${sourcePath} -> ${targetPath}`);
        } catch (copyError) {
            console.error(`[Python] Fehler beim Kopieren:`, copyError.message);
            return { success: false, error: `Fehler beim Kopieren: ${copyError.message}` };
        }
    }
    
    // Jetzt: Nur Sheets mit echten Änderungen modifizieren
    for (const sheet of sheets) {
        // Überspringe Sheets ohne Änderungen (fromFile: true und keine editedCells/data)
        if (sheet.fromFile && !sheet.changedCells && !sheet.data?.length && !sheet.fullRewrite) {
            console.log(`[Python] Sheet "${sheet.sheetName}" unverändert (fromFile)`);
            results.push(sheet.sheetName);
            continue;
        }
        
        try {
            const config = {
                filePath: targetPath,  // Jetzt immer vom Ziel lesen (wir haben es kopiert)
                outputPath: targetPath,
                sheetName: sheet.sheetName,
                changes: {
                    headers: sheet.headers || [],
                    data: sheet.data || [],
                    editedCells: sheet.changedCells || {},
                    cellStyles: sheet.cellStyles || {},
                    rowHighlights: sheet.rowHighlights || {},
                    deletedColumns: sheet.deletedColumnIndices || [],
                    insertedColumns: sheet.insertedColumnInfo || null,
                    hiddenColumns: sheet.hiddenColumns || [],
                    hiddenRows: sheet.hiddenRows || [],
                    rowMapping: sheet.rowMapping || null,
                    fromFile: sheet.fromFile || false
                }
            };
            
            const result = await writeExcel(config);
            
            if (!result.success) {
                hasError = true;
                errorMessage = result.error;
                console.error(`[Python] Sheet "${sheet.sheetName}" failed:`, result.error);
            } else {
                results.push(sheet.sheetName);
                console.log(`[Python] Sheet "${sheet.sheetName}" exported successfully`);
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
    
    return {
        success: true,
        message: `${results.length} Sheet(s) exportiert`,
        sheetsExported: results
    };
}

module.exports = {
    getPythonPath,
    callPython,
    listSheets,
    readSheet,
    writeExcel,
    exportMultipleSheets
};
