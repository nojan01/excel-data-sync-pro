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
 */
async function writeExcel(config) {
    const configJson = JSON.stringify(config);
    return callPython('excel_writer.py', ['write', configJson]);
}

module.exports = {
    getPythonPath,
    callPython,
    listSheets,
    readSheet,
    writeExcel
};
