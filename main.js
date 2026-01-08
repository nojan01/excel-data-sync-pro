const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');

// ============================================
// JSDoc TYPE DEFINITIONS
// ============================================

/**
 * @typedef {Object} FileDialogOptions
 * @property {string} [title] - Dialog title
 * @property {string} [defaultPath] - Default file path
 * @property {Array<{name: string, extensions: string[]}>} [filters] - File type filters
 */

/**
 * @typedef {Object} ExcelReadResult
 * @property {boolean} success - Whether the operation succeeded
 * @property {string} [filePath] - Path to the file
 * @property {string[]} [sheets] - List of sheet names
 * @property {string} [error] - Error message if failed
 */

/**
 * @typedef {Object} ExcelSheetData
 * @property {boolean} success - Whether the operation succeeded
 * @property {string[]} [headers] - Column headers
 * @property {Array<Array<string|number|Date>>} [data] - Row data
 * @property {number} [rowCount] - Total number of rows
 * @property {string} [error] - Error message if failed
 */

/**
 * @typedef {Object} TransferRow
 * @property {string} flag - Row flag (A/D/C or empty)
 * @property {string} comment - Row comment
 * @property {Object<number, string>} data - Column index to value mapping
 */

/**
 * @typedef {Object} InsertRowsParams
 * @property {string} filePath - Path to target Excel file
 * @property {string} sheetName - Target sheet name
 * @property {TransferRow[]} rows - Rows to insert
 * @property {number} startColumn - Starting column index
 */

/**
 * @typedef {Object} ConfigData
 * @property {string} [file1Path] - Source file path
 * @property {string} [file2Path] - Target file path
 * @property {string} [templatePath] - Template file path
 * @property {string} [sheet1Name] - Source sheet name
 * @property {string} [sheet2Name] - Target sheet name
 * @property {number} [startColumn] - Start column for insertion
 * @property {number} [checkColumn] - Column for duplicate checking
 * @property {number[]} [sourceColumns] - Source columns to copy
 */

/**
 * @typedef {Object} ExportParams
 * @property {string} filePath - Path to save the file
 * @property {string[]} headers - Column headers
 * @property {Array<Array<string|number>>} data - Row data
 */

let mainWindow = null;

// ============================================
// SICHERHEITSFUNKTIONEN
// ============================================

/**
 * Prüft ob ein Dateipfad sicher ist (keine Path Traversal-Angriffe)
 * @param {string} filePath - Der zu prüfende Pfad
 * @returns {boolean} true wenn der Pfad sicher ist
 */
function isValidFilePath(filePath) {
    if (!filePath || typeof filePath !== 'string') {
        console.warn('Ungültiger Dateipfad (nicht String):', typeof filePath);
        return false;
    }
    
    // Normalisiere den Pfad
    const normalized = path.normalize(filePath);
    
    // Prüfe auf Path Traversal-Muster
    if (normalized.includes('..')) {
        console.warn('Path Traversal-Versuch erkannt:', filePath);
        return false;
    }
    
    // Prüfe auf null-bytes (kann Sicherheitsprüfungen umgehen)
    if (filePath.includes('\0')) {
        console.warn('Null-Byte im Pfad erkannt:', filePath);
        return false;
    }
    
    return true;
}

// ============================================
// FENSTER ERSTELLEN
// ============================================
function createWindow() {
    // Plattformspezifisches Icon
    const iconFile = process.platform === 'darwin' ? 'icon.icns' : 
                     process.platform === 'win32' ? 'icon.ico' : 'icon.png';
    
    mainWindow = new BrowserWindow({
        width: 1600,
        height: 1000,
        minWidth: 800,
        minHeight: 600,
        resizable: true,
        maximizable: true,
        fullscreenable: true,
        fullscreen: false,  // Explizit kein Vollbild
        simpleFullscreen: false,  // Kein einfacher Vollbildmodus
        title: 'Excel Data Sync Pro',
        icon: path.join(__dirname, 'assets', iconFile),
        frame: true,
        // Wichtig für korrekte Dialog-Darstellung
        useContentSize: false,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });

    mainWindow.loadFile('src/index.html');
    
    // Kontextmen� f�r Texteingabefelder aktivieren
    mainWindow.webContents.on('context-menu', (event, params) => {
        const { isEditable, selectionText, editFlags } = params;
        
        if (isEditable) {
            const menuTemplate = [
                {
                    label: 'Ausschneiden',
                    role: 'cut',
                    enabled: editFlags.canCut
                },
                {
                    label: 'Kopieren',
                    role: 'copy',
                    enabled: editFlags.canCopy
                },
                {
                    label: 'Einf\u00FCgen',
                    role: 'paste',
                    enabled: editFlags.canPaste
                },
                { type: 'separator' },
                {
                    label: 'Alles ausw\u00E4hlen',
                    role: 'selectAll',
                    enabled: editFlags.canSelectAll
                }
            ];
            
            const menu = Menu.buildFromTemplate(menuTemplate);
            menu.popup({ window: mainWindow });
        } else if (selectionText) {
            // Kontextmen� f�r markierten Text (nicht editierbar)
            const menuTemplate = [
                {
                    label: 'Kopieren',
                    role: 'copy',
                    enabled: editFlags.canCopy
                }
            ];
            
            const menu = Menu.buildFromTemplate(menuTemplate);
            menu.popup({ window: mainWindow });
        }
    });
    
    // DevTools oeffnen (nur waehrend Entwicklung)
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }
    
    // Menueleiste ausblenden (optional)
    mainWindow.setMenuBarVisibility(false);
    
    // Schließen-Anfrage abfangen für Warteschlangen-Prüfung
    let closeConfirmed = false;
    
    mainWindow.on('close', (e) => {
        if (!closeConfirmed) {
            e.preventDefault();
            // Renderer fragen, ob Warteschlange leer ist
            mainWindow.webContents.send('app:beforeClose');
        }
    });
    
    ipcMain.on('app:confirmClose', (event, canClose) => {
        if (canClose) {
            closeConfirmed = true;
            mainWindow.close();
        }
    });
    
    // Fenster-Referenz aufräumen wenn geschlossen
    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    // Nur wenn App bereit ist und kein Fenster offen ist
    if (app.isReady() && BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// ============================================
// DATEI-DIALOGE - Windows-Workaround f�r Dialog-Problem
// ============================================

// Datei oeffnen Dialog
ipcMain.handle('dialog:openFile', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verf�gbar f�r Dialog');
        return null;
    }
    
    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();
    
    try {
        // Standard-Pfad setzen (hilft bei Dialog-Gr��enproblemen unter Windows)
        const defaultPath = options.defaultPath || app.getPath('documents');
        
        const result = await dialog.showOpenDialog({
            title: options.title || 'Datei oeffnen',
            defaultPath: defaultPath,
            filters: options.filters || [
                { name: 'Excel-Dateien', extensions: ['xlsx', 'xls'] },
                { name: 'Alle Dateien', extensions: ['*'] }
            ],
            properties: ['openFile']
        });
        
        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }
        
        if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
            return null;
        }
        return result.filePaths[0];
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// Datei speichern Dialog
ipcMain.handle('dialog:saveFile', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verf�gbar f�r Dialog');
        return null;
    }
    
    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();
    
    try {
        // Standard-Pfad setzen falls nicht angegeben
        const defaultPath = options.defaultPath || app.getPath('documents');
        
        const result = await dialog.showSaveDialog({
            title: options.title || 'Datei speichern',
            defaultPath: defaultPath,
            filters: options.filters || [
                { name: 'Excel-Dateien', extensions: ['xlsx'] }
            ]
        });
        
        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }
        
        if (result.canceled || !result.filePath) {
            return null;
        }
        return result.filePath;
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// Ordner oeffnen Dialog (fuer Arbeitsordner)
ipcMain.handle('dialog:openFolder', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verfuegbar fuer Dialog');
        return null;
    }
    
    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();
    
    try {
        const defaultPath = options.defaultPath || app.getPath('documents');
        
        const result = await dialog.showOpenDialog({
            title: options.title || 'Ordner auswaehlen',
            defaultPath: defaultPath,
            properties: ['openDirectory']
        });
        
        // Nach Dialog: Hauptfenster wieder fokussieren
        if (mainWindow && !mainWindow.isDestroyed()) {
            mainWindow.focus();
        }
        
        if (result.canceled || !result.filePaths || result.filePaths.length === 0) {
            return null;
        }
        return result.filePaths[0];
    } catch (err) {
        console.error('Dialog Fehler:', err);
        return null;
    }
});

// ============================================
// DATEISYSTEM OPERATIONEN
// ============================================

// Prüfen ob Datei existiert
ipcMain.handle('fs:checkFileExists', async (event, filePath) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { exists: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        const fs = require('fs');
        const exists = fs.existsSync(filePath);
        return { exists };
    } catch (err) {
        console.error('Dateiprüfung Fehler:', err);
        return { exists: false, error: err.message };
    }
});

// ============================================
// EXCEL OPERATIONEN (xlsx-populate - erhaelt Formatierung!)
// ============================================

// Excel-Datei lesen
ipcMain.handle('excel:readFile', async (event, filePath) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const sheets = workbook.sheets().map(ws => ws.name());
        
        return {
            success: true,
            fileName: path.basename(filePath),
            filePath: filePath,
            sheets: sheets
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Sheet-Daten lesen
ipcMain.handle('excel:readSheet', async (event, filePath, sheetName) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet(sheetName);
        
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Benutzte Range ermitteln
        const usedRange = worksheet.usedRange();
        if (!usedRange) {
            return { success: true, headers: [], data: [] };
        }
        
        const startRow = usedRange.startCell().rowNumber();
        const endRow = usedRange.endCell().rowNumber();
        const startCol = usedRange.startCell().columnNumber();
        const endCol = usedRange.endCell().columnNumber();
        
        const data = [];
        const headers = [];
        const hiddenColumns = []; // Indices der versteckten Spalten
        
        // Hilfsfunktion: Excel-Datum zu lesbarem String konvertieren
        function excelDateToString(excelDate) {
            // Excel-Datum: Tage seit 1.1.1900 (mit falschem Schaltjahr 1900)
            // JavaScript: Millisekunden seit 1.1.1970
            const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 30.12.1899 UTC
            const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
            
            // Pr�fen ob es ein reines Datum oder Datum mit Uhrzeit ist
            const hasTime = (excelDate % 1) !== 0;
            
            if (hasTime) {
                // Datum mit Uhrzeit
                const day = String(jsDate.getUTCDate()).padStart(2, '0');
                const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                const year = jsDate.getUTCFullYear();
                const hours = String(jsDate.getUTCHours()).padStart(2, '0');
                const minutes = String(jsDate.getUTCMinutes()).padStart(2, '0');
                return `${day}.${month}.${year} ${hours}:${minutes}`;
            } else {
                // Nur Datum
                const day = String(jsDate.getUTCDate()).padStart(2, '0');
                const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                const year = jsDate.getUTCFullYear();
                return `${day}.${month}.${year}`;
            }
        }
        
        // Hilfsfunktion: Prüfen ob ein Zellwert ein Datum ist
        function isExcelDate(cell, value) {
            if (typeof value !== 'number') return false;
            
            // Prüfe das Zahlenformat der Zelle
            let numFmt;
            try {
                numFmt = cell.style('numberFormat');
            } catch (e) {
                numFmt = null;
            }
            
            if (numFmt && typeof numFmt === 'string') {
                // Standard-Excel-Datumsformate (Format-IDs als Strings)
                // Diese werden von xlsx-populate oft als Strings zur�ckgegeben
                const dateFormatIds = [
                    '14', '15', '16', '17', '18', '19', '20', '21', '22',
                    '45', '46', '47', '27', '30', '36', '50', '57'
                ];
                
                // Pr�fe auf numerische Format-ID
                if (dateFormatIds.includes(String(numFmt))) {
                    return true;
                }
                
                // Explizite Nicht-Datum-Formate
                const nonDatePatterns = [
                    /^General$/i,
                    /^[#0,]+(\.[#0]+)?$/,        // Zahlenformat wie #,##0.00
                    /^[#0,]+(\.[#0]+)?%$/,       // Prozent
                    /%/,                          // Prozentzeichen
                    /�|EUR|\$/,                   // W�hrung
                    /^@$/,                        // Text
                    /^\[.*?\][#0]/,               // Buchhaltungsformat
                ];
                
                for (const pattern of nonDatePatterns) {
                    if (pattern.test(numFmt)) {
                        return false;
                    }
                }
                
                // Typische Datumsformate erkennen (Strings)
                const datePatterns = [
                    /d+[\/\-.\s]m+[\/\-.\s]y+/i,     // d.m.y, d/m/y, d-m-y, d m y
                    /m+[\/\-.\s]d+[\/\-.\s]y+/i,     // m/d/y (US-Format)
                    /y+[\/\-.\s]m+[\/\-.\s]d+/i,     // y-m-d (ISO-Format)
                    /dd\.mm\.yyyy/i,                  // Deutsches Format
                    /dd\/mm\/yyyy/i,                  // Britisches Format
                    /mm\/dd\/yyyy/i,                  // US-Format
                    /yyyy-mm-dd/i,                    // ISO-Format
                    /\[.*?\]dd/i,                     // Benutzerdefinierte Formate
                    /mmm/i,                           // Monatsname (mmm, mmmm)
                    /^d+$/i,                          // Nur "d" oder "dd"
                    /^[$-].*d.*m.*y/i,                // Locale-spezifische Formate
                    /[$-F800]/,                       // Windows Locale Format
                    /[$-407]/,                        // Deutsches Locale
                ];
                
                for (const pattern of datePatterns) {
                    if (pattern.test(numFmt)) {
                        return true;
                    }
                }
            }
            
            // Heuristik f�r Werte ohne explizites Format oder mit "General"
            // Excel-Datum: 1 = 1.1.1900, 44197 = 1.1.2021, 47848 = 1.1.2031
            // Typischer Bereich f�r aktuelle Daten: 35000 (1995) bis 55000 (2050)
            if (value >= 1 && value <= 73050) {
                // Nur ganzzahlige Werte oder Werte mit Zeitanteil pr�fen
                // Sehr kleine Zahlen (< 365) sind wahrscheinlich keine Daten
                if (value < 365) {
                    return false; // Wahrscheinlich eine normale Zahl (Tage im Jahr etc.)
                }
                
                // Pr�fe ob es vern�nftig aussieht
                // Moderne Daten liegen zwischen 30000 (1982) und 55000 (2050)
                if (value >= 30000 && value <= 55000) {
                    // Wenn kein explizites Nicht-Datum-Format, k�nnte es ein Datum sein
                    if (!numFmt || numFmt === 'General' || numFmt === 'general') {
                        // Zus�tzliche Heuristik: Ganzzahlige Werte in diesem Bereich 
                        // sind sehr wahrscheinlich Daten
                        if (Number.isInteger(value)) {
                            return true;
                        }
                    }
                }
            }
            
            return false;
        }
        
        for (let row = startRow; row <= endRow; row++) {
            const rowData = [];
            for (let col = startCol; col <= endCol; col++) {
                const cell = worksheet.cell(row, col);
                const value = cell.value();
                
                let textValue = '';
                if (value !== undefined && value !== null) {
                    // Pr�fe ob es ein Datum ist
                    if (value instanceof Date) {
                        // Falls xlsx-populate bereits ein Date-Objekt zur�ckgibt
                        const day = String(value.getDate()).padStart(2, '0');
                        const month = String(value.getMonth() + 1).padStart(2, '0');
                        const year = value.getFullYear();
                        textValue = `${day}.${month}.${year}`;
                    } else if (isExcelDate(cell, value)) {
                        textValue = excelDateToString(value);
                    } else {
                        textValue = String(value);
                    }
                }
                
                // Header-Zeile (erste Zeile)
                if (row === startRow) {
                    headers[col - 1] = textValue || `Spalte ${col}`;
                    // Prüfe ob die Spalte in Excel versteckt ist
                    try {
                        const column = worksheet.column(col);
                        if (column && column.hidden()) {
                            hiddenColumns.push(col - 1); // 0-basierter Index
                        }
                    } catch (e) {
                        // Spalte existiert möglicherweise nicht explizit
                    }
                }
                rowData[col - 1] = textValue;
            }
            
            // Zeilen auffuellen bis zur maximalen Spaltenanzahl
            while (rowData.length < headers.length) {
                rowData.push('');
            }
            
            data.push(rowData);
        }
        
        return {
            success: true,
            headers: headers,
            data: data,
            hiddenColumns: hiddenColumns
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Zeilen in Excel einfuegen (MIT Formatierungserhalt dank xlsx-populate!)
ipcMain.handle('excel:insertRows', async (event, { filePath, sheetName, rows, startColumn, enableFlag = true, enableComment = true, flagColumn = 1, commentColumn = 2, sourceFilePath = null, sourceSheetName = null, sourceColumns = [] }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const worksheet = workbook.sheet(sheetName);
        
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Hilfsfunktion: Deutsches Datum zu Excel-Datum konvertieren
        function parseGermanDateToExcel(dateStr) {
            if (!dateStr || typeof dateStr !== 'string') return null;
            
            // Deutsches Datum: dd.mm.yyyy oder dd.mm.yyyy hh:mm
            const dateTimeMatch = dateStr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})$/);
            const dateMatch = dateStr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
            
            if (dateTimeMatch) {
                const day = parseInt(dateTimeMatch[1], 10);
                const month = parseInt(dateTimeMatch[2], 10);
                const year = parseInt(dateTimeMatch[3], 10);
                const hours = parseInt(dateTimeMatch[4], 10);
                const minutes = parseInt(dateTimeMatch[5], 10);
                
                // Validierung
                if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900 || year > 2100) {
                    return null;
                }
                
                // Excel-Datum berechnen (UTC um Zeitzonenproblemen vorzubeugen)
                const jsDate = Date.UTC(year, month - 1, day, hours, minutes);
                const excelEpoch = Date.UTC(1899, 11, 30);
                const excelDate = (jsDate - excelEpoch) / 86400000;
                return excelDate;
            } else if (dateMatch) {
                const day = parseInt(dateMatch[1], 10);
                const month = parseInt(dateMatch[2], 10);
                const year = parseInt(dateMatch[3], 10);
                
                // Validierung
                if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900 || year > 2100) {
                    return null;
                }
                
                // Excel-Datum berechnen (UTC)
                const jsDate = Date.UTC(year, month - 1, day);
                const excelEpoch = Date.UTC(1899, 11, 30);
                const excelDate = Math.floor((jsDate - excelEpoch) / 86400000);
                return excelDate;
            }
            
            return null;
        }
        
        // Hilfsfunktion: Wert intelligent konvertieren
        function convertValue(value, targetCell) {
            if (value === null || value === undefined || value === '') {
                return { value: value, isDate: false };
            }
            
            // Pr�fe ob es ein deutsches Datum ist
            const excelDate = parseGermanDateToExcel(value);
            if (excelDate !== null) {
                return { value: excelDate, isDate: true };
            }
            
            // Pr�fe ob es eine Zahl ist (mit deutschen Dezimaltrennzeichen)
            if (typeof value === 'string') {
                // Deutsche Zahlen: 1.234,56 -> 1234.56
                const germanNumberMatch = value.match(/^-?\d{1,3}(\.\d{3})*(,\d+)?$/);
                if (germanNumberMatch) {
                    const normalized = value.replace(/\./g, '').replace(',', '.');
                    const num = parseFloat(normalized);
                    if (!isNaN(num)) {
                        return { value: num, isDate: false };
                    }
                }
                
                // Englische Zahlen oder einfache Zahlen
                const simpleNumber = value.match(/^-?\d+(\.\d+)?$/);
                if (simpleNumber) {
                    const num = parseFloat(value);
                    if (!isNaN(num)) {
                        return { value: num, isDate: false };
                    }
                }
            }
            
            // Als String belassen
            return { value: value, isDate: false };
        }
        
        // Formatvorlage aus Header-Zeile oder vorhandenen Zeilen ermitteln
        function getColumnFormat(colNumber) {
            // Suche nach dem ersten nicht-leeren Wert in dieser Spalte (ab Zeile 2)
            const usedRange = worksheet.usedRange();
            if (!usedRange) return null;
            
            const endRow = Math.min(usedRange.endCell().rowNumber(), 100); // Max 100 Zeilen pr�fen
            
            for (let row = 2; row <= endRow; row++) {
                const cell = worksheet.cell(row, colNumber);
                const value = cell.value();
                if (value !== undefined && value !== null && value !== '') {
                    const numFmt = cell.style('numberFormat');
                    if (numFmt && numFmt !== 'General') {
                        return numFmt;
                    }
                }
            }
            return null;
        }
        
        // Erste leere Zeile finden (ab Zeile 2, da Zeile 1 = Header)
        let insertRow = 2;  // Standard: direkt nach Header
        
        const usedRange = worksheet.usedRange();
        if (usedRange) {
            const endRow = usedRange.endCell().rowNumber();
            
            // Pr�fe ab Zeile 2, ob es Daten gibt
            for (let row = 2; row <= endRow; row++) {
                // Prüfe Flag-Spalte (wenn aktiviert) oder Datenspalte
                const checkCol = enableFlag ? flagColumn : startColumn;
                const flagCell = worksheet.cell(row, checkCol).value();
                const dataCell = worksheet.cell(row, startColumn).value();
                
                // Zeile ist leer wenn beide Zellen leer sind
                const flagEmpty = flagCell === undefined || flagCell === null || flagCell === '';
                const dataEmpty = dataCell === undefined || dataCell === null || dataCell === '';
                
                if (flagEmpty && dataEmpty) {
                    // Erste leere Zeile gefunden
                    insertRow = row;
                    break;
                }
                // Wenn wir hier sind, ist die Zeile nicht leer - gehe zur n�chsten
                insertRow = row + 1;
            }
        }
        
        console.log(`Einf�gen ab Zeile: ${insertRow}`);
        
        // Spaltenformate vorab ermitteln
        const columnFormats = {};
        
        // Formatierungsvorlage: Letzte belegte Zeile vor der Einfügeposition
        const templateRow = insertRow > 2 ? insertRow - 1 : 2;
        
        // Hilfsfunktion: Kopiere Zellformatierung von Template-Zeile der Zieldatei
        function copyStyleFromTemplate(targetCell, colNumber) {
            try {
                const templateCell = worksheet.cell(templateRow, colNumber);
                if (!templateCell) return;
                
                // Verfügbare Styles in xlsx-populate
                const styles = [
                    'bold', 'italic', 'underline', 'strikethrough',
                    'fontFamily', 'fontSize', 'fontColor',
                    'horizontalAlignment', 'verticalAlignment',
                    'wrapText', 'numberFormat'
                ];
                
                styles.forEach(styleName => {
                    try {
                        const styleValue = templateCell.style(styleName);
                        if (styleValue !== undefined && styleValue !== null) {
                            targetCell.style(styleName, styleValue);
                        }
                    } catch (e) {
                        // Ignoriere Fehler bei einzelnen Styles
                    }
                });
                
                // Fill (Hintergrundfarbe) separat behandeln - ist ein komplexes Objekt
                try {
                    const fillValue = templateCell.style('fill');
                    if (fillValue && typeof fillValue === 'object') {
                        // Deep copy des Fill-Objekts um Referenzprobleme zu vermeiden
                        const fillCopy = JSON.parse(JSON.stringify(fillValue));
                        targetCell.style('fill', fillCopy);
                    }
                } catch (e) {
                    // Ignoriere Fehler beim Kopieren der Fill-Formatierung
                }
            } catch (e) {
                // Ignoriere Fehler
            }
        }
        
        // Neue Zeilen einfuegen
        let insertedCount = 0;
        for (const row of rows) {
            const newRowNum = insertRow + insertedCount;
            
            // Bei Leerzeile: Zeile als "belegt" markieren
            if (row.flag === 'leer') {
                // Flag-Spalte setzen wenn aktiviert
                if (enableFlag) {
                    const flagCell = worksheet.cell(newRowNum, flagColumn);
                    copyStyleFromTemplate(flagCell, flagColumn);
                    flagCell.value(' ');
                } else {
                    // Wenn Flag deaktiviert, Leerzeichen in erste Datenspalte setzen
                    // damit die Zeile als "belegt" gilt und nicht überschrieben wird
                    const firstDataCell = worksheet.cell(newRowNum, startColumn);
                    copyStyleFromTemplate(firstDataCell, startColumn);
                    firstDataCell.value(' ');
                }
                // Kommentar trotzdem schreiben wenn vorhanden und aktiviert
                if (enableComment && row.comment) {
                    const commentCell = worksheet.cell(newRowNum, commentColumn);
                    copyStyleFromTemplate(commentCell, commentColumn);
                    commentCell.value(row.comment);
                }
                insertedCount++;
                continue;
            }
            
            // Flag in konfigurierter Spalte (nur wenn aktiviert)
            if (enableFlag && row.flag) {
                const flagCell = worksheet.cell(newRowNum, flagColumn);
                copyStyleFromTemplate(flagCell, flagColumn);
                flagCell.value(row.flag);
            }
            
            // Kommentar in konfigurierter Spalte (nur wenn aktiviert)
            if (enableComment && row.comment) {
                const commentCell = worksheet.cell(newRowNum, commentColumn);
                copyStyleFromTemplate(commentCell, commentColumn);
                commentCell.value(row.comment);
            }
            
            // Daten ab Startspalte - row.data ist ein Objekt mit Index als Key
            if (row.data) {
                const dataKeys = Object.keys(row.data);
                
                dataKeys.forEach(key => {
                    const index = parseInt(key);
                    const value = row.data[key];
                    if (value !== null && value !== undefined && value !== '') {
                        const colNumber = startColumn + index;
                        const targetCell = worksheet.cell(newRowNum, colNumber);
                        
                        // Formatierung von Template-Zeile der Zieldatei kopieren
                        // (Quelldatei kann bedingte Formatierungen nicht liefern)
                        copyStyleFromTemplate(targetCell, colNumber);
                        
                        const converted = convertValue(value, targetCell);
                        targetCell.value(converted.value);
                        
                        // Wenn es ein Datum ist, prüfe ob schon ein Format von Template kopiert wurde
                        if (converted.isDate) {
                            const currentFormat = targetCell.style('numberFormat');
                            // Nur setzen wenn kein Format vorhanden (Template hatte auch keins)
                            if (!currentFormat || currentFormat === 'General') {
                                // Spaltenformat aus Cache oder neu ermitteln
                                if (!(colNumber in columnFormats)) {
                                    columnFormats[colNumber] = getColumnFormat(colNumber);
                                }
                                
                                const colFormat = columnFormats[colNumber];
                                if (colFormat) {
                                    targetCell.style('numberFormat', colFormat);
                                } else {
                                    // Standard deutsches Datumsformat setzen
                                    targetCell.style('numberFormat', 'DD.MM.YYYY');
                                }
                            }
                            // Wenn Format vom Template/Source kopiert wurde, behalte es!
                        }
                    }
                });
            }
            
            insertedCount++;
        }
        
        // Speichern (xlsx-populate erh�lt die originale Formatierung!)
        await workbook.toFileAsync(filePath);
        
        return { 
            success: true, 
            message: `${insertedCount} Zeile(n) ab Zeile ${insertRow} eingefuegt`,
            insertedCount: insertedCount,
            startRow: insertRow
        };
    } catch (error) {
        console.error('Fehler beim Einf�gen:', error);
        return { success: false, error: error.message };
    }
});

// Datei kopieren (fuer "Neuer Monat") - BINÄRE KOPIE erhält 100% Formatierung!
ipcMain.handle('excel:copyFile', async (event, { sourcePath, targetPath, sheetName, keepHeader }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        // Prüfe ob Quelldatei existiert
        if (!fs.existsSync(sourcePath)) {
            return { success: false, error: `Quelldatei nicht gefunden: ${sourcePath}` };
        }
        
        // Wenn keepHeader false ist oder nicht gesetzt, einfach binär kopieren
        // Das erhält 100% der Formatierung!
        if (!keepHeader) {
            try {
                fs.copyFileSync(sourcePath, targetPath);
            } catch (copyErr) {
                return { success: false, error: `Kopieren fehlgeschlagen: ${copyErr.message}` };
            }
            return { success: true, message: `Datei kopiert: ${targetPath}` };
        }
        
        // Wenn keepHeader true ist, müssen wir die Daten löschen
        // Aber zuerst: Binär kopieren, dann nur die Werte löschen
        try {
            fs.copyFileSync(sourcePath, targetPath);
        } catch (copyErr) {
            return { success: false, error: `Kopieren fehlgeschlagen: ${copyErr.message}` };
        }
        
        // Jetzt die kopierte Datei �ffnen und nur die Datenwerte l�schen
        const workbook = await XlsxPopulate.fromFileAsync(targetPath);
        
        // Wenn sheetName angegeben und existiert, nutze dieses Sheet
        // Ansonsten nimm das erste Sheet
        let worksheet = null;
        if (sheetName) {
            worksheet = workbook.sheet(sheetName);
        }
        if (!worksheet) {
            // Erstes verfuegbares Sheet nehmen
            worksheet = workbook.sheets()[0];
        }
        
        if (!worksheet) {
            return { success: false, error: 'Keine Worksheets in der Template-Datei gefunden' };
        }
        
        // Nur die Werte ab Zeile 2 l�schen (Header in Zeile 1 bleibt)
        // Formatierung bleibt erhalten!
        const usedRange = worksheet.usedRange();
        if (usedRange) {
            const endRow = usedRange.endCell().rowNumber();
            const endCol = usedRange.endCell().columnNumber();
            
            // Alle Datenwerte ab Zeile 2 l�schen
            for (let row = 2; row <= endRow; row++) {
                for (let col = 1; col <= endCol; col++) {
                    worksheet.cell(row, col).value(undefined);
                }
            }
        }
        
        await workbook.toFileAsync(targetPath);
        
        return { success: true, message: `Datei erstellt: ${targetPath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Daten exportieren (fuer Datenexplorer) - nur ein Sheet
ipcMain.handle('excel:exportData', async (event, { filePath, headers, data }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        // Neue leere Workbook erstellen
        const workbook = await XlsxPopulate.fromBlankAsync();
        
        // Erstes Sheet umbenennen
        const worksheet = workbook.sheet(0);
        worksheet.name('Export');
        
        // Header-Zeile
        headers.forEach((header, colIndex) => {
            worksheet.cell(1, colIndex + 1).value(header);
        });
        
        // Daten-Zeilen (data ist ein Array von Arrays)
        data.forEach((row, rowIndex) => {
            row.forEach((value, colIndex) => {
                worksheet.cell(rowIndex + 2, colIndex + 1).value(value || '');
            });
        });
        
        await workbook.toFileAsync(filePath);
        
        return { success: true, message: `Export erstellt: ${filePath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Daten exportieren MIT allen Sheets (fuer Datenexplorer - Vollexport)
ipcMain.handle('excel:exportWithAllSheets', async (event, { sourcePath, targetPath, sheetName, headers, data, visibleColumns }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        // Originaldatei laden (mit allen Sheets und Formatierung)
        const workbook = await XlsxPopulate.fromFileAsync(sourcePath);
        const allSheets = workbook.sheets().map(s => s.name());
        
        // Das aktive Sheet finden
        const worksheet = workbook.sheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Alle vorhandenen Daten im Sheet löschen
        const usedRange = worksheet.usedRange();
        if (usedRange) {
            usedRange.clear();
        }
        
        // Wenn nur bestimmte Spalten sichtbar sind, diese exportieren
        if (visibleColumns && visibleColumns.length > 0 && visibleColumns.length < headers.length) {
            // Header-Zeile mit sichtbaren Spalten
            visibleColumns.forEach((colIdx, newColIdx) => {
                worksheet.cell(1, newColIdx + 1).value(headers[colIdx] || '');
            });
            
            // Daten-Zeilen mit sichtbaren Spalten
            data.forEach((row, rowIndex) => {
                visibleColumns.forEach((colIdx, newColIdx) => {
                    worksheet.cell(rowIndex + 2, newColIdx + 1).value(row[colIdx] || '');
                });
            });
        } else {
            // Alle Spalten exportieren
            // Header-Zeile
            headers.forEach((header, colIndex) => {
                worksheet.cell(1, colIndex + 1).value(header);
            });
            
            // Daten-Zeilen
            data.forEach((row, rowIndex) => {
                row.forEach((value, colIndex) => {
                    worksheet.cell(rowIndex + 2, colIndex + 1).value(value || '');
                });
            });
        }
        
        // Speichern (alle anderen Sheets bleiben unverändert)
        await workbook.toFileAsync(targetPath);
        
        return { success: true, message: `Export erstellt: ${targetPath}`, sheets: allSheets };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Export mit Auswahl der Arbeitsblätter (für Datenexplorer) - behält Formatierung bei
ipcMain.handle('excel:exportMultipleSheets', async (event, { sourcePath, targetPath, sheets }) => {
    // Sicherheitsprüfung: Pfade validieren
    if (!isValidFilePath(sourcePath) || !isValidFilePath(targetPath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        // Originaldatei laden (mit allen Sheets und Formatierung)
        const workbook = await XlsxPopulate.fromFileAsync(sourcePath);
        
        // Liste der ausgewählten Sheet-Namen
        const selectedSheetNames = sheets.map(s => s.sheetName);
        
        // Alle Sheets der Originaldatei durchgehen
        const allSheetNames = workbook.sheets().map(s => s.name());
        
        // Sheets entfernen, die nicht ausgewählt wurden (von hinten nach vorne, um Indexprobleme zu vermeiden)
        for (let i = allSheetNames.length - 1; i >= 0; i--) {
            const sheetName = allSheetNames[i];
            if (!selectedSheetNames.includes(sheetName)) {
                // Sheet nicht ausgewählt - entfernen
                const sheetToDelete = workbook.sheet(sheetName);
                if (sheetToDelete) {
                    workbook.deleteSheet(sheetToDelete);
                }
            }
        }
        
        let sheetsProcessed = 0;
        
        // Nur Sheets mit Änderungen aktualisieren
        for (const sheetData of sheets) {
            const worksheet = workbook.sheet(sheetData.sheetName);
            if (!worksheet) {
                console.warn(`Sheet "${sheetData.sheetName}" nicht gefunden - übersprungen`);
                continue;
            }
            
            // Wenn Sheet aus Datei kommt (keine Änderungen), nichts tun - Formatierung bleibt erhalten
            if (sheetData.fromFile) {
                sheetsProcessed++;
                continue;
            }
            
            // Sheet mit bearbeiteten Daten - nur Werte aktualisieren, Formatierung bleibt
            const headers = sheetData.headers;
            const data = sheetData.data;
            const visibleColumns = sheetData.visibleColumns;
            
            // Alle vorhandenen Daten im Sheet löschen
            const usedRange = worksheet.usedRange();
            if (usedRange) {
                usedRange.clear();
            }
            
            // Wenn nur bestimmte Spalten sichtbar sind, diese exportieren
            if (visibleColumns && visibleColumns.length > 0 && visibleColumns.length < headers.length) {
                // Header-Zeile mit sichtbaren Spalten
                visibleColumns.forEach((colIdx, newColIdx) => {
                    worksheet.cell(1, newColIdx + 1).value(headers[colIdx] || '');
                });
                
                // Daten-Zeilen mit sichtbaren Spalten
                data.forEach((row, rowIndex) => {
                    visibleColumns.forEach((colIdx, newColIdx) => {
                        worksheet.cell(rowIndex + 2, newColIdx + 1).value(row[colIdx] === null || row[colIdx] === undefined ? '' : row[colIdx]);
                    });
                });
            } else {
                // Alle Spalten exportieren
                // Header-Zeile schreiben (Zeile 1)
                headers.forEach((header, colIndex) => {
                    worksheet.cell(1, colIndex + 1).value(header);
                });
                
                // Daten-Zeilen schreiben (ab Zeile 2) - Formatierung bleibt erhalten
                data.forEach((row, rowIndex) => {
                    row.forEach((value, colIndex) => {
                        const cell = worksheet.cell(rowIndex + 2, colIndex + 1);
                        // Nur Wert setzen, Formatierung beibehalten
                        cell.value(value === null || value === undefined ? '' : value);
                    });
                });
            }
            
            sheetsProcessed++;
        }
        
        // Als neue Datei speichern (nicht Originaldatei überschreiben)
        await workbook.toFileAsync(targetPath);
        
        return { 
            success: true, 
            message: `${sheetsProcessed} Sheet(s) exportiert: ${targetPath}`,
            sheetsExported: sheetsProcessed
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Änderungen direkt in die Originaldatei speichern (für Datenexplorer)
ipcMain.handle('excel:saveFile', async (event, { filePath, sheets }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        // Originaldatei laden (mit allen Sheets und Formatierung)
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        
        let totalChanges = 0;
        
        // Jedes Sheet mit Änderungen aktualisieren
        for (const sheetData of sheets) {
            const worksheet = workbook.sheet(sheetData.sheetName);
            if (!worksheet) {
                console.warn(`Sheet "${sheetData.sheetName}" nicht gefunden - übersprungen`);
                continue;
            }
            
            const headers = sheetData.headers;
            const data = sheetData.data;
            const visibleColumns = sheetData.visibleColumns;
            
            // Alle vorhandenen Daten im Sheet löschen
            const usedRange = worksheet.usedRange();
            if (usedRange) {
                usedRange.clear();
            }
            
            // ALLE Spalten speichern (nicht nur sichtbare)
            // Header-Zeile schreiben (Zeile 1)
            headers.forEach((header, colIndex) => {
                worksheet.cell(1, colIndex + 1).value(header);
            });
            
            // Daten-Zeilen schreiben (ab Zeile 2)
            data.forEach((row, rowIndex) => {
                row.forEach((value, colIndex) => {
                    const cell = worksheet.cell(rowIndex + 2, colIndex + 1);
                    // Nur Wert setzen, Formatierung beibehalten
                    cell.value(value === null || value === undefined ? '' : value);
                });
                totalChanges++;
            });
            
            // Ausgeblendete Spalten in Excel als hidden markieren
            // visibleColumns enthält die Indices der sichtbaren Spalten
            if (visibleColumns && visibleColumns.length > 0 && visibleColumns.length < headers.length) {
                // Erstelle Set der sichtbaren Spalten für schnellen Lookup
                const visibleSet = new Set(visibleColumns);
                
                // Alle Spalten durchgehen und hidden-Status setzen
                headers.forEach((_, colIndex) => {
                    try {
                        const column = worksheet.column(colIndex + 1); // 1-basiert
                        if (visibleSet.has(colIndex)) {
                            // Spalte ist sichtbar -> unhide
                            column.hidden(false);
                        } else {
                            // Spalte ist ausgeblendet -> hide
                            column.hidden(true);
                        }
                    } catch (e) {
                        console.warn(`Konnte hidden-Status für Spalte ${colIndex + 1} nicht setzen:`, e.message);
                    }
                });
            } else {
                // Alle Spalten sichtbar -> alle unhide
                headers.forEach((_, colIndex) => {
                    try {
                        const column = worksheet.column(colIndex + 1);
                        column.hidden(false);
                    } catch (e) {
                        // Ignorieren
                    }
                });
            }
        }
        
        // Speichern (überschreibt die Originaldatei)
        await workbook.toFileAsync(filePath);
        
        return { 
            success: true, 
            message: `${sheets.length} Sheet(s) in ${filePath} gespeichert`,
            totalChanges 
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// ============================================
// KONFIGURATION
// ============================================

// Config speichern
ipcMain.handle('config:save', async (event, { filePath, config }) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        fs.writeFileSync(filePath, JSON.stringify(config, null, 2), 'utf8');
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Config laden
ipcMain.handle('config:load', async (event, filePath) => {
    // Sicherheitsprüfung: Pfad validieren
    if (!isValidFilePath(filePath)) {
        return { success: false, error: 'Ungültiger Dateipfad' };
    }
    
    try {
        if (!fs.existsSync(filePath)) {
            return { success: false, error: 'Datei nicht gefunden' };
        }
        const content = fs.readFileSync(filePath, 'utf8');
        let config;
        try {
            config = JSON.parse(content);
        } catch (parseError) {
            return { success: false, error: 'Ungültige JSON-Syntax' };
        }
        return { success: true, config: config };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// App-Pfad ermitteln (fuer Config im Programmordner)
ipcMain.handle('app:getPath', async (event) => {
    const exePath = app.getPath('exe');
    const exeDir = path.dirname(exePath);
    
    return {
        appPath: app.getAppPath(),
        userData: app.getPath('userData'),
        exe: exePath,
        exeDir: exeDir
    };
});

// Loading-State für config:loadFromAppDir um Race Conditions zu verhindern
let configLoadingState = {
    isLoading: false,
    pendingPromise: null
};

// Automatisch config.json im Programmordner oder Benutzerordnern suchen
ipcMain.handle('config:loadFromAppDir', async (event, workingDir) => {
    // Race Condition verhindern: Wenn bereits geladen wird, auf das Ergebnis warten
    if (configLoadingState.isLoading && configLoadingState.pendingPromise) {
        console.log('config:loadFromAppDir - Warte auf laufenden Ladevorgang...');
        return configLoadingState.pendingPromise;
    }
    
    // Ladevorgang starten
    configLoadingState.isLoading = true;
    
    const loadConfigAsync = async () => {
        console.log('=== config:loadFromAppDir aufgerufen ===');
        console.log('Arbeitsordner:', workingDir || '(nicht gesetzt)');
    
        try {
            const exePath = app.getPath('exe');
            const exeDir = path.dirname(exePath);
            const documentsDir = app.getPath('documents');
            const downloadsDir = app.getPath('downloads');
            
            // PORTABLE EXE: Der Ordner wo die portable EXE gestartet wurde
            const portableDir = process.env.PORTABLE_EXECUTABLE_DIR || '';
            
            console.log('EXE Pfad:', exePath);
            console.log('EXE Ordner:', exeDir);
            console.log('Dokumente Ordner:', documentsDir);
            console.log('Downloads Ordner:', downloadsDir);
            console.log('Portable Ordner:', portableDir || '(nicht gesetzt)');
            console.log('App gepackt:', app.isPackaged);
            
            // Suchpfade in Prioritätsreihenfolge
            const possiblePaths = [];
            
            // 1. ARBEITSORDNER (höchste Priorität - vom Benutzer festgelegt)
            if (workingDir && typeof workingDir === 'string') {
                possiblePaths.push(path.join(workingDir, 'config.json'));
            }
            
            // 2. Portable EXE: Neben der EXE (höchste Priorität für portable Version)
            if (portableDir) {
                possiblePaths.push(path.join(portableDir, 'config.json'));
            }
            
            // 3. Installationsordner (neben der EXE)
            possiblePaths.push(path.join(exeDir, 'config.json'));
            
            // 4. Dokumente-Ordner des Benutzers
            possiblePaths.push(path.join(documentsDir, 'config.json'));
            possiblePaths.push(path.join(documentsDir, 'Excel-Data-Sync-Pro', 'config.json'));
            
            // 5. Downloads-Ordner des Benutzers
            possiblePaths.push(path.join(downloadsDir, 'config.json'));
            
            // 6. Im Entwicklungsmodus: Projektordner
            if (process.argv.includes('--dev') || !app.isPackaged) {
                possiblePaths.push(path.join(__dirname, 'config.json'));
                possiblePaths.push(path.join(process.cwd(), 'config.json'));
            }
            
            console.log('Suche in folgenden Pfaden:');
            possiblePaths.forEach((p, i) => {
                const exists = fs.existsSync(p);
                console.log(`  ${i + 1}. ${p} - ${exists ? 'GEFUNDEN' : 'nicht vorhanden'}`);
            });
            
            // Schnelle Suche - bei erstem Treffer abbrechen
            for (const configPath of possiblePaths) {
                if (fs.existsSync(configPath)) {
                    console.log('>>> config.json gefunden:', configPath);
                    const content = fs.readFileSync(configPath, 'utf8');
                    let config;
                    try {
                        config = JSON.parse(content);
                    } catch (parseError) {
                        console.error('Ungültige JSON-Syntax in:', configPath, parseError);
                        continue; // Nächsten Pfad probieren
                    }
                    console.log('>>> Konfig geladen, Mapping:', config.mapping?.sourceColumns?.length || 0, 'Spalten');
                    return { 
                        success: true, 
                        config: config,
                        path: configPath
                    };
                }
            }
            
            // Keine config.json gefunden - kein Fehler, nur Info
            console.log('>>> Keine config.json in den Suchpfaden gefunden');
            return { 
                success: false, 
                error: 'Keine config.json gefunden',
                searchedPaths: possiblePaths
            };
        } catch (error) {
            console.error('Fehler beim Laden der config.json:', error);
            return { success: false, error: error.message };
        } finally {
            // Loading-State zurücksetzen
            configLoadingState.isLoading = false;
            configLoadingState.pendingPromise = null;
        }
    };
    
    // Promise speichern für parallele Aufrufe
    configLoadingState.pendingPromise = loadConfigAsync();
    return configLoadingState.pendingPromise;
});

// ============================================
// TEMPLATE AUS QUELLDATEI ERSTELLEN
// ============================================

/**
 * Hilfsfunktion: Konvertiert Spaltennummer zu Spaltenbuchstabe (1=A, 2=B, 27=AA, etc.)
 */
function numberToColumnLetter(num) {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

/**
 * Hilfsfunktion: Konvertiert Spaltenbuchstabe zu Nummer (A=1, B=2, AA=27, etc.)
 */
function columnLetterToNumber(col) {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
        result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
}

/**
 * Hilfsfunktion: Verschiebt alle Spaltenreferenzen in einer Zellreferenz um n Spalten
 * z.B. "A1" + 2 -> "C1", "H1:H100" + 2 -> "J1:J100"
 */
function shiftColumnReference(ref, shiftBy) {
    return ref.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
        const colNum = columnLetterToNumber(col);
        const newCol = numberToColumnLetter(colNum + shiftBy);
        return newCol + row;
    });
}

/**
 * Hilfsfunktion: Dekodiert XML-Entities (z.B. &amp; -> &)
 */
function decodeXmlEntities(str) {
    return str
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
}

/**
 * Hilfsfunktion: Enkodiert XML-Entities (z.B. & -> &amp;)
 */
function encodeXmlEntities(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Erstellt ein leeres Template aus einer Quelldatei.
 * - Behält nur ausgewählte Sheets
 * - Fügt optional Flag- und Kommentar-Spalten am Anfang ein
 * - Behält die erste Zeile (Header) pro Sheet
 * - Löscht alle Datenzeilen
 * - Erweitert Conditional Formatting Ranges auf ganze Spalten
 * - Erhält alle Formatierungen und Stile
 */
ipcMain.handle('excel:createTemplateFromSource', async (event, { sourcePath, outputPath, selectedSheets, addFlagColumn, addCommentColumn }) => {
    // Sicherheitsprüfungen
    if (!isValidFilePath(sourcePath)) {
        return { success: false, error: 'Ungültiger Quellpfad' };
    }
    if (!isValidFilePath(outputPath)) {
        return { success: false, error: 'Ungültiger Ausgabepfad' };
    }
    
    // Berechne wie viele Spalten eingefügt werden
    const extraColumnsCount = (addFlagColumn ? 1 : 0) + (addCommentColumn ? 1 : 0);
    console.log(`Extra-Spalten: Flag=${addFlagColumn}, Kommentar=${addCommentColumn}, Gesamt=${extraColumnsCount}`);
    
    try {
        const JSZip = require('jszip');
        
        // 1. Quelldatei als ZIP lesen
        const sourceBuffer = fs.readFileSync(sourcePath);
        const zip = await JSZip.loadAsync(sourceBuffer);
        
        // 2. workbook.xml lesen um Sheet-Zuordnungen zu bekommen
        const workbookXml = await zip.file('xl/workbook.xml').async('string');
        
        // Sheet-Namen und rId extrahieren (Namen sind XML-encoded in der Datei)
        const sheetMatches = [...workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>/g)];
        const sheetRels = {};        // XML-encoded Namen -> rId
        const sheetRelsDecoded = {}; // Dekodierte Namen -> rId
        const encodedToDecoded = {}; // Mapping encoded -> decoded
        sheetMatches.forEach(match => {
            const encodedName = match[1];
            const decodedName = decodeXmlEntities(encodedName);
            sheetRels[encodedName] = match[2];       // encoded name -> rId
            sheetRelsDecoded[decodedName] = match[2]; // decoded name -> rId
            encodedToDecoded[encodedName] = decodedName;
        });
        
        // 3. workbook.xml.rels lesen um rId -> sheetX.xml Zuordnung zu bekommen
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const rIdToFile = {};
        const relMatches = [...relsXml.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)];
        relMatches.forEach(match => {
            rIdToFile[match[1]] = match[2].replace(/^\//, ''); // rId -> worksheets/sheetX.xml
        });
        
        // 4. Mapping: SheetName (dekodiert) -> Dateiname erstellen
        const sheetToFile = {};
        const sheetToFileEncoded = {}; // Für das Entfernen aus workbook.xml
        for (const [encodedName, rId] of Object.entries(sheetRels)) {
            const target = rIdToFile[rId];
            const decodedName = encodedToDecoded[encodedName];
            if (target) {
                const filePath = 'xl/' + target.replace(/^xl\//, '');
                sheetToFile[decodedName] = filePath;  // Dekodierter Name -> Datei
                sheetToFileEncoded[encodedName] = filePath; // Encoded Name -> Datei (für XML-Operationen)
            }
        }
        
        console.log('Sheet-Zuordnung (dekodiert):', sheetToFile);
        console.log('Ausgewählte Sheets:', selectedSheets);
        
        // 5. Sheets identifizieren, die NICHT ausgewählt wurden (vergleiche mit dekodierten Namen)
        const allDecodedSheetNames = Object.keys(sheetToFile);
        const sheetsToRemove = allDecodedSheetNames.filter(name => !selectedSheets.includes(name));
        console.log('Zu entfernende Sheets:', sheetsToRemove);
        
        // 6. Nicht ausgewählte Sheets aus workbook.xml entfernen (verwende encoded Namen für XML)
        let modifiedWorkbookXml = workbookXml;
        for (const decodedName of sheetsToRemove) {
            // XML-encoded Name für Regex verwenden
            const encodedName = encodeXmlEntities(decodedName);
            const sheetRegex = new RegExp(`<sheet[^>]*name="${encodedName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`, 'g');
            modifiedWorkbookXml = modifiedWorkbookXml.replace(sheetRegex, '');
        }
        zip.file('xl/workbook.xml', modifiedWorkbookXml);
        
        // 7. Sheet-Dateien der nicht ausgewählten Sheets entfernen
        for (const decodedName of sheetsToRemove) {
            const sheetFile = sheetToFile[decodedName];
            if (sheetFile && zip.files[sheetFile]) {
                zip.remove(sheetFile);
                console.log(`Sheet entfernt: ${sheetFile}`);
            }
        }
        
        // 8. Ausgewählte Worksheets verarbeiten
        let totalCfRules = 0;
        let processedSheets = 0;
        
        for (const sheetName of selectedSheets) {
            const sheetFile = sheetToFile[sheetName];
            if (!sheetFile || !zip.files[sheetFile]) {
                console.warn(`Sheet-Datei nicht gefunden: ${sheetName} -> ${sheetFile}`);
                continue;
            }
            
            let sheetXml = await zip.file(sheetFile).async('string');
            
            // Anzahl der CF-Regeln zählen
            const cfMatches = sheetXml.match(/<conditionalFormatting[^>]*>/g) || [];
            totalCfRules += cfMatches.length;
            
            // Wenn Extra-Spalten hinzugefügt werden sollen
            if (extraColumnsCount > 0) {
                // A) Alle Zellreferenzen in <c r="..."> verschieben
                sheetXml = sheetXml.replace(
                    /<c\s+r="([A-Z]+)(\d+)"/g,
                    (match, col, row) => {
                        const newCol = numberToColumnLetter(columnLetterToNumber(col) + extraColumnsCount);
                        return `<c r="${newCol}${row}"`;
                    }
                );
                
                // B) Merge-Cells verschieben (falls vorhanden)
                sheetXml = sheetXml.replace(
                    /<mergeCell\s+ref="([A-Z]+\d+):([A-Z]+\d+)"/g,
                    (match, start, end) => {
                        const newStart = shiftColumnReference(start, extraColumnsCount);
                        const newEnd = shiftColumnReference(end, extraColumnsCount);
                        return `<mergeCell ref="${newStart}:${newEnd}"`;
                    }
                );
                
                // C) Dimension verschieben (falls vorhanden)
                sheetXml = sheetXml.replace(
                    /<dimension\s+ref="([A-Z]+\d+):([A-Z]+\d+)"/g,
                    (match, start, end) => {
                        // Start bei A1 lassen wenn Extra-Spalten, Ende verschieben
                        const newEnd = shiftColumnReference(end, extraColumnsCount);
                        return `<dimension ref="A1:${newEnd}"`;
                    }
                );
                
                // D) Neue Header-Zellen am Anfang von Zeile 1 einfügen
                const newCells = [];
                if (addFlagColumn) {
                    newCells.push('<c r="A1" t="inlineStr"><is><t>Flag</t></is></c>');
                }
                if (addCommentColumn) {
                    const col = addFlagColumn ? 'B' : 'A';
                    newCells.push(`<c r="${col}1" t="inlineStr"><is><t>Kommentar</t></is></c>`);
                }
                
                // Füge neue Zellen in Zeile 1 ein
                if (newCells.length > 0) {
                    sheetXml = sheetXml.replace(
                        /(<row[^>]*r="1"[^>]*>)/,
                        `$1${newCells.join('')}`
                    );
                }
            }
            
            // E) Datenzeilen löschen (behalte nur Zeile 1 = Header)
            sheetXml = sheetXml.replace(
                /(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/,
                (match, open, content, close) => {
                    const headerRowMatch = content.match(/<row[^>]*r="1"[^>]*>[\s\S]*?<\/row>/);
                    const headerRow = headerRowMatch ? headerRowMatch[0] : '';
                    return open + headerRow + close;
                }
            );
            
            // F) Conditional Formatting Ranges auf ganze Spalten erweitern UND verschieben
            sheetXml = sheetXml.replace(
                /sqref="([^"]+)"/g,
                (match, sqref) => {
                    const ranges = sqref.split(/\s+/);
                    const columns = new Set();
                    
                    for (const range of ranges) {
                        const colMatch = range.match(/^([A-Z]+)/);
                        if (colMatch) {
                            // Spalte um extraColumnsCount verschieben
                            const originalCol = colMatch[1];
                            const shiftedCol = extraColumnsCount > 0 
                                ? numberToColumnLetter(columnLetterToNumber(originalCol) + extraColumnsCount)
                                : originalCol;
                            columns.add(shiftedCol);
                        }
                    }
                    
                    if (columns.size > 0) {
                        const newSqref = Array.from(columns).map(col => `${col}:${col}`).join(' ');
                        return `sqref="${newSqref}"`;
                    }
                    return match;
                }
            );
            
            zip.file(sheetFile, sheetXml);
            processedSheets++;
        }
        
        // 9. Template speichern
        const outputBuffer = await zip.generateAsync({ 
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        
        fs.writeFileSync(outputPath, outputBuffer);
        
        console.log(`Template erstellt: ${outputPath}`);
        console.log(`Sheets verarbeitet: ${processedSheets}`);
        console.log(`CF-Regeln erhalten: ${totalCfRules}`);
        console.log(`Extra-Spalten hinzugefügt: ${extraColumnsCount}`);
        
        return { 
            success: true, 
            message: 'Template erfolgreich erstellt',
            fileName: path.basename(outputPath),
            stats: {
                sheetsProcessed: processedSheets,
                cfRulesPreserved: totalCfRules,
                extraColumnsAdded: extraColumnsCount
            }
        };
        
    } catch (error) {
        console.error('Fehler beim Erstellen des Templates:', error);
        return { success: false, error: error.message };
    }
});
