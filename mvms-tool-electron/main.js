const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');

let mainWindow = null;

// ============================================
// FENSTER ERSTELLEN
// ============================================
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 800,
        minHeight: 600,
        resizable: true,
        maximizable: true,
        fullscreenable: true,
        fullscreen: false,  // Explizit kein Vollbild
        simpleFullscreen: false,  // Kein einfacher Vollbildmodus
        title: 'MVMS-Tool',
        icon: path.join(__dirname, 'assets', 'icon.ico'),
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
    
    // DevTools oeffnen (nur waehrend Entwicklung)
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }
    
    // Menueleiste ausblenden (optional)
    mainWindow.setMenuBarVisibility(false);
    
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
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// ============================================
// DATEI-DIALOGE - Windows-Workaround für Dialog-Problem
// ============================================

// Datei oeffnen Dialog
ipcMain.handle('dialog:openFile', async (event, options) => {
    if (!mainWindow || mainWindow.isDestroyed()) {
        console.error('mainWindow nicht verfügbar für Dialog');
        return null;
    }
    
    // Fenster vorbereiten
    if (mainWindow.isMinimized()) {
        mainWindow.restore();
    }
    mainWindow.focus();
    
    try {
        // Standard-Pfad setzen (hilft bei Dialog-Größenproblemen unter Windows)
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
        console.error('mainWindow nicht verfügbar für Dialog');
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

// ============================================
// EXCEL OPERATIONEN (xlsx-populate - erhält Formatierung!)
// ============================================

// Excel-Datei lesen
ipcMain.handle('excel:readFile', async (event, filePath) => {
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
        
        // Hilfsfunktion: Excel-Datum zu lesbarem String konvertieren
        function excelDateToString(excelDate) {
            // Excel-Datum: Tage seit 1.1.1900 (mit falschem Schaltjahr 1900)
            // JavaScript: Millisekunden seit 1.1.1970
            const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // 30.12.1899 UTC
            const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
            
            // Prüfen ob es ein reines Datum oder Datum mit Uhrzeit ist
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
            const numFmt = cell.style('numberFormat');
            
            // Debug-Logging (kann später entfernt werden)
            // console.log(`Zelle: Wert=${value}, Format="${numFmt}"`);
            
            if (numFmt) {
                // Standard-Excel-Datumsformate (Format-IDs als Strings)
                // Diese werden von xlsx-populate oft als Strings zurückgegeben
                const dateFormatIds = [
                    '14', '15', '16', '17', '18', '19', '20', '21', '22',
                    '45', '46', '47', '27', '30', '36', '50', '57'
                ];
                
                // Prüfe auf numerische Format-ID
                if (dateFormatIds.includes(String(numFmt))) {
                    return true;
                }
                
                // Explizite Nicht-Datum-Formate
                const nonDatePatterns = [
                    /^General$/i,
                    /^[#0,]+(\.[#0]+)?$/,        // Zahlenformat wie #,##0.00
                    /^[#0,]+(\.[#0]+)?%$/,       // Prozent
                    /%/,                          // Prozentzeichen
                    /€|EUR|\$/,                   // Währung
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
            
            // Heuristik für Werte ohne explizites Format oder mit "General"
            // Excel-Datum: 1 = 1.1.1900, 44197 = 1.1.2021, 47848 = 1.1.2031
            // Typischer Bereich für aktuelle Daten: 35000 (1995) bis 55000 (2050)
            if (value >= 1 && value <= 73050) {
                // Nur ganzzahlige Werte oder Werte mit Zeitanteil prüfen
                // Sehr kleine Zahlen (< 365) sind wahrscheinlich keine Daten
                if (value < 365) {
                    return false; // Wahrscheinlich eine normale Zahl (Tage im Jahr etc.)
                }
                
                // Prüfe ob es vernünftig aussieht
                // Moderne Daten liegen zwischen 30000 (1982) und 55000 (2050)
                if (value >= 30000 && value <= 55000) {
                    // Wenn kein explizites Nicht-Datum-Format, könnte es ein Datum sein
                    if (!numFmt || numFmt === 'General' || numFmt === 'general') {
                        // Zusätzliche Heuristik: Ganzzahlige Werte in diesem Bereich 
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
                    // Prüfe ob es ein Datum ist
                    if (value instanceof Date) {
                        // Falls xlsx-populate bereits ein Date-Objekt zurückgibt
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
            data: data
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Zeilen in Excel einfuegen (MIT Formatierungserhalt dank xlsx-populate!)
ipcMain.handle('excel:insertRows', async (event, { filePath, sheetName, rows, startColumn }) => {
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
            
            // Prüfe ob es ein deutsches Datum ist
            const excelDate = parseGermanDateToExcel(value);
            if (excelDate !== null) {
                return { value: excelDate, isDate: true };
            }
            
            // Prüfe ob es eine Zahl ist (mit deutschen Dezimaltrennzeichen)
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
            
            const endRow = Math.min(usedRange.endCell().rowNumber(), 100); // Max 100 Zeilen prüfen
            
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
            
            // Prüfe ab Zeile 2, ob es Daten gibt
            for (let row = 2; row <= endRow; row++) {
                const flagCell = worksheet.cell(row, 1).value();
                const dataCell = worksheet.cell(row, startColumn).value();
                
                // Zeile ist leer wenn beide Zellen leer sind
                const flagEmpty = flagCell === undefined || flagCell === null || flagCell === '';
                const dataEmpty = dataCell === undefined || dataCell === null || dataCell === '';
                
                if (flagEmpty && dataEmpty) {
                    // Erste leere Zeile gefunden
                    insertRow = row;
                    break;
                }
                // Wenn wir hier sind, ist die Zeile nicht leer - gehe zur nächsten
                insertRow = row + 1;
            }
        }
        
        console.log(`Einfügen ab Zeile: ${insertRow}`);
        
        // Spaltenformate vorab ermitteln
        const columnFormats = {};
        
        // Neue Zeilen einfuegen
        let insertedCount = 0;
        for (const row of rows) {
            const newRowNum = insertRow + insertedCount;
            
            // Flag in Spalte A
            if (row.flag && row.flag !== 'leer') {
                worksheet.cell(newRowNum, 1).value(row.flag);
            }
            
            // Kommentar in Spalte B
            if (row.comment) {
                worksheet.cell(newRowNum, 2).value(row.comment);
            }
            
            // Daten ab Startspalte - row.data ist ein Objekt mit Index als Key
            if (row.data && row.flag !== 'leer') {
                const dataKeys = Object.keys(row.data);
                dataKeys.forEach(key => {
                    const index = parseInt(key);
                    const value = row.data[key];
                    if (value !== null && value !== undefined && value !== '') {
                        const colNumber = startColumn + index;
                        const targetCell = worksheet.cell(newRowNum, colNumber);
                        const converted = convertValue(value, targetCell);
                        
                        targetCell.value(converted.value);
                        
                        // Wenn es ein Datum ist und die Zelle kein Format hat, 
                        // versuche das Format aus der Spalte zu übernehmen
                        if (converted.isDate) {
                            const currentFormat = targetCell.style('numberFormat');
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
                        }
                    }
                });
            }
            
            insertedCount++;
        }
        
        // Speichern (xlsx-populate erhält die originale Formatierung!)
        await workbook.toFileAsync(filePath);
        
        return { 
            success: true, 
            message: `${insertedCount} Zeile(n) ab Zeile ${insertRow} eingefuegt`,
            insertedCount: insertedCount,
            startRow: insertRow
        };
    } catch (error) {
        console.error('Fehler beim Einfügen:', error);
        return { success: false, error: error.message };
    }
});

// Datei kopieren (fuer "Neuer Monat") - BINÄRE KOPIE erhält 100% Formatierung!
ipcMain.handle('excel:copyFile', async (event, { sourcePath, targetPath, sheetName, keepHeader }) => {
    try {
        // Wenn keepHeader false ist oder nicht gesetzt, einfach binär kopieren
        // Das erhält 100% der Formatierung!
        if (!keepHeader) {
            fs.copyFileSync(sourcePath, targetPath);
            return { success: true, message: `Datei kopiert: ${targetPath}` };
        }
        
        // Wenn keepHeader true ist, müssen wir die Daten löschen
        // Aber zuerst: Binär kopieren, dann nur die Werte löschen
        fs.copyFileSync(sourcePath, targetPath);
        
        // Jetzt die kopierte Datei öffnen und nur die Datenwerte löschen
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
        
        // Nur die Werte ab Zeile 2 löschen (Header in Zeile 1 bleibt)
        // Formatierung bleibt erhalten!
        const usedRange = worksheet.usedRange();
        if (usedRange) {
            const endRow = usedRange.endCell().rowNumber();
            const endCol = usedRange.endCell().columnNumber();
            
            // Alle Datenwerte ab Zeile 2 löschen
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

// Daten exportieren (fuer Datenexplorer)
ipcMain.handle('excel:exportData', async (event, { filePath, sheetName, headers, rows }) => {
    try {
        // Neue leere Workbook erstellen
        const workbook = await XlsxPopulate.fromBlankAsync();
        
        // Erstes Sheet umbenennen
        const worksheet = workbook.sheet(0);
        worksheet.name(sheetName.substring(0, 31));
        
        // Header-Zeile
        headers.forEach((header, colIndex) => {
            worksheet.cell(1, colIndex + 1).value(header);
        });
        
        // Daten-Zeilen
        rows.forEach((row, rowIndex) => {
            headers.forEach((header, colIndex) => {
                const value = row[header] || '';
                worksheet.cell(rowIndex + 2, colIndex + 1).value(value);
            });
        });
        
        await workbook.toFileAsync(filePath);
        
        return { success: true, message: `Export erstellt: ${filePath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// ============================================
// KONFIGURATION
// ============================================

// Config speichern
ipcMain.handle('config:save', async (event, { filePath, config }) => {
    try {
        fs.writeFileSync(filePath, JSON.stringify(config, null, 2), 'utf8');
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Config laden
ipcMain.handle('config:load', async (event, filePath) => {
    try {
        if (!fs.existsSync(filePath)) {
            return { success: false, error: 'Datei nicht gefunden' };
        }
        const content = fs.readFileSync(filePath, 'utf8');
        return { success: true, config: JSON.parse(content) };
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

// Automatisch config.json im Programmordner suchen
ipcMain.handle('config:loadFromAppDir', async (event) => {
    try {
        const exePath = app.getPath('exe');
        const exeDir = path.dirname(exePath);
        
        // PORTABLE EXE: Der Ordner wo die portable EXE gestartet wurde
        const portableDir = process.env.PORTABLE_EXECUTABLE_DIR || '';
        
        // Nur die wichtigsten Pfade prüfen (optimiert)
        const possiblePaths = [];
        
        // 1. Portable EXE: Neben der EXE
        if (portableDir) {
            possiblePaths.push(path.join(portableDir, 'config.json'));
        }
        
        // 2. Neben der EXE (Standard)
        possiblePaths.push(path.join(exeDir, 'config.json'));
        
        // 3. Im Entwicklungsmodus: Projektordner
        if (process.argv.includes('--dev') || !app.isPackaged) {
            possiblePaths.push(path.join(__dirname, 'config.json'));
            possiblePaths.push(path.join(process.cwd(), 'config.json'));
        }
        
        // Schnelle Suche - bei erstem Treffer abbrechen
        for (const configPath of possiblePaths) {
            if (fs.existsSync(configPath)) {
                console.log('? config.json gefunden:', configPath);
                const content = fs.readFileSync(configPath, 'utf8');
                return { 
                    success: true, 
                    config: JSON.parse(content),
                    path: configPath
                };
            }
        }
        
        // Keine config.json gefunden - kein Fehler, nur Info
        return { 
            success: false, 
            error: 'Keine config.json gefunden'
        };
    } catch (error) {
        console.error('Fehler beim Laden der config.json:', error);
        return { success: false, error: error.message };
    }
});
