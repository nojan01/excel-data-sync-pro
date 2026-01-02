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
            const excelEpoch = new Date(1899, 11, 30); // 30.12.1899
            const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
            
            // Prüfen ob es ein reines Datum oder Datum mit Uhrzeit ist
            const hasTime = (excelDate % 1) !== 0;
            
            if (hasTime) {
                // Datum mit Uhrzeit
                const day = String(jsDate.getDate()).padStart(2, '0');
                const month = String(jsDate.getMonth() + 1).padStart(2, '0');
                const year = jsDate.getFullYear();
                const hours = String(jsDate.getHours()).padStart(2, '0');
                const minutes = String(jsDate.getMinutes()).padStart(2, '0');
                return `${day}.${month}.${year} ${hours}:${minutes}`;
            } else {
                // Nur Datum
                const day = String(jsDate.getDate()).padStart(2, '0');
                const month = String(jsDate.getMonth() + 1).padStart(2, '0');
                const year = jsDate.getFullYear();
                return `${day}.${month}.${year}`;
            }
        }
        
        // Hilfsfunktion: Prüfen ob ein Zellwert ein Datum ist
        function isExcelDate(cell, value) {
            if (typeof value !== 'number') return false;
            
            // Prüfe das Zahlenformat der Zelle
            const numFmt = cell.style('numberFormat');
            if (numFmt) {
                // Typische Datumsformate erkennen
                const datePatterns = [
                    /d+[\/\-.]m+[\/\-.]y+/i,  // d.m.y, d/m/y, d-m-y
                    /m+[\/\-.]d+[\/\-.]y+/i,  // m/d/y (US-Format)
                    /y+[\/\-.]m+[\/\-.]d+/i,  // y-m-d (ISO-Format)
                    /\[.*?\]dd/i,              // Benutzerdefinierte Formate
                    /^d+$/i,                   // "d" oder "dd"
                    /mmm/i,                    // Monatsname (mmm, mmmm)
                    /^[$-F800]dddd/,           // Locale-spezifische Formate
                ];
                
                for (const pattern of datePatterns) {
                    if (pattern.test(numFmt)) {
                        return true;
                    }
                }
            }
            
            // Heuristik: Zahlen im typischen Excel-Datumsbereich (1900-2100)
            // Excel-Datum 1 = 1.1.1900, 73050 ? 1.1.2100
            if (value >= 1 && value <= 73050) {
                // Könnte ein Datum sein - prüfe auf vernünftigen Wert
                // Aber nur wenn kein offensichtliches Zahlenformat
                if (numFmt && /^[#0,]+(\.[#0]+)?$/.test(numFmt)) {
                    return false; // Explizites Zahlenformat
                }
                if (numFmt && numFmt.indexOf('%') !== -1) {
                    return false; // Prozentformat
                }
                if (numFmt && numFmt.indexOf('€') !== -1) {
                    return false; // Währungsformat
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
                    if (isExcelDate(cell, value)) {
                        textValue = excelDateToString(value);
                    } else if (value instanceof Date) {
                        // Falls xlsx-populate bereits ein Date-Objekt zurückgibt
                        const day = String(value.getDate()).padStart(2, '0');
                        const month = String(value.getMonth() + 1).padStart(2, '0');
                        const year = value.getFullYear();
                        textValue = `${day}.${month}.${year}`;
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
                
                // Excel-Datum berechnen
                const jsDate = new Date(year, month - 1, day, hours, minutes);
                const excelEpoch = new Date(1899, 11, 30);
                const excelDate = (jsDate.getTime() - excelEpoch.getTime()) / 86400000;
                return excelDate;
            } else if (dateMatch) {
                const day = parseInt(dateMatch[1], 10);
                const month = parseInt(dateMatch[2], 10);
                const year = parseInt(dateMatch[3], 10);
                
                // Excel-Datum berechnen
                const jsDate = new Date(year, month - 1, day);
                const excelEpoch = new Date(1899, 11, 30);
                const excelDate = Math.floor((jsDate.getTime() - excelEpoch.getTime()) / 86400000);
                return excelDate;
            }
            
            return null;
        }
        
        // Hilfsfunktion: Wert intelligent konvertieren
        function convertValue(value, targetCell) {
            if (value === null || value === undefined || value === '') {
                return value;
            }
            
            // Prüfe ob es ein deutsches Datum ist
            const excelDate = parseGermanDateToExcel(value);
            if (excelDate !== null) {
                return excelDate;
            }
            
            // Prüfe ob es eine Zahl ist (mit deutschen Dezimaltrennzeichen)
            if (typeof value === 'string') {
                // Deutsche Zahlen: 1.234,56 -> 1234.56
                const germanNumberMatch = value.match(/^-?\d{1,3}(\.\d{3})*(,\d+)?$/);
                if (germanNumberMatch) {
                    const normalized = value.replace(/\./g, '').replace(',', '.');
                    const num = parseFloat(normalized);
                    if (!isNaN(num)) {
                        return num;
                    }
                }
                
                // Englische Zahlen oder einfache Zahlen
                const simpleNumber = value.match(/^-?\d+(\.\d+)?$/);
                if (simpleNumber) {
                    const num = parseFloat(value);
                    if (!isNaN(num)) {
                        return num;
                    }
                }
            }
            
            // Als String belassen
            return value;
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
                        const targetCell = worksheet.cell(newRowNum, startColumn + index);
                        const convertedValue = convertValue(value, targetCell);
                        targetCell.value(convertedValue);
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
