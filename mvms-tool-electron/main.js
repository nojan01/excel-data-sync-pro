const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
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
// EXCEL OPERATIONEN
// ============================================

// Excel-Datei lesen
ipcMain.handle('excel:readFile', async (event, filePath) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        const sheets = workbook.worksheets.map(ws => ws.name);
        
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
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        const data = [];
        const headers = [];
        
        worksheet.eachRow((row, rowNumber) => {
            const rowData = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                // Header-Zeile
                if (rowNumber === 1) {
                    headers[colNumber - 1] = cell.text || `Spalte ${colNumber}`;
                }
                rowData[colNumber - 1] = cell.text || '';
            });
            
            // Zeilen auffuellen bis zur maximalen Spaltenanzahl
            while (rowData.length < headers.length) {
                rowData.push('');
            }
            
            data.push(rowData);
        });
        
        return {
            success: true,
            headers: headers,
            data: data
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Zeilen in Excel einfuegen (MIT Formatierungserhalt!)
ipcMain.handle('excel:insertRows', async (event, { filePath, sheetName, rows, startColumn }) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Letzte nicht-leere Zeile finden
        let lastRow = 1;
        worksheet.eachRow((row, rowNumber) => {
            let isEmpty = true;
            row.eachCell(cell => {
                if (cell.value !== null && cell.value !== '') {
                    isEmpty = false;
                }
            });
            if (!isEmpty) {
                lastRow = rowNumber;
            }
        });
        
        // Neue Zeilen einfuegen
        let insertedCount = 0;
        for (const row of rows) {
            const newRowNum = lastRow + insertedCount + 1;
            const newRow = worksheet.getRow(newRowNum);
            
            // Flag in Spalte A
            if (row.flag && row.flag !== 'leer') {
                newRow.getCell(1).value = row.flag;
            }
            
            // Kommentar in Spalte B
            if (row.comment) {
                newRow.getCell(2).value = row.comment;
            }
            
            // Daten ab Startspalte - row.data ist ein Objekt mit Index als Key
            if (row.data && row.flag !== 'leer') {
                const dataKeys = Object.keys(row.data);
                dataKeys.forEach(key => {
                    const index = parseInt(key);
                    const value = row.data[key];
                    if (value !== null && value !== undefined && value !== '') {
                        newRow.getCell(startColumn + index).value = value;
                    }
                });
            }
            
            newRow.commit();
            insertedCount++;
        }
        
        // Speichern
        await workbook.xlsx.writeFile(filePath);
        
        return { 
            success: true, 
            message: `${insertedCount} Zeile(n) eingefuegt`,
            insertedCount: insertedCount
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Datei kopieren (fuer "Neuer Monat")
ipcMain.handle('excel:copyFile', async (event, { sourcePath, targetPath, sheetName, keepHeader }) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(sourcePath);
        
        // Wenn sheetName angegeben und existiert, nutze dieses Sheet
        // Ansonsten nimm das erste Sheet
        let worksheet = null;
        if (sheetName) {
            worksheet = workbook.getWorksheet(sheetName);
        }
        if (!worksheet) {
            // Erstes verfuegbares Sheet nehmen
            worksheet = workbook.worksheets[0];
        }
        
        if (!worksheet) {
            return { success: false, error: 'Keine Worksheets in der Template-Datei gefunden' };
        }
        
        // Zeilen loeschen (ausser Header) wenn gewuenscht
        if (keepHeader) {
            const rowCount = worksheet.rowCount;
            for (let i = rowCount; i > 1; i--) {
                worksheet.spliceRows(i, 1);
            }
        }
        
        await workbook.xlsx.writeFile(targetPath);
        
        return { success: true, message: `Datei erstellt: ${targetPath}` };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Daten exportieren (fuer Datenexplorer)
ipcMain.handle('excel:exportData', async (event, { filePath, sheetName, headers, rows }) => {
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(sheetName.substring(0, 31));
        
        // Header-Zeile
        worksheet.addRow(headers);
        
        // Daten-Zeilen
        rows.forEach(row => {
            const rowValues = headers.map(h => row[h] || '');
            worksheet.addRow(rowValues);
        });
        
        // Spaltenbreiten automatisch anpassen
        worksheet.columns.forEach((column, i) => {
            let maxLength = headers[i] ? headers[i].length : 10;
            rows.forEach(row => {
                const cellValue = String(row[headers[i]] || '');
                if (cellValue.length > maxLength) {
                    maxLength = Math.min(cellValue.length, 50);
                }
            });
            column.width = maxLength + 2;
        });
        
        await workbook.xlsx.writeFile(filePath);
        
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
