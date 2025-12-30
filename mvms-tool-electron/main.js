const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

let mainWindow;

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
        title: 'MVMS-Tool',
        icon: path.join(__dirname, 'assets', 'icon.ico'),
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });

    // Fenster maximieren für bessere Übersicht
    // mainWindow.maximize();

    mainWindow.loadFile('src/index.html');
    
    // DevTools öffnen (nur während Entwicklung)
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }
    
    // Menüleiste ausblenden (optional)
    mainWindow.setMenuBarVisibility(false);
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
// DATEI-DIALOGE
// ============================================

// Datei öffnen Dialog
ipcMain.handle('dialog:openFile', async (event, options) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        title: options.title || 'Datei öffnen',
        filters: options.filters || [
            { name: 'Excel-Dateien', extensions: ['xlsx', 'xls'] },
            { name: 'Alle Dateien', extensions: ['*'] }
        ],
        properties: ['openFile']
    });
    
    if (result.canceled) return null;
    return result.filePaths[0];
});

// Datei speichern Dialog
ipcMain.handle('dialog:saveFile', async (event, options) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        title: options.title || 'Datei speichern',
        defaultPath: options.defaultPath,
        filters: options.filters || [
            { name: 'Excel-Dateien', extensions: ['xlsx'] }
        ]
    });
    
    if (result.canceled) return null;
    return result.filePath;
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
            
            // Zeilen auffüllen bis zur maximalen Spaltenanzahl
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

// Zeilen in Excel einfügen (MIT Formatierungserhalt!)
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
        
        // Neue Zeilen einfügen
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
            
            // Daten ab Startspalte
            if (row.data && row.flag !== 'leer') {
                row.data.forEach((value, index) => {
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
            message: `${insertedCount} Zeile(n) eingefügt`,
            insertedCount: insertedCount
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Datei kopieren (für "Neuer Monat")
ipcMain.handle('excel:copyFile', async (event, { sourcePath, targetPath, sheetName, keepHeader }) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(sourcePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
            return { success: false, error: `Sheet "${sheetName}" nicht gefunden` };
        }
        
        // Zeilen löschen (außer Header)
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

// Daten exportieren (für Datenexplorer)
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

// App-Pfad ermitteln (für Config im Programmordner)
ipcMain.handle('app:getPath', async (event) => {
    return {
        appPath: app.getAppPath(),
        userData: app.getPath('userData'),
        exe: app.getPath('exe')
    };
});
