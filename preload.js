const { contextBridge, ipcRenderer, webUtils } = require('electron');

// Sichere API fuer das Frontend bereitstellen
contextBridge.exposeInMainWorld('electronAPI', {
    // Dialoge
    openFileDialog: (options) => ipcRenderer.invoke('dialog:openFile', options),
    saveFileDialog: (options) => ipcRenderer.invoke('dialog:saveFile', options),
    openFolderDialog: (options) => ipcRenderer.invoke('dialog:openFolder', options),
    
    // Dateisystem
    checkFileExists: (filePath) => ipcRenderer.invoke('fs:checkFileExists', filePath),
    findFiles: (directory, pattern) => ipcRenderer.invoke('fs:findFiles', { directory, pattern }),
    
    // Drag & Drop - Dateipfad aus File-Objekt extrahieren
    getPathForFile: (file) => {
        try {
            // Electron 32+ verwendet webUtils.getPathForFile
            if (webUtils && webUtils.getPathForFile) {
                return webUtils.getPathForFile(file);
            }
            // Fallback für ältere Versionen
            return file.path || null;
        } catch (e) {
            console.error('getPathForFile error:', e);
            return null;
        }
    },
    
    // Excel-Operationen
    readExcelFile: (filePath, password) => ipcRenderer.invoke('excel:readFile', filePath, password),
    readExcelSheet: (filePath, sheetName, password) => ipcRenderer.invoke('excel:readSheet', filePath, sheetName, password),
    insertExcelRows: (params) => ipcRenderer.invoke('excel:insertRows', params),
    copyExcelFile: (params) => ipcRenderer.invoke('excel:copyFile', params),
    createTemplateFromSource: (params) => ipcRenderer.invoke('excel:createTemplateFromSource', params),
    
    // Python/openpyxl Writer (behält CF und Formatierungen)
    pythonExportMultipleSheets: (params) => ipcRenderer.invoke('python:exportMultipleSheets', params),
    
    // Excel-Engine Steuerung
    checkExcelAvailable: () => ipcRenderer.invoke('excel:checkAvailable'),
    
    // Sheet-Verwaltung
    addSheet: (params) => ipcRenderer.invoke('excel:addSheet', params),
    deleteSheet: (params) => ipcRenderer.invoke('excel:deleteSheet', params),
    renameSheet: (params) => ipcRenderer.invoke('excel:renameSheet', params),
    cloneSheet: (params) => ipcRenderer.invoke('excel:cloneSheet', params),
    moveSheet: (params) => ipcRenderer.invoke('excel:moveSheet', params),
    
    // Konfiguration
    saveConfig: (filePath, config) => ipcRenderer.invoke('config:save', { filePath, config }),
    loadConfig: (filePath) => ipcRenderer.invoke('config:load', filePath),
    loadConfigFromAppDir: (workingDir) => ipcRenderer.invoke('config:loadFromAppDir', workingDir),
    
    // Externe URLs öffnen
    openExternal: (url) => ipcRenderer.invoke('shell:openExternal', url),
    
    // Security-Logs
    getSecurityLogs: (options) => ipcRenderer.invoke('security:getLogs', options),
    verifySecurityLogs: () => ipcRenderer.invoke('security:verifyLogs'),
    clearSecurityLogs: () => ipcRenderer.invoke('security:clearLogs'),
    
    // Netzwerk-Logs
    isNetworkPath: (filePath) => ipcRenderer.invoke('network:isNetworkPath', filePath),
    getNetworkLogs: (filePath) => ipcRenderer.invoke('network:getLogs', filePath),
    checkNetworkConflict: (filePath, minutes) => ipcRenderer.invoke('network:checkConflict', filePath, minutes),
    createSessionLock: (filePath) => ipcRenderer.invoke('network:createSessionLock', filePath),
    removeSessionLock: (filePath) => ipcRenderer.invoke('network:removeSessionLock', filePath),
    
    // Event-Listener für App-Schließen
    onBeforeClose: (callback) => ipcRenderer.on('app:beforeClose', callback),
    confirmClose: (canClose) => ipcRenderer.send('app:confirmClose', canClose),
    
    // ==========================================================================
    // LIVE SESSION API - Excel bleibt offen für sofortige Operationen
    // ==========================================================================
    
    // Session-Management
    liveSessionStart: () => ipcRenderer.invoke('liveSession:start'),
    liveSessionInitApp: () => ipcRenderer.invoke('liveSession:initApp'),
    liveSessionOpenFile: (filePath, sheetName, password) => ipcRenderer.invoke('liveSession:openFile', filePath, sheetName, password),
    liveSessionSaveFile: (outputPath, password) => ipcRenderer.invoke('liveSession:saveFile', outputPath, password),
    liveSessionClose: () => ipcRenderer.invoke('liveSession:close'),
    
    // Passwort-Funktionen
    liveSessionSetPassword: (password) => ipcRenderer.invoke('liveSession:setPassword', password),
    
    // Zeilen-Operationen (werden SOFORT in Excel ausgeführt!)
    liveSessionDeleteRow: (rowIndex) => ipcRenderer.invoke('liveSession:deleteRow', rowIndex),
    liveSessionInsertRow: (rowIndex, count) => ipcRenderer.invoke('liveSession:insertRow', rowIndex, count || 1),
    liveSessionMoveRow: (fromIndex, toIndex) => ipcRenderer.invoke('liveSession:moveRow', fromIndex, toIndex),
    liveSessionHideRow: (rowIndex, hidden) => ipcRenderer.invoke('liveSession:hideRow', rowIndex, hidden !== false),
    liveSessionHideRowsBatch: (rowIndices, hidden) => ipcRenderer.invoke('liveSession:hideRowsBatch', rowIndices, hidden !== false),
    liveSessionHighlightRow: (rowIndex, color) => ipcRenderer.invoke('liveSession:highlightRow', rowIndex, color),
    
    // Spalten-Operationen (werden SOFORT in Excel ausgeführt!)
    liveSessionDeleteColumn: (colIndex) => ipcRenderer.invoke('liveSession:deleteColumn', colIndex),
    liveSessionInsertColumn: (colIndex, count, headers) => ipcRenderer.invoke('liveSession:insertColumn', colIndex, count || 1, headers),
    liveSessionMoveColumn: (fromIndex, toIndex) => ipcRenderer.invoke('liveSession:moveColumn', fromIndex, toIndex),
    liveSessionHideColumn: (colIndex, hidden) => ipcRenderer.invoke('liveSession:hideColumn', colIndex, hidden !== false),
    
    // Zell-Operationen
    liveSessionSetCellValue: (rowIndex, colIndex, value) => ipcRenderer.invoke('liveSession:setCellValue', rowIndex, colIndex, value),
    liveSessionSetColumnValues: (colIndex, values, startRow) => ipcRenderer.invoke('liveSession:setColumnValues', colIndex, values, startRow || 0),
    liveSessionSetCellsBatch: (cells) => ipcRenderer.invoke('liveSession:setCellsBatch', cells),
    liveSessionCopyCells: (sourceCells, targetRow, targetCol) => ipcRenderer.invoke('liveSession:copyCells', sourceCells, targetRow, targetCol),
    liveSessionSwitchSheet: (sheetName) => ipcRenderer.invoke('liveSession:switchSheet', sheetName),
    liveSessionSetRowValues: (rowIndex, values) => ipcRenderer.invoke('liveSession:setRowValues', rowIndex, values),
    liveSessionFindReplace: (searchText, replaceText, matchCase, wholeWord) => 
        ipcRenderer.invoke('liveSession:findReplace', searchText, replaceText, matchCase, wholeWord),
    
    // Daten lesen
    liveSessionGetData: () => ipcRenderer.invoke('liveSession:getData'),
    
    // Undo
    liveSessionUndo: () => ipcRenderer.invoke('liveSession:undo'),
    
    // Filter-Operationen
    liveSessionSetAutoFilter: (filters) => ipcRenderer.invoke('liveSession:setAutoFilter', filters),
    liveSessionClearAutoFilter: () => ipcRenderer.invoke('liveSession:clearAutoFilter'),
    
    // Sichtbarkeit & Status
    liveSessionSetVisible: (visible) => ipcRenderer.invoke('liveSession:setVisible', visible),
    liveSessionCheckAlive: () => ipcRenderer.invoke('liveSession:checkAlive'),
    
    // Recovery
    liveSessionGetRecoveryFiles: () => ipcRenderer.invoke('liveSession:getRecoveryFiles'),
    liveSessionOpenRecoveryFolder: () => ipcRenderer.invoke('liveSession:openRecoveryFolder')
});