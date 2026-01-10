const { contextBridge, ipcRenderer, webUtils } = require('electron');

// Sichere API fuer das Frontend bereitstellen
contextBridge.exposeInMainWorld('electronAPI', {
    // Dialoge
    openFileDialog: (options) => ipcRenderer.invoke('dialog:openFile', options),
    saveFileDialog: (options) => ipcRenderer.invoke('dialog:saveFile', options),
    openFolderDialog: (options) => ipcRenderer.invoke('dialog:openFolder', options),
    
    // Dateisystem
    checkFileExists: (filePath) => ipcRenderer.invoke('fs:checkFileExists', filePath),
    
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
    exportData: (params) => ipcRenderer.invoke('excel:exportData', params),
    exportWithAllSheets: (params) => ipcRenderer.invoke('excel:exportWithAllSheets', params),
    exportMultipleSheets: (params) => ipcRenderer.invoke('excel:exportMultipleSheets', params),
    saveExcelFile: (params) => ipcRenderer.invoke('excel:saveFile', params),
    createTemplateFromSource: (params) => ipcRenderer.invoke('excel:createTemplateFromSource', params),
    
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
    
    // App-Infos
    getAppPath: () => ipcRenderer.invoke('app:getPath'),
    
    // System-Infos (für Computer-spezifische Config)
    getComputerName: () => ipcRenderer.invoke('system:getComputerName'),
    
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
    confirmClose: (canClose) => ipcRenderer.send('app:confirmClose', canClose)
});
