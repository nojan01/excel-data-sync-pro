const { contextBridge, ipcRenderer } = require('electron');

// Sichere API fuer das Frontend bereitstellen
contextBridge.exposeInMainWorld('electronAPI', {
    // Dialoge
    openFileDialog: (options) => ipcRenderer.invoke('dialog:openFile', options),
    saveFileDialog: (options) => ipcRenderer.invoke('dialog:saveFile', options),
    
    // Excel-Operationen
    readExcelFile: (filePath) => ipcRenderer.invoke('excel:readFile', filePath),
    readExcelSheet: (filePath, sheetName) => ipcRenderer.invoke('excel:readSheet', filePath, sheetName),
    insertExcelRows: (params) => ipcRenderer.invoke('excel:insertRows', params),
    copyExcelFile: (params) => ipcRenderer.invoke('excel:copyFile', params),
    exportData: (params) => ipcRenderer.invoke('excel:exportData', params),
    exportWithAllSheets: (params) => ipcRenderer.invoke('excel:exportWithAllSheets', params),
    createTemplateFromSource: (params) => ipcRenderer.invoke('excel:createTemplateFromSource', params),
    
    // Konfiguration
    saveConfig: (filePath, config) => ipcRenderer.invoke('config:save', { filePath, config }),
    loadConfig: (filePath) => ipcRenderer.invoke('config:load', filePath),
    loadConfigFromAppDir: () => ipcRenderer.invoke('config:loadFromAppDir'),
    
    // App-Infos
    getAppPath: () => ipcRenderer.invoke('app:getPath'),
    
    // Event-Listener für App-Schließen
    onBeforeClose: (callback) => ipcRenderer.on('app:beforeClose', callback),
    confirmClose: (canClose) => ipcRenderer.send('app:confirmClose', canClose)
});
