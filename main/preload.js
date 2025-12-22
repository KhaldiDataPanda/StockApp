const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // File dialogs
    openFileDialog: (options) => ipcRenderer.invoke('open-file-dialog', options),
    openFolderDialog: () => ipcRenderer.invoke('open-folder-dialog'),
    saveFileDialog: (defaultName) => ipcRenderer.invoke('save-file-dialog', defaultName),
    saveExcelDialog: (defaultName) => ipcRenderer.invoke('save-excel-dialog', defaultName),
    
    // Python integration
    runPython: (script, args) => ipcRenderer.invoke('run-python', { script, args }),
    matchFiles: (unit, files) => ipcRenderer.invoke('match-files', { unit, files }),
    verifyFiles: (unit, matchedFiles) => ipcRenderer.invoke('verify-files', { unit, matchedFiles }),
    processFiles: (unit, stockFile, matchedFiles, month, overrides) => 
        ipcRenderer.invoke('process-files', { unit, stockFile, matchedFiles, month, overrides }),
    
    // Export
    exportCSV: (data, filePath) => ipcRenderer.invoke('export-csv', { data, filePath }),
    exportExcel: (data, filePath) => ipcRenderer.invoke('export-excel', { data, filePath })
});
