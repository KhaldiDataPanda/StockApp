const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // File dialogs
    openFileDialog: (options) => ipcRenderer.invoke('open-file-dialog', options),
    openFolderDialog: () => ipcRenderer.invoke('open-folder-dialog'),
    saveFileDialog: (defaultName) => ipcRenderer.invoke('save-file-dialog', defaultName),
    
    // Python integration
    runPython: (script, args) => ipcRenderer.invoke('run-python', { script, args }),
    matchFiles: (unit, files) => ipcRenderer.invoke('match-files', { unit, files }),
    processFiles: (unit, stockFile, matchedFiles, month) => 
        ipcRenderer.invoke('process-files', { unit, stockFile, matchedFiles, month }),
    
    // Export
    exportCSV: (data, filePath) => ipcRenderer.invoke('export-csv', { data, filePath })
});
