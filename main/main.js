const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { PythonShell } = require('python-shell');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 1200,
        minHeight: 700,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        icon: path.join(__dirname, '../renderer/assets/icon.png'),
        titleBarStyle: 'default',
        backgroundColor: '#f8fafc'
    });

    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
    
    // Open DevTools in development
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }
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

// IPC Handlers

// Get Python path
ipcMain.handle('get-python-path', async () => {
    return process.platform === 'win32' ? 'python' : 'python3';
});

// Open file dialog
ipcMain.handle('open-file-dialog', async (event, options) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: options.multiple ? ['openFile', 'multiSelections'] : ['openFile'],
        filters: [
            { name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] }
        ]
    });
    return result.filePaths;
});

// Open folder dialog
ipcMain.handle('open-folder-dialog', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory']
    });
    return result.filePaths[0];
});

// Save file dialog
ipcMain.handle('save-file-dialog', async (event, defaultName) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: defaultName,
        filters: [
            { name: 'CSV Files', extensions: ['csv'] }
        ]
    });
    return result.filePath;
});

// Save Excel file dialog
ipcMain.handle('save-excel-dialog', async (event, defaultName) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: defaultName,
        filters: [
            { name: 'Excel Files', extensions: ['xlsx'] }
        ]
    });
    return result.filePath;
});

// Run Python processor
ipcMain.handle('run-python', async (event, { script, args }) => {
    return new Promise((resolve, reject) => {
        const options = {
            mode: 'json',
            pythonPath: process.platform === 'win32' ? 'python' : 'python3',
            pythonOptions: ['-u'],
            scriptPath: path.join(__dirname, '../backend'),
            args: args ? [JSON.stringify(args)] : []
        };

        PythonShell.run(script, options)
            .then(results => {
                resolve(results[results.length - 1]);
            })
            .catch(err => {
                reject(err.message);
            });
    });
});

// Match files to ateliers
ipcMain.handle('match-files', async (event, { unit, files }) => {
    return new Promise((resolve, reject) => {
        const options = {
            mode: 'json',
            pythonPath: process.platform === 'win32' ? 'python' : 'python3',
            pythonOptions: ['-u'],
            scriptPath: path.join(__dirname, '../backend'),
            args: [JSON.stringify({ action: 'match', unit, files })]
        };

        PythonShell.run('processor.py', options)
            .then(results => {
                resolve(results[results.length - 1]);
            })
            .catch(err => {
                reject(err.message);
            });
    });
});

// Process files and calculate
ipcMain.handle('process-files', async (event, { unit, stockFile, matchedFiles, month, overrides }) => {
    return new Promise((resolve, reject) => {
        const options = {
            mode: 'json',
            pythonPath: process.platform === 'win32' ? 'python' : 'python3',
            pythonOptions: ['-u'],
            scriptPath: path.join(__dirname, '../backend'),
            args: [JSON.stringify({ 
                action: 'process', 
                unit, 
                stockFile, 
                matchedFiles,
                month,
                overrides
            })]
        };

        console.log('Processing with args:', JSON.stringify({ unit, stockFile: stockFile?.path, matchedFilesCount: Object.keys(matchedFiles).length, month }));

        const pyshell = new PythonShell('processor.py', options);
        let results = [];
        
        pyshell.on('message', function (message) {
            console.log('Python output:', message);
            results.push(message);
        });
        
        pyshell.on('stderr', function (stderr) {
            console.log('Python stderr (debug):', stderr);
        });
        
        pyshell.end(function (err, code, signal) {
            if (err) {
                console.error('Python error:', err);
                reject(err.message);
            } else {
                console.log('Python finished with code:', code);
                if (results.length > 0) {
                    resolve(results[results.length - 1]);
                } else {
                    resolve({});
                }
            }
        });
    });
});

// Export results
ipcMain.handle('export-csv', async (event, { data, filePath }) => {
    const fs = require('fs');
    return new Promise((resolve, reject) => {
        try {
            // Convert data array to CSV
            if (data.length === 0) {
                fs.writeFileSync(filePath, '', 'utf-8');
                resolve(true);
                return;
            }
            
            const headers = Object.keys(data[0]);
            const csvContent = [
                headers.join(','),
                ...data.map(row => headers.map(h => {
                    const val = row[h];
                    if (typeof val === 'string' && val.includes(',')) {
                        return `"${val}"`;
                    }
                    return val;
                }).join(','))
            ].join('\n');
            
            fs.writeFileSync(filePath, '\ufeff' + csvContent, 'utf-8'); // BOM for Excel UTF-8
            resolve(true);
        } catch (err) {
            reject(err.message);
        }
    });
});

// Verify files
ipcMain.handle('verify-files', async (event, { unit, matchedFiles }) => {
    return new Promise((resolve, reject) => {
        const options = {
            mode: 'json',
            pythonPath: process.platform === 'win32' ? 'python' : 'python3',
            pythonOptions: ['-u'],
            scriptPath: path.join(__dirname, '../backend'),
            args: [JSON.stringify({ 
                action: 'verify', 
                unit, 
                matchedFiles
            })]
        };

        console.log('Verifying files for unit:', unit);

        const pyshell = new PythonShell('processor.py', options);
        let results = [];
        
        pyshell.on('message', function (message) {
            console.log('Python verify output:', message);
            results.push(message);
        });
        
        pyshell.on('stderr', function (stderr) {
            console.log('Python stderr (debug):', stderr);
        });
        
        pyshell.end(function (err, code, signal) {
            if (err) {
                console.error('Python error:', err);
                reject(err.message);
            } else {
                console.log('Python verify finished with code:', code);
                if (results.length > 0) {
                    resolve(results[results.length - 1]);
                } else {
                    resolve({});
                }
            }
        });
    });
});

// Export to Excel
ipcMain.handle('export-excel', async (event, { data, filePath }) => {
    return new Promise((resolve, reject) => {
        const options = {
            mode: 'json',
            pythonPath: process.platform === 'win32' ? 'python' : 'python3',
            pythonOptions: ['-u'],
            scriptPath: path.join(__dirname, '../backend'),
            args: [JSON.stringify({ 
                action: 'export_excel', 
                data, 
                outputPath: filePath
            })]
        };

        console.log('Exporting to Excel:', filePath);

        const pyshell = new PythonShell('processor.py', options);
        let results = [];
        
        pyshell.on('message', function (message) {
            results.push(message);
        });
        
        pyshell.on('stderr', function (stderr) {
            console.log('Python stderr:', stderr);
        });
        
        pyshell.end(function (err, code, signal) {
            if (err) {
                console.error('Python error:', err);
                reject(err.message);
            } else {
                if (results.length > 0 && results[results.length - 1].success) {
                    resolve(true);
                } else {
                    reject('Export failed');
                }
            }
        });
    });
});
