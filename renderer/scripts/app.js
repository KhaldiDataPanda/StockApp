/**
 * Stock Reconciliation App - Main Application Logic
 */

// ============================================
// State Management
// ============================================
const AppState = {
    currentPage: 'page-unit-selection',
    selectedUnit: null,
    selectedMonth: '12',
    stockFile: null,
    droppedFiles: [],
    matchedFiles: {},      // { atelier: { path, filename } }
    unmatchedFiles: [],    // [{ path, filename }]
    ateliers: [],          // List of ateliers for selected unit
    skippedAteliers: new Set(),
    results: null          // Processing results
};

// Unit configurations with keywords
const UnitConfigs = {
    Fath1: {
        name: 'Fath 1',
        ateliers: [
            'femme 01', 'femme 02', 'bourde', 'coupage', '+croute', 'grattage+',
            'produt', '-grattage', 'rouler', 'conftection rouli', 'brodri',
            'conftection bourdi', 'fiber cardi', 'orielle', 'bloc',
            'magaza bourdi', 'secondaire', 'mov-com'
        ]
    },
    Fath2: {
        name: 'Fath 2',
        ateliers: [
            'ouate 01', 'ouate 02', 'sfifa', 'mgz plasic', 'secondaire',
            '-dechet', 'commercial', 'فيبر', 'plastique', 'tissu'
        ]
    },
    Fath5: {
        name: 'Fath 5',
        ateliers: [
            'bonda', 'orillier', 'confiction', 'couette fini', 'semi fini',
            'rouli', 'block', 'gratage', 'coupage', 'comersial', 'secondaire',
            'ouate', 'outin'
        ]
    },
    Larbaa: {
        name: 'Larbaa',
        ateliers: [
            'atelier découpage', 'coupage ', 'roulés.entré', 'roulés.sortie',
            'oreiller', 'couette', 'magasin blocs', 'magasin fibre',
            'magasin ouate', 'magasin mousse', 'magasin roules', 'grattage',
            'accessoire', 'piec', '(ouate', 'sortie mousse couture', ' pet '
        ]
    },
    Oran: {
        name: 'Oran',
        ateliers: ['block', 'mousse', 'rouléss', 'cardage']
    },
    Fibre: {
        name: 'Fibre',
        ateliers: [
            'drafter', 'extredeuse', 'filiére', 'carding',
            'magaisain pet', 'magaisain fibre', 'magaisain commercial'
        ]
    }
};

// ============================================
// DOM Elements
// ============================================
const elements = {
    // Pages
    pages: {
        unitSelection: document.getElementById('page-unit-selection'),
        fileUpload: document.getElementById('page-file-upload'),
        processing: document.getElementById('page-processing'),
        results: document.getElementById('page-results')
    },
    // Navigation
    breadcrumb: document.getElementById('breadcrumb'),
    btnBackToUnits: document.getElementById('btn-back-to-units'),
    btnBackToUpload: document.getElementById('btn-back-to-upload'),
    // Unit Selection
    unitCards: document.querySelectorAll('.unit-card'),
    selectedUnitName: document.getElementById('selected-unit-name'),
    // File Upload
    monthSelect: document.getElementById('month-select'),
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    stockFileDisplay: document.getElementById('stock-file-display'),
    stockStatus: document.getElementById('stock-status'),
    ateliersList: document.getElementById('ateliers-list'),
    unmatchedList: document.getElementById('unmatched-list'),
    matchedCount: document.getElementById('matched-count'),
    unmatchedCount: document.getElementById('unmatched-count'),
    btnClearFiles: document.getElementById('btn-clear-files'),
    btnProcess: document.getElementById('btn-process'),
    // Processing
    processingStatus: document.getElementById('processing-status'),
    progressFill: document.getElementById('progress-fill'),
    // Results
    totalMatches: document.getElementById('total-matches'),
    totalDiscrepancies: document.getElementById('total-discrepancies'),
    totalAteliers: document.getElementById('total-ateliers'),
    atelierSelect: document.getElementById('atelier-select'),
    toggleBtns: document.querySelectorAll('.toggle-btn'),
    btnExport: document.getElementById('btn-export'),
    resultsTbody: document.getElementById('results-tbody'),
    btnNewProcess: document.getElementById('btn-new-process')
};

// ============================================
// Page Navigation
// ============================================
function navigateTo(pageId) {
    // Hide all pages
    Object.values(elements.pages).forEach(page => {
        page.classList.remove('active');
    });
    // Show target page
    document.getElementById(pageId).classList.add('active');
    AppState.currentPage = pageId;
    updateBreadcrumb();
}

function updateBreadcrumb() {
    const breadcrumbItems = {
        'page-unit-selection': 'Select Unit',
        'page-file-upload': `${UnitConfigs[AppState.selectedUnit]?.name || ''} → Upload Files`,
        'page-processing': 'Processing...',
        'page-results': 'Results'
    };
    elements.breadcrumb.innerHTML = `
        <span class="breadcrumb-item active">${breadcrumbItems[AppState.currentPage]}</span>
    `;
}

// ============================================
// Unit Selection
// ============================================
function initUnitSelection() {
    elements.unitCards.forEach(card => {
        card.addEventListener('click', () => {
            const unit = card.dataset.unit;
            selectUnit(unit);
        });
    });
}

function selectUnit(unit) {
    AppState.selectedUnit = unit;
    AppState.ateliers = UnitConfigs[unit].ateliers;
    elements.selectedUnitName.textContent = UnitConfigs[unit].name;
    
    // Reset file state
    resetFileState();
    
    // Render ateliers list
    renderAteliersList();
    
    // Navigate to file upload
    navigateTo('page-file-upload');
}

// ============================================
// File Upload & Drop Zone
// ============================================
function initDropZone() {
    const dropZone = elements.dropZone;
    const fileInput = elements.fileInput;

    // Click to browse
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    // File input change
    fileInput.addEventListener('change', (e) => {
        handleFiles(Array.from(e.target.files));
    });

    // Drag events
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        
        const files = Array.from(e.dataTransfer.files).filter(file => {
            const ext = file.name.split('.').pop().toLowerCase();
            return ['xlsx', 'xls', 'csv'].includes(ext);
        });
        
        handleFiles(files);
    });

    // Month selection
    elements.monthSelect.addEventListener('change', (e) => {
        AppState.selectedMonth = e.target.value;
    });
}

function handleFiles(files) {
    files.forEach(file => {
        const fileInfo = {
            path: file.path,
            filename: file.name
        };
        
        // Check if it's a stock file
        if (isStockFile(file.name)) {
            AppState.stockFile = fileInfo;
            updateStockFileDisplay();
        } else {
            // Add to dropped files if not already present
            if (!AppState.droppedFiles.some(f => f.path === file.path)) {
                AppState.droppedFiles.push(fileInfo);
            }
        }
    });
    
    // Match files to ateliers
    matchFilesToAteliers();
    updateUI();
}

function isStockFile(filename) {
    const lower = filename.toLowerCase();
    return lower.includes('stock') && !lower.includes('mov');
}

function matchFilesToAteliers() {
    // Reset matching
    AppState.matchedFiles = {};
    AppState.unmatchedFiles = [];
    
    const usedFiles = new Set();
    
    // Try to match each atelier
    AppState.ateliers.forEach(atelier => {
        const keyword = atelier.toLowerCase().trim();
        
        for (const file of AppState.droppedFiles) {
            if (usedFiles.has(file.path)) continue;
            
            const filename = file.filename.toLowerCase();
            
            // Check if keyword matches filename
            if (filename.includes(keyword) || keywordMatch(keyword, filename)) {
                AppState.matchedFiles[atelier] = file;
                usedFiles.add(file.path);
                break;
            }
        }
    });
    
    // Find unmatched files
    AppState.unmatchedFiles = AppState.droppedFiles.filter(
        file => !usedFiles.has(file.path)
    );
}

function keywordMatch(keyword, filename) {
    // Additional fuzzy matching logic
    const keywordParts = keyword.split(/\s+/);
    return keywordParts.every(part => filename.includes(part));
}

// ============================================
// UI Rendering
// ============================================
function renderAteliersList() {
    const container = elements.ateliersList;
    container.innerHTML = '';
    
    AppState.ateliers.forEach(atelier => {
        const matched = AppState.matchedFiles[atelier];
        const isSkipped = AppState.skippedAteliers.has(atelier);
        
        const item = document.createElement('div');
        item.className = `atelier-item${matched ? ' matched' : ''}${isSkipped ? ' skipped' : ''}`;
        item.dataset.atelier = atelier;
        
        item.innerHTML = `
            <div class="atelier-info">
                <span class="atelier-name">${atelier}</span>
                <span class="atelier-file">${matched ? matched.filename : 'No file matched - drag file here or click upload'}</span>
            </div>
            <div class="atelier-actions">
                ${!matched ? `
                    <button class="btn-icon upload" title="Upload file manually" data-action="upload">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17,8 12,3 7,8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                    </button>
                ` : `
                    <button class="btn-icon remove" title="Remove file" data-action="remove">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                        </svg>
                    </button>
                `}
                <button class="btn-icon ${isSkipped ? 'skip-active' : ''}" title="${isSkipped ? 'Include' : 'Skip'}" data-action="skip">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M${isSkipped ? '5 12h14' : '9 18l6-6-6-6'}"/>
                    </svg>
                </button>
            </div>
        `;
        
        // Add event listeners for buttons
        item.querySelectorAll('.btn-icon').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const action = btn.dataset.action;
                handleAtelierAction(atelier, action);
            });
        });
        
        // Enable drop zone for unmatched ateliers
        if (!matched && !isSkipped) {
            item.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                item.classList.add('drag-over');
            });
            
            item.addEventListener('dragleave', () => {
                item.classList.remove('drag-over');
            });
            
            item.addEventListener('drop', (e) => {
                e.preventDefault();
                item.classList.remove('drag-over');
                
                try {
                    const fileData = JSON.parse(e.dataTransfer.getData('application/json'));
                    if (fileData && fileData.path) {
                        // Assign file to this atelier
                        AppState.matchedFiles[atelier] = fileData;
                        // Remove from unmatched files
                        AppState.unmatchedFiles = AppState.unmatchedFiles.filter(f => f.path !== fileData.path);
                        updateUI();
                        showToast(`Assigned "${fileData.filename}" to ${atelier}`, 'success');
                    }
                } catch (err) {
                    console.error('Drop error:', err);
                }
            });
        }
        
        container.appendChild(item);
    });
}

function renderUnmatchedFiles() {
    const container = elements.unmatchedList;
    
    // Get unmatched ateliers (ateliers without files)
    const unmatchedAteliers = AppState.ateliers.filter(a => 
        !AppState.matchedFiles[a] && !AppState.skippedAteliers.has(a)
    );
    
    if (AppState.unmatchedFiles.length === 0 && unmatchedAteliers.length === 0) {
        container.innerHTML = '<p class="placeholder-text">All files matched successfully</p>';
        return;
    }
    
    container.innerHTML = '';
    
    // Show unmatched files with ability to drag to ateliers
    if (AppState.unmatchedFiles.length > 0) {
        const filesHeader = document.createElement('div');
        filesHeader.className = 'unmatched-header';
        filesHeader.innerHTML = '<strong>Unmatched Files (drag to atelier):</strong>';
        container.appendChild(filesHeader);
        
        AppState.unmatchedFiles.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-item draggable';
            item.draggable = true;
            item.dataset.fileIndex = index;
            item.innerHTML = `
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14,2 14,8 20,8"/>
                </svg>
                <span title="${file.filename}">${file.filename}</span>
                <button class="btn-icon remove" title="Remove file" data-file-index="${index}">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <line x1="18" y1="6" x2="6" y2="18"/>
                        <line x1="6" y1="6" x2="18" y2="18"/>
                    </svg>
                </button>
            `;
            
            // Drag start
            item.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('application/json', JSON.stringify(file));
                e.dataTransfer.effectAllowed = 'move';
                item.classList.add('dragging');
            });
            
            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
            });
            
            // Remove button
            item.querySelector('.btn-icon.remove').addEventListener('click', (e) => {
                e.stopPropagation();
                AppState.unmatchedFiles.splice(index, 1);
                AppState.droppedFiles = AppState.droppedFiles.filter(f => f.path !== file.path);
                updateUI();
            });
            
            container.appendChild(item);
        });
    }
    
    // Show quick-assign dropdown for unmatched ateliers
    if (unmatchedAteliers.length > 0 && AppState.unmatchedFiles.length > 0) {
        const assignSection = document.createElement('div');
        assignSection.className = 'quick-assign-section';
        assignSection.innerHTML = `
            <div class="unmatched-header" style="margin-top: 16px;"><strong>Quick Assign:</strong></div>
            <p class="placeholder-text">Drop files on ateliers in left column, or use upload buttons</p>
        `;
        container.appendChild(assignSection);
    }
}

function updateStockFileDisplay() {
    const container = elements.stockFileDisplay;
    
    if (AppState.stockFile) {
        container.innerHTML = `
            <div class="file-item" style="background: var(--success-50); cursor: pointer;" id="stock-file-item">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                    <polyline points="22,4 12,14.01 9,11.01"/>
                </svg>
                <span>${AppState.stockFile.filename}</span>
                <button class="btn-icon remove" title="Change stock file" id="btn-change-stock">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="17,8 12,3 7,8"/>
                        <line x1="12" y1="3" x2="12" y2="15"/>
                    </svg>
                </button>
            </div>
        `;
        elements.stockStatus.textContent = 'Loaded';
        elements.stockStatus.className = 'status-badge success';
        
        // Add change stock handler
        document.getElementById('btn-change-stock').addEventListener('click', async (e) => {
            e.stopPropagation();
            await loadStockFileManually();
        });
    } else {
        container.innerHTML = `
            <div class="stock-upload-zone" id="stock-upload-zone">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 32px; height: 32px; margin-bottom: 8px;">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17,8 12,3 7,8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                </svg>
                <p style="margin: 0; font-weight: 500;">Click to load Stock file</p>
                <p class="placeholder-text" style="margin: 4px 0 0 0;">Or drop it in the main drop zone</p>
            </div>
        `;
        elements.stockStatus.textContent = 'Not Loaded';
        elements.stockStatus.className = 'status-badge warning';
        
        // Add click handler for stock upload zone
        document.getElementById('stock-upload-zone').addEventListener('click', loadStockFileManually);
    }
}

async function loadStockFileManually() {
    const files = await window.electronAPI.openFileDialog({ multiple: false });
    if (files && files.length > 0) {
        const filename = files[0].split(/[/\\]/).pop();
        AppState.stockFile = {
            path: files[0],
            filename: filename
        };
        updateStockFileDisplay();
        validateProcessButton();
        showToast('Stock file loaded successfully', 'success');
    }
}

function updateUI() {
    renderAteliersList();
    renderUnmatchedFiles();
    updateCounts();
    validateProcessButton();
}

function updateCounts() {
    const matchedCount = Object.keys(AppState.matchedFiles).length;
    const unmatchedCount = AppState.unmatchedFiles.length;
    
    elements.matchedCount.textContent = `${matchedCount} matched`;
    elements.unmatchedCount.textContent = `${unmatchedCount} files`;
}

function validateProcessButton() {
    // At least stock file and 1 movement file required
    const hasStock = AppState.stockFile !== null;
    const hasAtLeastOneFile = Object.keys(AppState.matchedFiles).length > 0;
    
    elements.btnProcess.disabled = !(hasStock && hasAtLeastOneFile);
}

// ============================================
// Atelier Actions
// ============================================
async function handleAtelierAction(atelier, action) {
    switch (action) {
        case 'upload':
            const files = await window.electronAPI.openFileDialog({ multiple: false });
            if (files && files.length > 0) {
                const filename = files[0].split(/[/\\]/).pop();
                AppState.matchedFiles[atelier] = {
                    path: files[0],
                    filename: filename
                };
                updateUI();
            }
            break;
            
        case 'remove':
            const removedFile = AppState.matchedFiles[atelier];
            delete AppState.matchedFiles[atelier];
            // Add back to unmatched
            if (removedFile) {
                AppState.unmatchedFiles.push(removedFile);
            }
            updateUI();
            break;
            
        case 'skip':
            if (AppState.skippedAteliers.has(atelier)) {
                AppState.skippedAteliers.delete(atelier);
            } else {
                AppState.skippedAteliers.add(atelier);
            }
            renderAteliersList();
            break;
    }
}

// ============================================
// File Processing
// ============================================
async function processFiles() {
    navigateTo('page-processing');
    
    try {
        // Prepare matched files (exclude skipped)
        const filesToProcess = {};
        for (const [atelier, file] of Object.entries(AppState.matchedFiles)) {
            if (!AppState.skippedAteliers.has(atelier)) {
                filesToProcess[atelier] = file;
            }
        }
        
        updateProgress(10, 'Loading files...');
        
        // Call Python processor
        const result = await window.electronAPI.processFiles(
            AppState.selectedUnit,
            AppState.stockFile,
            filesToProcess,
            AppState.selectedMonth
        );
        
        updateProgress(100, 'Complete!');
        
        // Extract results from response
        console.log('Raw processor result:', result);
        if (result && result.success && result.results) {
            AppState.results = result.results;
        } else if (result && !result.success) {
            throw new Error(result.error || 'Unknown processing error');
        } else {
            // Assume result is already the results object
            AppState.results = result;
        }
        
        // Short delay then show results
        setTimeout(() => {
            showResults();
        }, 500);
        
    } catch (error) {
        showToast(`Error: ${error}`, 'error');
        navigateTo('page-file-upload');
    }
}

function updateProgress(percent, status) {
    elements.progressFill.style.width = `${percent}%`;
    elements.processingStatus.textContent = status;
}

// ============================================
// Results Display
// ============================================
function showResults() {
    navigateTo('page-results');
    
    console.log('Results received:', AppState.results);
    
    if (!AppState.results) {
        showToast('No results received from processor', 'error');
        return;
    }
    
    const results = AppState.results;
    
    // Check for errors in results
    for (const [atelier, data] of Object.entries(results)) {
        if (data.error) {
            console.error(`Error in ${atelier}:`, data.error);
        }
    }
    
    // Update summary
    let totalMatches = 0;
    let totalDiscrepancies = 0;
    
    for (const atelier of Object.keys(results)) {
        const matchCount = results[atelier].matches?.length || 0;
        const discCount = results[atelier].discrepancies?.length || 0;
        console.log(`${atelier}: ${matchCount} matches, ${discCount} discrepancies`);
        totalMatches += matchCount;
        totalDiscrepancies += discCount;
    }
    
    elements.totalMatches.textContent = totalMatches;
    elements.totalDiscrepancies.textContent = totalDiscrepancies;
    elements.totalAteliers.textContent = Object.keys(results).length;
    
    // Populate atelier dropdown
    elements.atelierSelect.innerHTML = '<option value="">Select an atelier...</option>';
    for (const atelier of Object.keys(results)) {
        const option = document.createElement('option');
        option.value = atelier;
        const matchCount = results[atelier].matches?.length || 0;
        const discCount = results[atelier].discrepancies?.length || 0;
        option.textContent = `${atelier} (${matchCount} / ${discCount})`;
        elements.atelierSelect.appendChild(option);
    }
    
    // Auto-select first atelier
    if (Object.keys(results).length > 0) {
        elements.atelierSelect.value = Object.keys(results)[0];
        renderResultsTable();
    }
}

function renderResultsTable() {
    const atelier = elements.atelierSelect.value;
    const viewType = document.querySelector('.toggle-btn.active').dataset.view;
    
    if (!atelier || !AppState.results || !AppState.results[atelier]) {
        elements.resultsTbody.innerHTML = `
            <tr class="empty-row">
                <td colspan="4">Select an atelier to view results</td>
            </tr>
        `;
        return;
    }
    
    const atelierData = AppState.results[atelier];
    
    // Check for error
    if (atelierData.error) {
        elements.resultsTbody.innerHTML = `
            <tr class="empty-row error-row">
                <td colspan="4" style="color: var(--error-600);">
                    <strong>Error:</strong> ${atelierData.error}
                </td>
            </tr>
        `;
        return;
    }
    
    const data = atelierData[viewType] || [];
    
    if (data.length === 0) {
        elements.resultsTbody.innerHTML = `
            <tr class="empty-row">
                <td colspan="4">No ${viewType} found for this atelier</td>
            </tr>
        `;
        return;
    }
    
    elements.resultsTbody.innerHTML = data.map(row => {
        const diff = parseFloat(row.Difference) || 0;
        const rowClass = diff > 0 ? 'positive' : diff < 0 ? 'negative' : '';
        return `
            <tr class="${rowClass}">
                <td>${row.Ref || ''}</td>
                <td>${formatNumber(row.Stock_Qty)}</td>
                <td>${formatNumber(row.Calc_Mov_Qty)}</td>
                <td>${formatNumber(row.Difference)}</td>
            </tr>
        `;
    }).join('');
}

function formatNumber(val) {
    if (val === null || val === undefined || val === '') return '0';
    const num = parseFloat(val);
    return isNaN(num) ? '0' : num.toFixed(2);
}

// ============================================
// Export
// ============================================
async function exportResults() {
    const atelier = elements.atelierSelect.value;
    const viewType = document.querySelector('.toggle-btn.active').dataset.view;
    
    if (!atelier || !AppState.results || !AppState.results[atelier]) {
        showToast('Please select an atelier first', 'warning');
        return;
    }
    
    const data = AppState.results[atelier][viewType] || [];
    
    if (data.length === 0) {
        showToast('No data to export', 'warning');
        return;
    }
    
    const defaultName = `${viewType}_${atelier.replace(/\s+/g, '_')}.csv`;
    const filePath = await window.electronAPI.saveFileDialog(defaultName);
    
    if (filePath) {
        try {
            await window.electronAPI.exportCSV(data, filePath);
            showToast('Export successful!', 'success');
        } catch (error) {
            showToast(`Export failed: ${error}`, 'error');
        }
    }
}

// ============================================
// Reset State
// ============================================
function resetFileState() {
    AppState.stockFile = null;
    AppState.droppedFiles = [];
    AppState.matchedFiles = {};
    AppState.unmatchedFiles = [];
    AppState.skippedAteliers.clear();
    AppState.results = null;
    
    updateStockFileDisplay();
    elements.fileInput.value = '';
}

// ============================================
// Toast Notifications
// ============================================
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 4000);
}

// ============================================
// Event Listeners
// ============================================
function initEventListeners() {
    // Navigation
    elements.btnBackToUnits.addEventListener('click', () => {
        navigateTo('page-unit-selection');
    });
    
    elements.btnBackToUpload.addEventListener('click', () => {
        navigateTo('page-file-upload');
    });
    
    // File actions
    elements.btnClearFiles.addEventListener('click', () => {
        resetFileState();
        updateUI();
    });
    
    elements.btnProcess.addEventListener('click', processFiles);
    
    // Results
    elements.atelierSelect.addEventListener('change', renderResultsTable);
    
    elements.toggleBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            elements.toggleBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderResultsTable();
        });
    });
    
    elements.btnExport.addEventListener('click', exportResults);
    
    elements.btnNewProcess.addEventListener('click', () => {
        resetFileState();
        navigateTo('page-unit-selection');
    });
}

// ============================================
// Initialization
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    initUnitSelection();
    initDropZone();
    initEventListeners();
});
