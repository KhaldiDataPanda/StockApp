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
    prevStockFile: null,
    droppedFiles: [],
    matchedFiles: {},      // { atelier: { path, filename } }
    unmatchedFiles: [],    // [{ path, filename }]
    ateliers: [],          // List of ateliers for selected unit
    skippedAteliers: new Set(),
    results: null,         // Processing results
    verificationResults: null, // Verification results
    fileOverrides: {},     // { atelier: { sheetName, refCol, qtyCol } }
    currentSortColumn: null,
    currentSortDirection: 'asc',
    searchQuery: '',
    // Report Generation State
    reportData: {},           // { atelierName: discrepanciesArray }
    // Verify ± State
    verifyMode: false,        // true when in "Verify ±" mode
    oppositeMatches: [],      // Array of matched opposite pairs
    selectedForElimination: new Set(), // Set of row refs selected for elimination
    contextMenuRowRef: null,  // Reference of row for context menu
    contextMenuAtelier: null,
    contextMenuViewType: null
};

function escapeHtml(text) {
    return (text ?? '').toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function highlightMatchSafe(text, query) {
    const raw = (text ?? '').toString();
    if (!query) return escapeHtml(raw);
    const safeQuery = escapeRegex(query);
    const regex = new RegExp(safeQuery, 'gi');

    let result = '';
    let lastIndex = 0;
    for (const match of raw.matchAll(regex)) {
        const start = match.index ?? 0;
        const end = start + match[0].length;
        result += escapeHtml(raw.slice(lastIndex, start));
        result += '<mark style="background: var(--warning-200); padding: 0 2px; border-radius: 2px;">' + escapeHtml(raw.slice(start, end)) + '</mark>';
        lastIndex = end;
    }
    result += escapeHtml(raw.slice(lastIndex));
    return result;
}


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
    },
    Fath3: {
        name: 'Fath 3',
        ateliers: ['pet', 'triage']
    },
    Mdoukal: {
        name: "M'doukal",
        ateliers: ['couture femmes', 'orielles', 'magasin']
    },
    Mags: {
        name: 'Mags',
        ateliers: ['magz']
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
    btnVerify: document.getElementById('btn-verify'),
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
    btnExportExcel: document.getElementById('btn-export-excel'),
    resultsSearch: document.getElementById('results-search'),
    resultsTbody: document.getElementById('results-tbody'),
    btnNewProcess: document.getElementById('btn-new-process'),
    // Verification Modal
    verificationModal: document.getElementById('verification-modal'),
    closeVerificationModal: document.getElementById('close-verification-modal'),
    verificationStatus: document.getElementById('verification-status'),
    verificationResults: document.getElementById('verification-results'),
    btnCancelVerification: document.getElementById('btn-cancel-verification'),
    btnApplyFixes: document.getElementById('btn-apply-fixes'),
    // Report Features
    reportToolbar: document.getElementById('report-toolbar'),
    btnGenerateReport: document.getElementById('btn-generate-report'),
    reportCount: document.getElementById('report-count'),
    btnClearReport: document.getElementById('btn-clear-report'),
    btnVerifyOpposite: document.getElementById('btn-verify-opposite'),
    btnInsertReport: document.getElementById('btn-insert-report'),
    // Report Modal
    reportModal: document.getElementById('report-modal'),
    closeReportModal: document.getElementById('close-report-modal'),
    reportEditor: document.getElementById('report-editor'),
    btnCancelReport: document.getElementById('btn-cancel-report'),
    btnExportMd: document.getElementById('btn-export-md'),
    btnExportPdf: document.getElementById('btn-export-pdf'),
    // Context Menu
    contextMenu: document.getElementById('context-menu'),
    ctxEditRow: document.getElementById('ctx-edit-row'),
    ctxDeleteRow: document.getElementById('ctx-delete-row')
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
            if (AppState.selectedUnit === 'Mags') {
                assignMagsStockFile(fileInfo);
            } else {
                AppState.stockFile = fileInfo;
            }
            updateStockFileDisplay();

            // Auto-detect month from stock filename
            const parsed = parseStockMonthYear(file.name);
            if (parsed && elements.monthSelect) {
                AppState.selectedMonth = parsed.month;
                elements.monthSelect.value = parsed.month;
                showToast(`Month auto-set to ${getMonthName(parsed.month)}`, 'info');
            }
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

// Helper to get month name
function getMonthName(monthStr) {
    const months = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    return months[parseInt(monthStr, 10) - 1] || monthStr;
}

function isStockFile(filename) {
    const lower = filename.toLowerCase();
    return lower.includes('stock') && !lower.includes('mov');
}

function parseStockMonthYear(filename) {
    const match = filename.match(/stock\s*(\d{1,2})\s*[-_\s]\s*(\d{4})/i);
    if (!match) return null;
    const monthNum = parseInt(match[1], 10);
    const yearNum = parseInt(match[2], 10);
    if (!monthNum || monthNum < 1 || monthNum > 12 || !yearNum) return null;
    return {
        month: String(monthNum).padStart(2, '0'),
        year: yearNum
    };
}

function prevMonthFrom(monthStr) {
    const m = parseInt(monthStr, 10);
    if (!m || m < 1 || m > 12) return null;
    return m === 1 ? '12' : String(m - 1).padStart(2, '0');
}

function assignMagsStockFile(fileInfo) {
    const meta = parseStockMonthYear(fileInfo.filename);
    const selectedMonth = AppState.selectedMonth;
    const prevMonth = prevMonthFrom(selectedMonth);

    // Best effort assignment based on filename month
    if (meta && meta.month === selectedMonth) {
        AppState.stockFile = fileInfo;
        return;
    }
    if (meta && prevMonth && meta.month === prevMonth) {
        AppState.prevStockFile = fileInfo;
        return;
    }

    // Fallback: fill current first, then previous
    if (!AppState.stockFile || AppState.stockFile.path === fileInfo.path) {
        AppState.stockFile = fileInfo;
        return;
    }
    if (!AppState.prevStockFile || AppState.prevStockFile.path === fileInfo.path) {
        AppState.prevStockFile = fileInfo;
        return;
    }
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

    if (AppState.selectedUnit === 'Mags') {
        const current = AppState.stockFile;
        const prev = AppState.prevStockFile;

        container.innerHTML = `
            <div style="display: flex; gap: 12px; align-items: flex-start;">
                <div style="flex: 1; min-width: 260px;">
                    <div style="font-size: 12px; color: var(--gray-600); margin-bottom: 6px;">Current month stock</div>
                    ${current ? `
                        <div class="file-item" style="background: var(--success-50); cursor: pointer;">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                                <polyline points="22,4 12,14.01 9,11.01"/>
                            </svg>
                            <span>${current.filename}</span>
                            <button class="btn-icon remove" title="Change current stock file" id="btn-change-stock">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                    <polyline points="17,8 12,3 7,8"/>
                                    <line x1="12" y1="3" x2="12" y2="15"/>
                                </svg>
                            </button>
                        </div>
                    ` : `
                        <div class="stock-upload-zone" id="stock-upload-zone">
                            <p style="margin: 0; font-weight: 500;">Click to load current Stock file</p>
                            <p class="placeholder-text" style="margin: 4px 0 0 0;">Or drop it in the main drop zone</p>
                        </div>
                    `}
                </div>

                <div style="flex: 1; min-width: 260px;">
                    <div style="font-size: 12px; color: var(--gray-600); margin-bottom: 6px;">Previous month stock (opening inventory)</div>
                    ${prev ? `
                        <div class="file-item" style="background: var(--success-50); cursor: pointer;">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                                <polyline points="22,4 12,14.01 9,11.01"/>
                            </svg>
                            <span>${prev.filename}</span>
                            <button class="btn-icon remove" title="Change previous stock file" id="btn-change-prev-stock">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                    <polyline points="17,8 12,3 7,8"/>
                                    <line x1="12" y1="3" x2="12" y2="15"/>
                                </svg>
                            </button>
                        </div>
                    ` : `
                        <div class="stock-upload-zone" id="prev-stock-upload-zone">
                            <p style="margin: 0; font-weight: 500;">Click to load previous Stock file</p>
                            <p class="placeholder-text" style="margin: 4px 0 0 0;">Or drop it in the main drop zone</p>
                        </div>
                    `}
                </div>
            </div>
        `;

        const hasCurrent = !!AppState.stockFile;
        if (hasCurrent) {
            elements.stockStatus.textContent = 'Loaded';
            elements.stockStatus.className = 'status-badge success';
        } else {
            elements.stockStatus.textContent = 'Not Loaded';
            elements.stockStatus.className = 'status-badge warning';
        }

        // Wire buttons/zones
        const btnChangeCurrent = document.getElementById('btn-change-stock');
        if (btnChangeCurrent) {
            btnChangeCurrent.addEventListener('click', async (e) => {
                e.stopPropagation();
                await loadStockFileManually();
            });
        }

        const btnChangePrev = document.getElementById('btn-change-prev-stock');
        if (btnChangePrev) {
            btnChangePrev.addEventListener('click', async (e) => {
                e.stopPropagation();
                await loadPrevStockFileManually();
            });
        }

        const currentZone = document.getElementById('stock-upload-zone');
        if (currentZone) {
            currentZone.addEventListener('click', loadStockFileManually);
        }

        const prevZone = document.getElementById('prev-stock-upload-zone');
        if (prevZone) {
            prevZone.addEventListener('click', loadPrevStockFileManually);
        }

        return;
    }

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
        const fileInfo = {
            path: files[0],
            filename: filename
        };
        if (AppState.selectedUnit === 'Mags') {
            AppState.stockFile = fileInfo;
        } else {
            AppState.stockFile = fileInfo;
        }
        updateStockFileDisplay();
        validateProcessButton();
        showToast('Stock file loaded successfully', 'success');
    }
}

async function loadPrevStockFileManually() {
    const files = await window.electronAPI.openFileDialog({ multiple: false });
    if (files && files.length > 0) {
        const filename = files[0].split(/[/\\]/).pop();
        AppState.prevStockFile = {
            path: files[0],
            filename: filename
        };
        updateStockFileDisplay();
        validateProcessButton();
        showToast('Previous stock file loaded successfully', 'success');
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
    let hasStock = AppState.stockFile !== null;
    const hasAtLeastOneFile = Object.keys(AppState.matchedFiles).length > 0;

    elements.btnProcess.disabled = !(hasStock && hasAtLeastOneFile);
    elements.btnVerify.disabled = !(hasStock && hasAtLeastOneFile);
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
// File Verification
// ============================================
async function verifyFiles() {
    // Show modal
    elements.verificationModal.classList.remove('hidden');
    elements.verificationStatus.classList.remove('hidden');
    elements.verificationResults.classList.add('hidden');
    elements.btnApplyFixes.disabled = true;

    try {
        // Prepare matched files (exclude skipped)
        const filesToVerify = {};
        for (const [atelier, file] of Object.entries(AppState.matchedFiles)) {
            if (!AppState.skippedAteliers.has(atelier)) {
                filesToVerify[atelier] = file;
            }
        }

        // Call Python verifier
        const result = await window.electronAPI.verifyFiles(
            AppState.selectedUnit,
            filesToVerify
        );

        if (result && result.success && result.verification) {
            AppState.verificationResults = result.verification;
            renderVerificationResults(result.verification);
        } else {
            throw new Error(result?.error || 'Verification failed');
        }

    } catch (error) {
        showToast(`Verification error: ${error}`, 'error');
        closeVerificationModal();
    }
}

function renderVerificationResults(verification) {
    elements.verificationStatus.classList.add('hidden');
    elements.verificationResults.classList.remove('hidden');

    const container = elements.verificationResults;
    container.innerHTML = '';

    const ateliers = Object.keys(verification);
    let hasErrors = false;

    // Check if all files are valid
    const allValid = ateliers.every(atelier => verification[atelier].valid);

    if (allValid) {
        container.innerHTML = `
            <div class="all-valid-message">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                    <polyline points="22,4 12,14.01 9,11.01"/>
                </svg>
                <h3>All Files Verified Successfully!</h3>
                <p>All sheet names and column configurations are correct. You can proceed with processing.</p>
            </div>
        `;
        elements.btnApplyFixes.disabled = false;
        elements.btnApplyFixes.textContent = 'Continue to Process';
        return;
    }

    elements.btnApplyFixes.textContent = 'Apply & Continue';

    ateliers.forEach(atelier => {
        const data = verification[atelier];
        const fileDiv = document.createElement('div');
        fileDiv.className = `verification-file ${data.valid ? 'valid' : 'has-errors'}`;
        fileDiv.dataset.atelier = atelier;

        const statusIcon = data.valid
            ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22,4 12,14.01 9,11.01"/></svg>'
            : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';

        fileDiv.innerHTML = `
            <div class="verification-file-header">
                <div class="file-info">
                    ${statusIcon}
                    <span class="file-name">${data.filename}</span>
                    <span class="atelier-name">(${atelier})</span>
                </div>
                <span class="status-indicator">${data.valid ? 'Valid' : data.errors.length + ' Issue(s)'}</span>
            </div>
            <div class="verification-file-body">
                <div class="verification-field">
                    <label>Sheet Name:</label>
                    <select class="sheet-select" data-atelier="${atelier}">
                        ${data.availableSheets.map(sheet =>
            `<option value="${sheet}" ${sheet === data.expectedSheet ? 'selected' : ''}>${sheet}</option>`
        ).join('')}
                    </select>
                    <div class="field-status ${data.availableSheets.includes(data.expectedSheet) ? 'valid' : 'error'}">
                        ${data.availableSheets.includes(data.expectedSheet)
                ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20,6 9,17 4,12"/></svg>'
                : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>'}
                    </div>
                </div>
                ${!data.availableSheets.includes(data.expectedSheet) ? `
                    <div class="available-options">Expected: "${data.expectedSheet}"</div>
                ` : ''}
                
                <div class="verification-field">
                    <label>Reference Column:</label>
                    <select class="ref-col-select" data-atelier="${atelier}">
                        <option value="">-- Select Column --</option>
                        ${data.availableColumns.map(col =>
                    `<option value="${col}" ${col === data.detectedRefCol ? 'selected' : ''}>${col}</option>`
                ).join('')}
                    </select>
                    <div class="field-status ${data.detectedRefCol ? 'valid' : 'error'}">
                        ${data.detectedRefCol
                ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20,6 9,17 4,12"/></svg>'
                : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>'}
                    </div>
                </div>
                
                <div class="verification-field">
                    <label>Quantity Column:</label>
                    <select class="qty-col-select" data-atelier="${atelier}">
                        <option value="">-- Select Column --</option>
                        ${data.availableColumns.map(col =>
                    `<option value="${col}" ${col === data.detectedQtyCol ? 'selected' : ''}>${col}</option>`
                ).join('')}
                    </select>
                    <div class="field-status ${data.detectedQtyCol ? 'valid' : 'error'}">
                        ${data.detectedQtyCol
                ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20,6 9,17 4,12"/></svg>'
                : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>'}
                    </div>
                </div>
            </div>
        `;

        container.appendChild(fileDiv);

        if (!data.valid) hasErrors = true;
    });

    // Enable apply button - user can apply fixes
    elements.btnApplyFixes.disabled = false;

    // Add change listeners to update overrides
    container.querySelectorAll('select').forEach(select => {
        select.addEventListener('change', updateOverridesFromUI);
    });
}

function updateOverridesFromUI() {
    const container = elements.verificationResults;

    container.querySelectorAll('.verification-file').forEach(fileDiv => {
        const atelier = fileDiv.dataset.atelier;
        const sheetSelect = fileDiv.querySelector('.sheet-select');
        const refColSelect = fileDiv.querySelector('.ref-col-select');
        const qtyColSelect = fileDiv.querySelector('.qty-col-select');

        AppState.fileOverrides[atelier] = {
            sheetName: sheetSelect?.value || '',
            refCol: refColSelect?.value || '',
            qtyCol: qtyColSelect?.value || ''
        };
    });
}

function applyFixesAndProcess() {
    // Collect all overrides from the UI
    updateOverridesFromUI();

    // Close modal
    closeVerificationModal();

    // Start processing with overrides
    processFiles();
}

function closeVerificationModal() {
    elements.verificationModal.classList.add('hidden');
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

        // Prepare overrides if any
        const overrides = Object.keys(AppState.fileOverrides).length > 0 ? AppState.fileOverrides : null;

        // Call Python processor
        const stockPayload = (AppState.selectedUnit === 'Mags' && AppState.stockFile)
            ? {
                ...AppState.stockFile,
                prevPath: AppState.prevStockFile?.path,
                prevFilename: AppState.prevStockFile?.filename
            }
            : AppState.stockFile;

        const result = await window.electronAPI.processFiles(
            AppState.selectedUnit,
            stockPayload,
            filesToProcess,
            AppState.selectedMonth,
            overrides
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

    let data = [...(atelierData[viewType] || [])];

    // Apply search filter
    if (AppState.searchQuery) {
        const query = AppState.searchQuery.toLowerCase();
        data = data.filter(row => {
            const ref = (row.Ref || '').toString().toLowerCase();
            return ref.includes(query);
        });
    }

    // Apply sorting
    if (AppState.currentSortColumn) {
        data.sort((a, b) => {
            let aVal = a[AppState.currentSortColumn];
            let bVal = b[AppState.currentSortColumn];

            // Handle numeric values
            if (typeof aVal === 'number' || typeof bVal === 'number') {
                aVal = parseFloat(aVal) || 0;
                bVal = parseFloat(bVal) || 0;
            } else {
                aVal = (aVal || '').toString().toLowerCase();
                bVal = (bVal || '').toString().toLowerCase();
            }

            if (aVal < bVal) return AppState.currentSortDirection === 'asc' ? -1 : 1;
            if (aVal > bVal) return AppState.currentSortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }

    if (data.length === 0) {
        const message = AppState.searchQuery
            ? `No results matching "${AppState.searchQuery}"`
            : `No ${viewType} found for this atelier`;
        elements.resultsTbody.innerHTML = `
            <tr class="empty-row">
                <td colspan="${AppState.verifyMode ? 5 : 4}">${message}</td>
            </tr>
        `;
        return;
    }

    // Build highlighted refs set for verify mode
    const highlightedRefs = new Map(); // ref -> 'high' or 'low'
    if (AppState.verifyMode && viewType === 'discrepancies') {
        AppState.oppositeMatches.forEach(match => {
            const level = match.highSimilarity ? 'high' : 'low';
            highlightedRefs.set(match.ref1, level);
            highlightedRefs.set(match.ref2, level);
        });
    }

    elements.resultsTbody.innerHTML = data.map(row => {
        const diff = parseFloat(row.Difference) || 0;
        let rowClass = diff > 0 ? 'positive' : diff < 0 ? 'negative' : '';
        const refHighlight = highlightMatchSafe(row.Ref || '', AppState.searchQuery);

        // Add highlighting classes in verify mode
        const highlightLevel = highlightedRefs.get(row.Ref);
        if (highlightLevel === 'high') {
            rowClass += ' opposite-high-similarity';
        } else if (highlightLevel === 'low') {
            rowClass += ' opposite-low-similarity';
        }

        // Check if selected for elimination
        const isSelected = AppState.selectedForElimination.has(row.Ref);
        if (isSelected) {
            rowClass += ' marked-for-deletion';
        }

        // Checkbox column for verify mode
        let checkboxCell = '';
        if (AppState.verifyMode && viewType === 'discrepancies') {
            if (highlightLevel) {
                const checked = isSelected ? 'checked' : '';
                checkboxCell = `<td class="checkbox-cell">
                    <input class="row-select-checkbox" type="checkbox" ${checked}>
                </td>`;
            } else {
                checkboxCell = '<td class="checkbox-cell"></td>';
            }
        }

        const rowRefAttr = escapeHtml(row.Ref || '');

        return `
            <tr class="${rowClass}" data-ref="${rowRefAttr}">
                ${checkboxCell}
                <td>${refHighlight}</td>
                <td>${formatNumber(row.Stock_Qty)}</td>
                <td>${formatNumber(row.Calc_Mov_Qty)}</td>
                <td>${formatNumber(row.Difference)}</td>
            </tr>
        `;
    }).join('');

    // Update table header for checkbox column
    const thead = document.querySelector('.results-table thead tr');
    const hasCheckboxHeader = thead.querySelector('.checkbox-header');
    if (AppState.verifyMode && viewType === 'discrepancies' && !hasCheckboxHeader) {
        thead.insertAdjacentHTML('afterbegin', '<th class="checkbox-header"><input type="checkbox" id="verify-select-all" title="Check all"></th>');
    } else if ((!AppState.verifyMode || viewType !== 'discrepancies') && hasCheckboxHeader) {
        hasCheckboxHeader.remove();
    }

    updateVerifySelectAllCheckboxState();

    // Update sort indicators
    updateSortIndicators();
}

function updateVerifySelectAllCheckboxState() {
    const selectAll = document.getElementById('verify-select-all');
    if (!selectAll) return;

    if (!AppState.verifyMode) {
        selectAll.checked = false;
        selectAll.indeterminate = false;
        return;
    }

    const all = getAllHighlightedOppositeRefs();
    if (all.length === 0) {
        selectAll.checked = false;
        selectAll.indeterminate = false;
        selectAll.disabled = true;
        return;
    }

    selectAll.disabled = false;
    const selectedCount = all.reduce((count, ref) => count + (AppState.selectedForElimination.has(ref) ? 1 : 0), 0);
    selectAll.checked = selectedCount === all.length;
    selectAll.indeterminate = selectedCount > 0 && selectedCount < all.length;
}

function getAllHighlightedOppositeRefs() {
    const uniq = new Set();
    AppState.oppositeMatches.forEach(m => {
        if (m?.ref1) uniq.add(m.ref1);
        if (m?.ref2) uniq.add(m.ref2);
    });
    return Array.from(uniq);
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function handleSort(column) {
    if (AppState.currentSortColumn === column) {
        // Toggle direction
        AppState.currentSortDirection = AppState.currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        AppState.currentSortColumn = column;
        AppState.currentSortDirection = 'asc';
    }
    renderResultsTable();
}

function updateSortIndicators() {
    document.querySelectorAll('.results-table th.sortable').forEach(th => {
        th.classList.remove('asc', 'desc');
        if (th.dataset.column === AppState.currentSortColumn) {
            th.classList.add(AppState.currentSortDirection);
        }
    });
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

async function exportResultsExcel() {
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

    const defaultName = `${viewType}_${atelier.replace(/\s+/g, '_')}.xlsx`;
    const filePath = await window.electronAPI.saveExcelDialog(defaultName);

    if (filePath) {
        try {
            await window.electronAPI.exportExcel(data, filePath);
            showToast('Excel export successful!', 'success');
        } catch (error) {
            showToast(`Excel export failed: ${error}`, 'error');
        }
    }
}

// ============================================
// Reset State
// ============================================
function resetFileState() {
    AppState.stockFile = null;
    AppState.prevStockFile = null;
    AppState.droppedFiles = [];
    AppState.matchedFiles = {};
    AppState.unmatchedFiles = [];
    AppState.skippedAteliers.clear();
    AppState.results = null;
    AppState.verificationResults = null;
    AppState.fileOverrides = {};
    AppState.currentSortColumn = null;
    AppState.currentSortDirection = 'asc';
    AppState.searchQuery = '';

    updateStockFileDisplay();
    elements.fileInput.value = '';
    if (elements.resultsSearch) {
        elements.resultsSearch.value = '';
    }
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
// String Similarity (Levenshtein Distance)
// ============================================
function levenshteinSimilarity(str1, str2) {
    const s1 = (str1 || '').toString().toLowerCase().trim();
    const s2 = (str2 || '').toString().toLowerCase().trim();

    if (s1 === s2) return 1;
    if (s1.length === 0 || s2.length === 0) return 0;

    const matrix = [];
    for (let i = 0; i <= s1.length; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= s2.length; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= s1.length; i++) {
        for (let j = 1; j <= s2.length; j++) {
            const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
            matrix[i][j] = Math.min(
                matrix[i - 1][j] + 1,
                matrix[i][j - 1] + 1,
                matrix[i - 1][j - 1] + cost
            );
        }
    }

    const maxLen = Math.max(s1.length, s2.length);
    return 1 - (matrix[s1.length][s2.length] / maxLen);
}

// ============================================
// Verify Opposite Discrepancies
// ============================================
function findOppositeDiscrepancies() {
    const atelier = elements.atelierSelect.value;
    const discrepancies = AppState.results[atelier]?.discrepancies || [];
    const matches = [];

    for (let i = 0; i < discrepancies.length; i++) {
        const diff1 = parseFloat(discrepancies[i].Difference);
        if (diff1 === 0 || isNaN(diff1)) continue;

        for (let j = i + 1; j < discrepancies.length; j++) {
            const diff2 = parseFloat(discrepancies[j].Difference);
            if (isNaN(diff2)) continue;

            // Check if values are opposite (sum to ~0)
            if (Math.abs(diff1 + diff2) < 0.01) {
                const similarity = levenshteinSimilarity(
                    discrepancies[i].Ref,
                    discrepancies[j].Ref
                );

                matches.push({
                    ref1: discrepancies[i].Ref,
                    ref2: discrepancies[j].Ref,
                    diff1: diff1,
                    diff2: diff2,
                    similarity: similarity,
                    highSimilarity: similarity >= 0.8
                });
            }
        }
    }

    return matches;
}

function toggleVerifyMode() {
    const viewType = document.querySelector('.toggle-btn.active').dataset.view;
    if (viewType !== 'discrepancies') {
        showToast('Switch to Discrepancies view first', 'warning');
        return;
    }

    if (!AppState.verifyMode) {
        // Enter verify mode
        AppState.oppositeMatches = findOppositeDiscrepancies();

        if (AppState.oppositeMatches.length === 0) {
            showToast('No opposite discrepancies found', 'info');
            return;
        }

        AppState.verifyMode = true;
        AppState.selectedForElimination.clear();

        // Auto-select high similarity matches
        AppState.oppositeMatches.forEach(match => {
            if (match.highSimilarity) {
                AppState.selectedForElimination.add(match.ref1);
                AppState.selectedForElimination.add(match.ref2);
            }
        });

        elements.btnVerifyOpposite.innerHTML = `
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="3,6 5,6 21,6"/>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
            </svg>
            Eliminate ±
        `;
        elements.btnVerifyOpposite.classList.remove('btn-warning');
        elements.btnVerifyOpposite.classList.add('btn-danger');

        showToast(`Found ${AppState.oppositeMatches.length} opposite pair(s)`, 'success');
    } else {
        // Eliminate selected rows
        eliminateSelectedOpposites();
    }

    renderResultsTable();
}

function eliminateSelectedOpposites() {
    const atelier = elements.atelierSelect.value;
    if (!atelier || !AppState.results[atelier]) return;

    const toRemove = AppState.selectedForElimination;
    if (toRemove.size === 0) {
        showToast('No items selected for elimination', 'warning');
        exitVerifyMode();
        return;
    }

    // Filter out selected references (trim to avoid whitespace mismatches)
    const toRemoveNormalized = new Set(Array.from(toRemove).map(r => (r ?? '').toString().trim()));
    AppState.results[atelier].discrepancies = AppState.results[atelier].discrepancies.filter(
        row => !toRemoveNormalized.has((row.Ref ?? '').toString().trim())
    );

    // Update totals
    updateResultsSummary();

    showToast(`Eliminated ${toRemove.size} items`, 'success');
    exitVerifyMode();
}

function exitVerifyMode() {
    AppState.verifyMode = false;
    AppState.oppositeMatches = [];
    AppState.selectedForElimination.clear();

    elements.btnVerifyOpposite.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
        </svg>
        Verify ±
    `;
    elements.btnVerifyOpposite.classList.remove('btn-danger');
    elements.btnVerifyOpposite.classList.add('btn-warning');

    renderResultsTable();
}

function toggleRowSelection(ref) {
    if (AppState.selectedForElimination.has(ref)) {
        AppState.selectedForElimination.delete(ref);
    } else {
        AppState.selectedForElimination.add(ref);
    }
    renderResultsTable();
}

function updateResultsSummary() {
    if (!AppState.results) return;

    let totalMatches = 0;
    let totalDiscrepancies = 0;

    for (const atelier of Object.keys(AppState.results)) {
        totalMatches += AppState.results[atelier].matches?.length || 0;
        totalDiscrepancies += AppState.results[atelier].discrepancies?.length || 0;
    }

    elements.totalMatches.textContent = totalMatches;
    elements.totalDiscrepancies.textContent = totalDiscrepancies;
}

// ============================================
// Context Menu for Row Deletion
// ============================================
function showContextMenu(event, ref, ctx) {
    event.preventDefault();
    event.stopPropagation();
    AppState.contextMenuRowRef = ref;
    AppState.contextMenuAtelier = ctx?.atelier ?? elements.atelierSelect.value;
    AppState.contextMenuViewType = ctx?.viewType ?? document.querySelector('.toggle-btn.active')?.dataset?.view;

    const menu = elements.contextMenu;
    menu.style.left = `${event.clientX}px`;
    menu.style.top = `${event.clientY}px`;
    menu.classList.remove('hidden');

    // Close menu when clicking outside
    document.addEventListener('click', hideContextMenu);
}

function hideContextMenu() {
    elements.contextMenu.classList.add('hidden');
    document.removeEventListener('click', hideContextMenu);
}

function deleteRowFromContext() {
    const ref = AppState.contextMenuRowRef;
    const atelier = AppState.contextMenuAtelier;
    const viewType = AppState.contextMenuViewType;

    if (!ref || !atelier || !AppState.results[atelier]) return;

    // Remove from the current view (trim to avoid whitespace mismatches)
    const refNorm = (ref ?? '').toString().trim();
    AppState.results[atelier][viewType] = AppState.results[atelier][viewType].filter(
        row => (row.Ref ?? '').toString().trim() !== refNorm
    );

    updateResultsSummary();
    renderResultsTable();
    hideContextMenu();
    showToast('Row deleted', 'success');
}

function editRowFromContext() {
    const oldRef = AppState.contextMenuRowRef;
    const atelier = AppState.contextMenuAtelier;
    const viewType = AppState.contextMenuViewType;

    if (!oldRef || !atelier || !viewType || !AppState.results?.[atelier]?.[viewType]) return;

    const newRefRaw = window.prompt('Edit reference:', oldRef);
    if (newRefRaw === null) {
        hideContextMenu();
        return;
    }

    const newRef = newRefRaw.toString().trim();
    if (!newRef) {
        showToast('Reference cannot be empty', 'warning');
        hideContextMenu();
        return;
    }

    const oldRefNorm = oldRef.toString().trim();

    // Update first matching row in the selected table
    const rows = AppState.results[atelier][viewType];
    const rowToEdit = rows.find(r => (r.Ref ?? '').toString().trim() === oldRefNorm);
    if (!rowToEdit) {
        showToast('Row not found', 'warning');
        hideContextMenu();
        return;
    }
    rowToEdit.Ref = newRef;

    // Keep verify-mode selection consistent
    if (AppState.verifyMode) {
        if (AppState.selectedForElimination.has(oldRef)) {
            AppState.selectedForElimination.delete(oldRef);
            AppState.selectedForElimination.add(newRef);
        }
        AppState.oppositeMatches = AppState.oppositeMatches.map(m => {
            if (!m) return m;
            return {
                ...m,
                ref1: m.ref1 === oldRef ? newRef : m.ref1,
                ref2: m.ref2 === oldRef ? newRef : m.ref2
            };
        });
    }

    updateResultsSummary();
    renderResultsTable();
    hideContextMenu();
    showToast('Reference updated', 'success');
}

// ============================================
// Report Generation
// ============================================
function addToReport() {
    const atelier = elements.atelierSelect.value;
    const viewType = document.querySelector('.toggle-btn.active').dataset.view;

    if (!atelier || viewType !== 'discrepancies') {
        showToast('Select an atelier and switch to Discrepancies view', 'warning');
        return;
    }

    const discrepancies = AppState.results[atelier]?.discrepancies || [];
    if (discrepancies.length === 0) {
        showToast('No discrepancies to add', 'warning');
        return;
    }

    // Deep copy the discrepancies
    AppState.reportData[atelier] = JSON.parse(JSON.stringify(discrepancies));

    updateReportCount();
    showToast(`Added ${atelier} to report`, 'success');
}

function updateReportCount() {
    const count = Object.keys(AppState.reportData).length;
    elements.reportCount.textContent = `${count} atelier${count !== 1 ? 's' : ''} added to report`;
    elements.btnGenerateReport.disabled = count === 0;
    elements.btnClearReport.style.display = count > 0 ? 'inline-flex' : 'none';
}

function clearReport() {
    AppState.reportData = {};
    updateReportCount();
    showToast('Report cleared', 'info');
}

function getUnitLocation(unit) {
    const locations = {
        'Fath1': 'بئر خادم - الفتح 1',
        'Fath2': 'بئر خادم - الفتح 2',
        'Fath3': 'بئر خادم - الفتح 3',
        'Fath5': 'بئر خادم - الفتح 5',
        'Larbaa': 'الأربعاء',
        'Oran': 'وهران',
        'Fibre': 'الفيبر',
        'Mdoukal': "مضوكل",
        'Mags': 'المغازن'
    };
    return locations[unit] || unit;
}

function generateReportMarkdown() {
    const unit = AppState.selectedUnit;
    const today = new Date();
    const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
    const monthNames = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو',
        'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
    const selectedMonthName = monthNames[parseInt(AppState.selectedMonth) - 1];

    let markdown = `# SPA ELFATH

**شركة ذات الأسهم الفتح**  
FABRICATION DE MOUSSE ET DE LITERIE

---

في: ${getUnitLocation(unit)}                                                                                           **${dateStr}**  
**المديرية العامة**  
**مصلحة المراقبة**

---

بعد مراقبتنا لملفات المخزون ربطته مع ملف حركة المخزون شهر ${selectedMonthName} ${today.getFullYear()}، وجدنا بعض الاختلافات المبينة في الجداول التالية:

---

`;

    // Add each atelier table
    for (const [atelierName, discrepancies] of Object.entries(AppState.reportData)) {
        markdown += `## ${atelierName}\n\n`;
        markdown += `| REFERENCE | ETAT STOCKS | FICHIER DE MOV | ECARTS |\n`;
        markdown += `|-----------|-------------|----------------|--------|\n`;

        for (const row of discrepancies) {
            const stockQty = parseFloat(row.Stock_Qty || 0).toFixed(2);
            const movQty = parseFloat(row.Calc_Mov_Qty || 0).toFixed(2);
            const diff = parseFloat(row.Difference || 0).toFixed(2);
            markdown += `| ${row.Ref} | ${stockQty} | ${movQty} | ${diff} |\n`;
        }
        markdown += `\n---\n\n`;
    }

    return markdown;
}

function showReportModal() {
    if (Object.keys(AppState.reportData).length === 0) {
        showToast('Add at least one atelier to the report first', 'warning');
        return;
    }

    const markdown = generateReportMarkdown();
    elements.reportEditor.value = markdown;
    elements.reportModal.classList.remove('hidden');
}

function closeReportModal() {
    elements.reportModal.classList.add('hidden');
}

async function exportReportMd() {
    const content = elements.reportEditor.value;
    if (!content.trim()) {
        showToast('Report is empty', 'warning');
        return;
    }

    const defaultName = `report_${AppState.selectedUnit}_${new Date().toISOString().split('T')[0]}.md`;

    try {
        const filePath = await window.electronAPI.saveMarkdownDialog(defaultName);
        if (filePath) {
            await window.electronAPI.saveMarkdown(content, filePath);
            showToast('Markdown exported successfully!', 'success');
            closeReportModal();
        }
    } catch (error) {
        // Fallback: create a blob and trigger download
        const blob = new Blob([content], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = defaultName;
        a.click();
        URL.revokeObjectURL(url);
        showToast('Markdown downloaded!', 'success');
        closeReportModal();
    }
}

async function exportReportPdf() {
    if (Object.keys(AppState.reportData).length === 0) {
        showToast('Report is empty', 'warning');
        return;
    }

    try {
        showToast('Generating PDF...', 'info');

        const today = new Date();
        const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
        const monthNames = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو',
            'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
        const selectedMonthName = monthNames[parseInt(AppState.selectedMonth) - 1];

        // Build atelier tables HTML
        let ateliersHtml = '';
        for (const [atelierName, discrepancies] of Object.entries(AppState.reportData)) {
            let rowsHtml = '';
            for (const row of discrepancies) {
                const stockQty = parseFloat(row.Stock_Qty || 0).toFixed(2);
                const movQty = parseFloat(row.Calc_Mov_Qty || 0).toFixed(2);
                const diff = parseFloat(row.Difference || 0).toFixed(2);
                const diffClass = parseFloat(diff) !== 0 ? 'font-bold text-red-600' : '';
                rowsHtml += `
                    <tr class="hover:bg-gray-50">
                        <td class="border border-gray-400 px-4 py-2 font-medium">${row.Ref}</td>
                        <td class="border border-gray-400 px-4 py-2 text-right">${stockQty}</td>
                        <td class="border border-gray-400 px-4 py-2 text-right">${movQty}</td>
                        <td class="border border-gray-400 px-4 py-2 text-right ${diffClass}">${diff}</td>
                    </tr>`;
            }

            ateliersHtml += `
            <div class="mb-8" dir="ltr">
                <h3 class="text-lg font-bold mb-3 uppercase border-l-4 border-blue-800 pl-2 latin-text">${atelierName.toUpperCase()} :</h3>
                <div class="overflow-x-auto">
                    <table class="min-w-full border-collapse border border-gray-400 text-sm md:text-base latin-text">
                        <thead class="bg-gray-200">
                            <tr>
                                <th class="border border-gray-400 px-4 py-2 text-left">REFERENCE</th>
                                <th class="border border-gray-400 px-4 py-2 text-right">ETAT STOCKS</th>
                                <th class="border border-gray-400 px-4 py-2 text-right">FICHIER DE MOV</th>
                                <th class="border border-gray-400 px-4 py-2 text-right">ECARTS</th>
                            </tr>
                        </thead>
                        <tbody>${rowsHtml}</tbody>
                    </table>
                </div>
            </div>`;
        }

        // Generate full HTML using the template format
        const htmlContent = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stock Movement Report - El Fath</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Roboto:wght@400;500;700&display=swap');
        body { font-family: 'Amiri', serif; }
        .latin-text { font-family: 'Roboto', sans-serif; }
        @media print {
            body { -webkit-print-color-adjust: exact; }
            .no-print { display: none; }
        }
    </style>
</head>
<body class="bg-gray-100 min-h-screen p-4 md:p-8 text-gray-900">
    <div class="max-w-4xl mx-auto bg-white shadow-lg p-8 md:p-12 border border-gray-200">
        <!-- Header Section -->
        <div class="flex flex-col md:flex-row justify-between items-start mb-8 border-b-2 border-gray-800 pb-4">
            <div class="text-right w-full md:w-1/3 mb-4 md:mb-0 order-1 md:order-2">
                <p class="text-lg font-bold">${getUnitLocation(AppState.selectedUnit)} في : ${dateStr}</p>
            </div>
            <div class="text-left w-full md:w-2/3 order-2 md:order-1 latin-text" dir="ltr">
                <h1 class="text-2xl font-bold uppercase tracking-wider text-blue-900">el.Fath</h1>
                <h2 class="text-xl font-bold uppercase">SPA ELFATH</h2>
                <p class="text-sm text-gray-600 font-medium">FABRICATION DE MOUSSE ET DE LITERIE</p>
            </div>
        </div>

        <!-- Title Section -->
        <div class="text-center mb-10">
            <h1 class="text-3xl font-bold underline decoration-2 underline-offset-8 mb-2">ملف حركة المخزون</h1>
            <h2 class="text-xl font-semibold mt-4">شركة ذات الأسهم الفتح</h2>
        </div>

        <!-- Departments -->
        <div class="flex flex-col items-start gap-2 mb-8 text-xl">
            <div class="font-bold">المديرية العامة</div>
            <div class="font-bold">مصلحة المراقبة</div>
        </div>

        <!-- Body Text -->
        <div class="mb-8 text-lg leading-relaxed text-justify">
            <p>
                وجدنا بعض بعد مراقبتنا لملفات المخزون شهر ${selectedMonthName} ${today.getFullYear()} في وحدة ${getUnitLocation(AppState.selectedUnit)} ومقارنته مع الاختلافات المبينة في الجداول التالية :
            </p>
        </div>

        ${ateliersHtml}
    </div>
</body>
</html>`;

        const printWindow = window.open('', '_blank');
        printWindow.document.write(htmlContent);
        printWindow.document.close();

        setTimeout(() => {
            printWindow.print();
            showToast('Print dialog opened', 'success');
            closeReportModal();
        }, 500);

    } catch (error) {
        showToast(`PDF export failed: ${error}`, 'error');
    }
}

function markdownToHtml(markdown) {
    // Simple markdown to HTML conversion
    return markdown
        .replace(/^# (.*$)/gm, '<h1>$1</h1>')
        .replace(/^## (.*$)/gm, '<h2>$1</h2>')
        .replace(/^---$/gm, '<hr>')
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\n\n/g, '</p><p>')
        .replace(/\n/g, '<br>')
        .replace(/\|(.+)\|/g, (match, content) => {
            const cells = content.split('|').map(c => c.trim());
            if (cells.every(c => c.match(/^-+$/))) return ''; // Skip separator row
            const tag = cells[0].match(/^[A-Z]/) ? 'td' : 'td';
            return '<tr>' + cells.map(c => `<${tag}>${c}</${tag}>`).join('') + '</tr>';
        })
        .replace(/(<tr>.*<\/tr>)+/g, '<table>$&</table>')
        .replace(/<p><\/p>/g, '')
        .replace(/^<br>/gm, '');
}

function updateVerifyButtonVisibility() {
    const viewType = document.querySelector('.toggle-btn.active').dataset.view;
    const isDiscrepancies = viewType === 'discrepancies';

    if (elements.btnVerifyOpposite) {
        elements.btnVerifyOpposite.style.display = isDiscrepancies ? 'inline-flex' : 'none';
    }
    if (elements.btnInsertReport) {
        elements.btnInsertReport.style.display = isDiscrepancies ? 'inline-flex' : 'none';
    }
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

    elements.btnVerify.addEventListener('click', verifyFiles);
    elements.btnProcess.addEventListener('click', processFiles);

    // Verification modal
    elements.closeVerificationModal.addEventListener('click', closeVerificationModal);
    elements.btnCancelVerification.addEventListener('click', closeVerificationModal);
    elements.btnApplyFixes.addEventListener('click', applyFixesAndProcess);
    elements.verificationModal.querySelector('.modal-overlay').addEventListener('click', closeVerificationModal);

    // Results
    elements.atelierSelect.addEventListener('change', () => {
        AppState.searchQuery = '';
        AppState.currentSortColumn = null;
        if (elements.resultsSearch) elements.resultsSearch.value = '';
        exitVerifyMode();
        renderResultsTable();
    });

    elements.toggleBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            elements.toggleBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            AppState.searchQuery = '';
            AppState.currentSortColumn = null;
            if (elements.resultsSearch) elements.resultsSearch.value = '';
            exitVerifyMode();
            updateVerifyButtonVisibility();
            renderResultsTable();
        });
    });

    // Search
    if (elements.resultsSearch) {
        elements.resultsSearch.addEventListener('input', (e) => {
            AppState.searchQuery = e.target.value;
            renderResultsTable();
        });
    }

    // Sortable headers
    document.querySelectorAll('.results-table th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            handleSort(th.dataset.column);
        });
    });

    // Verify-mode: checkbox changes (row selection)
    if (elements.resultsTbody) {
        elements.resultsTbody.addEventListener('change', (e) => {
            const target = e.target;
            if (!(target instanceof HTMLInputElement)) return;
            if (!target.classList.contains('row-select-checkbox')) return;
            const tr = target.closest('tr');
            const ref = tr?.dataset?.ref;
            if (!ref) return;
            window.toggleRowSelection(ref);
        });

        // Right-click context menu for any row (matches or discrepancies)
        elements.resultsTbody.addEventListener('contextmenu', (e) => {
            const tr = e.target?.closest?.('tr[data-ref]');
            if (!tr) return;
            const ref = tr.dataset.ref;
            const atelier = elements.atelierSelect.value;
            const viewType = document.querySelector('.toggle-btn.active')?.dataset?.view;
            showContextMenu(e, ref, { atelier, viewType });
        });
    }

    // Verify-mode: select all
    document.addEventListener('change', (e) => {
        const target = e.target;
        if (!(target instanceof HTMLInputElement)) return;
        if (target.id !== 'verify-select-all') return;
        if (!AppState.verifyMode) return;

        const all = getAllHighlightedOppositeRefs();
        if (target.checked) {
            all.forEach(ref => AppState.selectedForElimination.add(ref));
        } else {
            all.forEach(ref => AppState.selectedForElimination.delete(ref));
        }
        renderResultsTable();
    });

    elements.btnExport.addEventListener('click', exportResults);

    if (elements.btnExportExcel) {
        elements.btnExportExcel.addEventListener('click', exportResultsExcel);
    }

    elements.btnNewProcess.addEventListener('click', () => {
        resetFileState();
        AppState.reportData = {};
        updateReportCount();
        navigateTo('page-unit-selection');
    });

    // Verify Opposite Button
    if (elements.btnVerifyOpposite) {
        elements.btnVerifyOpposite.addEventListener('click', toggleVerifyMode);
    }

    // Insert Report Button
    if (elements.btnInsertReport) {
        elements.btnInsertReport.addEventListener('click', addToReport);
    }

    // Generate Report Button
    if (elements.btnGenerateReport) {
        elements.btnGenerateReport.addEventListener('click', showReportModal);
    }

    // Clear Report Button
    if (elements.btnClearReport) {
        elements.btnClearReport.addEventListener('click', clearReport);
    }

    // Report Modal
    if (elements.closeReportModal) {
        elements.closeReportModal.addEventListener('click', closeReportModal);
    }
    if (elements.btnCancelReport) {
        elements.btnCancelReport.addEventListener('click', closeReportModal);
    }
    if (elements.reportModal) {
        elements.reportModal.querySelector('.modal-overlay').addEventListener('click', closeReportModal);
    }
    if (elements.btnExportMd) {
        elements.btnExportMd.addEventListener('click', exportReportMd);
    }
    if (elements.btnExportPdf) {
        elements.btnExportPdf.addEventListener('click', exportReportPdf);
    }

    // Context Menu
    if (elements.ctxDeleteRow) {
        elements.ctxDeleteRow.addEventListener('click', deleteRowFromContext);
    }
    if (elements.ctxEditRow) {
        elements.ctxEditRow.addEventListener('click', editRowFromContext);
    }
}

// ============================================
// Initialization
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    initUnitSelection();
    initDropZone();
    initEventListeners();
});

// Expose functions to global scope for inline event handlers
window.toggleRowSelection = function (ref) {
    if (AppState.selectedForElimination.has(ref)) {
        AppState.selectedForElimination.delete(ref);
    } else {
        AppState.selectedForElimination.add(ref);
    }
    renderResultsTable();
};

window.showContextMenu = function (event, ref) {
    showContextMenu(event, ref, { atelier: elements.atelierSelect.value, viewType: document.querySelector('.toggle-btn.active')?.dataset?.view });
};
