// app.js - Client UI State Orchestrator & Exporter

// Global State
let fileABuffer = null;
let fileBBuffer = null;
let fileAName = '';
let fileBName = '';
let comparisonResult = null;
let activeSheetName = null;
let currentPage = 1;
const rowsPerPage = 100;
let currentSearchQuery = '';
let currentFilter = 'diff'; // 'all', 'diff', 'added', 'missing', 'modified'
let worker = null;

// Initialize Lucide Icons
function initIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

document.addEventListener('DOMContentLoaded', () => {
  // Theme Management
  const themeToggle = document.getElementById('theme-toggle');
  let currentTheme = 'light';
  try {
    currentTheme = localStorage.getItem('theme') || 'light';
  } catch (e) {
    console.warn('localStorage is blocked or not supported in this environment. Defaulting to light theme.', e);
  }
  document.documentElement.setAttribute('data-theme', currentTheme);
  updateThemeIcon(currentTheme);
  
  themeToggle.addEventListener('click', () => {
    const theme = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', theme);
    try {
      localStorage.setItem('theme', theme);
    } catch (e) {
      console.warn('localStorage is blocked or not supported in this environment. Theme preference not saved.', e);
    }
    updateThemeIcon(theme);
  });

  // Setup drag and drop for files
  setupDragAndDrop('upload-card-a', 'file-input-a', 'file-info-a', 'file-name-a', (buf, name) => {
    fileABuffer = buf;
    fileAName = name;
    checkReadyToCompare();
  });
  
  setupDragAndDrop('upload-card-b', 'file-input-b', 'file-info-b', 'file-name-b', (buf, name) => {
    fileBBuffer = buf;
    fileBName = name;
    checkReadyToCompare();
  });

  // Remove buttons
  document.getElementById('remove-file-a').addEventListener('click', (e) => {
    e.stopPropagation();
    resetFileInput('file-input-a', 'file-info-a');
    fileABuffer = null;
    fileAName = '';
    checkReadyToCompare();
  });
  
  document.getElementById('remove-file-b').addEventListener('click', (e) => {
    e.stopPropagation();
    resetFileInput('file-input-b', 'file-info-b');
    fileBBuffer = null;
    fileBName = '';
    checkReadyToCompare();
  });

  // Compare Button
  const compareBtn = document.getElementById('compare-btn');
  compareBtn.addEventListener('click', startComparison);

  // Clear Button
  const clearBtn = document.getElementById('clear-btn');
  if (clearBtn) clearBtn.addEventListener('click', clearAll);

  // Export Buttons
  const expExcel = document.getElementById('export-excel');
  if (expExcel) expExcel.addEventListener('click', exportExcel);
  const expPdf = document.getElementById('export-pdf');
  if (expPdf) expPdf.addEventListener('click', exportPdf);
  const expCsv = document.getElementById('export-csv');
  if (expCsv) expCsv.addEventListener('click', exportCsv);

  // Search & Filter Listeners
  const searchInput = document.getElementById('sheet-search');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      currentSearchQuery = e.target.value.toLowerCase().trim();
      currentPage = 1;
      renderCurrentSheet();
    });
  }

  // Filter Pills
  document.querySelectorAll('.filter-pill').forEach(pill => {
    pill.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      currentFilter = pill.dataset.filter;
      currentPage = 1;
      renderCurrentSheet();
    });
  });

  // Mobile Sheet dropdown change
  const mobSelect = document.getElementById('mobile-sheet-select');
  if (mobSelect) {
    mobSelect.addEventListener('change', (e) => {
      selectSheet(e.target.value);
    });
  }

  // View Switcher logic
  const btnDashboard = document.getElementById('btn-view-dashboard');
  const btnExcel = document.getElementById('btn-view-excel');
  const containerDashboard = document.getElementById('view-dashboard-container');
  const containerExcel = document.getElementById('view-excel-container');
  
  btnDashboard.addEventListener('click', () => {
    btnDashboard.classList.add('active');
    btnExcel.classList.remove('active');
    containerDashboard.classList.add('active');
    containerExcel.classList.remove('active');
    renderCurrentSheet();
  });
  
  btnExcel.addEventListener('click', () => {
    btnDashboard.classList.remove('active');
    btnExcel.classList.add('active');
    containerDashboard.classList.remove('active');
    containerExcel.classList.add('active');
    renderExcelView();
  });

  // Synchronized scroll listener bindings
  setupSynchronizedScrolling();
  
  // Excel view tool listeners
  document.getElementById('excel-show-diffs-only').addEventListener('change', () => {
    renderExcelSheetGrids();
  });

  initIcons();
});

// Update light/dark toggle icon
function updateThemeIcon(theme) {
  const btn = document.getElementById('theme-toggle');
  if (theme === 'dark') {
    btn.innerHTML = `<i data-lucide="sun"></i>`;
  } else {
    btn.innerHTML = `<i data-lucide="moon"></i>`;
  }
  initIcons();
}

// Drag and drop helper
function setupDragAndDrop(cardId, inputId, infoId, nameId, callback) {
  const card = document.getElementById(cardId);
  const input = document.getElementById(inputId);
  const info = document.getElementById(infoId);
  const nameEl = document.getElementById(nameId);

  card.addEventListener('dragover', (e) => {
    e.preventDefault();
    card.classList.add('dragover');
  });

  card.addEventListener('dragleave', () => {
    card.classList.remove('dragover');
  });

  card.addEventListener('drop', (e) => {
    e.preventDefault();
    card.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      input.files = files;
      handleFileSelect(files[0], info, nameEl, callback);
    }
  });

  input.addEventListener('change', (e) => {
    const files = e.target.files;
    if (files.length > 0) {
      handleFileSelect(files[0], info, nameEl, callback);
    }
  });
}

function handleFileSelect(file, info, nameEl, callback) {
  nameEl.textContent = `${file.name} (${formatBytes(file.size)})`;
  info.classList.add('active');
  
  const reader = new FileReader();
  reader.onload = (e) => {
    callback(e.target.result, file.name);
  };
  reader.readAsArrayBuffer(file);
}

function resetFileInput(inputId, infoId) {
  const input = document.getElementById(inputId);
  const info = document.getElementById(infoId);
  input.value = '';
  info.classList.remove('active');
}

function clearAll() {
  fileABuffer = null;
  fileBBuffer = null;
  fileAName = '';
  fileBName = '';
  comparisonResult = null;
  activeSheetName = null;
  excelActiveSheetName = null;
  currentPage = 1;
  currentSearchQuery = '';
  currentFilter = 'diff';

  resetFileInput('file-input-a', 'file-info-a');
  resetFileInput('file-input-b', 'file-info-b');
  
  document.getElementById('compare-btn').setAttribute('disabled', 'true');
  document.getElementById('results-container').classList.remove('active');
  document.getElementById('loading-container').classList.remove('active');
  document.getElementById('sheet-search').value = '';
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}



function checkReadyToCompare() {
  const compareBtn = document.getElementById('compare-btn');
  if (fileABuffer && fileBBuffer) {
    compareBtn.removeAttribute('disabled');
  } else {
    compareBtn.setAttribute('disabled', 'true');
  }
}

// Toast notification system helper
function showToast(message, type = 'info', iconName = 'zap') {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <i data-lucide="${iconName}"></i>
    <span>${escapeHtml(message)}</span>
  `;
  container.appendChild(toast);
  if (typeof initIcons === 'function') initIcons();

  // Trigger smooth enter animation
  setTimeout(() => toast.classList.add('show'), 10);

  // Auto remove after 3.5 seconds
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// Main comparison dispatcher
function startComparison() {
  if (!fileABuffer || !fileBBuffer) return;

  // Trigger Toast Notification on click
  showToast('Initializing report comparison...', 'info', 'zap');

  const compareBtn = document.getElementById('compare-btn');
  const loadingContainer = document.getElementById('loading-container');
  const resultsContainer = document.getElementById('results-container');
  const logEl = document.getElementById('progress-log');
  
  // Reset views
  resultsContainer.classList.remove('active');
  loadingContainer.classList.add('active');
  compareBtn.setAttribute('disabled', 'true');
  logEl.textContent = 'Initializing comparison engine...';

  // Config parameters
  const config = {
    ignoreEmptyRows: document.getElementById('ignore-empty-rows').checked,
    ignoreCase: document.getElementById('ignore-case').checked,
    ignoreWhitespace: document.getElementById('ignore-whitespace').checked,
    ignoreDateFormatting: document.getElementById('ignore-dates').checked
  };

  // Determine if we should attempt using Web Worker
  const useWorker = window.location.protocol !== 'file:';
  
  if (useWorker) {
    try {
      if (worker) {
        worker.terminate();
      }
      
      worker = new Worker('worker.js?v=18');
      
      worker.onmessage = function (e) {
        const data = e.data;
        if (data.status === 'progress') {
          logEl.textContent = data.message;
        } else if (data.status === 'complete') {
          comparisonResult = data.result;
          loadingContainer.classList.remove('active');
          resultsContainer.classList.add('active');
          compareBtn.removeAttribute('disabled');
          renderDashboard(comparisonResult);
        } else if (data.status === 'error') {
          loadingContainer.classList.remove('active');
          compareBtn.removeAttribute('disabled');
          alert(data.message);
        }
      };

      worker.onerror = function (err) {
        console.warn('Web Worker error, falling back to main thread:', err);
        try { worker.terminate(); } catch(e){}
        runMainThreadComparison(config, compareBtn, loadingContainer, resultsContainer, logEl);
      };

      worker.postMessage({
        fileABuffer,
        fileBBuffer,
        config
      });
    } catch (e) {
      console.warn('Worker creation failed, falling back to main thread:', e);
      runMainThreadComparison(config, compareBtn, loadingContainer, resultsContainer, logEl);
    }
  } else {
    console.info('Local file protocol detected. Running comparison in main thread fallback.');
    runMainThreadComparison(config, compareBtn, loadingContainer, resultsContainer, logEl);
  }
}

// Fallback synchronous execution on the main thread
function runMainThreadComparison(config, compareBtn, loadingContainer, resultsContainer, logEl) {
  // Using setTimeout so the UI thread renders the loading screen first
  setTimeout(() => {
    try {
      const startTime = performance.now();
      
      const result = compareWorkbooks(fileABuffer, fileBBuffer, config, (message) => {
        logEl.textContent = message;
      });
      
      const endTime = performance.now();
      const elapsedTime = ((endTime - startTime) / 1000).toFixed(2);
      result.summary.elapsedTime = `${elapsedTime} sec`;
      
      comparisonResult = result;
      loadingContainer.classList.remove('active');
      resultsContainer.classList.add('active');
      compareBtn.removeAttribute('disabled');
      renderDashboard(comparisonResult);
    } catch (err) {
      loadingContainer.classList.remove('active');
      compareBtn.removeAttribute('disabled');
      alert('Comparison failed: ' + err.message);
    }
  }, 100);
}

// Dashboard metrics display
function renderDashboard(result) {
  const summary = result.summary;
  
  // Trigger completion success Toast
  showToast(`Comparison complete! (${summary.cellDifferences} cell diffs found)`, 'success', 'check-circle-2');

  // Write stats
  document.getElementById('stat-sheets-compared').textContent = summary.sheetsCompared;
  document.getElementById('stat-matched-sheets').textContent = summary.matchedSheets;
  document.getElementById('stat-modified-sheets').textContent = summary.modifiedSheets;
  document.getElementById('stat-missing-sheets').textContent = summary.missingSheets;
  document.getElementById('stat-additional-sheets').textContent = summary.additionalSheets;
  document.getElementById('stat-cell-changes').textContent = summary.cellDifferences;
  document.getElementById('stat-elapsed-time').textContent = summary.elapsedTime;

  // Dynamic show/hide summary cards based on values
  const wMatched = document.getElementById('wrapper-matched-sheets');
  const wModified = document.getElementById('wrapper-modified-sheets');
  const wMissing = document.getElementById('wrapper-missing-sheets');
  const wAdditional = document.getElementById('wrapper-additional-sheets');
  const wCells = document.getElementById('wrapper-cell-changes');

  if (wMatched) wMatched.style.display = summary.matchedSheets > 0 ? 'flex' : 'none';
  if (wModified) wModified.style.display = summary.modifiedSheets > 0 ? 'flex' : 'none';
  if (wMissing) wMissing.style.display = summary.missingSheets > 0 ? 'flex' : 'none';
  if (wAdditional) wAdditional.style.display = summary.additionalSheets > 0 ? 'flex' : 'none';
  if (wCells) wCells.style.display = summary.cellDifferences > 0 ? 'flex' : 'none';

  // Render sheet list (header dropdown)
  renderSheetNavigator(result.sheetStatuses);
  
  // Activate Side-by-Side Excel view by default
  const btnDashboard = document.getElementById('btn-view-dashboard');
  const btnExcel = document.getElementById('btn-view-excel');
  const containerDashboard = document.getElementById('view-dashboard-container');
  const containerExcel = document.getElementById('view-excel-container');
  
  btnExcel.classList.add('active');
  btnDashboard.classList.remove('active');
  containerExcel.classList.add('active');
  containerDashboard.classList.remove('active');

  renderExcelView();
}

// Helper to format tab name HTML with strikethrough for renamed sheets
function formatTabNameHTML(sheet) {
  if (sheet.nameA && sheet.nameB) {
    if (sheet.nameA === sheet.nameB) {
      return `<span>${escapeHtml(sheet.nameA)}</span>`;
    } else {
      return `<del class="old-tab-name">${escapeHtml(sheet.nameA)}</del> <span class="new-tab-name">${escapeHtml(sheet.nameB)}</span>`;
    }
  } else if (sheet.nameA && !sheet.nameB) {
    return `<del class="old-tab-name">${escapeHtml(sheet.nameA)}</del> <span class="deleted-tag">(Deleted)</span>`;
  } else if (!sheet.nameA && sheet.nameB) {
    return `<span class="added-tag">(Added)</span> <span class="new-tab-name">${escapeHtml(sheet.nameB)}</span>`;
  }
  return '';
}

// Helper to format tab name for dropdown select (option elements)
function formatTabNameText(sheet) {
  if (sheet.nameA && sheet.nameB) {
    if (sheet.nameA === sheet.nameB) {
      return sheet.nameA;
    } else {
      return `${sheet.nameA} ➔ ${sheet.nameB}`;
    }
  } else if (sheet.nameA && !sheet.nameB) {
    return `${sheet.nameA} [Deleted]`;
  } else if (!sheet.nameA && sheet.nameB) {
    return `[Added] ${sheet.nameB}`;
  }
  return '';
}

function renderSheetNavigator(sheetStatuses) {
  const headerSelect = document.getElementById('header-sheet-select');
  if (!headerSelect) return;
  
  headerSelect.innerHTML = '';
  
  sheetStatuses.forEach((sheet) => {
    let badgeText = 'Match';
    if (sheet.status === 'Differences') badgeText = 'Diffs';
    else if (sheet.status === 'Missing') badgeText = 'Missing';
    else if (sheet.status === 'Additional') badgeText = 'Added';
    
    const opt = document.createElement('option');
    opt.value = sheet.displayName;
    opt.textContent = `${formatTabNameText(sheet)} (${badgeText})`;
    headerSelect.appendChild(opt);
  });
  
  // Re-bind change listener
  const newSelect = headerSelect.cloneNode(true);
  headerSelect.parentNode.replaceChild(newSelect, headerSelect);
  newSelect.addEventListener('change', (e) => {
    selectSheet(e.target.value);
  });
  
  // Select first sheet by default
  if (sheetStatuses.length > 0) {
    selectSheet(sheetStatuses[0].displayName);
  }
}

function selectSheet(displayName) {
  activeSheetName = displayName;
  currentPage = 1;
  currentSearchQuery = '';
  currentFilter = 'diff';
  
  // Reset UI search input & filter active class
  const searchEl = document.getElementById('sheet-search');
  if (searchEl) searchEl.value = '';
  document.querySelectorAll('.filter-pill').forEach(p => {
    if (p.dataset.filter === 'diff') p.classList.add('active');
    else p.classList.remove('active');
  });
  
  const headerSelect = document.getElementById('header-sheet-select');
  if (headerSelect) headerSelect.value = displayName;

  renderCurrentSheet();
}

// Render the sheet details (diff sections & grid)
function renderCurrentSheet() {
  if (!activeSheetName || !comparisonResult) return;
  
  const detail = comparisonResult.sheetDetails[activeSheetName];
  
  // Update Status Badge
  const statusBadge = document.getElementById('detail-sheet-status');
  if (statusBadge) {
    statusBadge.className = 'badge';
    if (detail.status === 'Match') {
      statusBadge.classList.add('match');
      statusBadge.textContent = 'Matched';
    } else if (detail.status === 'Differences') {
      statusBadge.classList.add('diff');
      statusBadge.textContent = 'Differences Found';
    } else if (detail.status === 'Missing') {
      statusBadge.classList.add('missing');
      statusBadge.textContent = 'Missing in Comparison';
    } else if (detail.status === 'Additional') {
      statusBadge.classList.add('additional');
      statusBadge.textContent = 'Added in Comparison';
    }
  }

  // Cell Differences Section
  const cellAccordion = document.getElementById('sec-cell-diffs');
  const cellCount = document.getElementById('count-cell-diffs');
  const cellTableBody = document.getElementById('cell-diff-table-body');
  
  cellTableBody.innerHTML = '';
  
  // Display tools & pagination if sheet is compared
  const toolsEl = document.querySelector('.sheet-tools');
  const paginationEl = document.getElementById('pagination-container');
  const gridWrapperEl = document.getElementById('grid-container');

  if (detail.status === 'Missing') {
    // Entirely missing sheet - show warning banner, hide details
    cellAccordion.style.display = 'none';
    toolsEl.style.display = 'none';
    paginationEl.style.display = 'none';
    gridWrapperEl.innerHTML = `
      <div style="padding: 3rem; text-align: center; color: var(--color-danger); font-weight: 600; background-color: var(--bg-secondary); border-radius: 8px; border: 1px dashed var(--color-danger-border);">
        <i data-lucide="x-circle" style="width: 48px; height: 48px; margin-bottom: 0.5rem; color: var(--color-danger);"></i>
        <div>Worksheet is entirely missing in Comparison File B.</div>
      </div>
    `;
    initIcons();
    setupAccordions();
    return;
  } else if (detail.status === 'Additional') {
    // Entirely additional sheet - show info banner, hide details
    cellAccordion.style.display = 'none';
    toolsEl.style.display = 'none';
    paginationEl.style.display = 'none';
    gridWrapperEl.innerHTML = `
      <div style="padding: 3rem; text-align: center; color: var(--color-primary); font-weight: 600; background-color: var(--bg-secondary); border-radius: 8px; border: 1px dashed #c7d2fe;">
        <i data-lucide="plus-circle" style="width: 48px; height: 48px; margin-bottom: 0.5rem; color: var(--color-primary);"></i>
        <div>Worksheet is entirely additional in Comparison File B.</div>
      </div>
    `;
    initIcons();
    setupAccordions();
    return;
  }

  // Normal Compared sheet
  toolsEl.style.display = 'flex';
  paginationEl.style.display = 'flex';

  if (detail.cells) {
    cellCount.textContent = detail.cells.length;
    cellAccordion.style.display = detail.cells.length > 0 ? 'block' : 'none';
    
    // Fill top 100 cell diffs in details panel table to prevent browser crash
    const topCells = detail.cells.slice(0, 100);
    topCells.forEach(c => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${c.cellRef}</strong></td>
        <td>Row ${c.row}</td>
        <td>Column ${c.colLetter}</td>
        <td class="color-danger" style="text-decoration: line-through;">${escapeHtml(c.expected)}</td>
        <td class="color-success" style="font-weight:600;">${escapeHtml(c.actual)}</td>
      `;
      cellTableBody.appendChild(tr);
    });
    
    if (detail.cells.length > 100) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td colspan="5" style="text-align: center; color: var(--text-muted);">... Showing first 100 cell differences. Export to Excel or PDF to view all ${detail.cells.length} differences.</td>`;
      cellTableBody.appendChild(tr);
    }
  } else {
    cellCount.textContent = 0;
    cellAccordion.style.display = 'none';
  }

  // Render Side-by-Side Comparison Grid
  renderComparisonGrid(detail);
  initIcons();
  
  // Accordion Expand/Collapse Bindings
  setupAccordions();
}

function setupAccordions() {
  document.querySelectorAll('.diff-section-header').forEach(header => {
    // Avoid double-binding
    if (header.dataset.bound) return;
    header.dataset.bound = "true";
    
    header.addEventListener('click', () => {
      const section = header.parentElement;
      section.classList.toggle('open');
    });
  });
}

// Render the grid values with styling, search, filters, pagination
function renderComparisonGrid(detail) {
  const container = document.getElementById('grid-container');
  container.innerHTML = '';
  
  if (detail.maxRows === 0) {
    container.innerHTML = '<div style="padding: 2rem; text-align: center; color: var(--text-muted);">No data available in this sheet.</div>';
    document.getElementById('pagination-container').style.display = 'none';
    return;
  }

  const matchedRowIndices = [];
  
  const rowSlots = detail.rowSlots || [];
  for (let r = 0; r < detail.maxRows; r++) {
    const rowSlot = rowSlots[r] || { type: 'match', mapText: `${r+1}:${r+1}` };
    
    // 1. Filter pill rules
    let rowHasDiff = false;
    if (detail.diffMask[r]) {
      rowHasDiff = detail.diffMask[r].some(val => val === true);
    }
    
    let isAddedRow = rowSlot.type === 'added';
    let isMissingRow = rowSlot.type === 'deleted';
    let isModifiedRow = rowHasDiff && !isAddedRow && !isMissingRow;
    
    let matchesFilter = true;
    if (currentFilter === 'diff' && !rowHasDiff) matchesFilter = false;
    if (currentFilter === 'added' && !isAddedRow) matchesFilter = false;
    if (currentFilter === 'missing' && !isMissingRow) matchesFilter = false;
    if (currentFilter === 'modified' && !isModifiedRow) matchesFilter = false;
    
    // 2. Search query rules
    let matchesSearch = true;
    if (currentSearchQuery) {
      matchesSearch = false;
      const searchLower = currentSearchQuery.toLowerCase();
      if (String(rowSlot.mapText).toLowerCase().includes(searchLower)) {
        matchesSearch = true;
      } else {
        const colsA = detail.gridA[r] || [];
        const colsB = detail.gridB[r] || [];
        for (let c = 0; c < detail.maxCols; c++) {
          const valA = colsA[c];
          const valB = colsB[c];
          if ((valA !== undefined && String(valA).toLowerCase().includes(searchLower)) ||
              (valB !== undefined && String(valB).toLowerCase().includes(searchLower))) {
            matchesSearch = true;
            break;
          }
        }
      }
    }
    
    if (matchesFilter && matchesSearch) {
      matchedRowIndices.push({
        rowIndex: r,
        rowSlot,
        isAddedRow,
        isMissingRow,
        isModifiedRow,
        rowHasDiff
      });
    }
  }

  const totalFiltered = matchedRowIndices.length;
  
  if (totalFiltered === 0) {
    container.innerHTML = '<div style="padding: 2rem; text-align: center; color: var(--text-muted);">No records match your filters.</div>';
    document.getElementById('pagination-container').style.display = 'none';
    return;
  }

  // Pagination bounds
  const totalPages = Math.ceil(totalFiltered / rowsPerPage);
  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;
  
  const startIdx = (currentPage - 1) * rowsPerPage;
  const endIdx = Math.min(startIdx + rowsPerPage, totalFiltered);
  const paginatedRowMeta = matchedRowIndices.slice(startIdx, endIdx);

  // Setup Pagination UI
  const pagContainer = document.getElementById('pagination-container');
  pagContainer.style.display = 'flex';
  document.getElementById('pagination-info').textContent = `Showing ${startIdx + 1} - ${endIdx} of ${totalFiltered} records`;
  
  const prevBtn = document.getElementById('pagination-prev');
  const nextBtn = document.getElementById('pagination-next');
  
  prevBtn.disabled = currentPage === 1;
  nextBtn.disabled = currentPage === totalPages;

  // Clear older listeners to avoid loops
  const newPrev = prevBtn.cloneNode(true);
  const newNext = nextBtn.cloneNode(true);
  prevBtn.parentNode.replaceChild(newPrev, prevBtn);
  nextBtn.parentNode.replaceChild(newNext, nextBtn);
  
  newPrev.addEventListener('click', () => {
    currentPage--;
    renderComparisonGrid(detail);
  });
  newNext.addEventListener('click', () => {
    currentPage++;
    renderComparisonGrid(detail);
  });

  // Build Table
  const table = document.createElement('table');
  table.className = 'grid-comparison-table';
  
  // Table Head
  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  
  // Frozen Header 1: Row Mapping (@:@)
  const thMap = document.createElement('th');
  thMap.textContent = '@:@';
  thMap.title = 'Row Mapping';
  headerTr.appendChild(thMap);
  
  // Frozen Header 2: Status Indicator (!)
  const thStatus = document.createElement('th');
  thStatus.textContent = '!';
  thStatus.title = 'Diff Status Indicator';
  headerTr.appendChild(thStatus);
  
  const columnSlots = detail.columnSlots || [];
  for (let c = 0; c < detail.maxCols; c++) {
    const th = document.createElement('th');
    const slot = columnSlots[c] || { mapText: getExcelColLetter(c), type: 'match' };
    
    th.textContent = slot.mapText;
    if (slot.type === 'added') {
      th.className = 'col-added-header';
      th.title = `Column is Added in Comparison File`;
    } else if (slot.type === 'deleted') {
      th.className = 'col-deleted-header';
      th.title = `Column is Deleted in Comparison File`;
    }
    
    headerTr.appendChild(th);
  }
  thead.appendChild(headerTr);
  table.appendChild(thead);
  
  // Table Body
  const tbody = document.createElement('tbody');
  paginatedRowMeta.forEach(meta => {
    const r = meta.rowIndex;
    const slotInfo = meta.rowSlot || { mapText: `${r+1}:${r+1}` };
    const tr = document.createElement('tr');
    
    // Row classes
    if (meta.isMissingRow) {
      tr.className = 'grid-row-missing';
    } else if (meta.isAddedRow) {
      tr.className = 'grid-row-added';
    }
    
    // Cell 1: Row mapping index
    const tdMap = document.createElement('td');
    tdMap.textContent = slotInfo.mapText;
    tr.appendChild(tdMap);
    
    // Cell 2: Status tag symbol
    const tdStatus = document.createElement('td');
    if (meta.isMissingRow) {
      tdStatus.innerHTML = '<span class="tag-deleted">---</span>';
    } else if (meta.isAddedRow) {
      tdStatus.innerHTML = '<span class="tag-added">+++</span>';
    } else if (meta.isModifiedRow) {
      tdStatus.innerHTML = '<span class="tag-modified">➔</span>';
    } else {
      tdStatus.innerHTML = '<span style="color:var(--text-muted); opacity:0.7;">=</span>';
    }
    tr.appendChild(tdStatus);
    
    // Data Values
    for (let c = 0; c < detail.maxCols; c++) {
      const td = document.createElement('td');
      const slot = columnSlots[c] || { type: 'match' };
      const valA = (detail.gridA[r] && detail.gridA[r][c] !== undefined) ? detail.gridA[r][c] : '';
      const valB = (detail.gridB[r] && detail.gridB[r][c] !== undefined) ? detail.gridB[r][c] : '';
      const isDiff = detail.diffMask[r] ? detail.diffMask[r][c] : false;
      
      if (meta.isMissingRow) {
        td.textContent = String(valA);
      } else if (meta.isAddedRow) {
        td.textContent = String(valB);
      } else {
        if (slot.type === 'added') {
          td.className = 'grid-col-added';
          td.textContent = String(valB);
        } else if (slot.type === 'deleted') {
          td.className = 'grid-col-deleted';
          td.textContent = String(valA);
        } else if (isDiff) {
          td.className = 'grid-cell-diff';
          const wrapper = document.createElement('div');
          wrapper.className = 'cell-diff-val-wrapper';
          
          const oldSpan = document.createElement('span');
          oldSpan.className = 'cell-diff-old';
          oldSpan.textContent = valA !== '' ? String(valA) : '(Empty)';
          
          const arrowSpan = document.createElement('span');
          arrowSpan.className = 'cell-diff-arrow';
          arrowSpan.textContent = '➔';

          const newSpan = document.createElement('span');
          newSpan.className = 'cell-diff-new';
          newSpan.textContent = valB !== '' ? String(valB) : '(Empty)';
          
          wrapper.appendChild(oldSpan);
          wrapper.appendChild(arrowSpan);
          wrapper.appendChild(newSpan);
          td.appendChild(wrapper);
        } else {
          td.textContent = String(valB);
        }
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  
  table.appendChild(tbody);
  container.appendChild(table);
}

// HTML escape helper
function escapeHtml(str) {
  if (str === undefined || str === null) return '';
  const val = String(str);
  return val
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* ==========================================================================
   EXPORT UTILITIES
   ========================================================================== */

/// 1. Export Excel (using SheetJS)
function exportExcel() {
  if (!comparisonResult) return;
  
  try {
    const wb = XLSX.utils.book_new();
    
    // Tab 1: Summary Sheet
    const summaryData = [
      ['Workbook Comparison Summary Report'],
      ['Report Generated:', new Date().toLocaleString()],
      ['Baseline File (File A):', fileAName],
      ['Comparison File (File B):', fileBName],
      [],
      ['Metric', 'Value'],
      ['Sheets Compared', comparisonResult.summary.sheetsCompared],
      ['Matched Sheets', comparisonResult.summary.matchedSheets],
      ['Sheets with Differences', comparisonResult.summary.modifiedSheets],
      ['Missing Sheets (in A only)', comparisonResult.summary.missingSheets],
      ['Additional Sheets (in B only)', comparisonResult.summary.additionalSheets],
      ['Total Cell Differences', comparisonResult.summary.cellDifferences],
      ['Comparison Elapsed Time', comparisonResult.summary.elapsedTime]
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    // Tab 2: Sheet Differences (Missing / Added / Renamed)
    const structureData = [['Baseline Tab Name', 'Comparison Tab Name', 'Status', 'Message']];
    comparisonResult.sheetStatuses.forEach(s => {
      let msg = '';
      if (s.status === 'Match') {
        msg = 'Tabs are matched and identical';
      } else if (s.status === 'Differences') {
        if (s.nameA !== s.nameB) {
          msg = `Tab was renamed from '${s.nameA}' to '${s.nameB}' with value differences`;
        } else {
          msg = `Values in tab '${s.nameA}' are different`;
        }
      } else if (s.status === 'Missing') {
        msg = `Tab '${s.nameA}' is present in File A but missing in File B`;
      } else if (s.status === 'Additional') {
        msg = `Tab '${s.nameB}' is present in File B but missing in File A`;
      }
      structureData.push([s.nameA || '(N/A)', s.nameB || '(N/A)', s.status, msg]);
    });
    const wsStructure = XLSX.utils.aoa_to_sheet(structureData);
    XLSX.utils.book_append_sheet(wb, wsStructure, 'Tab Mapping & Status');

    // Tab 3: Cell Differences
    const cellData = [['Sheet/Tab Map', 'Cell Reference', 'Row', 'Column Letter', 'Expected Value (File A)', 'Actual Value (File B)']];
    for (const sheetName in comparisonResult.sheetDetails) {
      const det = comparisonResult.sheetDetails[sheetName];
      if (det.cells) {
        det.cells.forEach(c => {
          cellData.push([sheetName, c.cellRef, c.row, c.colLetter, c.expected, c.actual]);
        });
      }
    }
    const wsCell = XLSX.utils.aoa_to_sheet(cellData);
    XLSX.utils.book_append_sheet(wb, wsCell, 'Cell Value Diffs');

    // Write file
    XLSX.writeFile(wb, `Comparison_Report_${Date.now()}.xlsx`);
  } catch (err) {
    alert('Failed to generate Excel export: ' + err.message);
  }
}

// 2. Export PDF (using jsPDF & jsPDF-AutoTable)
function exportPdf() {
  if (!comparisonResult) return;
  
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const summary = comparisonResult.summary;
    
    // PDF Styling colors
    const primaryColor = [79, 70, 229]; // Indigo
    const successColor = [16, 185, 129];
    const dangerColor = [239, 68, 68];
    
    // Cover Header Title
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(22);
    doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
    doc.text('Excel Comparison Report', 14, 20);
    
    doc.setFontSize(10);
    doc.setFont('Helvetica', 'normal');
    doc.setTextColor(100);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 27);
    
    // Metadata lines
    doc.setDrawColor(220);
    doc.line(14, 32, 196, 32);
    
    doc.setFontSize(11);
    doc.setFont('Helvetica', 'bold');
    doc.setTextColor(40);
    doc.text('Baseline (File A):', 14, 40);
    doc.setFont('Helvetica', 'normal');
    doc.text(fileAName, 50, 40);
    
    doc.setFont('Helvetica', 'bold');
    doc.text('Comparison (File B):', 14, 46);
    doc.setFont('Helvetica', 'normal');
    doc.text(fileBName, 50, 46);
    
    doc.line(14, 52, 196, 52);
    
    // Summary Cards (Workbook status)
    doc.setFontSize(14);
    doc.setFont('Helvetica', 'bold');
    doc.setTextColor(50);
    doc.text('Workbook Summary Metrics', 14, 62);
    
    // Summary details in simple list or table
    const summaryRows = [
      ['Metric', 'Value'],
      ['Sheets Compared', String(summary.sheetsCompared)],
      ['Matched Sheets', String(summary.matchedSheets)],
      ['Sheets with Differences', String(summary.modifiedSheets)],
      ['Missing Sheets', String(summary.missingSheets)],
      ['Additional Sheets', String(summary.additionalSheets)],
      ['Cell Value Differences', String(summary.cellDifferences)],
      ['Processing Time', summary.elapsedTime]
    ];
    
    doc.autoTable({
      startY: 68,
      head: [summaryRows[0]],
      body: summaryRows.slice(1),
      theme: 'striped',
      headStyles: { fillColor: primaryColor },
      styles: { fontSize: 9 }
    });
    
    // Structure sheet changes if any
    let nextY = doc.lastAutoTable.finalY + 15;
    const structureDiffs = [];
    comparisonResult.sheetStatuses.forEach(s => {
      if (s.status !== 'Match') {
        let renameText = s.displayName;
        let changeDesc = s.nameA && s.nameB && s.nameA !== s.nameB ? 'Renamed' : (s.status === 'Missing' ? 'Deleted' : (s.status === 'Additional' ? 'Added' : 'Modified values'));
        structureDiffs.push([renameText, s.status, changeDesc]);
      }
    });
    
    if (structureDiffs.length > 0) {
      if (nextY > 250) { doc.addPage(); nextY = 20; }
      doc.setFontSize(13);
      doc.setFont('Helvetica', 'bold');
      doc.text('Sheet/Tab Differences', 14, nextY);
      
      doc.autoTable({
        startY: nextY + 5,
        head: [['Sheet Map', 'Status', 'Description']],
        body: structureDiffs,
        theme: 'grid',
        headStyles: { fillColor: dangerColor },
        styles: { fontSize: 8 }
      });
      nextY = doc.lastAutoTable.finalY + 15;
    }
    
    // Add page for details
    doc.addPage();
    nextY = 20;
    
    // Cell Differences details limit 150 records to prevent bloating pdf
    doc.setFontSize(14);
    doc.setFont('Helvetica', 'bold');
    doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
    doc.text('Detailed Cell Differences (First 150 Diffs)', 14, nextY);
    
    const cellRows = [];
    let cellDiffLogged = 0;
    
    for (const sheetName in comparisonResult.sheetDetails) {
      const det = comparisonResult.sheetDetails[sheetName];
      if (det.cells) {
        for (const c of det.cells) {
          if (cellDiffLogged >= 150) break;
          cellRows.push([sheetName, c.cellRef, String(c.row), c.colLetter, c.expected, c.actual]);
          cellDiffLogged++;
        }
      }
      if (cellDiffLogged >= 150) break;
    }
    
    if (cellRows.length > 0) {
      doc.autoTable({
        startY: nextY + 5,
        head: [['Sheet Map', 'Cell', 'Row', 'Column', 'Expected (A)', 'Actual (B)']],
        body: cellRows,
        theme: 'striped',
        headStyles: { fillColor: primaryColor },
        styles: { fontSize: 7, cellPadding: 1 },
        columnStyles: {
          4: { textColor: [220, 50, 50] },
          5: { textColor: [20, 150, 50] }
        }
      });
      
      if (summary.cellDifferences > 150) {
        doc.setFontSize(9);
        doc.setFont('Helvetica', 'italic');
        doc.setTextColor(120);
        doc.text(`* Showing first 150 cell differences out of ${summary.cellDifferences} total. Export to Excel to view all.`, 14, doc.lastAutoTable.finalY + 8);
      }
    } else {
      doc.setFontSize(10);
      doc.setFont('Helvetica', 'normal');
      doc.text('No value cell differences detected in the sheets compared.', 14, nextY + 10);
    }
    
    doc.save(`Comparison_Report_${Date.now()}.pdf`);
  } catch (err) {
    alert('Failed to generate PDF: ' + err.message);
  }
}

// 3. Export CSV consolidated report
function exportCsv() {
  if (!comparisonResult) return;
  
  try {
    let csvContent = 'Sheet/Tab Map,Cell Reference,Row,Column Letter,Expected Value (A),Actual Value (B),Diff Type\n';
    
    for (const sheetName in comparisonResult.sheetDetails) {
      const det = comparisonResult.sheetDetails[sheetName];
      
      if (det.status === 'Missing') {
        csvContent += `"${csvEscape(sheetName)}",,"Sheet Level",,"Sheet Missing in B",,"Tab Missing"\n`;
      } else if (det.status === 'Additional') {
        csvContent += `"${csvEscape(sheetName)}",,"Sheet Level",,,"Sheet Added in B","Tab Added"\n`;
      }
      
      if (det.cells) {
        det.cells.forEach(c => {
          csvContent += `"${csvEscape(sheetName)}","${csvEscape(c.cellRef)}","${csvEscape(c.row)}","${csvEscape(c.colLetter)}","${csvEscape(c.expected)}","${csvEscape(c.actual)}","Cell Mismatch"\n`;
        });
      }
    }
    
    // Create Blob
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `Comparison_Summary_${Date.now()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  } catch (err) {
    alert('Failed to generate CSV: ' + err.message);
  }
}
function csvEscape(str) {
  if (str === undefined || str === null) return '';
  const val = String(str);
  return val.replace(/"/g, '""');
}

/* ==========================================================================
   SIDE-BY-SIDE EXCEL WORKBOOK VIEWER ENGINE
   ========================================================================== */

let excelActiveSheetName = null;

// Synchronized scrolling engine
function setupSynchronizedScrolling() {
  const vpA = document.getElementById('excel-viewport-a');
  const vpB = document.getElementById('excel-viewport-b');
  const syncCheckbox = document.getElementById('excel-sync-scroll');
  
  let isSyncingA = false;
  let isSyncingB = false;
  
  vpA.addEventListener('scroll', () => {
    if (!syncCheckbox.checked) return;
    if (isSyncingB) {
      isSyncingB = false;
      return;
    }
    isSyncingA = true;
    vpB.scrollTop = vpA.scrollTop;
    vpB.scrollLeft = vpA.scrollLeft;
  });
  
  vpB.addEventListener('scroll', () => {
    if (!syncCheckbox.checked) return;
    if (isSyncingA) {
      isSyncingA = false;
      return;
    }
    isSyncingB = true;
    vpA.scrollTop = vpB.scrollTop;
    vpA.scrollLeft = vpB.scrollLeft;
  });
}

// Translate 0-indexed integer to Excel Column letters (0 -> A, 27 -> AB)
function getExcelColLetter(index) {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

// Render the sheet tabs at the bottom and pane headers
function renderExcelView() {
  if (!comparisonResult) return;
  
  // Update pane titles
  document.getElementById('excel-pane-name-a').textContent = fileAName || 'fileA_baseline.xlsx';
  document.getElementById('excel-pane-name-b').textContent = fileBName || 'fileB_comparison.xlsx';
  
  const tabsContainer = document.getElementById('excel-sheet-tabs');
  tabsContainer.innerHTML = '';
  
  const sheetStatuses = comparisonResult.sheetStatuses;
  
  sheetStatuses.forEach((sheet) => {
    let tabClass = 'match';
    if (sheet.status === 'Differences') tabClass = 'diff';
    else if (sheet.status === 'Missing') tabClass = 'missing';
    else if (sheet.status === 'Additional') tabClass = 'additional';
    
    const tab = document.createElement('div');
    tab.className = `excel-tab ${tabClass}`;
    tab.dataset.sheetName = sheet.displayName;
    if (sheet.displayName === excelActiveSheetName) {
      tab.classList.add('active');
    }
    
    tab.innerHTML = `
      <span class="excel-tab-status"></span>
      ${formatTabNameHTML(sheet)}
    `;
    
    tab.addEventListener('click', () => selectExcelSheet(sheet.displayName));
    tabsContainer.appendChild(tab);
  });
  
  // Choose default active tab if none selected
  const hasActive = sheetStatuses.some(s => s.displayName === excelActiveSheetName);
  if (!excelActiveSheetName || !hasActive) {
    if (sheetStatuses.length > 0) {
      selectExcelSheet(sheetStatuses[0].displayName);
    }
  } else {
    renderExcelSheetGrids();
  }
}

// Tab switcher
function selectExcelSheet(displayName) {
  excelActiveSheetName = displayName;
  
  // Set tab active state
  const tabs = document.querySelectorAll('.excel-tab');
  tabs.forEach(tab => {
    if (tab.dataset.sheetName === displayName) {
      tab.classList.add('active');
    } else {
      tab.classList.remove('active');
    }
  });
  
  renderExcelSheetGrids();
}

// Grid generator (dual aligned layout)
function renderExcelSheetGrids() {
  if (!excelActiveSheetName || !comparisonResult) return;
  
  const viewportA = document.getElementById('excel-viewport-a');
  const viewportB = document.getElementById('excel-viewport-b');
  
  const detail = comparisonResult.sheetDetails[excelActiveSheetName];
  const showDiffsOnly = document.getElementById('excel-show-diffs-only').checked;
  
  if (!detail) {
    viewportA.innerHTML = '<div style="padding:2rem; text-align:center; color:var(--text-muted);">No sheet selected.</div>';
    viewportB.innerHTML = '<div style="padding:2rem; text-align:center; color:var(--text-muted);">No sheet selected.</div>';
    return;
  }
  
  // Case 1: Sheet is entirely missing in B
  if (detail.status === 'Missing') {
    viewportA.innerHTML = renderSingleRawSheet(detail, 'baseline');
    viewportB.innerHTML = `
      <div style="padding:3rem; text-align:center; color:var(--color-danger); font-weight:600; background-color:var(--bg-secondary); height:100%; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:0.5rem; border-radius:8px;">
        <i data-lucide="x-circle" style="width:40px; height:40px;"></i>
        <span>Sheet is Missing in Comparison File B</span>
      </div>`;
    initIcons();
    return;
  }
  
  // Case 2: Sheet is entirely additional in B
  if (detail.status === 'Additional') {
    viewportA.innerHTML = `
      <div style="padding:3rem; text-align:center; color:var(--color-primary); font-weight:600; background-color:var(--bg-secondary); height:100%; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:0.5rem; border-radius:8px;">
        <i data-lucide="plus-circle" style="width:40px; height:40px;"></i>
        <span>Sheet is Additional in Comparison File B</span>
      </div>`;
    viewportB.innerHTML = renderSingleRawSheet(detail, 'comparison');
    initIcons();
    return;
  }
  
  // Case 3: Common sheet comparison side-by-side (aligned)
  const rowsToRender = [];
  for (let r = 0; r < detail.maxRows; r++) {
    const hasDiff = detail.diffMask[r] ? detail.diffMask[r].some(val => val === true) : false;
    if (!showDiffsOnly || hasDiff) {
      rowsToRender.push(r);
    }
  }
  
  if (rowsToRender.length === 0) {
    viewportA.innerHTML = '<div style="padding:3rem; text-align:center; color:var(--text-muted); background:var(--bg-secondary); border-radius:8px;">No difference rows to display.</div>';
    viewportB.innerHTML = '<div style="padding:3rem; text-align:center; color:var(--text-muted); background:var(--bg-secondary); border-radius:8px;">No difference rows to display.</div>';
    return;
  }
  
  // Render Baseline Grid Table (Left)
  let htmlA = '<table class="excel-table">';
  htmlA += '<thead><tr><th class="excel-col-letter">Row</th>';
  for (let c = 0; c < detail.maxCols; c++) {
    htmlA += `<th class="excel-col-letter">${getExcelColLetter(c)}</th>`;
  }
  htmlA += '</tr></thead><tbody>';
  
  // Render Comparison Grid Table (Right)
  let htmlB = '<table class="excel-table">';
  htmlB += '<thead><tr><th class="excel-col-letter">Row</th>';
  for (let c = 0; c < detail.maxCols; c++) {
    htmlB += `<th class="excel-col-letter">${getExcelColLetter(c)}</th>`;
  }
  htmlB += '</tr></thead><tbody>';
  
  // Aligned rows rendering loop
  rowsToRender.forEach((r) => {
    // Check row states
    let isRowEmptyA = !detail.gridA[r] || detail.gridA[r].every(val => val === undefined || val === null || val === '');
    let isRowEmptyB = !detail.gridB[r] || detail.gridB[r].every(val => val === undefined || val === null || val === '');
    let isAddedRow = isRowEmptyA && !isRowEmptyB;
    let isMissingRow = !isRowEmptyA && isRowEmptyB;
    
    // Left Grid Cells
    htmlA += '<tr>';
    htmlA += `<td class="excel-row-num">${r + 1}</td>`;
    for (let c = 0; c < detail.maxCols; c++) {
      const valA = (detail.gridA[r] && detail.gridA[r][c] !== undefined) ? detail.gridA[r][c] : '';
      const isDiff = detail.diffMask[r] ? detail.diffMask[r][c] : false;
      
      let cellClass = '';
      if (isMissingRow) {
        cellClass = 'excel-cell-removed';
      } else if (isAddedRow) {
        cellClass = 'excel-cell-added'; // blank placeholder
      } else if (isDiff) {
        cellClass = 'excel-cell-modified';
      }
      
      htmlA += `<td class="excel-cell ${cellClass}">${escapeHtml(valA)}</td>`;
    }
    htmlA += '</tr>';
    
    // Right Grid Cells
    htmlB += '<tr>';
    htmlB += `<td class="excel-row-num">${r + 1}</td>`;
    for (let c = 0; c < detail.maxCols; c++) {
      const valB = (detail.gridB[r] && detail.gridB[r][c] !== undefined) ? detail.gridB[r][c] : '';
      const isDiff = detail.diffMask[r] ? detail.diffMask[r][c] : false;
      
      let cellClass = '';
      if (isMissingRow) {
        cellClass = 'excel-cell-removed'; // blank placeholder
      } else if (isAddedRow) {
        cellClass = 'excel-cell-added';
      } else if (isDiff) {
        cellClass = 'excel-cell-modified';
      }
      
      htmlB += `<td class="excel-cell ${cellClass}">${escapeHtml(valB)}</td>`;
    }
    htmlB += '</tr>';
  });
  
  htmlA += '</tbody></table>';
  htmlB += '</tbody></table>';
  
  viewportA.innerHTML = htmlA;
  viewportB.innerHTML = htmlB;
}

// Render an un-aligned, raw sheet (when sheet is entirely missing or added in B)
function renderSingleRawSheet(detail, fileType) {
  const grid = fileType === 'baseline' ? detail.gridA : detail.gridB;
  const rowsCount = grid.length;
  
  let html = '<table class="excel-table">';
  html += '<thead><tr><th class="excel-col-letter">Row</th>';
  for (let c = 0; c < detail.maxCols; c++) {
    html += `<th class="excel-col-letter">${getExcelColLetter(c)}</th>`;
  }
  html += '</tr></thead><tbody>';
  
  for (let r = 0; r < rowsCount; r++) {
    html += '<tr>';
    html += `<td class="excel-row-num">${r + 1}</td>`;
    for (let c = 0; c < detail.maxCols; c++) {
      const val = (grid[r] && grid[r][c] !== undefined) ? grid[r][c] : '';
      html += `<td class="excel-cell">${escapeHtml(val)}</td>`;
    }
    html += '</tr>';
  }
  
  html += '</tbody></table>';
  return html;
}
