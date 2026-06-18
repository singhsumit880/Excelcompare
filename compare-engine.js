// compare-engine.js - Core comparison logic (tab-by-tab and cell-by-cell)

function compareWorkbooks(fileABuffer, fileBBuffer, config, onProgress) {
  if (typeof onProgress !== 'function') {
    onProgress = () => {};
  }
  
  onProgress('Parsing Baseline File (File A)...');
  const wbA = XLSX.read(fileABuffer, { type: 'array', cellDates: true, cellNF: true, cellText: true });
  
  onProgress('Parsing Comparison File (File B)...');
  const wbB = XLSX.read(fileBBuffer, { type: 'array', cellDates: true, cellNF: true, cellText: true });
  
  onProgress('Analyzing workbook sheets by tab position...');
  
  const sheetsA = wbA.SheetNames;
  const sheetsB = wbB.SheetNames;
  
  const maxSheets = Math.max(sheetsA.length, sheetsB.length);
  
  const sheetStatuses = [];
  const sheetDetails = {};
  
  let totalSheetsCompared = 0;
  let totalMatchedSheets = 0;
  let totalModifiedSheets = 0;
  let totalMissingSheets = 0;
  let totalAdditionalSheets = 0;
  let totalCellDifferences = 0;

  // Helper to check if value is empty
  function isEmpty(val) {
    return val === undefined || val === null || val === '';
  }

  // Translate index to Excel Column letters (0 -> A, 27 -> AB)
  function getExcelColLetter(index) {
    let letter = '';
    let temp = index;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }

  // Normalization helper
  function normalizeValue(val, rules) {
    if (isEmpty(val)) return '';
    
    if (val instanceof Date) {
      if (rules.ignoreDateFormatting) {
        return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
      }
      return val.toISOString();
    }
    
    let str = String(val);
    
    if (rules.ignoreWhitespace) {
      str = str.trim().replace(/\s+/g, ' ');
    }
    
    if (rules.ignoreCase) {
      str = str.toLowerCase();
    }
    
    if (rules.ignoreDateFormatting) {
      // Inline timezone-insensitive date parsing
      const clean = str.trim();
      if (!clean.includes(':')) {
        const parsed = Date.parse(clean);
        if (!isNaN(parsed)) {
          const d = new Date(parsed);
          const isIso = /^\d{4}-\d{2}-\d{2}$/.test(clean);
          if (isIso) {
            return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
          } else {
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          }
        }
      }
    }
    
    const num = Number(str);
    if (!isNaN(num) && str.trim() !== '') {
      return num;
    }
    
    return str;
  }

  // Compare Tab by Tab based on index
  for (let i = 0; i < maxSheets; i++) {
    const nameA = sheetsA[i];
    const nameB = sheetsB[i];
    
    let status = 'Match';
    let displayName = '';
    
    // Case 1: Tab exists in both files (compared by position)
    if (nameA && nameB) {
      totalSheetsCompared++;
      onProgress(`Comparing Tab ${i + 1}: ${nameA} vs ${nameB}...`);
      
      const sheetA = wbA.Sheets[nameA];
      const sheetB = wbB.Sheets[nameB];
      
      const dataA = XLSX.utils.sheet_to_json(sheetA, { header: 1, defval: '' });
      const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1, defval: '' });
      
      // Filter out empty rows if config specifies
      let rowsA = dataA;
      let rowsB = dataB;
      if (config.ignoreEmptyRows) {
        rowsA = rowsA.filter(row => row.some(cell => !isEmpty(cell)));
        rowsB = rowsB.filter(row => row.some(cell => !isEmpty(cell)));
      }
      
      const maxRows = Math.max(rowsA.length, rowsB.length);
      
      // Find maximum columns in any row
      let maxCols = 0;
      for (let r = 0; r < maxRows; r++) {
        const lenA = (rowsA[r] || []).length;
        const lenB = (rowsB[r] || []).length;
        maxCols = Math.max(maxCols, lenA, lenB);
      }
      
      const gridA = [];
      const gridB = [];
      const diffMask = [];
      const cells = [];
      
      let sheetHasDiff = nameA !== nameB; // Different tab name is a difference
      
      for (let r = 0; r < maxRows; r++) {
        const rowA = rowsA[r] || [];
        const rowB = rowsB[r] || [];
        
        const gridRowA = [];
        const gridRowB = [];
        const diffRowMask = [];
        
        for (let c = 0; c < maxCols; c++) {
          const valA = rowA[c] !== undefined ? rowA[c] : '';
          const valB = rowB[c] !== undefined ? rowB[c] : '';
          
          gridRowA.push(valA);
          gridRowB.push(valB);
          
          // Compare normalized cell values
          const normA = normalizeValue(valA, config);
          const normB = normalizeValue(valB, config);
          
          const isDiff = normA !== normB;
          diffRowMask.push(isDiff);
          
          if (isDiff) {
            sheetHasDiff = true;
            totalCellDifferences++;
            const cellRef = XLSX.utils.encode_cell({ c: c, r: r });
            cells.push({
              cellRef,
              row: r + 1,
              col: c,
              colLetter: getExcelColLetter(c),
              expected: isEmpty(valA) ? '(Empty)' : String(valA),
              actual: isEmpty(valB) ? '(Empty)' : String(valB)
            });
          }
        }
        
        gridA.push(gridRowA);
        gridB.push(gridRowB);
        diffMask.push(diffRowMask);
      }
      
      if (sheetHasDiff) {
        status = 'Differences';
        totalModifiedSheets++;
      } else {
        totalMatchedSheets++;
      }
      
      displayName = nameA === nameB ? nameA : `${nameA} ➔ ${nameB}`;
      sheetStatuses.push({ nameA, nameB, status, displayName });
      
      sheetDetails[displayName] = {
        status,
        nameA,
        nameB,
        maxRows,
        maxCols,
        cells,
        gridA,
        gridB,
        diffMask
      };
      
    } 
    // Case 2: Tab exists in File A but not File B (Missing in B)
    else if (nameA && !nameB) {
      totalMissingSheets++;
      displayName = `${nameA} ➔ (Deleted)`;
      sheetStatuses.push({ nameA, nameB: null, status: 'Missing', displayName });
      
      const sheetA = wbA.Sheets[nameA];
      const dataA = XLSX.utils.sheet_to_json(sheetA, { header: 1, defval: '' });
      let rowsA = dataA;
      if (config.ignoreEmptyRows) {
        rowsA = rowsA.filter(row => row.some(cell => !isEmpty(cell)));
      }
      
      const maxRows = rowsA.length;
      let maxCols = 0;
      rowsA.forEach(row => { maxCols = Math.max(maxCols, row.length); });
      
      const gridA = [];
      const gridB = [];
      const diffMask = [];
      
      for (let r = 0; r < maxRows; r++) {
        const row = rowsA[r] || [];
        const gridRowA = [];
        const gridRowB = [];
        const diffRowMask = [];
        
        for (let c = 0; c < maxCols; c++) {
          const valA = row[c] !== undefined ? row[c] : '';
          gridRowA.push(valA);
          gridRowB.push('');
          diffRowMask.push(true); // All cells are differences
        }
        gridA.push(gridRowA);
        gridB.push(gridRowB);
        diffMask.push(diffRowMask);
      }
      
      sheetDetails[displayName] = {
        status: 'Missing',
        nameA,
        nameB: null,
        maxRows,
        maxCols,
        cells: [],
        gridA,
        gridB,
        diffMask
      };
    } 
    // Case 3: Tab exists in File B but not File A (Additional in B)
    else if (!nameA && nameB) {
      totalAdditionalSheets++;
      displayName = `(Added) ➔ ${nameB}`;
      sheetStatuses.push({ nameA: null, nameB, status: 'Additional', displayName });
      
      const sheetB = wbB.Sheets[nameB];
      const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1, defval: '' });
      let rowsB = dataB;
      if (config.ignoreEmptyRows) {
        rowsB = rowsB.filter(row => row.some(cell => !isEmpty(cell)));
      }
      
      const maxRows = rowsB.length;
      let maxCols = 0;
      rowsB.forEach(row => { maxCols = Math.max(maxCols, row.length); });
      
      const gridA = [];
      const gridB = [];
      const diffMask = [];
      
      for (let r = 0; r < maxRows; r++) {
        const row = rowsB[r] || [];
        const gridRowA = [];
        const gridRowB = [];
        const diffRowMask = [];
        
        for (let c = 0; c < maxCols; c++) {
          const valB = row[c] !== undefined ? row[c] : '';
          gridRowA.push('');
          gridRowB.push(valB);
          diffRowMask.push(true); // All cells are differences
        }
        gridA.push(gridRowA);
        gridB.push(gridRowB);
        diffMask.push(diffRowMask);
      }
      
      sheetDetails[displayName] = {
        status: 'Additional',
        nameA: null,
        nameB,
        maxRows,
        maxCols,
        cells: [],
        gridA,
        gridB,
        diffMask
      };
    }
  }
  
  return {
    summary: {
      sheetsCompared: totalSheetsCompared,
      matchedSheets: totalMatchedSheets,
      modifiedSheets: totalModifiedSheets,
      missingSheets: totalMissingSheets,
      additionalSheets: totalAdditionalSheets,
      cellDifferences: totalCellDifferences
    },
    sheetStatuses,
    sheetDetails
  };
}
