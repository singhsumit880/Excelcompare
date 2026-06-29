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

  // Universal Column & Row Alignment Algorithms (ts-excel-compare model)
  function getColumnSlots(rowsA, rowsB) {
    let maxCols = 0;
    for (let r = 0; r < rowsA.length; r++) {
      if (rowsA[r] && rowsA[r].length > maxCols) maxCols = rowsA[r].length;
    }
    for (let r = 0; r < rowsB.length; r++) {
      if (rowsB[r] && rowsB[r].length > maxCols) maxCols = rowsB[r].length;
    }

    const slots = [];
    for (let c = 0; c < maxCols; c++) {
      const letter = getExcelColLetter(c);
      slots.push({
        type: 'match',
        colA: c,
        colB: c,
        labelA: letter,
        labelB: letter,
        mapText: `${letter}:${letter}`
      });
    }
    return slots;
  }

  function alignRows(rowsA, rowsB, rules, normFn) {
    const lenA = rowsA.length;
    const lenB = rowsB.length;
    if (lenA === 0 && lenB === 0) return [];

    function getRowSig(row) {
      if (!row) return '';
      return row.map(val => String(normFn(val, rules))).join('|');
    }

    function getRowKey(row) {
      if (!row) return '';
      for (let i = 0; i < row.length; i++) {
        if (row[i] !== undefined && row[i] !== null && row[i] !== '') {
          const norm = normFn(row[i], rules);
          if (!isEmpty(norm)) return String(norm);
        }
      }
      return '';
    }

    const sigsA = rowsA.map(getRowSig);
    const sigsB = rowsB.map(getRowSig);
    const keysA = rowsA.map(getRowKey);
    const keysB = rowsB.map(getRowKey);

    const dp = Array.from({ length: lenA + 1 }, () => Array(lenB + 1).fill(0));
    for (let i = 0; i < lenA; i++) {
      for (let j = 0; j < lenB; j++) {
        let score = 0;
        if (sigsA[i] !== '' && sigsA[i] === sigsB[j]) score = 2;
        else if (keysA[i] !== '' && keysA[i] === keysB[j]) score = 1;

        if (score > 0) {
          dp[i + 1][j + 1] = dp[i][j] + score;
        } else {
          dp[i + 1][j + 1] = Math.max(dp[i + 1][j], dp[i][j + 1]);
        }
      }
    }

    let i = lenA, j = lenB;
    const matched = [];
    while (i > 0 && j > 0) {
      let score = 0;
      if (sigsA[i - 1] !== '' && sigsA[i - 1] === sigsB[j - 1]) score = 2;
      else if (keysA[i - 1] !== '' && keysA[i - 1] === keysB[j - 1]) score = 1;

      if (score > 0 && dp[i][j] === dp[i - 1][j - 1] + score) {
        matched.unshift({ rowA: i - 1, rowB: j - 1 });
        i--; j--;
      } else if (dp[i - 1][j] >= dp[i][j - 1]) {
        i--;
      } else {
        j--;
      }
    }

    if (matched.length === 0 && lenA > 0 && lenB > 0) {
      const max = Math.max(lenA, lenB);
      const slots = [];
      for (let k = 0; k < max; k++) {
        const rA = k < lenA ? k : null;
        const rB = k < lenB ? k : null;
        let type = 'match';
        if (rA !== null && rB === null) type = 'deleted';
        if (rA === null && rB !== null) type = 'added';
        const labelA = rA !== null ? String(rA + 1) : '-';
        const labelB = rB !== null ? String(rB + 1) : '-';
        let mapText = `${labelA}:${labelB}`;
        if (type === 'added') mapText = `-:${labelB}`;
        if (type === 'deleted') mapText = `${labelA}:-`;
        slots.push({ type, rowA: rA, rowB: rB, labelA, labelB, mapText });
      }
      return slots;
    }

    const slots = [];
    let curA = 0, curB = 0;

    for (const pair of matched) {
      while (curA < pair.rowA) {
        const labelA = String(curA + 1);
        slots.push({ type: 'deleted', rowA: curA, rowB: null, labelA, labelB: '-', mapText: `${labelA}:-` });
        curA++;
      }
      while (curB < pair.rowB) {
        const labelB = String(curB + 1);
        slots.push({ type: 'added', rowA: null, rowB: curB, labelA: '-', labelB, mapText: `-:${labelB}` });
        curB++;
      }
      const labelA = String(pair.rowA + 1);
      const labelB = String(pair.rowB + 1);
      slots.push({ type: 'match', rowA: pair.rowA, rowB: pair.rowB, labelA, labelB, mapText: `${labelA}:${labelB}` });
      curA++; curB++;
    }

    while (curA < lenA) {
      const labelA = String(curA + 1);
      slots.push({ type: 'deleted', rowA: curA, rowB: null, labelA, labelB: '-', mapText: `${labelA}:-` });
      curA++;
    }
    while (curB < lenB) {
      const labelB = String(curB + 1);
      slots.push({ type: 'added', rowA: null, rowB: curB, labelA: '-', labelB, mapText: `-:${labelB}` });
      curB++;
    }

    return slots;
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
      
      // Standardize Columns mapping
      const columnSlots = getColumnSlots(rowsA, rowsB);
      const maxCols = columnSlots.length;
      
      // Align Rows dynamically via signatures
      const rowSlots = alignRows(rowsA, rowsB, config, normalizeValue);
      const maxRows = rowSlots.length;

      const gridA = [];
      const gridB = [];
      const diffMask = [];
      const cells = [];
      
      let sheetHasDiff = nameA !== nameB; // Different tab name is a difference
      
      for (let r = 0; r < maxRows; r++) {
        const rowSlot = rowSlots[r];
        const rowA = rowSlot.rowA !== null ? (rowsA[rowSlot.rowA] || []) : [];
        const rowB = rowSlot.rowB !== null ? (rowsB[rowSlot.rowB] || []) : [];
        
        const gridRowA = [];
        const gridRowB = [];
        const diffRowMask = [];
        
        for (let c = 0; c < maxCols; c++) {
          const slot = columnSlots[c];
          const valA = (slot.colA !== null && rowA[slot.colA] !== undefined) ? rowA[slot.colA] : '';
          const valB = (slot.colB !== null && rowB[slot.colB] !== undefined) ? rowB[slot.colB] : '';
          
          gridRowA.push(valA);
          gridRowB.push(valB);
          
          let isDiff = false;
          if (rowSlot.type === 'added') {
            isDiff = !isEmpty(valB);
          } else if (rowSlot.type === 'deleted') {
            isDiff = !isEmpty(valA);
          } else {
            const normA = normalizeValue(valA, config);
            const normB = normalizeValue(valB, config);
            isDiff = normA !== normB;
          }
          
          diffRowMask.push(isDiff);
          
          if (isDiff) {
            sheetHasDiff = true;
            totalCellDifferences++;
            const colLet = slot.labelB !== '-' ? slot.labelB : slot.labelA;
            const displayRow = rowSlot.labelB !== '-' ? rowSlot.labelB : rowSlot.labelA;
            const cellRef = `${colLet}${displayRow}`;
            cells.push({
              cellRef,
              row: displayRow,
              col: c,
              colLetter: colLet,
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
        diffMask,
        columnSlots,
        rowSlots
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
