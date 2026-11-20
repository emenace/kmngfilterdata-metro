const XLSX = require('xlsx');
const { DEFAULT_CONFIG } = require('../config');

/**
 * Error class for client-visible upload/schema validation failures (HTTP 400).
 */
class ValidationError extends Error {
  constructor(message) {
    super(message);
    this.name = 'ValidationError';
    this.statusCode = 400;
  }
}

/**
 * Extract a cell's canonical exact string and raw descriptor without mutating content.
 */
function extractCell(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell || cell.v === undefined || cell.v === null) {
    return { canonical: '', raw: null };
  }

  if (cell.t === 's' || cell.t === 'str') {
    const val = String(cell.v);
    return { canonical: val, raw: { t: 's', v: val } };
  }

  if (cell.t === 'n') {
    const canonical = cell.w !== undefined ? String(cell.w) : String(cell.v);
    const raw = { t: 'n', v: cell.v };
    if (cell.z) {
      raw.z = cell.z;
    }
    return { canonical, raw };
  }

  if (cell.t === 'b') {
    const boolVal = Boolean(cell.v);
    return { canonical: boolVal ? 'TRUE' : 'FALSE', raw: { t: 'b', v: boolVal } };
  }

  if (cell.t === 'e') {
    const errText = cell.w !== undefined ? String(cell.w) : String(cell.v);
    return { canonical: errText, raw: { t: 's', v: errText } };
  }

  const fallback = cell.w !== undefined ? String(cell.w) : String(cell.v);
  return { canonical: fallback, raw: { t: 's', v: fallback } };
}

/**
 * Parse an uploaded Excel buffer (.xls or .xlsx), filter rows where SATKER_3 contains
 * "Kantor Kementerian Agama Kota Metro", and remove all KODE_* columns without altering any cell value.
 */
function processExcelBuffer(buffer, originalFilename, options = {}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new ValidationError('Uploaded file is empty.');
  }

  const filterColumn = options.filterColumn || DEFAULT_CONFIG.FILTER_COLUMN;
  const targetSatker = options.targetSatker || DEFAULT_CONFIG.TARGET_SATKER;
  const removeColumnPattern = options.removeColumnPattern || DEFAULT_CONFIG.REMOVE_COLUMN_PATTERN;

  let workbook;
  try {
    workbook = XLSX.read(buffer, {
      type: 'buffer',
      raw: true,
      cellDates: false,
      cellNF: true,
      cellStyles: false,
      cellFormula: false,
    });
  } catch (err) {
    throw new ValidationError(`Failed to read Excel file "${originalFilename}": ${err.message}`);
  }

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new ValidationError('The uploaded Excel file does not contain any worksheets.');
  }

  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws['!ref']) {
    throw new ValidationError(`Worksheet "${sheetName}" is empty.`);
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  if (range.e.r < range.s.r || range.e.c < range.s.c) {
    throw new ValidationError(`Worksheet "${sheetName}" contains no readable cells.`);
  }

  // Extract header row (first row in range).
  const headerRowIndex = range.s.r;
  const headerCols = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const { canonical } = extractCell(ws, headerRowIndex, c);
    headerCols.push({
      colIndex: c,
      colLetter: XLSX.utils.encode_col(c),
      name: canonical,
    });
  }

  // Trim trailing phantom columns whose header and all data cells are empty.
  while (headerCols.length > 0 && headerCols[headerCols.length - 1].name === '') {
    const lastCol = headerCols[headerCols.length - 1].colIndex;
    let hasAnyValue = false;
    for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
      if (extractCell(ws, r, lastCol).canonical !== '') {
        hasAnyValue = true;
        break;
      }
    }
    if (hasAnyValue) {
      break;
    }
    headerCols.pop();
  }

  if (headerCols.length === 0) {
    throw new ValidationError('No header columns were found in the first row.');
  }

  const normalizedTargetHeader = filterColumn.trim().toUpperCase();
  const satkerCol = headerCols.find((col) => col.name.trim().toUpperCase() === normalizedTargetHeader);
  if (!satkerCol) {
    throw new ValidationError(
      `Column "${filterColumn}" was not found in the header row of "${sheetName}".`
    );
  }

  const removedCols = [];
  const keptCols = [];
  for (const col of headerCols) {
    if (removeColumnPattern.test(col.name.trim())) {
      removedCols.push(col);
    } else {
      keptCols.push(col);
    }
  }

  if (keptCols.length === 0) {
    throw new ValidationError('No columns remain after removing KODE_* columns.');
  }

  const needle = targetSatker.toLowerCase();
  const rows = [];
  const rawRows = [];
  let totalSourceRows = 0;

  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    // Skip completely empty rows (all cells empty across source columns).
    let rowHasContent = false;
    for (let i = 0; i < headerCols.length; i += 1) {
      if (extractCell(ws, r, headerCols[i].colIndex).canonical !== '') {
        rowHasContent = true;
        break;
      }
    }
    if (!rowHasContent) {
      continue;
    }

    totalSourceRows += 1;

    const satkerCell = extractCell(ws, r, satkerCol.colIndex);
    if (!satkerCell.canonical.toLowerCase().includes(needle)) {
      continue;
    }

    const canonicalRow = new Array(keptCols.length);
    const rawRow = new Array(keptCols.length);
    for (let i = 0; i < keptCols.length; i += 1) {
      const extracted = extractCell(ws, r, keptCols[i].colIndex);
      canonicalRow[i] = extracted.canonical;
      rawRow[i] = extracted.raw;
    }

    rows.push(canonicalRow);
    rawRows.push(rawRow);
  }

  return {
    sheetName,
    filterColumn: {
      name: satkerCol.name,
      letter: satkerCol.colLetter,
    },
    targetSatker,
    totalSourceRows,
    filteredRowCount: rows.length,
    originalColumnCount: headerCols.length,
    removedColumns: removedCols.map((c) => c.name),
    keptColumns: keptCols.map((c) => c.name),
    rows,
    rawRows,
  };
}

module.exports = {
  ValidationError,
  extractCell,
  processExcelBuffer,
};

