/**
 * PivotTableBuilder for Google Sheets (Apps Script)
 * - Builds a pivot table from any tabular data range into a destination sheet.
 * - Works with headers or direct column references (letter or index).
 * - Uses the Advanced Sheets Service (enable Services > Sheets API in Apps Script).
 *
 * Author: Copilot
 * Notes:
 * - This script relies on the Advanced Sheets service: Services > + > Google Sheets API
 * - The "Sheets" service here refers to Advanced Sheets API, not SpreadsheetApp.
 */

/**
 * Example: Creates a pivot table from a data sheet to a pivot sheet.
 * - Rows: "Region"
 * - Columns: "Status"
 * - Values: Sum of "Amount", Count of "ID"
 */
function createExamplePivot() {
  const config = {
    // Leave spreadsheetId empty to use bound spreadsheet
    spreadsheetId: SpreadsheetApp.getActive().getId(),

    // Source data
    sourceSheet: 'Data',                 // Sheet name containing the data
    sourceRangeA1: 'A1:F',               // Must include header row

    // Destination
    destinationSheet: 'Pivot',           // Will be created if missing
    anchorCellA1: 'A1',                  // Top-left cell for pivot output

    // Define groups using header names, column letters, or 1-based indexes
    rows: [
      { column: 'Region', showTotals: true, sortOrder: 'ASC' }
    ],
    columns: [
      { column: 'Status', showTotals: true, sortOrder: 'ASC' }
    ],
    values: [
      { column: 'Amount', summarizeFunction: 'SUM', name: 'Total Amount' },
      { column: 'ID', summarizeFunction: 'COUNT', name: 'Count of ID' }
    ],

    // Optional: group rules, filters can be added later if needed
  };

  createPivotTable(config);
}

/**
 * Main entry: create a pivot table based on a flexible config.
 *
 * Config schema:
 * {
 *   spreadsheetId?: string (defaults to active spreadsheet)
 *   sourceSheet: string
 *   sourceRangeA1: string  // Must include headers
 *   destinationSheet: string
 *   anchorCellA1?: string   // default "A1"
 *   rows?: Array<{ column: HeaderOrRef, showTotals?: boolean, sortOrder?: 'ASC'|'DESC', groupRule?: object }>
 *   columns?: Array<{ column: HeaderOrRef, showTotals?: boolean, sortOrder?: 'ASC'|'DESC', groupRule?: object }>
 *   values?: Array<{ column: HeaderOrRef, summarizeFunction?: SummarizeFn, name?: string, formula?: string }>
 * }
 *
 * HeaderOrRef: string header name | string column letter like "C" | number 1-based column index
 * SummarizeFn: 'SUM'|'COUNTA'|'COUNT'|'MAX'|'MIN'|'AVERAGE'|'MEDIAN'|'PRODUCT'|'STDEV'|'STDEVP'|'VAR'|'VARP'|'CUSTOM'
 */
function createPivotTable(config) {
  const ss = SpreadsheetApp.getActive();
  const spreadsheetId = config.spreadsheetId || ss.getId();

  // Resolve source range
  const sourceSheet = ss.getSheetByName(config.sourceSheet);
  if (!sourceSheet) {
    throw new Error(`Source sheet "${config.sourceSheet}" not found.`);
  }
  const sourceRange = sourceSheet.getRange(config.sourceRangeA1);
  const headerValues = sourceRange.offset(0, 0, 1, sourceRange.getNumColumns()).getValues()[0];

  // Resolve destination sheet (create if missing)
  const destSheet = ensureSheet_(ss, config.destinationSheet);

  const sourceGridRange = rangeToGridRange_(sourceRange);
  const headerMap = buildHeaderMap_(headerValues); // name -> 0-based offset within sourceRange
  const toOffset = (colRef) => colRefToOffset_(colRef, sourceRange, headerMap);

  // Build PivotGroups (rows/columns)
  const rows = (config.rows || []).map(r => toPivotGroup_(toOffset(r.column), r));
  const columns = (config.columns || []).map(c => toPivotGroup_(toOffset(c.column), c));

  // Build PivotValues
  const values = (config.values || []).map(v => toPivotValue_(toOffset(v.column), v));

  const destSheetId = destSheet.getSheetId();
  const anchorA1 = config.anchorCellA1 || 'A1';
  const anchor = destSheet.getRange(anchorA1);
  const anchorStart = {
    sheetId: destSheetId,
    rowIndex: anchor.getRow() - 1,
    columnIndex: anchor.getColumn() - 1,
  };

  const pivotTable = {
    source: sourceGridRange,
    rows,
    columns,
    values
    // filters: [], // You can extend with filters via the Sheets API if needed
  };

  const requests = [
    {
      updateCells: {
        start: anchorStart,
        rows: [{ values: [{ pivotTable }] }],
        fields: 'pivotTable'
      }
    }
  ];

  Sheets.Spreadsheets.batchUpdate(
    { requests },
    spreadsheetId
  );
}

/**
 * Ensures a sheet exists by name; creates it if missing.
 */
function ensureSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Convert SpreadsheetApp Range to Sheets API GridRange (0-based, end-exclusive).
 */
function rangeToGridRange_(range) {
  const sheet = range.getSheet();
  const sheetId = sheet.getSheetId();
  const startRowIndex = range.getRow() - 1;
  const startColumnIndex = range.getColumn() - 1;
  const endRowIndex = startRowIndex + range.getNumRows();
  const endColumnIndex = startColumnIndex + range.getNumColumns();
  return {
    sheetId,
    startRowIndex,
    endRowIndex,
    startColumnIndex,
    endColumnIndex
  };
}

/**
 * Build header name -> 0-based offset map.
 */
function buildHeaderMap_(headerRowValues) {
  const map = {};
  for (let i = 0; i < headerRowValues.length; i++) {
    const name = String(headerRowValues[i] || '').trim();
    if (name) {
      map[name] = i;
    }
  }
  return map;
}

/**
 * Convert a column reference to 0-based offset within the source range.
 * Accepts:
 * - header name: "Amount"
 * - column letter: "C" or "AA"
 * - 1-based index: 3
 */
function colRefToOffset_(colRef, sourceRange, headerMap) {
  if (colRef == null) throw new Error('Column reference is required.');
  let absoluteColIndex; // 1-based absolute column index in sheet

  if (typeof colRef === 'number') {
    if (colRef < 1) throw new Error(`Column index must be 1-based. Got ${colRef}`);
    absoluteColIndex = colRef;
  } else if (typeof colRef === 'string') {
    const trimmed = colRef.trim();
    if (headerMap.hasOwnProperty(trimmed)) {
      // header name
      const offset = headerMap[trimmed];
      return offset;
    }
    // Try as column letters
    const m = /^[A-Za-z]+$/.test(trimmed);
    if (m) {
      absoluteColIndex = letterToColIndex1_(trimmed);
    } else {
      throw new Error(`Unknown column reference "${colRef}". Use header name, column letter (e.g., "C"), or 1-based index.`);
    }
  } else {
    throw new Error(`Unsupported column reference type: ${typeof colRef}`);
  }

  const sourceStartCol = sourceRange.getColumn(); // 1-based absolute
  const offset = absoluteColIndex - sourceStartCol;
  if (offset < 0 || offset >= sourceRange.getNumColumns()) {
    throw new Error(`Column reference "${colRef}" is outside the source range ${sourceRange.getA1Notation()}.`);
  }
  return offset;
}

/**
 * Convert column letters to 1-based index. e.g., A=1, Z=26, AA=27
 */
function letterToColIndex1_(letters) {
  let n = 0;
  const s = letters.toUpperCase();
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64); // 'A' => 65
  }
  return n;
}

/**
 * Build a PivotGroup object.
 */
function toPivotGroup_(offset, spec) {
  const pg = {
    sourceColumnOffset: offset,
    showTotals: spec.showTotals !== false, // default true
    sortOrder: normalizeSortOrder_(spec.sortOrder)
  };
  if (spec.groupRule) {
    pg.groupRule = spec.groupRule;
  }
  return pg;
}

/**
 * Build a PivotValue object.
 */
function toPivotValue_(offset, spec) {
  const pv = {
    sourceColumnOffset: offset,
    summarizeFunction: normalizeSummarizeFunction_(spec.summarizeFunction),
  };
  if (spec.name) pv.name = spec.name;
  if (spec.formula) pv.formula = spec.formula; // Use with summarizeFunction: 'CUSTOM'
  return pv;
}

function normalizeSortOrder_(order) {
  if (!order) return 'ASCENDING';
  const v = String(order).toUpperCase();
  if (v === 'ASC' || v === 'ASCENDING') return 'ASCENDING';
  if (v === 'DESC' || v === 'DESCENDING') return 'DESCENDING';
  throw new Error(`Unsupported sortOrder "${order}". Use 'ASC' or 'DESC'.`);
}

function normalizeSummarizeFunction_(fn) {
  if (!fn) return 'SUM';
  const v = String(fn).toUpperCase();
  const allowed = new Set(['SUM','COUNTA','COUNT','MAX','MIN','AVERAGE','MEDIAN','PRODUCT','STDEV','STDEVP','VAR','VARP','CUSTOM']);
  if (!allowed.has(v)) {
    throw new Error(`Unsupported summarizeFunction "${fn}". Allowed: ${Array.from(allowed).join(', ')}`);
  }
  return v;
}

/**
 * Optional: Add a simple menu to run the example quickly.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pivot Tools')
    .addItem('Create Example Pivot', 'createExamplePivot')
    .addToUi();
}
