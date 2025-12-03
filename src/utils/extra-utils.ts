/**
 * Utility functions for ExcelTS
 * Provides convenient helper functions for common spreadsheet operations
 * compatible with xlsx library's utils API but built on excelts native types
 */

import { Workbook } from "../doc/workbook.js";
import type { Worksheet } from "../doc/worksheet.js";
import { colCache } from "./col-cache.js";
import type { CellValue } from "../types.js";

// =============================================================================
// Types
// =============================================================================

/**
 * Cell address object (0-indexed)
 */
export interface CellAddress {
  /** 0-indexed column number */
  c: number;
  /** 0-indexed row number */
  r: number;
}

/**
 * Range object with start and end addresses
 */
export interface Range {
  /** Start cell (top-left) */
  s: CellAddress;
  /** End cell (bottom-right) */
  e: CellAddress;
}

/**
 * Row data for JSON conversion
 */
export type JSONRow = Record<string, CellValue>;

// =============================================================================
// Cell Address Encoding/Decoding
// =============================================================================

/**
 * Decode column string to 0-indexed number
 * @example decode_col("A") => 0, decode_col("Z") => 25, decode_col("AA") => 26
 */
export function decode_col(colstr: string): number {
  return colCache.l2n(colstr.toUpperCase()) - 1;
}

/**
 * Encode 0-indexed column number to string
 * @example encode_col(0) => "A", encode_col(25) => "Z", encode_col(26) => "AA"
 */
export function encode_col(col: number): string {
  return colCache.n2l(col + 1);
}

/**
 * Decode row string to 0-indexed number
 * @example decode_row("1") => 0, decode_row("10") => 9
 */
export function decode_row(rowstr: string): number {
  return parseInt(rowstr, 10) - 1;
}

/**
 * Encode 0-indexed row number to string
 * @example encode_row(0) => "1", encode_row(9) => "10"
 */
export function encode_row(row: number): string {
  return String(row + 1);
}

/**
 * Decode cell address string to CellAddress object
 * @example decode_cell("A1") => {c: 0, r: 0}, decode_cell("B2") => {c: 1, r: 1}
 */
export function decode_cell(cstr: string): CellAddress {
  const addr = colCache.decodeAddress(cstr.toUpperCase());
  return { c: addr.col - 1, r: addr.row - 1 };
}

/**
 * Encode CellAddress object to cell address string
 * @example encode_cell({c: 0, r: 0}) => "A1", encode_cell({c: 1, r: 1}) => "B2"
 */
export function encode_cell(cell: CellAddress): string {
  return colCache.encodeAddress(cell.r + 1, cell.c + 1);
}

/**
 * Decode range string to Range object
 * @example decode_range("A1:B2") => {s: {c: 0, r: 0}, e: {c: 1, r: 1}}
 */
export function decode_range(range: string): Range {
  const idx = range.indexOf(":");
  if (idx === -1) {
    const cell = decode_cell(range);
    return { s: cell, e: { ...cell } };
  }
  return {
    s: decode_cell(range.slice(0, idx)),
    e: decode_cell(range.slice(idx + 1))
  };
}

/**
 * Encode Range object to range string
 */
export function encode_range(range: Range): string;
export function encode_range(start: CellAddress, end: CellAddress): string;
export function encode_range(startOrRange: CellAddress | Range, end?: CellAddress): string {
  if (end === undefined) {
    const range = startOrRange as Range;
    return encode_range(range.s, range.e);
  }
  const start = startOrRange as CellAddress;
  const startStr = encode_cell(start);
  const endStr = encode_cell(end);
  return startStr === endStr ? startStr : `${startStr}:${endStr}`;
}

// =============================================================================
// Sheet/JSON Conversion
// =============================================================================

/** Origin can be cell address string, cell object, row number, or -1 to append */
export type Origin = string | CellAddress | number;

export interface JSON2SheetOpts {
  /** Use specified field order (default Object.keys) */
  header?: string[];
  /** Use specified date format in string output */
  dateNF?: string;
  /** Store dates as type d (default is n) */
  cellDates?: boolean;
  /** If true, do not include header row in output */
  skipHeader?: boolean;
  /** If true, emit #NULL! error cells for null values */
  nullError?: boolean;
}

export interface SheetAddJSONOpts extends JSON2SheetOpts {
  /** Use specified cell as starting point */
  origin?: Origin;
}

/**
 * Create a worksheet from an array of objects (xlsx compatible)
 * @example
 * const ws = json_to_sheet([{name: "Alice", age: 30}, {name: "Bob", age: 25}])
 */
export function json_to_sheet(data: JSONRow[], opts?: JSON2SheetOpts): Worksheet {
  const o = opts || {};
  // Create a temporary workbook to get a worksheet
  const tempWb = new Workbook();
  const worksheet = tempWb.addWorksheet("Sheet1");

  if (data.length === 0) {
    return worksheet;
  }

  // Determine headers - use provided header or Object.keys from first object
  const allKeys = new Set<string>();
  data.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));
  const headers = o.header ? [...o.header] : [...allKeys];

  // Add any missing keys from data that weren't in header
  if (o.header) {
    allKeys.forEach(k => {
      if (!headers.includes(k)) {
        headers.push(k);
      }
    });
  }

  let rowNum = 1;

  // Write header row
  if (!o.skipHeader) {
    headers.forEach((h, colIdx) => {
      worksheet.getCell(rowNum, colIdx + 1).value = h;
    });
    rowNum++;
  }

  // Write data rows
  for (const row of data) {
    headers.forEach((key, colIdx) => {
      const val = row[key];
      if (val === null && o.nullError) {
        worksheet.getCell(rowNum, colIdx + 1).value = { error: "#NULL!" };
      } else if (val !== undefined && val !== null) {
        worksheet.getCell(rowNum, colIdx + 1).value = val;
      }
    });
    rowNum++;
  }

  return worksheet;
}

/**
 * Add data from an array of objects to an existing worksheet (xlsx compatible)
 */
export function sheet_add_json(
  worksheet: Worksheet,
  data: JSONRow[],
  opts?: SheetAddJSONOpts
): Worksheet {
  const o = opts || {};

  if (data.length === 0) {
    return worksheet;
  }

  // Determine starting position
  let startRow = 1;
  let startCol = 1;

  if (o.origin !== undefined) {
    if (typeof o.origin === "string") {
      const addr = decode_cell(o.origin);
      startRow = addr.r + 1;
      startCol = addr.c + 1;
    } else if (typeof o.origin === "number") {
      if (o.origin === -1) {
        // Append to bottom
        startRow = worksheet.rowCount + 1;
      } else {
        startRow = o.origin + 1; // 0-indexed row
      }
    } else {
      startRow = o.origin.r + 1;
      startCol = o.origin.c + 1;
    }
  }

  // Determine headers
  const allKeys = new Set<string>();
  data.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));
  const headers = o.header ? [...o.header] : [...allKeys];

  if (o.header) {
    allKeys.forEach(k => {
      if (!headers.includes(k)) {
        headers.push(k);
      }
    });
  }

  let rowNum = startRow;

  // Write header row
  if (!o.skipHeader) {
    headers.forEach((h, colIdx) => {
      worksheet.getCell(rowNum, startCol + colIdx).value = h;
    });
    rowNum++;
  }

  // Write data rows
  for (const row of data) {
    headers.forEach((key, colIdx) => {
      const val = row[key];
      if (val === null && o.nullError) {
        worksheet.getCell(rowNum, startCol + colIdx).value = { error: "#NULL!" };
      } else if (val !== undefined && val !== null) {
        worksheet.getCell(rowNum, startCol + colIdx).value = val;
      }
    });
    rowNum++;
  }

  return worksheet;
}

export interface Sheet2JSONOpts {
  /**
   * Control output format:
   * - 1: Generate an array of arrays
   * - "A": Row object keys are literal column labels (A, B, C, ...)
   * - string[]: Use specified strings as keys in row objects
   * - undefined: Read and disambiguate first row as keys
   */
  header?: 1 | "A" | string[];
  /**
   * Override Range:
   * - number: Use worksheet range but set starting row to the value
   * - string: Use specified range (A1-Style bounded range string)
   * - undefined: Use worksheet range
   */
  range?: number | string;
  /** Use raw values (true) or formatted strings (false). Default: true */
  raw?: boolean;
  /** Default value for empty cells */
  defval?: CellValue;
  /** Include blank lines in the output */
  blankrows?: boolean;
}

/**
 * Convert worksheet to JSON array (xlsx compatible)
 * @example
 * // Default: array of objects with first row as headers
 * const data = sheet_to_json(worksheet)
 * // => [{name: "Alice", age: 30}, {name: "Bob", age: 25}]
 *
 * // Array of arrays
 * const aoa = sheet_to_json(worksheet, { header: 1 })
 * // => [["name", "age"], ["Alice", 30], ["Bob", 25]]
 *
 * // Column letters as keys
 * const cols = sheet_to_json(worksheet, { header: "A" })
 * // => [{A: "name", B: "age"}, {A: "Alice", B: 30}]
 */
export function sheet_to_json<T = JSONRow>(worksheet: Worksheet, opts?: Sheet2JSONOpts): T[] {
  const o = opts || {};

  // Determine range
  let startRow = 1;
  let endRow = worksheet.rowCount;
  let startCol = 1;
  let endCol = worksheet.columnCount;

  if (o.range !== undefined) {
    if (typeof o.range === "number") {
      startRow = o.range + 1; // 0-indexed to 1-indexed
    } else if (typeof o.range === "string") {
      const r = decode_range(o.range);
      startRow = r.s.r + 1;
      endRow = r.e.r + 1;
      startCol = r.s.c + 1;
      endCol = r.e.c + 1;
    }
  }

  if (endRow < startRow || endCol < startCol) {
    return [];
  }

  // Handle header option
  const headerOpt = o.header;

  // header: 1 - return array of arrays
  if (headerOpt === 1) {
    const result: CellValue[][] = [];
    // Default for header:1 is to include blank rows
    const includeBlank = o.blankrows !== false;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: CellValue[] = [];
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        const val = cell.value;

        if (val != null) {
          rowData[col - startCol] = val;
          isEmpty = false;
        } else if (o.defval !== undefined) {
          rowData[col - startCol] = o.defval;
        } else {
          rowData[col - startCol] = null;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result as T[];
  }

  // header: "A" - use column letters as keys
  if (headerOpt === "A") {
    const result: Record<string, CellValue>[] = [];
    // Default for header:"A" is to skip blank rows
    const includeBlank = o.blankrows === true;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: Record<string, CellValue> = {};
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        const val = cell.value;
        const key = encode_col(col - 1); // 0-indexed for encode_col

        if (val != null) {
          rowData[key] = val;
          isEmpty = false;
        } else if (o.defval !== undefined) {
          rowData[key] = o.defval;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result as T[];
  }

  // header: string[] - use provided array as keys
  if (Array.isArray(headerOpt)) {
    const result: Record<string, CellValue>[] = [];
    const includeBlank = o.blankrows === true;

    for (let row = startRow; row <= endRow; row++) {
      const rowData: Record<string, CellValue> = {};
      let isEmpty = true;

      for (let col = startCol; col <= endCol; col++) {
        const colIdx = col - startCol;
        const key = headerOpt[colIdx] ?? `__EMPTY_${colIdx}`;
        const cell = worksheet.getCell(row, col);
        const val = cell.value;

        if (val != null) {
          rowData[key] = val;
          isEmpty = false;
        } else if (o.defval !== undefined) {
          rowData[key] = o.defval;
        }
      }

      if (!isEmpty || includeBlank) {
        result.push(rowData);
      }
    }

    return result as T[];
  }

  // Default: first row as header, disambiguate duplicates
  const headers: string[] = [];
  const headerCounts: Record<string, number> = {};

  for (let col = startCol; col <= endCol; col++) {
    const cell = worksheet.getCell(startRow, col);
    const val = cell.value;
    let header = val != null ? String(val) : `__EMPTY_${col - startCol}`;

    // Disambiguate duplicate headers
    if (headerCounts[header] !== undefined) {
      headerCounts[header]++;
      header = `${header}_${headerCounts[header]}`;
    } else {
      headerCounts[header] = 0;
    }

    headers.push(header);
  }

  // Read data rows (skip header row)
  const result: Record<string, CellValue>[] = [];
  const dataStartRow = startRow + 1;
  // Default for objects is to skip blank rows
  const includeBlank = o.blankrows === true;

  for (let row = dataStartRow; row <= endRow; row++) {
    const rowData: Record<string, CellValue> = {};
    let isEmpty = true;

    for (let col = startCol; col <= endCol; col++) {
      const cell = worksheet.getCell(row, col);
      const val = cell.value;
      const key = headers[col - startCol];

      if (val != null) {
        rowData[key] = val;
        isEmpty = false;
      } else if (o.defval !== undefined) {
        rowData[key] = o.defval;
      }
    }

    if (!isEmpty || includeBlank) {
      result.push(rowData);
    }
  }

  return result as T[];
}

// =============================================================================
// Sheet to CSV
// =============================================================================

export interface Sheet2CSVOpts {
  /** Field separator (default: ",") */
  FS?: string;
  /** Record separator (default: "\n") */
  RS?: string;
  /** Skip blank rows */
  blankrows?: boolean;
  /** Force quote all fields */
  forceQuotes?: boolean;
}

/**
 * Convert worksheet to CSV string
 */
export function sheet_to_csv(worksheet: Worksheet, opts?: Sheet2CSVOpts): string {
  const o = opts || {};
  const FS = o.FS ?? ",";
  const RS = o.RS ?? "\n";
  const rows: string[] = [];

  worksheet.eachRow({ includeEmpty: o.blankrows !== false }, (row, rowNumber) => {
    const cells: string[] = [];
    let isEmpty = true;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      let val = "";
      if (cell.value != null) {
        if (cell.value instanceof Date) {
          val = cell.value.toISOString();
        } else if (typeof cell.value === "object") {
          // Handle rich text, formula results, etc.
          if ("result" in cell.value) {
            val = String(cell.value.result ?? "");
          } else if ("text" in cell.value) {
            val = String(cell.value.text ?? "");
          } else if ("richText" in cell.value) {
            val = (cell.value.richText as Array<{ text: string }>).map(r => r.text).join("");
          } else {
            val = String(cell.value);
          }
        } else {
          val = String(cell.value);
        }
        isEmpty = false;
      }

      // Quote if necessary
      const needsQuote =
        o.forceQuotes ||
        val.includes(FS) ||
        val.includes('"') ||
        val.includes("\n") ||
        val.includes("\r");

      if (needsQuote) {
        val = `"${val.replace(/"/g, '""')}"`;
      }

      cells.push(val);
    });

    // Pad cells to match column count
    while (cells.length < worksheet.columnCount) {
      cells.push("");
    }

    if (!isEmpty || o.blankrows !== false) {
      rows.push(cells.join(FS));
    }
  });

  return rows.join(RS);
}

// =============================================================================
// Workbook Functions
// =============================================================================

/**
 * Create a new workbook
 */
export function book_new(): Workbook {
  return new Workbook();
}

/**
 * Append worksheet to workbook (xlsx compatible)
 * @example
 * const wb = book_new();
 * const ws = json_to_sheet([{a: 1, b: 2}]);
 * book_append_sheet(wb, ws, "Sheet1");
 */
export function book_append_sheet(workbook: Workbook, worksheet: Worksheet, name?: string): void {
  // Copy the worksheet data to a new sheet in the workbook
  const newWs = workbook.addWorksheet(name);

  // Copy all cells from source worksheet
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newWs.getCell(rowNumber, colNumber);
      newCell.value = cell.value;
      if (cell.style) {
        newCell.style = cell.style;
      }
    });
  });

  // Copy column properties
  worksheet.columns?.forEach((col, idx) => {
    if (col && newWs.columns[idx]) {
      if (col.width) {
        newWs.getColumn(idx + 1).width = col.width;
      }
    }
  });
}

// =============================================================================
// Array of Arrays Conversion
// =============================================================================

export interface AOA2SheetOpts {
  /** Use specified cell as starting point */
  origin?: Origin;
  /** Use specified date format in string output */
  dateNF?: string;
  /** Store dates as type d (default is n) */
  cellDates?: boolean;
}

/**
 * Create a worksheet from an array of arrays (xlsx compatible)
 * @example
 * const ws = aoa_to_sheet([["Name", "Age"], ["Alice", 30], ["Bob", 25]])
 */
export function aoa_to_sheet(data: CellValue[][], opts?: AOA2SheetOpts): Worksheet {
  const tempWb = new Workbook();
  const worksheet = tempWb.addWorksheet("Sheet1");

  if (data.length === 0) {
    return worksheet;
  }

  // Determine starting position
  let startRow = 1;
  let startCol = 1;

  if (opts?.origin !== undefined) {
    if (typeof opts.origin === "string") {
      const addr = decode_cell(opts.origin);
      startRow = addr.r + 1;
      startCol = addr.c + 1;
    } else if (typeof opts.origin === "number") {
      startRow = opts.origin + 1; // 0-indexed row
    } else {
      startRow = opts.origin.r + 1;
      startCol = opts.origin.c + 1;
    }
  }

  data.forEach((row, rowIdx) => {
    if (!row) {
      return;
    }
    row.forEach((val, colIdx) => {
      if (val !== undefined && val !== null) {
        worksheet.getCell(startRow + rowIdx, startCol + colIdx).value = val;
      }
    });
  });

  return worksheet;
}

/**
 * Add data from an array of arrays to an existing worksheet (xlsx compatible)
 */
export function sheet_add_aoa(
  worksheet: Worksheet,
  data: CellValue[][],
  opts?: AOA2SheetOpts
): Worksheet {
  if (data.length === 0) {
    return worksheet;
  }

  // Determine starting position
  let startRow = 1;
  let startCol = 1;

  if (opts?.origin !== undefined) {
    if (typeof opts.origin === "string") {
      const addr = decode_cell(opts.origin);
      startRow = addr.r + 1;
      startCol = addr.c + 1;
    } else if (typeof opts.origin === "number") {
      if (opts.origin === -1) {
        // Append to bottom
        startRow = worksheet.rowCount + 1;
      } else {
        startRow = opts.origin + 1; // 0-indexed row
      }
    } else {
      startRow = opts.origin.r + 1;
      startCol = opts.origin.c + 1;
    }
  }

  data.forEach((row, rowIdx) => {
    if (!row) {
      return;
    }
    row.forEach((val, colIdx) => {
      if (val !== undefined && val !== null) {
        worksheet.getCell(startRow + rowIdx, startCol + colIdx).value = val;
      }
    });
  });

  return worksheet;
}

/**
 * Convert worksheet to array of arrays
 */
export function sheet_to_aoa(worksheet: Worksheet): CellValue[][] {
  const result: CellValue[][] = [];

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowData: CellValue[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      rowData[colNumber - 1] = cell.value;
    });
    result[rowNumber - 1] = rowData;
  });

  return result;
}

// =============================================================================
// Export utils object
// =============================================================================

export const utils = {
  // Cell encoding/decoding
  decode_col,
  encode_col,
  decode_row,
  encode_row,
  decode_cell,
  encode_cell,
  decode_range,
  encode_range,

  // Sheet/JSON conversion
  json_to_sheet,
  sheet_add_json,
  sheet_to_json,
  sheet_to_csv,
  aoa_to_sheet,
  sheet_add_aoa,
  sheet_to_aoa,

  // Workbook functions
  book_new,
  book_append_sheet
};
