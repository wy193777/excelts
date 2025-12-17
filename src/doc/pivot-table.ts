import { objectFromProps, range, toSortedArray } from "../utils/utils.js";
import { colCache } from "../utils/col-cache.js";
import type { Table } from "./table.js";

/**
 * Interface representing the source data abstraction for pivot tables.
 * This allows both Worksheet and Table to be used as pivot table data sources.
 */
interface PivotTableSource {
  name: string;
  getRow(rowNumber: number): { values: any[] };
  getColumn(columnNumber: number): { values: any[] };
  getSheetValues(): any[][];
  dimensions: { shortRange: string };
}

interface PivotTableModel {
  /**
   * Source worksheet for the pivot table data.
   * Either sourceSheet or sourceTable must be provided (mutually exclusive).
   */
  sourceSheet?: any;
  /**
   * Source table for the pivot table data.
   * Either sourceSheet or sourceTable must be provided (mutually exclusive).
   * The table must have headerRow=true and contain at least one data row.
   */
  sourceTable?: Table;
  /** Column names to use as row fields in the pivot table */
  rows: string[];
  /**
   * Column names to use as column fields in the pivot table.
   * If omitted or empty, Excel will use "Values" as the column field.
   * @default []
   */
  columns?: string[];
  /** Column names to aggregate as values in the pivot table */
  values: string[];
  /**
   * Aggregation metric for the pivot table values.
   * - 'sum': Sum of values (default)
   * - 'count': Count of values
   * @default 'sum'
   */
  metric?: "sum" | "count";
  /**
   * Controls whether pivot table style overrides worksheet column widths.
   * - '0': Preserve worksheet column widths (useful for custom sizing)
   * - '1': Apply pivot table style width/height (default Excel behavior)
   * @default '1'
   */
  applyWidthHeightFormats?: "0" | "1";
}

interface CacheField {
  name: string;
  sharedItems: any[] | null;
}

/**
 * Data field configuration for pivot table aggregation
 */
interface DataField {
  name: string;
  fld: number;
  baseField?: number;
  baseItem?: number;
  subtotal?:
    | "sum"
    | "count"
    | "average"
    | "max"
    | "min"
    | "product"
    | "countNums"
    | "stdDev"
    | "stdDevP"
    | "var"
    | "varP";
}

interface PivotTable {
  source?: PivotTableSource;
  rows: number[];
  columns: number[];
  values: number[];
  metric: "sum" | "count";
  cacheFields: CacheField[];
  cacheId: string;
  applyWidthHeightFormats: "0" | "1";
  /** 1-indexed table number for file naming (pivotTable1.xml, pivotTable2.xml, etc.) */
  tableNumber: number;
  /** Flag indicating this pivot table was loaded from file (not newly created) */
  isLoaded?: boolean;
  /** Data fields for loaded pivot tables */
  dataFields?: DataField[];
  /** Cache definition for loaded pivot tables */
  cacheDefinition?: any;
  /** Cache records for loaded pivot tables */
  cacheRecords?: any;
}

// TK(2023-10-10): turn this into a class constructor.

/**
 * Creates a PivotTableSource adapter from a Table object.
 * This allows Tables to be used as pivot table data sources with the same interface as Worksheets.
 */
function createTableSourceAdapter(table: Table): PivotTableSource {
  const tableModel = table.model;

  // Validate that table has headerRow enabled (required for pivot table column names)
  if (tableModel.headerRow === false) {
    throw new Error(
      "Cannot create pivot table from a table without headers. Set headerRow: true on the table."
    );
  }

  // Validate table has data rows
  if (!tableModel.rows || tableModel.rows.length === 0) {
    throw new Error("Cannot create pivot table from an empty table. Add data rows to the table.");
  }

  const columnNames = tableModel.columns.map(col => col.name);

  // Check for duplicate column names
  const nameSet = new Set<string>();
  for (const name of columnNames) {
    if (nameSet.has(name)) {
      throw new Error(
        `Duplicate column name "${name}" found in table. Pivot tables require unique column names.`
      );
    }
    nameSet.add(name);
  }

  // Build the full data array: headers + rows
  const headerRow = [undefined, ...columnNames]; // sparse array starting at index 1
  const dataRows = tableModel.rows.map(row => [undefined, ...row]); // sparse array starting at index 1

  // Calculate the range reference for the table
  const tl = tableModel.tl!;
  const startRow = tl.row;
  const startCol = tl.col;
  const endRow = startRow + tableModel.rows.length; // header row + data rows
  const endCol = startCol + columnNames.length - 1;

  const shortRange = colCache.encode(startRow, startCol, endRow, endCol);

  return {
    name: tableModel.name,
    getRow(rowNumber: number): { values: any[] } {
      if (rowNumber === 1) {
        return { values: headerRow };
      }
      const dataIndex = rowNumber - 2; // rowNumber 2 maps to index 0
      if (dataIndex >= 0 && dataIndex < dataRows.length) {
        return { values: dataRows[dataIndex] };
      }
      return { values: [] };
    },
    getColumn(columnNumber: number): { values: any[] } {
      // Validate column number is within bounds
      if (columnNumber < 1 || columnNumber > columnNames.length) {
        return { values: [] };
      }
      // Values should be sparse array with header at index 1, data starting at index 2
      const values: any[] = [];
      values[1] = columnNames[columnNumber - 1];
      for (let i = 0; i < tableModel.rows.length; i++) {
        values[i + 2] = tableModel.rows[i][columnNumber - 1];
      }
      return { values };
    },
    getSheetValues(): any[][] {
      // Return sparse array where index 1 is header row, and subsequent indices are data rows
      const result: any[][] = [];
      result[1] = headerRow;
      for (let i = 0; i < dataRows.length; i++) {
        result[i + 2] = dataRows[i];
      }
      return result;
    },
    dimensions: { shortRange }
  };
}

/**
 * Resolves the data source from the model, supporting both sourceSheet and sourceTable.
 */
function resolveSource(model: PivotTableModel): PivotTableSource {
  if (model.sourceTable) {
    return createTableSourceAdapter(model.sourceTable);
  }
  // For sourceSheet, it already implements the required interface
  return model.sourceSheet as PivotTableSource;
}

function makePivotTable(worksheet: any, model: PivotTableModel): PivotTable {
  // Example `model`:
  // {
  //   // Source of data (either sourceSheet OR sourceTable):
  //   sourceSheet: worksheet1,  // Use entire sheet range
  //   // OR
  //   sourceTable: table,       // Use table data
  //
  //   // Pivot table fields: values indicate field names;
  //   // they come from the first row in `worksheet1` or table column names.
  //   rows: ['A', 'B'],
  //   columns: ['C'],
  //   values: ['E'], // only 1 item possible for now
  //   metric: 'sum', // only 'sum' possible for now
  // }

  // Validate source exists before trying to resolve it
  if (!model.sourceSheet && !model.sourceTable) {
    throw new Error("Either sourceSheet or sourceTable must be provided.");
  }
  if (model.sourceSheet && model.sourceTable) {
    throw new Error("Cannot specify both sourceSheet and sourceTable. Choose one.");
  }

  // Resolve source first to avoid creating adapter multiple times
  const source = resolveSource(model);

  validate(worksheet, model, source);

  const { rows, values } = model;
  const columns = model.columns ?? [];

  const cacheFields = makeCacheFields(source, [...rows, ...columns]);

  const nameToIndex = cacheFields.reduce(
    (result: Record<string, number>, cacheField: CacheField, index: number) => {
      result[cacheField.name] = index;
      return result;
    },
    {} as Record<string, number>
  );
  const rowIndices = rows.map(row => nameToIndex[row]);
  const columnIndices = columns.map(column => nameToIndex[column]);
  const valueIndices = values.map(value => nameToIndex[value]);

  // Calculate tableNumber based on existing pivot tables (1-indexed)
  const tableNumber = worksheet.workbook.pivotTables.length + 1;

  // Base cache ID starts at 10 (Excel convention), each subsequent table increments
  const BASE_CACHE_ID = 10;

  // form pivot table object
  return {
    source,
    rows: rowIndices,
    columns: columnIndices,
    values: valueIndices,
    metric: model.metric ?? "sum",
    cacheFields,
    // Dynamic cacheId: 10 for first table, 11 for second, etc.
    // Used in <pivotTableDefinition> and xl/workbook.xml
    cacheId: String(BASE_CACHE_ID + tableNumber - 1),
    // Control whether pivot table style overrides worksheet column widths
    // '0' = preserve worksheet column widths (useful for custom sizing)
    // '1' = apply pivot table style width/height (default Excel behavior)
    applyWidthHeightFormats: model.applyWidthHeightFormats ?? "1",
    // Table number for file naming (pivotTable1.xml, pivotTable2.xml, etc.)
    tableNumber
  };
}

function validate(worksheet: any, model: PivotTableModel, source: PivotTableSource): void {
  if (model.metric && model.metric !== "sum" && model.metric !== "count") {
    throw new Error('Only the "sum" and "count" metrics are supported at this time.');
  }

  const columns = model.columns ?? [];

  // Get header names from source (already resolved)
  const headerNames = source.getRow(1).values.slice(1);
  const isInHeaderNames = objectFromProps(headerNames, true);
  for (const name of [...model.rows, ...columns, ...model.values]) {
    if (!isInHeaderNames[name]) {
      throw new Error(`The header name "${name}" was not found in ${source.name}.`);
    }
  }

  if (!model.rows.length) {
    throw new Error("No pivot table rows specified.");
  }

  // Allow empty columns - Excel will use "Values" as column field
  // But can't have multiple values with columns specified
  if (model.values.length < 1) {
    throw new Error("Must have at least one value.");
  }

  if (model.values.length > 1 && columns.length > 0) {
    throw new Error(
      "It is currently not possible to have multiple values when columns are specified. Please either supply an empty array for columns or a single value."
    );
  }
}

function makeCacheFields(
  source: PivotTableSource,
  fieldNamesWithSharedItems: string[]
): CacheField[] {
  // Cache fields are used in pivot tables to reference source data.
  //
  // Example
  // -------
  // Turn
  //
  //  `source` sheet values [
  //    ['A', 'B', 'C', 'D', 'E'],
  //    ['a1', 'b1', 'c1', 4, 5],
  //    ['a1', 'b2', 'c1', 4, 5],
  //    ['a2', 'b1', 'c2', 14, 24],
  //    ['a2', 'b2', 'c2', 24, 35],
  //    ['a3', 'b1', 'c3', 34, 45],
  //    ['a3', 'b2', 'c3', 44, 45]
  //  ];
  //  fieldNamesWithSharedItems = ['A', 'B', 'C'];
  //
  // into
  //
  //  [
  //    { name: 'A', sharedItems: ['a1', 'a2', 'a3'] },
  //    { name: 'B', sharedItems: ['b1', 'b2'] },
  //    { name: 'C', sharedItems: ['c1', 'c2', 'c3'] },
  //    { name: 'D', sharedItems: null },
  //    { name: 'E', sharedItems: null }
  //  ]

  const names = source.getRow(1).values;
  const nameToHasSharedItems = objectFromProps(fieldNamesWithSharedItems, true);

  const aggregate = (columnIndex: number): any[] => {
    // Use slice (not splice) to avoid mutating the original array
    const columnValues = source.getColumn(columnIndex).values.slice(2);
    // Filter out null/undefined values - they are represented as <m /> in XML,
    // not as sharedItems entries
    const validValues = columnValues.filter((v: any) => v !== null && v !== undefined);
    const columnValuesAsSet = new Set(validValues);
    return toSortedArray(columnValuesAsSet);
  };

  // make result
  const result: CacheField[] = [];
  for (const columnIndex of range(1, names.length)) {
    const name = names[columnIndex];
    const sharedItems = nameToHasSharedItems[name] ? aggregate(columnIndex) : null;
    result.push({ name, sharedItems });
  }
  return result;
}

export { makePivotTable, type PivotTable, type PivotTableModel, type PivotTableSource };
