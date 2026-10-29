/**
 * TabularJS - Spreadsheet file parser supporting 16+ formats
 * @module tabularjs
 */

/**
 * Column definition for a worksheet
 */
export interface Column {
    /** Column title (e.g., "A", "B", "C") */
    title: string;
    /** Column name (optional) */
    name?: string;
    /** Column width in pixels */
    width?: number;
    /** Column type (e.g., "text", "numeric", "calendar", "checkbox") */
    type?: string;
    /** Number of decimal places for numeric columns */
    decimal?: string;
    /** Format mask (e.g., "#,##0.00") */
    mask?: string;
    /** Options for specific column types */
    options?: Record<string, any>;
}

/**
 * Row definition for a worksheet
 */
export interface Row {
    /** Row height in pixels */
    height?: number;
}

/**
 * Cell definition with specific properties
 */
export interface Cell {
    /** Cell format (e.g., "MM/DD/YY") */
    format?: string;
    /** Horizontal alignment */
    align?: 'left' | 'center' | 'right' | 'fill' | 'justify';
    /** Vertical alignment */
    valign?: 'top' | 'middle' | 'bottom';
    /** Text wrapping */
    wrap?: boolean;
    /** Formula (if cell contains a formula) */
    formula?: string;
    /** Any other cell properties */
    [key: string]: any;
}

/**
 * Worksheet data structure
 */
export interface Worksheet {
    /** Name of the worksheet */
    worksheetName?: string;
    /** 2D array of cell values */
    data: any[][];
    /** Column definitions */
    columns: Column[];
    /** Row definitions (keyed by row index) */
    rows?: Record<number, Row>;
    /** Cell-specific properties (keyed by cell address like "A1") */
    cells?: Record<string, Cell>;
    /** Style of each cell (keyed by cell address): an index into ParseResult.style */
    style?: Record<string, string | number>;
    /** Merged cells (keyed by cell address, value is [colspan, rowspan]) */
    mergeCells?: Record<string, [number, number]>;
    /** Cell comments (keyed by cell address) */
    comments?: Record<string, string>;
    /** Minimum dimensions [columns, rows] */
    minDimensions?: [number, number];
    /** Sheet visibility (omitted when visible) */
    worksheetState?: 'hidden' | 'veryHidden';
    /** false when the sheet hides its gridlines */
    gridline?: boolean;
    /** Images, charts and shapes (XLSX, ODS) */
    media?: any[];
    /** Format-specific metadata */
    meta?: Record<string, any>;
}

/**
 * Result object returned by the parser
 */
export interface ParseResult {
    /** Array of worksheets */
    worksheets: Worksheet[];
    /** Global styles (array of CSS strings) */
    style?: string[];
    /** Named ranges (name -> reference, e.g. "Sheet1!$A$1:$B$10") */
    definedNames?: Record<string, string>;
}

/**
 * Options for parsing files
 */
export interface ParseOptions {
    /** Force a format (e.g. "xlsx" or ".xlsx") instead of detecting it from the file name or content */
    extension?: string;
    /** Delimiter for CSV and TXT files (default: "," for CSV, "\t" for TXT; TSV always uses tabs) */
    delimiter?: string;
    /** Text encoding for CSV, TSV, TXT, SYLK, DIF and DBF (default: "utf-8" for CSV/TSV/TXT/SYLK, auto-detected for DIF, the file's codepage for DBF) */
    encoding?: string;
    /** HTML: return only this table (0-based); by default every <table> becomes a worksheet */
    tableIndex?: number;
    /** Use first row as header for HTML tables (default: true) */
    firstRowAsHeader?: boolean;
    /** Called with the result when parsing succeeds (the promise also resolves with it) */
    onload?: (result: ParseResult) => void;
    /** Called with the error when parsing fails (the promise also rejects with it) */
    onerror?: (error: Error) => void;
    /** Any other format-specific options */
    [key: string]: any;
}

/**
 * Input file type - can be a file path (Node.js), File object (Browser), raw bytes (Uint8Array, ArrayBuffer or
 * Node Buffer), HTML string, or DOM element
 */
export type FileInput = string | File | Blob | Uint8Array | ArrayBuffer | HTMLElement;

/**
 * Parse a spreadsheet file and convert it to JSON format
 *
 * @param file - File path (Node.js), File object (Browser), Blob, raw bytes, HTML string, or DOM element
 * @param options - Parsing options
 * @returns Promise resolving to parsed data in Jspreadsheet Pro format
 *
 * @example
 * ```javascript
 * // Node.js - with file path
 * import tabularjs from 'tabularjs';
 * const result = await tabularjs('path/to/file.xlsx');
 * console.log(result.worksheets[0].data);
 * ```
 *
 * @example
 * ```javascript
 * // Browser - with File object
 * import tabularjs from 'tabularjs';
 *
 * const fileInput = document.querySelector('input[type="file"]');
 * fileInput.addEventListener('change', async (e) => {
 *   const file = e.target.files[0];
 *   const result = await tabularjs(file);
 *   console.log(result.worksheets[0].data);
 * });
 * ```
 *
 * @example
 * ```javascript
 * // Browser - with DOM element
 * const table = document.querySelector('table');
 * const result = await tabularjs(table);
 * console.log(result.worksheets[0].data);
 * ```
 *
 * @example
 * ```javascript
 * // Node.js - with HTML string
 * const htmlString = '<table><tr><td>A1</td><td>B1</td></tr></table>';
 * const result = await tabularjs(htmlString);
 * console.log(result.worksheets[0].data);
 * ```
 *
 * @example
 * ```javascript
 * // With try/catch
 * try {
 *   const result = await tabularjs(file, { encoding: 'utf-8' });
 *   console.log(result);
 * } catch (error) {
 *   console.error('Parse error:', error);
 * }
 * ```
 */
declare function tabularjs(file: FileInput, options?: ParseOptions): Promise<ParseResult>;

export default tabularjs;
