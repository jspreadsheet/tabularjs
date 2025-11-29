/**
 * TabularJS - Universal spreadsheet parser for JavaScript
 * @module tabularjs
 */

import { parseXLS } from './parser/xls.js';
import { parseXLSX } from './parser/xlsx.js';
import { parseODS } from './parser/ods.js';
import { parseLotus } from './parser/lotus.js';
import { parseCSVFile, parseCSVString } from './parser/csv.js';
import { parseXMLSpreadsheet } from './parser/xml.js';
import { parseDIF } from './parser/dif.js';
import { parseSLK } from './parser/slk.js';
import { parseHTMLTable} from './parser/table.js';
import { parseDBF } from './parser/dbf.js';

/**
 * Parse spreadsheet files with automatic format detection
 *
 * Supports 16+ file formats including Excel (.xls, .xlsx), OpenDocument (.ods),
 * CSV, TSV, HTML tables, XML Spreadsheet, Lotus 1-2-3, SYLK, DIF, and dBase.
 *
 * @async
 * @param {string|File|Uint8Array} file - File path (Node.js), File object (Browser), or buffer
 * @param {Object} [options={}] - Parser options
 * @param {string} [options.delimiter=','] - CSV delimiter (for CSV/TSV files)
 * @param {number} [options.tableIndex=0] - Table index to parse (for HTML with multiple tables)
 * @param {boolean} [options.firstRowAsHeader=true] - Treat first row as header (for HTML/CSV)
 * @param {number} [options.worksheetIndex] - Specific worksheet to parse (0-based)
 * @param {Function} [options.onload] - Callback when parsing completes
 * @param {Function} [options.onerror] - Callback when parsing fails
 *
 * @returns {Promise<Object>} Parsed spreadsheet data in Jspreadsheet format
 * @returns {Array<Object>} return.worksheets - Array of worksheet objects
 * @returns {Array<Array>} return.worksheets[].data - 2D array of cell values
 * @returns {Array<Object>} return.worksheets[].columns - Column definitions with titles and widths
 * @returns {string} [return.worksheets[].worksheetName] - Name of the worksheet
 * @returns {Object} [return.worksheets[].mergeCells] - Merged cell definitions (e.g., {A1: [2, 1]})
 * @returns {Object} [return.worksheets[].style] - Cell styles by address (e.g., {A1: "color: red"})
 * @returns {Object} [return.worksheets[].comments] - Cell comments by address
 * @returns {Object} [return.worksheets[].cells] - Cell metadata (formulas, masks, etc.)
 * @returns {Array<Object>} [return.worksheets[].rows] - Row properties (height, visibility)
 * @returns {Object} [return.definedNames] - Named ranges (e.g., {MyRange: "Sheet1!A1:B10"})
 *
 * @throws {Error} If file is null/undefined or unsupported format
 *
 * @example
 * // Node.js
 * import tabularjs from 'tabularjs';
 * const result = await tabularjs('path/to/file.xlsx');
 * console.log(result.worksheets[0].data);
 *
 * @example
 * // Browser with File object
 * import tabularjs from 'tabularjs';
 * const file = document.getElementById('fileInput').files[0];
 * const result = await tabularjs(file);
 *
 * @example
 * // CSV with custom delimiter
 * const result = await tabularjs('data.csv', { delimiter: ';' });
 *
 * @example
 * // HTML table (second table)
 * const result = await tabularjs('page.html', { tableIndex: 1 });
 *
 * @example
 * // Direct integration with Jspreadsheet
 * import jspreadsheet from 'jspreadsheet-ce';
 * const result = await tabularjs(file);
 * jspreadsheet(document.getElementById('spreadsheet'), result);
 */
export default async function parser(file, options = {}) {
    if (!file) {
        throw new Error('Invalid file');
    }

    // Handle both file paths (string) and file objects (with .name property)
    let fileName;
    if (typeof file === 'string') {
        // File path
        fileName = file;
    } else if (file.name) {
        // File object (browser File API)
        fileName = file.name;
    } else {
        throw new Error('Invalid file: must be a file path or file object with name property');
    }

    const ext = fileName.toLowerCase().split('.').pop();

    switch (ext) {
        case 'xls':
            return await parseXLS(file, options);
        case 'xlsx':
            return await parseXLSX(file, options);
        case 'ods':
            return await parseODS(file, options);
        case 'wks':
        case 'wk1':
        case 'wk3':
        case 'wk4':
        case '123':
            return await parseLotus(file, options);
        case 'csv':
            return await parseCSVFile(file, options);
        case 'tsv':
        case 'tab':
            return await parseCSVFile(file, { ...options, delimiter: '\t' });
        case 'txt':
            return await parseCSVFile(file, { ...options, delimiter: '\t' });
        case 'xml':
            return await parseXMLSpreadsheet(file, options);
        case 'dif':
            return await parseDIF(file, options);
        case 'slk':
        case 'sylk':
            return await parseSLK(file, options);
        case 'html':
        case 'htm':
            return await parseHTMLTable(file, options);
        case 'dbf':
            return await parseDBF(file, options);
        default:
            throw new Error(`Unsupported file type: ${ext}`);
    }
}