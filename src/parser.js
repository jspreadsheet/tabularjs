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
import { parseNumbers } from './parser/numbers.js';
import { loadAsBuffer } from './utils/loader.js';
import { readZipEntry } from './utils/zip-utils.js';
import JSZip from 'jszip';

/**
 * Parse spreadsheet files with automatic format detection
 *
 * Supports 16+ file formats including Excel (.xls, .xlsx), OpenDocument (.ods),
 * CSV, TSV, HTML tables, XML Spreadsheet, Lotus 1-2-3, SYLK, DIF, and dBase.
 *
 * @async
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file - File path (Node.js), File/Blob (Browser), or raw bytes
 * @param {Object} [options={}] - Parser options
 * @param {string} [options.extension] - Force a format (e.g. 'xlsx') instead of detecting it
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

    // Normalize raw byte inputs so the loaders can consume them
    if (typeof ArrayBuffer !== 'undefined' && file instanceof ArrayBuffer) {
        file = new Uint8Array(file);
    }

    const isBinaryInput = file instanceof Uint8Array
        || (typeof Buffer !== 'undefined' && Buffer.isBuffer(file))
        || (typeof Blob !== 'undefined' && file instanceof Blob);

    // Handle file paths (string) and file objects (with .name property)
    let fileName = '';
    if (typeof file === 'string') {
        fileName = file;
    } else if (file.name && typeof file.name === 'string') {
        fileName = file.name;
    } else if (!isBinaryInput) {
        throw new Error('Invalid file: must be a file path, File, Blob, Uint8Array, ArrayBuffer, or Buffer');
    }

    // Format resolution: explicit option first, then the file extension
    let ext = null;
    if (options.extension) {
        ext = String(options.extension).toLowerCase().replace(/^\./, '');
    } else {
        const dot = fileName.lastIndexOf('.');
        if (dot > -1) {
            ext = fileName.toLowerCase().slice(dot + 1);
        }
    }

    if (ext) {
        const result = dispatch(ext, file, options);
        if (result) {
            return await result;
        }
    }

    // Unknown or missing extension: detect the format from the content
    let buffer = null;
    let detected = null;
    try {
        buffer = await loadAsBuffer(file);
        detected = await detectFormat(buffer);
    } catch (e) {
        // Input could not be read or sniffed; report as unsupported below
    }

    if (detected) {
        const result = dispatch(detected, buffer, options);
        if (result) {
            return await result;
        }
    }

    throw new Error(`Unsupported file type: ${ext || fileName || '(buffer)'}`);
}

/**
 * Route a resolved format to its parser
 * @returns {Promise<Object>|null} Parser promise, or null when the format is unknown
 */
function dispatch(ext, file, options) {
    switch (ext) {
        case 'xls':
            return parseXLS(file, options);
        case 'xlsx':
            return parseXLSX(file, options);
        case 'ods':
            return parseODS(file, options);
        case 'numbers':
            return parseNumbers(file, options);
        case 'wks':
        case 'wk1':
        case 'wk3':
        case 'wk4':
        case '123':
            return parseLotus(file, options);
        case 'csv':
            return parseCSVFile(file, options);
        case 'tsv':
        case 'tab':
            return parseCSVFile(file, { ...options, delimiter: '\t' });
        case 'txt':
            return parseCSVFile(file, { ...options, delimiter: '\t' });
        case 'xml':
            return parseXMLSpreadsheet(file, options);
        case 'dif':
            return parseDIF(file, options);
        case 'slk':
        case 'sylk':
            return parseSLK(file, options);
        case 'html':
        case 'htm':
            return parseHTMLTable(file, options);
        case 'dbf':
            return parseDBF(file, options);
        default:
            return null;
    }
}

/**
 * Detect a spreadsheet format from file content (magic bytes)
 *
 * @param {Uint8Array} buffer - File content
 * @returns {Promise<string|null>} Detected extension (e.g. 'xlsx') or null
 */
export async function detectFormat(buffer) {
    if (!buffer || buffer.length < 4) {
        return null;
    }

    // ZIP container: xlsx, ods or Apple Numbers
    if (buffer[0] === 0x50 && buffer[1] === 0x4B && (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07)) {
        try {
            const zip = await JSZip.loadAsync(buffer);
            const names = Object.keys(zip.files);
            if (names.some((n) => n === '[Content_Types].xml' || n.startsWith('xl/'))) {
                return 'xlsx';
            }
            if (names.some((n) => n === 'Index.zip' || n.startsWith('Index/') || n.endsWith('.iwa'))) {
                return 'numbers';
            }
            if (names.includes('mimetype')) {
                const mimetype = await readZipEntry(zip.file('mimetype'), 'string', 4096);
                if (mimetype.includes('opendocument.spreadsheet')) {
                    return 'ods';
                }
            }
            if (names.includes('content.xml')) {
                return 'ods';
            }
        } catch (e) {
            // Corrupt archive; fall through to the most common zip format
        }
        return 'xlsx';
    }

    // OLE2 compound file (legacy Excel)
    if (buffer[0] === 0xD0 && buffer[1] === 0xCF && buffer[2] === 0x11 && buffer[3] === 0xE0) {
        return 'xls';
    }

    // Lotus 1-2-3 BOF record (0x0002/0x001A with 16-bit length)
    if (buffer[0] === 0x00 && buffer[1] === 0x00 && (buffer[2] === 0x02 || buffer[2] === 0x1A) && buffer[3] === 0x00) {
        return 'wk1';
    }

    // dBase: known version byte + plausible header
    if ([0x02, 0x03, 0x30, 0x31, 0x83, 0x8B, 0xF5].includes(buffer[0]) && buffer.length > 32) {
        const month = buffer[2];
        const day = buffer[3];
        const headerLength = buffer[8] | (buffer[9] << 8);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && headerLength > 32) {
            return 'dbf';
        }
    }

    // Text-based formats
    const sample = new TextDecoder('utf-8').decode(buffer.slice(0, 4096));
    if (sample.includes('\u0000')) {
        return null; // Binary content we do not recognize
    }
    const text = sample.replace(/^\uFEFF/, '').trimStart();
    const lower = text.toLowerCase();

    if (text.startsWith('ID;')) {
        return 'slk';
    }
    if (text.startsWith('TABLE') && text.includes('VECTORS')) {
        return 'dif';
    }
    if (lower.startsWith('<?xml')) {
        if (lower.includes('urn:schemas-microsoft-com:office:spreadsheet') || lower.includes('<workbook')) {
            return 'xml';
        }
        if (lower.includes('<html') || lower.includes('<table')) {
            return 'html';
        }
        return 'xml';
    }
    if (lower.startsWith('<!doctype html') || lower.includes('<html') || lower.includes('<table')) {
        return 'html';
    }
    if (text.length) {
        const firstLine = text.split(/\r?\n/, 1)[0];
        if (firstLine.includes('\t') && !firstLine.includes(',')) {
            return 'tsv';
        }
        return 'csv';
    }

    return null;
}