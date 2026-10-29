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
import { parseHTMLTable, parseAllHTMLTables, isHTMLString } from './parser/table.js';
import { parseDBF } from './parser/dbf.js';
import { parseNumbers } from './parser/numbers.js';
import { loadAsBuffer } from './utils/loader.js';
import { readZipEntry } from './utils/zip-utils.js';
import CFB from './utils/cfb.js';
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
 * @param {number} [options.tableIndex] - HTML: return only this table (0-based); by default one worksheet per table
 * @param {boolean} [options.firstRowAsHeader=true] - Treat first row as header (for HTML/CSV)
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

    // HTML markup or a DOM element: not a file path
    if (isHTMLString(file) || (typeof HTMLElement !== 'undefined' && file instanceof HTMLElement)) {
        return dispatch('html', file, options);
    }

    // Normalize raw byte inputs so the loaders can consume them
    if (typeof ArrayBuffer !== 'undefined' && file instanceof ArrayBuffer) {
        file = new Uint8Array(file);
    }

    const isBlob = typeof Blob !== 'undefined' && file instanceof Blob;
    const isBinaryInput = file instanceof Uint8Array
        || (typeof Buffer !== 'undefined' && Buffer.isBuffer(file))
        || isBlob;

    // Handle file paths (string) and file objects (with .name property)
    let fileName = '';
    if (typeof file === 'string') {
        fileName = file;
    } else if (file.name && typeof file.name === 'string') {
        fileName = file.name;
    } else if (!isBinaryInput) {
        throw new Error('Invalid file: must be a file path, File, Blob, Uint8Array, ArrayBuffer, or Buffer');
    }

    // The callbacks fire once for the whole call, not once per parser attempt
    const { onload, onerror, ...opts } = options;
    try {
        const result = await resolve(file, fileName, isBlob, opts);
        if (onload) onload(result);
        return result;
    } catch (error) {
        if (onerror) onerror(error);
        throw error;
    }
}

/**
 * Extensions dispatch() knows, mapped to a canonical format (aliases share a parser)
 */
const FORMATS = {
    xls: 'xls', xlsx: 'xlsx', ods: 'ods', numbers: 'numbers',
    wks: 'wk1', wk1: 'wk1', wk3: 'wk1', wk4: 'wk1', '123': 'wk1',
    csv: 'csv', tsv: 'tsv', tab: 'tsv', txt: 'txt', xml: 'xml', dif: 'dif',
    slk: 'slk', sylk: 'slk', html: 'html', htm: 'html', dbf: 'dbf'
};

/**
 * Pick the parser and run it
 *
 * - options.extension is authoritative and never second-guessed.
 * - Otherwise the content is sniffed once. A container signature (OLE2, or a
 *   zip whose entries identify xlsx/ods/numbers) is unambiguous and beats the
 *   file extension, so a renamed .xlsx/.xls/.ods goes straight to its parser.
 * - Without such a signature the extension's parser runs; if it throws and
 *   the content looks like another format (CSV, HTML or XML Spreadsheet saved
 *   as .xls, ...), that parser is tried. If it fails too, the error of the
 *   extension's parser is rethrown.
 */
async function resolve(file, fileName, isBlob, options) {
    let ext = null;
    if (options.extension) {
        ext = String(options.extension).toLowerCase().replace(/^\./, '');
        // JSZip cannot read a Blob/File in Node.js (no FileReader): hand over bytes
        const result = dispatch(ext, isBlob ? await loadAsBuffer(file) : file, options);
        if (result) {
            return await result;
        }
        // Unknown forced format: fall through to content detection
    } else {
        const dot = fileName.lastIndexOf('.');
        if (dot > -1) {
            ext = fileName.toLowerCase().slice(dot + 1);
        }
    }

    // Known extension: read once, sniff, and let a container signature win
    if (Object.hasOwn(FORMATS, ext)) {
        const buffer = await loadAsBuffer(file);
        const { format, certain } = await sniffFormat(buffer);

        if (certain && format !== FORMATS[ext]) {
            return await run(format, buffer, options);
        }

        try {
            return await run(ext, buffer, options);
        } catch (error) {
            if (format && format !== FORMATS[ext]) {
                try {
                    return await dispatch(format, buffer, options);
                } catch (e) {
                    // Not the detected format either: report the original problem
                }
            }
            throw error;
        }
    }

    // Unknown or missing extension: detect the format from the content
    let buffer = null;
    let detected = null;
    try {
        buffer = await loadAsBuffer(file);
        detected = (await sniffFormat(buffer)).format;
    } catch (e) {
        // Input could not be read or sniffed; report as unsupported below
    }

    if (detected) {
        return await run(detected, buffer, options);
    }

    throw new Error(`Unsupported file type: ${ext || fileName || '(buffer)'}`);
}

/**
 * dispatch(), except that an OLE2 file holding an encrypted OOXML package
 * (a password-protected .xlsx) reports the password, not a missing stream
 */
async function run(ext, buffer, options) {
    try {
        return await dispatch(ext, buffer, options);
    } catch (error) {
        if (FORMATS[ext] === 'xls' && isEncryptedPackage(buffer)) {
            throw new Error('Workbook is encrypted with a password');
        }
        throw error;
    }
}

function isEncryptedPackage(buffer) {
    try {
        return !!CFB.find(CFB.read(buffer), 'EncryptedPackage');
    } catch (e) {
        return false;
    }
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
            return parseCSVFile(file, { delimiter: '\t', ...options }); // tab unless the caller picks one
        case 'xml':
            return parseXMLSpreadsheet(file, options);
        case 'dif':
            return parseDIF(file, options);
        case 'slk':
        case 'sylk':
            return parseSLK(file, options);
        case 'html':
        case 'htm':
            // One worksheet per <table>, or only the one tableIndex picks
            return options.tableIndex === undefined
                ? parseAllHTMLTables(file, options)
                : parseHTMLTable(file, options);
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
    return (await sniffFormat(buffer)).format;
}

/**
 * Detect a format and tell whether the evidence is a container signature
 *
 * @param {Uint8Array} buffer - File content
 * @returns {Promise<{format: string|null, certain: boolean}>} certain is true for
 * OLE2 and for zip archives whose entries identify xlsx, ods or numbers
 */
async function sniffFormat(buffer) {
    if (!buffer || buffer.length < 4) {
        return { format: null, certain: false };
    }

    // ZIP container: xlsx, ods or Apple Numbers
    if (buffer[0] === 0x50 && buffer[1] === 0x4B && (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07)) {
        let format = null;
        try {
            format = await detectZipFormat(await JSZip.loadAsync(buffer));
        } catch (e) {
            // Corrupt archive; fall through to the most common zip format
        }
        return { format: format || 'xlsx', certain: !!format };
    }

    // OLE2 compound file (legacy Excel)
    if (buffer[0] === 0xD0 && buffer[1] === 0xCF && buffer[2] === 0x11 && buffer[3] === 0xE0) {
        return { format: 'xls', certain: true };
    }

    return { format: detectContentFormat(buffer), certain: false };
}

/**
 * Identify a zip-based format from its entries
 * @returns {Promise<string|null>}
 */
async function detectZipFormat(zip) {
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
    return null;
}

/**
 * Identify the non-container formats (Lotus, dBase, text based)
 * @returns {string|null}
 */
function detectContentFormat(buffer) {
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