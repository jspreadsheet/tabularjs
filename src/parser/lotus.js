import { loadAsBuffer, parse } from '../utils/loader.js';
import {
    getColumnName,
    getCellNameFromCoords,
    readUInt8,
    readUInt16LE,
    readInt16LE,
    readFloat64LE
} from '../utils/helpers.js';

// Lotus Record Types
const LOTUS_RECORDS = {
    BOF: 0x00,          // Beginning of File
    EOF: 0x01,          // End of File
    CALCMODE: 0x02,     // Calculation mode
    CALCORDER: 0x03,    // Calculation order
    SPLIT: 0x04,        // Split window
    SYNC: 0x05,         // Synchronize windows
    RANGE: 0x06,        // Active range
    WINDOW1: 0x07,      // Window settings
    COLW1: 0x08,        // Column width
    WINTWO: 0x09,       // Window 2 settings
    NAME: 0x0B,         // Range name
    BLANK: 0x0C,        // Blank cell
    INTEGER: 0x0D,      // Integer cell
    NUMBER: 0x0E,       // Number cell
    LABEL: 0x0F,        // Label/String cell
    FORMULA: 0x10,      // Formula cell
    TABLE: 0x18,        // Data table
    QRANGE: 0x19,       // Query range
    PRANGE: 0x1A,       // Print range
    SRANGE: 0x1B,       // Sort range
    FRANGE: 0x1C,       // Fill range
    KRANGE1: 0x1D,      // Primary sort key
    HRANGE: 0x20,       // Distribution range
    KRANGE2: 0x23,      // Secondary sort key
    PROTEC: 0x24,       // Protection
    FOOTER: 0x25,       // Page footer
    HEADER: 0x26,       // Page header
    SETUP: 0x27,        // Print setup
    MARGINS: 0x28,      // Print margins
    LABELFMT: 0x29,     // Label format
    TITLES: 0x2A,       // Print titles
    SHEETJS: 0x2B,      // Sheet settings
    GRAPH: 0x2D,        // Graph settings
    NGRAPH: 0x2E,       // Named graph
    CALCCOUNT: 0x2F,    // Calculation count
    UNFORMATTED: 0x30,  // Unformatted mode
    CURSORW12: 0x31,    // Cursor position
    WINDOW: 0x32,       // Window settings
    STRING: 0x33,       // String formula result
    PASSWORD: 0x37,     // File password
    LOCKED: 0x38,       // Locked ranges
    QUERY: 0x3C,        // Query settings
    QUERYNAME: 0x3D,    // Query name
    PRINT: 0x3E,        // Print record
    PRINTNAME: 0x3F,    // Print range name
    GRAPH2: 0x40,       // Graph settings 2
    GRAPHNAME: 0x41,    // Graph name
    ZOOM: 0x42,         // Zoom factor
    SYMSPLIT: 0x43,     // Symmetric split
    NSROWS: 0x44,       // Number of rows
    NSCOLS: 0x45,       // Number of columns
    RULER: 0x46,        // Ruler settings
    NNAME: 0x47,        // Named range
    ACOMM: 0x48,        // Attached comment
    AMACRO: 0x49,       // Attached macro
    PARSE: 0x4A,        // Parse information

    // WK3/WK4 Extended records
    FONTFACE: 0x7E,     // Font face name
    STYLE: 0x7F,        // Style definition
    SHEETNAME: 0x8D,    // Sheet name (WK3+)
    SHEETINFO: 0x8E,    // Sheet information (WK3+)
};

// Lotus format codes
const FORMAT_CODES = {
    0: 'Fixed',          // Fixed decimal
    1: 'Scientific',     // Scientific notation
    2: 'Currency',       // Currency
    3: 'Percent',        // Percent
    4: 'Comma',          // Comma separated
    5: 'General',        // General
    6: '+/-',            // +/- format
    7: 'General',        // General
    15: 'Date',          // Date formats (various)
    16: 'Time',          // Time formats
};

// Parse Lotus records
function parseLotusRecords(data) {
    const records = [];
    let offset = 0;

    while (offset + 4 <= data.length) {
        const type = readUInt16LE(data, offset);
        const length = readUInt16LE(data, offset + 2);

        if (offset + 4 + length > data.length) break;

        records.push({
            type,
            typeName: Object.keys(LOTUS_RECORDS).find(k => LOTUS_RECORDS[k] === type) || `UNKNOWN_${type.toString(16)}`,
            length,
            offset: offset + 4,
            data: data.slice(offset + 4, offset + 4 + length)
        });

        offset += 4 + length;
    }

    return records;
}

// Parse label/string format
function parseLabelFormat(format) {
    const alignment = {
        0x27: 'left',      // ' (apostrophe) - left align
        0x22: 'right',     // " (quote) - right align
        0x5E: 'center',    // ^ (caret) - center
        0x5C: 'fill',      // \ (backslash) - fill/repeat
    };
    return alignment[format] || 'left';
}

// Parse cell formatting attributes
function parseFormatting(attrs) {
    const style = {
        format: null,
        font: {},
        alignment: {},
        colors: {},
        borders: {}
    };

    if (!attrs) return style;

    // Format code (bits 0-6)
    const formatCode = attrs & 0x7F;
    style.format = FORMAT_CODES[formatCode] || 'General';

    // Protection (bit 7)
    style.protected = (attrs & 0x80) !== 0;

    return style;
}

// Parse column widths
function parseColumnWidths(records) {
    const columns = {};

    records.forEach(record => {
        if (record.type === LOTUS_RECORDS.COLW1) {
            const data = record.data;
            const col = readUInt8(data, 0);
            const width = readUInt8(data, 1);
            columns[col] = {
                width: width * 8.43, // Convert to approximate pixels
                hidden: width === 0
            };
        }
    });

    return columns;
}

// Parse cells
function parseCells(records) {
    const cells = [];
    const sheetNames = {};
    let currentSheet = 0;

    records.forEach(record => {
        const data = record.data;

        // Sheet name (WK3+)
        if (record.type === LOTUS_RECORDS.SHEETNAME) {
            const sheetNum = readUInt16LE(data, 0);
            const nameLen = readUInt8(data, 2);
            const name = String.fromCharCode(...data.slice(3, 3 + nameLen));
            sheetNames[sheetNum] = name;
        }

        // Blank cell
        if (record.type === LOTUS_RECORDS.BLANK) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value: '',
                type: 'blank',
                format: parseFormatting(format)
            });
        }

        // Integer cell
        if (record.type === LOTUS_RECORDS.INTEGER) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);
            const value = readInt16LE(data, 5);

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value,
                type: 'number',
                format: parseFormatting(format)
            });
        }

        // Number cell
        if (record.type === LOTUS_RECORDS.NUMBER) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);
            const value = readFloat64LE(data, 5);

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value,
                type: 'number',
                format: parseFormatting(format)
            });
        }

        // Label/String cell
        if (record.type === LOTUS_RECORDS.LABEL) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);
            const labelFormat = String.fromCharCode(readUInt8(data, 5));
            const value = String.fromCharCode(...data.slice(6)).replace(/\0/g, '');

            const alignment = parseLabelFormat(labelFormat.charCodeAt(0));

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value,
                type: 'string',
                format: parseFormatting(format),
                alignment
            });
        }

        // Formula cell
        if (record.type === LOTUS_RECORDS.FORMULA) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);
            const value = readFloat64LE(data, 5);
            const formulaLen = readUInt16LE(data, 13);
            const formulaData = data.slice(15, 15 + formulaLen);

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value,
                type: 'formula',
                formula: Array.from(formulaData),
                format: parseFormatting(format)
            });
        }

        // String result from formula
        if (record.type === LOTUS_RECORDS.STRING) {
            const value = String.fromCharCode(...data).replace(/\0/g, '');
            // Update the last formula cell's value
            if (cells.length > 0 && cells[cells.length - 1].type === 'formula') {
                cells[cells.length - 1].value = value;
            }
        }
    });

    return { cells, sheetNames };
}

// Convert format to CSS
function formatToCSS(format, alignment) {
    const css = [];

    if (alignment) {
        css.push(`text-align: ${alignment}`);
    }

    // Add more CSS based on format attributes
    if (format && format.protected) {
        css.push('pointer-events: none');
    }

    return css.join('; ');
}

// Convert to Jspreadsheet format
function convertToJspreadsheet(cells, columns, sheetNames) {
    const sheets = {};

    // Group cells by sheet
    cells.forEach(cell => {
        if (!sheets[cell.sheet]) {
            sheets[cell.sheet] = [];
        }
        sheets[cell.sheet].push(cell);
    });

    const worksheets = [];

    Object.keys(sheets).forEach(sheetNum => {
        const sheetCells = sheets[sheetNum];

        if (sheetCells.length === 0) return;

        const maxRow = Math.max(...sheetCells.map(c => c.row));
        const maxCol = Math.max(...sheetCells.map(c => c.col));

        // Create data array
        const data = [];
        for (let r = 0; r <= maxRow; r++) {
            data[r] = [];
            for (let c = 0; c <= maxCol; c++) {
                data[r][c] = '';
            }
        }

        const cellsObj = {};
        const styleObj = {};

        // Fill data and styles
        sheetCells.forEach(cell => {
            const address = getCellNameFromCoords(cell.col, cell.row);
            data[cell.row][cell.col] = cell.value;

            // Add formula if present
            if (cell.type === 'formula' && cell.formula) {
                cellsObj[address] = { formula: cell.formula };
            }

            // Add style if present
            const cssStyle = formatToCSS(cell.format, cell.alignment);
            if (cssStyle) {
                styleObj[address] = cssStyle;
            }
        });

        // Create columns array
        const columnsArray = [];
        for (let c = 0; c <= maxCol; c++) {
            const col = columns[c] || {};
            columnsArray.push({
                width: col.width || 100,
                title: getColumnName(c),
                type: 'text',
                hidden: col.hidden || false
            });
        }

        worksheets.push({
            worksheetName: sheetNames[sheetNum] || `Sheet${parseInt(sheetNum) + 1}`,
            data,
            columns: columnsArray,
            rows: {},
            cells: cellsObj,
            style: styleObj,
            mergeCells: {}
        });
    });

    return worksheets;
}

// Main parser function
/**
 * Parse Lotus from buffer
 */
export function parseLotusBuffer(buffer) {
    const data = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);

    // Check file signature
    const bof = readUInt16LE(data, 0);
    if (bof !== LOTUS_RECORDS.BOF) {
        throw new Error('Invalid Lotus file format');
    }

    // Parse all records
    const records = parseLotusRecords(data);

    // Extract data
    const columns = parseColumnWidths(records);
    const { cells, sheetNames } = parseCells(records);

    // Convert to Jspreadsheet format
    const worksheets = convertToJspreadsheet(cells, columns, sheetNames);

    return { worksheets };
}

/**
 * Parse Lotus file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseLotus(input, options = {}) {
    return parse(async (inp) => {
        const buffer = await loadAsBuffer(inp);
        return parseLotusBuffer(buffer);
    }, input, options);
}

// Export format-specific parsers
export const parseWK1 = parseLotus;
export const parseWK3 = parseLotus;
export const parseWK4 = parseLotus;
