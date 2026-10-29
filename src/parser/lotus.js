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

        // Formula cell: only the calculated value is kept; the formula tokens (bytes 15+) are not decoded
        if (record.type === LOTUS_RECORDS.FORMULA) {
            const format = readUInt8(data, 0);
            const col = readUInt16LE(data, 1);
            const row = readUInt16LE(data, 3);
            const value = readFloat64LE(data, 5);

            cells.push({
                row,
                col,
                sheet: currentSheet,
                value,
                type: 'formula',
                format: parseFormatting(format)
            });
        }

        // String result from formula
        if (record.type === LOTUS_RECORDS.STRING) {
            // format (1), column (2), row (2), then the null-terminated string
            const value = String.fromCharCode(...data.slice(5)).replace(/\0/g, '');
            // Update the last formula cell's value
            if (cells.length > 0 && cells[cells.length - 1].type === 'formula') {
                cells[cells.length - 1].value = value;
            }
        }
    });

    return { cells, sheetNames };
}

/*
 * Lotus 1-2-3 release 3+ (WK3, WK4) and 1-2-3 97 / Millennium (.123)
 *
 * These files open with a 26-byte BOF whose first word is the file version:
 * 0x1000 (WK3), 0x1002 (WK4), 0x1003 (1-2-3 97) and 0x1005 (Millennium).
 * Cell records address a cell as row u16, sheet u8, column u8 and the two
 * families differ only in how numbers are stored.
 */

// Records shared by WK3/WK4 and .123
const WK3_RECORDS = {
    EOF: 0x01,
    PASSWORD: 0x02,       // WK3/WK4: file is password protected
    COLUMNWIDTH: 0x07,    // WK3/WK4: sheet u8, window u8, 2 bytes, then (col u8, width u8) pairs
    HIDDENCOLUMN: 0x08,   // WK3/WK4: sheet u8, window u8, 2 bytes, then col u8 list
    ERRCELL: 0x14,        // WK3/WK4: address only, cell shows ERR
    NACELL: 0x15,         // WK3/WK4: address only, cell shows NA
    LABEL: 0x16,          // address, then null-terminated text whose first byte is the alignment prefix
    NUMBER: 0x17,         // WK3/WK4: address, 10-byte extended float
    SMALLNUMBER: 0x18,    // WK3/WK4: address, 16-bit packed number
    FORMULA: 0x19,        // WK3/WK4: address, 10-byte cached result, then the formula tokens
    STRING: 0x1A,         // address, null-terminated text: string result of the formula at that address
    EXTENDED: 0x1B,       // u16 subtype, then the subtype's payload
    SHEETNAME: 0x23,      // .123: u16 subtype (14000), sheet u16, null-terminated name
    NUMBER32: 0x25,       // .123: address, 32-bit packed number
    NOTE: 0x26,           // .123: address, null-terminated comment text
    IEEENUMBER: 0x27,     // .123: address, IEEE 754 double
    IEEEFORMULA: 0x28,    // .123: address, IEEE 754 double cached result, then the formula tokens
};

// EXTENDED (0x1B) subtype carrying a sheet name: sheet u16, null-terminated name
const EXTENDED_NAMED_SHEET = 14000;

const WK3_VERSIONS = [0x1000, 0x1001, 0x1002];
const WK123_VERSIONS = [0x1003, 0x1004, 0x1005, 0x1006];

// Total cells (rows x columns, all sheets) a WK3+ file may expand to; a full 65536 x 256 sheet is 16.7M
const MAX_WK3_GRID_CELLS = 1 << 25;

// Upper half of code page 850, the default (group 1) character set of LMBCS
const CP850_HIGH = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐' +
    '└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´­±‗¾¶§÷¸°¨·¹³²■ ';

/**
 * Detect the record layout from the BOF record
 * @return {'wk3'|'123'|'wk1'}
 */
function detectLotusFamily(data) {
    if (data.length >= 6 && readUInt16LE(data, 2) >= 2) {
        const version = readUInt16LE(data, 4);
        if (WK3_VERSIONS.includes(version)) return 'wk3';
        if (WK123_VERSIONS.includes(version)) return '123';
    }
    return 'wk1';
}

/**
 * Decode a null-terminated LMBCS string (Lotus Multi-Byte Character Set).
 * Plain bytes are ASCII / code page 850; 0x01-0x1F introduce a group byte sequence.
 */
function decodeLMBCS(bytes, start) {
    let text = '';
    let i = start;
    while (i < bytes.length && bytes[i] !== 0) {
        const b = bytes[i++];
        if (b >= 0x80) {
            text += CP850_HIGH[b - 0x80];
        } else if (b >= 0x20) {
            text += String.fromCharCode(b);
        } else if (b === 0x01) {
            // Group 1 (code page 850)
            if (i >= bytes.length || bytes[i] === 0) break;
            const c = bytes[i++];
            text += c >= 0x80 ? CP850_HIGH[c - 0x80] : String.fromCharCode(c);
        } else if (b === 0x0F) {
            // Control characters: C0 controls are shifted up by 0x20
            if (i >= bytes.length || bytes[i] === 0) break;
            const c = bytes[i++];
            text += String.fromCharCode(c >= 0x20 && c < 0x40 ? c - 0x20 : c);
        } else if (b === 0x14) {
            // Unicode: UTF-16 code unit, big-endian
            if (i + 1 >= bytes.length) break;
            text += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
            i += 2;
        } else if (b >= 0x10 && b <= 0x13) {
            // Double-byte Asian groups: no table available
            i += 2;
            text += '�';
        } else if (b >= 0x02 && b <= 0x0B) {
            // Other single-byte groups: no table available, keep ASCII as is
            if (i >= bytes.length || bytes[i] === 0) break;
            const c = bytes[i++];
            text += c < 0x80 ? String.fromCharCode(c) : '�';
        } else {
            text += String.fromCharCode(b);
        }
    }
    return text;
}

/**
 * Read an 80-bit x87 extended precision float (little-endian)
 */
function readFloat80LE(data, offset) {
    const lo = data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
    const hi = data[offset + 4] | (data[offset + 5] << 8) | (data[offset + 6] << 16) | (data[offset + 7] << 24);
    const signExp = data[offset + 8] | (data[offset + 9] << 8);
    const sign = signExp & 0x8000 ? -1 : 1;
    const exp = signExp & 0x7FFF;
    const mantHi = hi >>> 0;
    const mantLo = lo >>> 0;

    if (exp === 0x7FFF) {
        // Infinity when the fraction (all bits but the explicit integer bit) is zero, NaN otherwise
        return (mantHi & 0x7FFFFFFF) === 0 && mantLo === 0 ? sign * Infinity : NaN;
    }
    if (mantHi === 0 && mantLo === 0) {
        return sign * 0;
    }

    // value = mantissa / 2^63 * 2^(exp - 16383)
    const mantissa = mantHi * 2 ** -31 + mantLo * 2 ** -63;
    let e = (exp === 0 ? 1 : exp) - 16383;
    if (e > 1100) return sign * Infinity;
    if (e < -1200) return sign * 0;
    // Scale in two steps so the intermediate power of two stays representable
    const half = Math.trunc(e / 2);
    return sign * mantissa * 2 ** half * 2 ** (e - half);
}

/**
 * 16-bit packed number (WK3/WK4 SMALLNUMBER)
 * bit 0 clear: the value is the signed integer in bits 1-15
 * bit 0 set: the signed integer in bits 4-15 times a factor chosen by bits 1-3
 */
function smallNumberToDouble(raw) {
    const value = raw > 32767 ? raw - 65536 : raw;
    if (value & 0x0001) {
        const factors = [5000, 500, 0.05, 0.005, 0.0005, 0.00005, 0.0625, 0.015625];
        return factors[(value >> 1) & 0x07] * (value >> 4);
    }
    return value >> 1;
}

/**
 * 32-bit packed number (.123 NUMBER32)
 * bits 6-31: unsigned mantissa, bits 0-3: power of ten, bit 4: divide instead of multiply, bit 5: negative
 */
function number32ToDouble(raw) {
    let value = raw >>> 6;
    const power = raw & 0x0F;
    if (power) {
        value = raw & 0x10 ? value / 10 ** power : value * 10 ** power;
    }
    return raw & 0x20 ? -value : value;
}

// Strip the label alignment prefix
function splitLabel(text) {
    const prefix = text.charCodeAt(0);
    if (prefix === 0x27 || prefix === 0x22 || prefix === 0x5E || prefix === 0x5C) {
        return { value: text.slice(1), alignment: parseLabelFormat(prefix) };
    }
    if (prefix === 0x7C) {
        // '|' marks a printer command row: not cell content
        return null;
    }
    return { value: text, alignment: null };
}

/**
 * Parse the cells of a WK3/WK4 ('wk3') or 1-2-3 97/Millennium ('123') file
 */
function parseWK3Records(data, family) {
    const is123 = family === '123';
    const cells = new Map();          // "sheet:row:col" -> cell
    const sheetNames = {};
    const sheetColumns = {};
    const notes = [];

    const setCell = (cell) => {
        cells.set(`${cell.sheet}:${cell.row}:${cell.col}`, cell);
    };
    const address = (rec) => ({
        row: readUInt16LE(rec, 0),
        sheet: rec[2],
        col: rec[3]
    });
    const columnsOf = (sheet) => sheetColumns[sheet] || (sheetColumns[sheet] = {});

    // The BOF record is skipped; records are read up to EOF or the first truncated record
    let offset = 4 + readUInt16LE(data, 2);

    while (offset + 4 <= data.length) {
        const type = readUInt16LE(data, offset);
        const length = readUInt16LE(data, offset + 2);
        if (offset + 4 + length > data.length) break;
        const rec = data.subarray(offset + 4, offset + 4 + length);
        offset += 4 + length;

        if (type === WK3_RECORDS.EOF) break;

        if (!is123) {
            switch (type) {
                case WK3_RECORDS.PASSWORD:
                    throw new Error('Password-protected Lotus files are not supported');

                case WK3_RECORDS.COLUMNWIDTH:
                    if (length >= 4 && rec[1] === 0) {
                        const columns = columnsOf(rec[0]);
                        for (let i = 4; i + 1 < length; i += 2) {
                            const width = rec[i + 1];
                            columns[rec[i]] = Object.assign(columns[rec[i]] || {}, {
                                width: width * 8.43
                            });
                        }
                    }
                    continue;

                case WK3_RECORDS.HIDDENCOLUMN:
                    if (length >= 4 && rec[1] === 0) {
                        const columns = columnsOf(rec[0]);
                        for (let i = 4; i < length; i++) {
                            columns[rec[i]] = Object.assign(columns[rec[i]] || {}, { hidden: true });
                        }
                    }
                    continue;

                case WK3_RECORDS.ERRCELL:
                case WK3_RECORDS.NACELL:
                    if (length >= 4) {
                        setCell({
                            ...address(rec),
                            value: type === WK3_RECORDS.ERRCELL ? '#ERR!' : '#N/A',
                            type: 'error'
                        });
                    }
                    continue;

                case WK3_RECORDS.NUMBER:
                    if (length >= 14) {
                        setCell({ ...address(rec), value: readFloat80LE(rec, 4), type: 'number' });
                    }
                    continue;

                case WK3_RECORDS.SMALLNUMBER:
                    if (length >= 6) {
                        setCell({ ...address(rec), value: smallNumberToDouble(readUInt16LE(rec, 4)), type: 'number' });
                    }
                    continue;

                case WK3_RECORDS.FORMULA:
                    if (length >= 14) {
                        const value = readFloat80LE(rec, 4);
                        // A NaN result stands for a string result (given by the STRING record that follows) or an error
                        setCell({ ...address(rec), value: Number.isNaN(value) ? '' : value, type: 'formula' });
                    }
                    continue;
            }
        } else {
            switch (type) {
                case WK3_RECORDS.SHEETNAME:
                    if (length > 4) {
                        sheetNames[readUInt16LE(rec, 2)] = decodeLMBCS(rec, 4);
                    }
                    continue;

                case WK3_RECORDS.NUMBER32:
                    if (length >= 8) {
                        const raw = (rec[4] | (rec[5] << 8) | (rec[6] << 16) | (rec[7] << 24)) >>> 0;
                        setCell({ ...address(rec), value: number32ToDouble(raw), type: 'number' });
                    }
                    continue;

                case WK3_RECORDS.IEEENUMBER:
                    if (length >= 12) {
                        setCell({ ...address(rec), value: readFloat64LE(rec, 4), type: 'number' });
                    }
                    continue;

                case WK3_RECORDS.IEEEFORMULA:
                    if (length >= 12) {
                        const value = readFloat64LE(rec, 4);
                        setCell({ ...address(rec), value: Number.isNaN(value) ? '' : value, type: 'formula' });
                    }
                    continue;

                case WK3_RECORDS.NOTE:
                    if (length >= 4) {
                        notes.push({ ...address(rec), text: decodeLMBCS(rec, 4) });
                    }
                    continue;
            }
        }

        // Records common to both families
        switch (type) {
            case WK3_RECORDS.LABEL:
                if (length >= 4) {
                    const label = splitLabel(decodeLMBCS(rec, 4));
                    if (label) {
                        setCell({ ...address(rec), value: label.value, type: 'string', alignment: label.alignment });
                    }
                }
                break;

            case WK3_RECORDS.STRING:
                if (length >= 4) {
                    const addr = address(rec);
                    const cell = cells.get(`${addr.sheet}:${addr.row}:${addr.col}`);
                    if (cell && cell.type === 'formula') {
                        cell.value = decodeLMBCS(rec, 4);
                    }
                }
                break;

            case WK3_RECORDS.EXTENDED:
                if (length > 4 && readUInt16LE(rec, 0) === EXTENDED_NAMED_SHEET) {
                    sheetNames[readUInt16LE(rec, 2)] = decodeLMBCS(rec, 4);
                }
                break;
        }
    }

    return { cells: [...cells.values()], sheetNames, sheetColumns, notes };
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
/**
 * @param {Array} cells - parsed cells
 * @param {Object} columns - column settings of sheet 0, or of every sheet when no sheetColumns is given
 * @param {Object} sheetNames - sheet number -> name
 * @param {Object} [extra] - WK3+ only: sheetColumns (sheet -> columns), notes ([{sheet,row,col,text}])
 */
function convertToJspreadsheet(cells, columns, sheetNames, extra = {}) {
    const sheets = {};
    const notesBySheet = {};

    // Group cells by sheet
    cells.forEach(cell => {
        if (!sheets[cell.sheet]) {
            sheets[cell.sheet] = [];
        }
        sheets[cell.sheet].push(cell);
    });

    (extra.notes || []).forEach(note => {
        if (!sheets[note.sheet]) {
            sheets[note.sheet] = [];
        }
        (notesBySheet[note.sheet] = notesBySheet[note.sheet] || []).push(note);
    });

    // Named sheets are kept even when they are empty
    if (extra.sheetColumns) {
        Object.keys(sheetNames).forEach(sheetNum => {
            if (!sheets[sheetNum]) {
                sheets[sheetNum] = [];
            }
        });
    }

    const worksheets = [];

    Object.keys(sheets).forEach(sheetNum => {
        const sheetCells = sheets[sheetNum];
        const sheetNotes = notesBySheet[sheetNum] || [];
        const sheetColumns = extra.sheetColumns ? (extra.sheetColumns[sheetNum] || {}) : columns;

        if (sheetCells.length === 0 && sheetNotes.length === 0) {
            if (extra.sheetColumns && sheetNames[sheetNum] !== undefined) {
                worksheets.push({
                    worksheetName: sheetNames[sheetNum],
                    data: [],
                    columns: [],
                    rows: {},
                    cells: {},
                    style: {},
                    mergeCells: {}
                });
            }
            return;
        }

        // Loop instead of Math.max(...): spreading large cell lists overflows the call stack
        let maxRow = -1;
        let maxCol = -1;
        for (const c of sheetCells.concat(sheetNotes)) {
            if (c.row > maxRow) maxRow = c.row;
            if (c.col > maxCol) maxCol = c.col;
        }

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

            // Add style if present
            const cssStyle = formatToCSS(cell.format, cell.alignment);
            if (cssStyle) {
                styleObj[address] = cssStyle;
            }
        });

        // Create columns array
        const columnsArray = [];
        for (let c = 0; c <= maxCol; c++) {
            const col = sheetColumns[c] || {};
            columnsArray.push({
                width: col.width || 100,
                title: getColumnName(c),
                type: 'text',
                hidden: col.hidden || false
            });
        }

        const worksheet = {
            worksheetName: sheetNames[sheetNum] || `Sheet${parseInt(sheetNum) + 1}`,
            data,
            columns: columnsArray,
            rows: {},
            cells: cellsObj,
            style: styleObj,
            mergeCells: {}
        };

        if (sheetNotes.length > 0) {
            worksheet.comments = {};
            sheetNotes.forEach(note => {
                worksheet.comments[getCellNameFromCoords(note.col, note.row)] = note.text;
            });
        }

        worksheets.push(worksheet);
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

    // WK3/WK4/.123 use their own record layout, told apart by the BOF version
    const family = detectLotusFamily(data);
    if (family !== 'wk1') {
        const { cells, sheetNames, sheetColumns, notes } = parseWK3Records(data, family);

        // Sheets are returned as dense grids: refuse damaged files whose stray addresses would need billions of cells
        const extents = {};
        for (const c of cells.concat(notes)) {
            const e = extents[c.sheet] || (extents[c.sheet] = { rows: 0, cols: 0 });
            if (c.row >= e.rows) e.rows = c.row + 1;
            if (c.col >= e.cols) e.cols = c.col + 1;
        }
        const gridSize = Object.values(extents).reduce((sum, e) => sum + e.rows * e.cols, 0);
        if (gridSize > MAX_WK3_GRID_CELLS) {
            throw new Error('Invalid Lotus file: cell addresses exceed the supported sheet size');
        }

        return { worksheets: convertToJspreadsheet(cells, {}, sheetNames, { sheetColumns, notes }) };
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
