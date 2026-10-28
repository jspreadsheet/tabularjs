import { loadAsBuffer, parse } from '../utils/loader.js';
import CFB from '../utils/cfb.js';
import { decodePTG, BUILTIN_NAMES } from '../utils/ptg-decoder.js';
import { decryptWorkbook } from '../utils/xls-crypto.js';
import {
    getColumnName,
    getCellNameFromCoords,
    readUInt16LE,
    readUInt32LE,
    readFloat64LE,
    convertWidthToPixels
} from '../utils/helpers.js';

// BIFF Record Types
const BIFF_RECORDS = {
    BOF: 0x0809,
    EOF: 0x000A,
    BOUNDSHEET: 0x0085,
    SST: 0x00FC,
    LABELSST: 0x00FD,
    RK: 0x027E,
    MULRK: 0x00BD,
    NUMBER: 0x0203,
    LABEL: 0x0204,
    BOOLERR: 0x0205,
    FORMULA: 0x0006,
    SHRFMLA: 0x04BC,
    ARRAY: 0x0221,
    NAME: 0x0018,
    EXTERNSHEET: 0x0017,
    SUPBOOK: 0x01AE,
    EXTERNNAME: 0x0023,
    CONTINUE: 0x003C,
    ROW: 0x0208,
    DIMENSIONS: 0x0200,
    XF: 0x00E0,
    FORMAT: 0x041E,
    FONT: 0x0031,
    STYLE: 0x0293,
    COLINFO: 0x007D,
    MERGECELLS: 0x00E5,
    HLINK: 0x01B8,
    NOTE: 0x001C,
    TXO: 0x01B6,
    MULBLANK: 0x00BE,
    BLANK: 0x0201,
    STRING: 0x0207
};

// BErr codes of BOOLERR records and formula results
const BIFF_ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };

/**
 * Parse the Shared String Table. [MS-XLS] 2.4.265 SST + 2.5.293 XLUnicodeRichExtendedString: each string is
 * cch (2), flags (1: fHighByte 0x01, fExtSt 0x04, fRichSt 0x08), [cRun (2)], [cbExtRst (4)], characters,
 * [cRun * 4 bytes of formatting runs], [cbExtRst bytes of phonetic data]. The table continues in CONTINUE
 * records (records[index + 1...]); a string's characters may break across them, and the continuation then
 * starts with a new flags byte that can switch between 8-bit and 16-bit characters.
 */
function parseSST(records, index) {
    const segments = [records[index].data];
    for (let k = index + 1; k < records.length && records[k].type === BIFF_RECORDS.CONTINUE; k++) {
        segments.push(records[k].data);
    }
    let seg = 0;
    let pos = 8; // cstTotal (4) + cstUnique (4)
    const unique = segments[0].length >= 8 ? readUInt32LE(segments[0], 4) : 0;

    // move to the next segment when the current one is exhausted; false at the end of the table
    const ensure = () => {
        while (seg < segments.length && pos >= segments[seg].length) {
            seg++;
            pos = 0;
        }
        return seg < segments.length;
    };
    const byte = () => {
        if (!ensure()) throw new RangeError('SST truncated');
        return segments[seg][pos++];
    };
    const u16 = () => byte() | (byte() << 8);
    const u32 = () => (u16() | (u16() << 16)) >>> 0;
    const skip = (n) => {
        while (n > 0 && ensure()) {
            const take = Math.min(n, segments[seg].length - pos);
            pos += take;
            n -= take;
        }
    };

    const strings = [];
    try {
        while (strings.length < unique && ensure()) {
            const cch = u16();
            const flags = byte();
            let high = flags & 0x01;
            const runs = flags & 0x08 ? u16() : 0;
            const ext = flags & 0x04 ? u32() : 0;
            const chars = [];
            while (chars.length < cch) {
                if (pos >= segments[seg].length) {
                    // the characters continue in the next CONTINUE record, after a new flags byte
                    seg++;
                    pos = 0;
                    if (seg >= segments.length) break;
                    high = segments[seg][pos++] & 0x01;
                    continue;
                }
                const data = segments[seg];
                if (high) {
                    if (pos + 1 >= data.length) {
                        pos = data.length;
                        continue;
                    }
                    chars.push(data[pos] | (data[pos + 1] << 8));
                    pos += 2;
                } else {
                    chars.push(data[pos++]);
                }
            }
            let str = '';
            for (let i = 0; i < chars.length; i += 8192) {
                str += String.fromCharCode.apply(null, chars.slice(i, i + 8192));
            }
            strings.push(str);
            skip(runs * 4 + ext);
        }
    } catch (err) {
        // a truncated table: keep the strings read so far
    }

    return strings;
}

/**
 * Assign every record to the sheet whose substream holds it (record.sheet = index into the BOUNDSHEET list,
 * -1 for the workbook globals and for chart substreams). A substream starts with a BOF at the stream offset that
 * its BOUNDSHEET records; charts embedded in a worksheet nest their own BOF ... EOF inside it.
 */
function assignSheets(records, sheetOffsets) {
    const byOffset = new Map(sheetOffsets.map((off, i) => [off, i]));
    const stack = []; // sheet index of each open substream
    let nextByOrder = 0;
    for (const record of records) {
        if (record.type === BIFF_RECORDS.BOF) {
            const dt = record.data.length >= 4 ? readUInt16LE(record.data, 2) : 0;
            let sheet = -1;
            if (stack.length === 0 && dt !== 0x0005) {
                // a top-level substream after the globals: find it by offset, else by order
                sheet = byOffset.has(record.offset - 4) ? byOffset.get(record.offset - 4) : nextByOrder;
                nextByOrder = sheet + 1;
                if (dt === 0x0020) sheet = -1; // chart sheet: no cells
            }
            stack.push(sheet);
            record.sheet = -1;
            continue;
        }
        record.sheet = stack.length === 1 ? stack[0] : -1;
        if (record.type === BIFF_RECORDS.EOF && stack.length) {
            stack.pop();
            record.sheet = -1;
        }
    }
}

function parseBIFFRecords(data) {
    const records = [];
    let offset = 0;

    while (offset + 4 <= data.length) {
        const type = readUInt16LE(data, offset);
        const length = readUInt16LE(data, offset + 2);

        if (offset + 4 + length > data.length) break;

        records.push({
            type,
            typeName: Object.keys(BIFF_RECORDS).find(k => BIFF_RECORDS[k] === type) || `UNKNOWN_${type.toString(16)}`,
            length,
            offset: offset + 4,
            data: data.slice(offset + 4, offset + 4 + length)
        });

        offset += 4 + length;
    }

    return records;
}

function decodeRK(rk) {
    // RK is an encoded number format
    const num = readUInt32LE(new Uint8Array([rk[0], rk[1], rk[2], rk[3]]), 0);
    const isInteger = (num & 0x02) !== 0;
    const div100 = (num & 0x01) !== 0;

    let value;
    if (isInteger) {
        value = num >> 2;
    } else {
        // IEEE float - need to reconstruct the double
        const bytes = new Uint8Array(8);
        bytes[4] = rk[0] & 0xFC;
        bytes[5] = rk[1];
        bytes[6] = rk[2];
        bytes[7] = rk[3];
        value = new DataView(bytes.buffer).getFloat64(0, true);
    }

    return div100 ? value / 100 : value;
}

function parseFonts(records) {
    const fonts = [];

    for (const record of records) {
        if (record.type === BIFF_RECORDS.FONT) {
            const data = record.data;
            const height = readUInt16LE(data, 0);
            const options = readUInt16LE(data, 2);
            const colorIndex = readUInt16LE(data, 4);
            const weight = readUInt16LE(data, 6);
            const escapement = readUInt16LE(data, 8);
            const underline = data[10];
            const family = data[11];
            const charset = data[12];
            const nameLen = data[14];
            const flags = data[15]; // String flags byte (after length)

            // Parse font name based on encoding flag
            let name = '';
            const isCompressed = (flags & 0x01) === 0;
            let pos = 16; // Characters start after length and flags

            if (isCompressed) {
                // 8-bit characters
                for (let i = 0; i < nameLen && pos < data.length; i++, pos++) {
                    const byte = data[pos];
                    // Filter out control characters (0x00-0x1F)
                    if (byte >= 0x20) {
                        name += String.fromCharCode(byte);
                    }
                }
            } else {
                // 16-bit characters
                for (let i = 0; i < nameLen && pos + 1 < data.length; i++, pos += 2) {
                    const char = readUInt16LE(data, pos);
                    // Filter out control characters (0x00-0x1F)
                    if (char >= 0x20) {
                        name += String.fromCharCode(char);
                    }
                }
            }

            // Note: Font weight in XLS FONT records is often unreliable or incorrectly encoded by Excel.
            // The xlsx library intentionally skips parsing font styles from FONT records for XLS files.
            // We follow the same approach and do not set bold=true from FONT records.
            fonts.push({
                name,
                height: height / 20, // Convert to points
                bold: false, // Don't trust FONT record weight field - it's unreliable in XLS files
                italic: (options & 0x0002) !== 0,
                strikeout: (options & 0x0008) !== 0,
                underline,
                colorIndex,
                family,
                charset
            });
        }
    }

    return fonts;
}

function parseFormats(records) {
    const formats = {};

    for (const record of records) {
        if (record.type === BIFF_RECORDS.FORMAT) {
            const data = record.data;
            const formatIndex = readUInt16LE(data, 0);
            const strLen = readUInt16LE(data, 2);

            // Read format string - BIFF8 uses Unicode (2 bytes per char) or ASCII (1 byte per char)
            // Check if it's Unicode or ASCII by looking at the options byte
            const options = data[4];
            const isUnicode = (options & 0x01) !== 0;

            let formatString;
            if (isUnicode) {
                // Unicode: 2 bytes per character
                const chars = [];
                for (let i = 0; i < strLen; i++) {
                    chars.push(readUInt16LE(data, 5 + i * 2));
                }
                formatString = String.fromCharCode(...chars);
            } else {
                // ASCII: 1 byte per character
                formatString = String.fromCharCode(...data.slice(5, 5 + strLen));
            }

            formats[formatIndex] = formatString;
        }
    }

    return formats;
}

function parseXF(records, formats, fonts) {
    const xfStyles = [];

    for (const record of records) {
        if (record.type === BIFF_RECORDS.XF) {
            const data = record.data;
            const fontIndex = readUInt16LE(data, 0);
            const formatIndex = readUInt16LE(data, 2);

            // Parse XF structure for BIFF8 (following xlrd implementation)
            // The structure is: font(2), format(2), type_par(2), align1(1), rotation(1), align2(1), used(1), brdbkg1(4), brdbkg2(4), brdbkg3(2)
            // Bytes 0-1: font index
            // Bytes 2-3: format index
            // Bytes 4-5: type/parent style
            // Byte 6: alignment 1
            // Byte 7: rotation
            // Byte 8: alignment 2
            // Byte 9: used attributes flags
            // Bytes 10-13: Border styles and colors (brdbkg1)
            // Bytes 14-17: Border colors continued (brdbkg2)
            // Bytes 18-19: Pattern colors (brdbkg3)

            const align1 = data[6];
            const align2 = data[8];
            const brdbkg1 = readUInt32LE(data, 10);  // Border field is at bytes 10-13!
            const brdbkg2 = readUInt32LE(data, 14);
            const brdbkg3 = readUInt16LE(data, 18);

            // Alignment
            const horizontalAlign = align1 & 0x07;
            const wrapText = ((align1 >> 3) & 0x01) !== 0;
            const verticalAlign = (align1 >> 4) & 0x07;

            // Borders (from brdbkg1, bytes 10-13)
            const borders = {
                left: brdbkg1 & 0x0F,
                right: (brdbkg1 >> 4) & 0x0F,
                top: (brdbkg1 >> 8) & 0x0F,
                bottom: (brdbkg1 >> 12) & 0x0F
            };

            // Fill pattern (from brdbkg2, bits 26-31)
            const patternType = (brdbkg2 >> 26) & 0x3F;

            // Pattern colors from brdbkg3 (bytes 18-19)
            const icvFore = brdbkg3 & 0x7F;
            const icvBack = (brdbkg3 >> 7) & 0x7F;

            xfStyles.push({
                fontIndex,
                font: fonts[fontIndex] ? { ...fonts[fontIndex] } : null, // Clone font object to avoid shared references
                formatIndex,
                format: formats[formatIndex] || null,
                alignment: {
                    horizontal: horizontalAlign,
                    vertical: verticalAlign,
                    wrapText
                },
                borders,
                fill: {
                    patternType,
                    fgColor: icvFore,
                    bgColor: icvBack
                }
            });
        }
    }

    return xfStyles;
}

function parseColumnInfo(records) {
    const columnsBySheet = {};
    for (const record of records) {
        const currentSheet = record.sheet;
        if (currentSheet === undefined || currentSheet < 0) continue;
        if (!columnsBySheet[currentSheet]) {
            columnsBySheet[currentSheet] = [];
        }

        if (record.type === BIFF_RECORDS.COLINFO) {
            const data = record.data;
            const colFirst = readUInt16LE(data, 0);
            const colLast = readUInt16LE(data, 2);
            const width = readUInt16LE(data, 4);
            const xfIndex = readUInt16LE(data, 6);
            const options = readUInt16LE(data, 8);

            for (let col = colFirst; col <= colLast; col++) {
                columnsBySheet[currentSheet][col] = {
                    width: width / 256, // Convert to character units
                    hidden: (options & 0x0001) !== 0,
                    xfIndex
                };
            }
        }
    }

    return columnsBySheet;
}

function parseRowInfo(records) {
    const rowsBySheet = {};
    for (const record of records) {
        const currentSheet = record.sheet;
        if (currentSheet === undefined || currentSheet < 0) continue;
        if (!rowsBySheet[currentSheet]) {
            rowsBySheet[currentSheet] = {};
        }

        if (record.type === BIFF_RECORDS.ROW) {
            const data = record.data;
            const rowIndex = readUInt16LE(data, 0);
            const height = readUInt16LE(data, 6);
            const options = readUInt16LE(data, 12);

            rowsBySheet[currentSheet][rowIndex] = {
                height: height / 20, // Convert to points
                hidden: (options & 0x0020) !== 0
            };
        }
    }

    return rowsBySheet;
}

function parseMergedCells(records) {
    const mergedCellsBySheet = {};
    for (const record of records) {
        const currentSheet = record.sheet;
        if (currentSheet === undefined || currentSheet < 0) continue;
        if (!mergedCellsBySheet[currentSheet]) {
            mergedCellsBySheet[currentSheet] = [];
        }

        if (record.type === BIFF_RECORDS.MERGECELLS) {
            const data = record.data;
            const count = readUInt16LE(data, 0);

            for (let i = 0; i < count; i++) {
                const offset = 2 + i * 8;
                // BIFF MERGECELLS format: each range is 8 bytes
                const rowFirst = readUInt16LE(data, offset);
                const rowLast = readUInt16LE(data, offset + 2);
                const colFirst = readUInt16LE(data, offset + 4);
                const colLast = readUInt16LE(data, offset + 6);

                mergedCellsBySheet[currentSheet].push({
                    rowFirst,
                    rowLast,
                    colFirst,
                    colLast
                });
            }
        }
    }

    return mergedCellsBySheet;
}

/**
 * Read an XLUnicodeString (cch: 2 bytes) or ShortXLUnicodeString (cch: 1 byte) at pos -> [string, next pos]
 */
function readXLString(data, pos, cchBytes) {
    const cch = cchBytes === 2 ? readUInt16LE(data, pos) : data[pos];
    pos += cchBytes;
    const unicode = data[pos] & 0x01;
    pos++;
    let str = '';
    for (let i = 0; i < cch && pos < data.length; i++) {
        str += String.fromCharCode(unicode ? readUInt16LE(data, pos) : data[pos]);
        pos += unicode ? 2 : 1;
    }
    return [str, pos];
}

/**
 * Workbook-globals records that formulas refer to: SUPBOOK (+ EXTERNNAME), EXTERNSHEET and NAME.
 * -> { sheets, supbooks: [{type, name, sheetNames, externNames}], externSheets: [{supbook, first, last}], names: [{name}] }
 */
function parseFormulaContext(records, sheetNames) {
    const ctx = { sheets: sheetNames, supbooks: [], externSheets: [], names: [] };
    for (let i = 0; i < records.length; i++) {
        const record = records[i];
        if (record.type === BIFF_RECORDS.BOF && readUInt16LE(record.data, 2) === 0x0010) {
            break; // first worksheet: the globals are over
        }
        let data = record.data;
        try {
            if (record.type === BIFF_RECORDS.SUPBOOK) {
                const ctab = readUInt16LE(data, 0);
                const cch = readUInt16LE(data, 2);
                if (cch === 0x0401) {
                    ctx.supbooks.push({ type: 'self', externNames: [] });
                } else if (cch === 0x3A01) {
                    ctx.supbooks.push({ type: 'addin', externNames: [] });
                } else {
                    let [path, pos] = readXLString(data, 2, 2);
                    const sheets = [];
                    for (let k = 0; k < ctab && pos < data.length; k++) {
                        let name;
                        [name, pos] = readXLString(data, pos, 2);
                        sheets.push(name);
                    }
                    // VirtualPath: control characters separate volume / directories; keep the file name
                    const name = path.split(/[\x00-\x08\\/]/).filter(Boolean).pop() || path;
                    ctx.supbooks.push({ type: 'external', name, sheetNames: sheets, externNames: [] });
                }
            } else if (record.type === BIFF_RECORDS.EXTERNNAME && ctx.supbooks.length) {
                // flags (2) + reserved / ixals (4) + ShortXLUnicodeString name (add-in UDF or external name)
                const [name] = readXLString(data, 6, 1);
                ctx.supbooks[ctx.supbooks.length - 1].externNames.push(name);
            } else if (record.type === BIFF_RECORDS.EXTERNSHEET) {
                // the XTI array may continue in CONTINUE records
                const parts = [data];
                for (let k = i + 1; k < records.length && records[k].type === BIFF_RECORDS.CONTINUE; k++) {
                    parts.push(records[k].data);
                }
                if (parts.length > 1) {
                    data = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
                    let off = 0;
                    for (const p of parts) {
                        data.set(p, off);
                        off += p.length;
                    }
                }
                const n = readUInt16LE(data, 0);
                for (let k = 0; k < n && 2 + 6 * k + 6 <= data.length; k++) {
                    ctx.externSheets.push({
                        supbook: readUInt16LE(data, 2 + 6 * k),
                        first: readUInt16LE(data, 4 + 6 * k),
                        last: readUInt16LE(data, 6 + 6 * k)
                    });
                }
            } else if (record.type === BIFF_RECORDS.NAME) {
                // flags (2), chKey (1), cch (1), cce (2), reserved (2), itab (2), reserved (4), name (flags byte + chars)
                const flags = readUInt16LE(data, 0);
                const cch = data[3];
                const unicode = data[14] & 0x01;
                let name = '';
                for (let k = 0, pos = 15; k < cch && pos < data.length; k++, pos += unicode ? 2 : 1) {
                    name += String.fromCharCode(unicode ? readUInt16LE(data, pos) : data[pos]);
                }
                if (flags & 0x0020) {
                    name = BUILTIN_NAMES[name.charCodeAt(0)] || name;
                }
                ctx.names.push({ name });
            }
        } catch (err) {
            // a malformed globals record only makes the formulas that use it undecodable
        }
    }
    return ctx;
}

function parseCellRecords(records, sst, xfStyles, workbookData) {
    const cells = [];
    const shrfmlaRecords = []; // Store SHRFMLA records for second pass
    let lastFormulaCell = null;

    // First pass: collect all cells and SHRFMLA records
    for (const record of records) {
        const data = record.data;
        const sheetIndex = record.sheet;
        if (sheetIndex === undefined || sheetIndex < 0) continue;

        // Parse LABEL (inline string; older writers use it instead of LABELSST)
        if (record.type === BIFF_RECORDS.LABEL) {
            cells.push({
                row: readUInt16LE(data, 0),
                col: readUInt16LE(data, 2),
                value: readXLString(data, 6, 2)[0],
                type: 'string',
                style: xfStyles[readUInt16LE(data, 4)] || null,
                sheetIndex
            });
        }

        // Parse LABELSST (string from SST)
        if (record.type === BIFF_RECORDS.LABELSST) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);
            const sstIndex = readUInt32LE(data, 6);

            cells.push({
                row,
                col,
                value: sst[sstIndex] || '',
                type: 'string',
                style: xfStyles[xfIndex] || null,
                sheetIndex
            });
        }

        // Parse RK (encoded number)
        if (record.type === BIFF_RECORDS.RK) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);
            const value = decodeRK(data.slice(6, 10));

            cells.push({
                row,
                col,
                value,
                type: 'number',
                style: xfStyles[xfIndex] || null,
                sheetIndex
            });
        }

        // Parse NUMBER (IEEE float)
        if (record.type === BIFF_RECORDS.NUMBER) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);
            const value = readFloat64LE(data, 6);

            cells.push({
                row,
                col,
                value,
                type: 'number',
                style: xfStyles[xfIndex] || null,
                sheetIndex
            });
        }

        // Parse FORMULA
        if (record.type === BIFF_RECORDS.FORMULA) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);

            // Formula result is stored in bytes 6-13
            const resultBytes = data.slice(6, 14);
            let value = null;
            let formulaType = 'number';

            // FormulaValue: fExprO = 0xFFFF in bytes 6-7 marks a non-number; byte 0 is its type
            // (0 string in the next STRING record, 1 boolean, 2 error, 3 empty string)
            if (resultBytes[6] === 0xFF && resultBytes[7] === 0xFF) {
                const kind = resultBytes[0];
                if (kind === 0) {
                    formulaType = 'string';
                } else if (kind === 1) {
                    formulaType = 'boolean';
                    value = resultBytes[2] !== 0;
                } else if (kind === 2) {
                    formulaType = 'error';
                    value = BIFF_ERRORS[resultBytes[2]] || '#N/A';
                } else {
                    formulaType = 'string';
                    value = '';
                }
            } else {
                // Numeric result
                value = readFloat64LE(resultBytes, 0);
            }

            // BIFF8 FORMULA record structure:
            // Offset 14-15: grbit (options)
            // Offset 16-19: (chn, reserved)
            // Offset 20-21: cce (formula length)
            // Offset 22+: formula tokens (rgce)
            const formulaLen = readUInt16LE(data, 20);
            const formulaTokens = data.slice(22, 22 + formulaLen);
            const formulaExtra = data.slice(22 + formulaLen); // RgbExtra: tArray constants

            lastFormulaCell = {
                row,
                col,
                value,
                type: 'formula',
                formulaType,
                formula: Array.from(formulaTokens), // Store raw tokens
                formulaExtra,
                style: xfStyles[xfIndex] || null,
                sheetIndex
            };

            cells.push(lastFormulaCell);
        }

        // Parse STRING (follows FORMULA for string results)
        if (record.type === BIFF_RECORDS.STRING && lastFormulaCell) {
            const strLen = readUInt16LE(data, 0);
            const flags = data[2];
            const isCompressed = (flags & 0x01) === 0;

            let str = '';
            let pos = 3;

            if (isCompressed) {
                for (let i = 0; i < strLen && pos < data.length; i++, pos++) {
                    str += String.fromCharCode(data[pos]);
                }
            } else {
                for (let i = 0; i < strLen && pos + 1 < data.length; i++, pos += 2) {
                    str += String.fromCharCode(readUInt16LE(data, pos));
                }
            }

            lastFormulaCell.value = str;
            lastFormulaCell = null;
        }

        // Parse SHRFMLA (Shared Formula)
        // Store for second pass after all cells are collected
        if (record.type === BIFF_RECORDS.SHRFMLA) {
            // SHRFMLA structure (BIFF8):
            // Offset 0-1: First row of range
            // Offset 2-3: Last row of range
            // Offset 4: First column (1 byte)
            // Offset 5: Last column (1 byte)
            // Offset 6: Reserved (1 byte)
            // Offset 7: Number of unused fields following (cce)
            // Offset 8-9: Formula length (2 bytes)
            // Offset 10+: Formula tokens (rgce)
            const firstRow = readUInt16LE(data, 0);
            const lastRow = readUInt16LE(data, 2);
            const firstCol = data[4];
            const lastCol = data[5];
            const formulaLen = readUInt16LE(data, 8);
            const sharedFormula = data.slice(10, 10 + formulaLen);
            const sharedExtra = data.slice(10 + formulaLen);

            // Store for later processing
            shrfmlaRecords.push({
                sheetIndex,
                firstRow,
                lastRow,
                firstCol,
                lastCol,
                formula: Array.from(sharedFormula),
                extra: sharedExtra
            });
        }

        // Parse ARRAY (array formula entered with Ctrl+Shift+Enter); cells point to it with tExp like SHRFMLA
        // RefU (6) + flags (2) + reserved (4) + cce (2) + rgce + RgbExtra
        if (record.type === BIFF_RECORDS.ARRAY) {
            const formulaLen = readUInt16LE(data, 12);
            shrfmlaRecords.push({
                sheetIndex,
                firstRow: readUInt16LE(data, 0),
                lastRow: readUInt16LE(data, 2),
                firstCol: data[4],
                lastCol: data[5],
                formula: Array.from(data.slice(14, 14 + formulaLen)),
                extra: data.slice(14 + formulaLen),
                array: true
            });
        }

        // Parse BLANK
        if (record.type === BIFF_RECORDS.BLANK) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);

            cells.push({
                row,
                col,
                value: '',
                type: 'blank',
                style: xfStyles[xfIndex] || null,
                sheetIndex
            });
        }

        // Parse MULBLANK (multiple blank cells)
        if (record.type === BIFF_RECORDS.MULBLANK) {
            const row = readUInt16LE(data, 0);
            const colFirst = readUInt16LE(data, 2);
            const colLast = readUInt16LE(data, data.length - 2);

            let pos = 4;
            for (let col = colFirst; col <= colLast; col++) {
                const xfIndex = readUInt16LE(data, pos);

                cells.push({
                    row,
                    col,
                    value: '',
                    type: 'blank',
                    style: xfStyles[xfIndex] || null,
                    sheetIndex
                });

                pos += 2;
            }
        }

        // Parse MULRK (multiple RK values in one row)
        if (record.type === BIFF_RECORDS.MULRK) {
            const row = readUInt16LE(data, 0);
            const colFirst = readUInt16LE(data, 2);
            const colLast = readUInt16LE(data, data.length - 2);

            let pos = 4;
            for (let col = colFirst; col <= colLast; col++) {
                const xfIndex = readUInt16LE(data, pos);
                const rkValue = data.slice(pos + 2, pos + 6);
                const value = decodeRK(rkValue);

                cells.push({
                    row,
                    col,
                    value,
                    type: 'number',
                    style: xfStyles[xfIndex] || null,
                    sheetIndex
                });

                pos += 6;
            }
        }

        // Parse BOOLERR (boolean or error)
        if (record.type === BIFF_RECORDS.BOOLERR) {
            const row = readUInt16LE(data, 0);
            const col = readUInt16LE(data, 2);
            const xfIndex = readUInt16LE(data, 4);
            const boolErr = data[6];
            const isError = data[7];

            cells.push({
                row,
                col,
                value: isError ? (BIFF_ERRORS[boolErr] || '#N/A') : (boolErr !== 0),
                type: isError ? 'error' : 'boolean',
                style: xfStyles[xfIndex] || null,
                sheetIndex
            });
        }
    }

    // Second pass: a cell whose formula is tExp (row, col of the base cell) takes the SHRFMLA / ARRAY formula
    // anchored there. Indexed by sheet + base cell: one pass over the cells, not cells x records.
    const shrByBase = new Map();
    for (const shr of shrfmlaRecords) {
        shrByBase.set(`${shr.sheetIndex}:${shr.firstRow}:${shr.firstCol}`, shr);
    }
    for (const cell of cells) {
        if (cell.type === 'formula' && cell.formula && cell.formula.length >= 5 && cell.formula[0] === 0x01) {
            const baseRow = cell.formula[1] | (cell.formula[2] << 8);
            const baseCol = cell.formula[3] | (cell.formula[4] << 8);
            const shr = shrByBase.get(`${cell.sheetIndex}:${baseRow}:${baseCol}`);
            if (shr && cell.row >= shr.firstRow && cell.row <= shr.lastRow &&
                cell.col >= shr.firstCol && cell.col <= shr.lastCol) {
                cell.formula = shr.formula;
                cell.formulaExtra = shr.extra;
                if (shr.array) {
                    // array formulas store absolute references (RgceLoc), so no host-cell adjustment
                    cell.arrayFormula = true;
                } else {
                    cell.sharedFormulaBase = { row: shr.firstRow, col: shr.firstCol };
                }
            }
        }
    }

    return cells;
}

// Convert XF style to CSS string
function styleToCss(style) {
    if (!style) return '';

    const css = [];

    // Font
    if (style.font) {
        if (style.font.name) css.push(`font-family: ${style.font.name}`);
        if (style.font.height) css.push(`font-size: ${style.font.height}pt`);
        if (style.font.bold) css.push('font-weight: bold');
        if (style.font.italic) css.push('font-style: italic');
        if (style.font.underline) css.push('text-decoration: underline');
        if (style.font.strikeout) css.push('text-decoration: line-through');
    }

    // Alignment
    if (style.alignment) {
        const hAlign = ['left', 'center', 'right', 'fill', 'justify', 'center-across-selection', 'distributed'][style.alignment.horizontal];
        const vAlign = ['top', 'middle', 'bottom', 'justify', 'distributed'][style.alignment.vertical];
        if (hAlign) css.push(`text-align: ${hAlign}`);
        if (vAlign) css.push(`vertical-align: ${vAlign}`);
        if (style.alignment.wrapText) css.push('white-space: pre-wrap');
    }

    // Excel default color palette (indexes 0-71)
    // Indexes 0-7: Fixed system colors
    // Indexes 8-63: Default colors (can be overridden by PALETTE record)
    // Indexes 64+: System/automatic colors
    const excelDefaultPalette = [
        '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
        '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
        '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
        '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
        '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
        '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
        '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
        '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
        // System/automatic colors (64+)
        // Index 64: Window Text/Foreground (typically black, but for fills often renders as light color)
        // Index 65: Window Background (typically white)
        '#000000', '#FFFFFF', '#000000', '#000000', '#000000', '#000000', '#000000', '#000000'
    ];

    // Fill pattern / background color
    if (style.fill) {
        const patternType = style.fill.patternType;
        const fgColorIdx = style.fill.fgColor;
        const bgColorIdx = style.fill.bgColor;

        // Pattern type 1 = solid fill, use foreground color as the background
        // Index 64+ are automatic/system colors - only apply fill if real color (< 64)
        if (patternType === 1 && fgColorIdx >= 0 && fgColorIdx < 64 && fgColorIdx < excelDefaultPalette.length) {
            css.push(`background-color: ${excelDefaultPalette[fgColorIdx]}`);
        }
        // Other pattern types could use bgColor or create patterns
        else if (patternType > 1 && bgColorIdx >= 0 && bgColorIdx < 64 && bgColorIdx < excelDefaultPalette.length) {
            css.push(`background-color: ${excelDefaultPalette[bgColorIdx]}`);
        }
    }

    // Font color (from font object, not fill)
    if (style.font && style.font.color) {
        const colorIdx = style.font.color;
        if (colorIdx >= 0 && colorIdx < excelDefaultPalette.length) {
            css.push(`color: ${excelDefaultPalette[colorIdx]}`);
        }
    }

    // Borders
    if (style.borders) {
        if (style.borders.left) css.push('border-left: 1px solid #000');
        if (style.borders.right) css.push('border-right: 1px solid #000');
        if (style.borders.top) css.push('border-top: 1px solid #000');
        if (style.borders.bottom) css.push('border-bottom: 1px solid #000');
    }

    return css.join('; ');
}

function cellsToJspreadsheet(cells, columns, rows, mergedCells, formats, formulaContext = null) {
    if (cells.length === 0) return {
        data: [],
        columns: [],
        rows: {},
        cells: {},
        style: {},
        mergeCells: {}
    };

    // Find dimensions from cells
    let maxRow = 0;
    let maxCol = 0;
    for (const c of cells) {
        if (c.row > maxRow) maxRow = c.row;
        if (c.col > maxCol) maxCol = c.col;
    }

    // Update dimensions based on merged cells
    mergedCells.forEach(merge => {
        if (merge.rowLast > maxRow) maxRow = merge.rowLast;
        if (merge.colLast > maxCol) maxCol = merge.colLast;
    });

    // Create data array
    const data = [];
    for (let r = 0; r <= maxRow; r++) {
        data[r] = [];
        for (let c = 0; c <= maxCol; c++) {
            data[r][c] = '';
        }
    }

    // Jspreadsheet format objects
    const cellsObj = {};
    const styleObj = {};

    // Fill cells and build style/cells objects
    cells.forEach(cell => {
        const address = getCellNameFromCoords(cell.col, cell.row);

        // ALWAYS check for formulas and prefer them over calculated values
        if (cell.type === 'formula') {
            // This cell contains a formula
            if (cell.formula && cell.formula.length > 0) {
                try {
                    // Prepare cell context for relative reference adjustment (for shared formulas)
                    const cellContext = cell.sharedFormulaBase ? {
                        row: cell.row,
                        col: cell.col,
                        baseRow: cell.sharedFormulaBase.row,
                        baseCol: cell.sharedFormulaBase.col
                    } : null;

                    const formulaStr = decodePTG(new Uint8Array(cell.formula), false, cellContext,
                        { ...(formulaContext || {}), extra: cell.formulaExtra });

                    // Store formula in data with = prefix
                    data[cell.row][cell.col] = formulaStr ? '=' + formulaStr : '=';

                    // Store formula in cells metadata (Jspreadsheet Pro format)
                    if (!cellsObj[address]) {
                        cellsObj[address] = {};
                    }
                    cellsObj[address].formula = formulaStr;
                } catch (err) {
                    // Undecodable formula: keep the value Excel cached instead of emitting a wrong formula
                    data[cell.row][cell.col] = cell.value === null || cell.value === undefined ? '' : cell.value;
                }
            } else {
                // Formula cell but no tokens
                data[cell.row][cell.col] = `=FORMULA(${JSON.stringify(cell.value)})`;
            }
        } else {
            // Regular value cell
            data[cell.row][cell.col] = cell.value;
        }

        // Convert style to CSS
        if (cell.style) {
            const cssString = styleToCss(cell.style);
            if (cssString) {
                styleObj[address] = cssString;
            }

            // Extract cell-level properties (format, alignment, wrap)
            const cellProps = {};

            // Extract number format (skip GENERAL format as it's the default)
            if (cell.style.format && !cell.style.format.toUpperCase().startsWith('GENERAL')) {
                cellProps.format = cell.style.format;
            }

            // Extract alignment
            if (cell.style.alignment) {
                const hAlign = ['left', 'center', 'right', 'fill', 'justify'][cell.style.alignment.horizontal];
                if (hAlign) {
                    cellProps.align = hAlign;
                }
                if (cell.style.alignment.wrapText) {
                    cellProps.wrap = true;
                }
            }

            // Store cell properties if any (merge with existing properties like formula)
            if (Object.keys(cellProps).length > 0) {
                if (cellsObj[address]) {
                    // Merge with existing properties (e.g., formula)
                    cellsObj[address] = { ...cellsObj[address], ...cellProps };
                } else {
                    cellsObj[address] = cellProps;
                }
            }
        }
    });

    // Convert columns to Jspreadsheet format
    const columnsArray = [];
    for (let c = 0; c <= maxCol; c++) {
        const col = columns[c] || {};
        const column = {
            width: col.width ? convertWidthToPixels(col.width, 'char') : 100,
            title: col.title || getColumnName(c),
            type: 'text'
        };

        // Add hidden state if column is hidden
        if (col.hidden) {
            column.visible = false;
        }

        columnsArray.push(column);
    }

    // Convert rows to Jspreadsheet format
    const rowsObj = {};
    for (let r = 0; r <= maxRow; r++) {
        if (rows[r]) {
            const row = {
                height: rows[r].height || 21
            };

            // Convert hidden property to visible
            if (rows[r].hidden) {
                row.visible = false;
            }

            rowsObj[r] = row;
        }
    }

    // Convert merged cells to Jspreadsheet format
    const mergeCellsObj = {};
    mergedCells.forEach(merge => {
        const address = getCellNameFromCoords(merge.colFirst, merge.rowFirst);
        const colspan = merge.colLast - merge.colFirst + 1;
        const rowspan = merge.rowLast - merge.rowFirst + 1;

        mergeCellsObj[address] = [colspan, rowspan];
    });

    return {
        data,
        columns: columnsArray,
        rows: rowsObj,
        cells: cellsObj,
        style: styleObj,
        mergeCells: mergeCellsObj,
        minDimensions: [maxCol + 1, maxRow + 1]
    };
}

/**
 * Parse XLS from buffer
 */
export function parseXLSBuffer(buffer) {
    const cfb = CFB.read(buffer, { type: 'buffer' });

    // Find the Workbook stream
    const workbookEntry = CFB.find(cfb, 'Workbook') || CFB.find(cfb, 'Book');

    if (!workbookEntry) {
        throw new Error('No Workbook stream found in XLS file');
    }

    return parseWorkbookStream(new Uint8Array(workbookEntry.content));
}

/**
 * Parse a BIFF8 Workbook stream (the content of the CFB "Workbook" entry) -> Jspreadsheet Pro format
 * @param {Uint8Array} stream
 */
export function parseWorkbookStream(stream) {
    // write-protected workbooks are RC4-encrypted with Excel's default password
    const workbookData = decryptWorkbook(stream);
    const records = parseBIFFRecords(workbookData);

    // Parse Shared String Table
    let sst = [];
    const sstIndex = records.findIndex(r => r.type === BIFF_RECORDS.SST);
    if (sstIndex >= 0) {
        sst = parseSST(records, sstIndex);
    }

    // Map every record to its sheet substream (BOUNDSHEET lbPlyPos = stream offset of the sheet's BOF)
    const boundsheetRecords = records.filter(r => r.type === BIFF_RECORDS.BOUNDSHEET);
    assignSheets(records, boundsheetRecords.map(r => readUInt32LE(r.data, 0)));

    // Parse all styling and formatting information
    const fonts = parseFonts(records);
    const formats = parseFormats(records);
    const xfStyles = parseXF(records, formats, fonts);
    const columnsBySheet = parseColumnInfo(records);
    const rowsBySheet = parseRowInfo(records);
    const mergedCellsBySheet = parseMergedCells(records);

    // Find all sheets with their properties
    const sheets = [];
    const boundsheets = records.filter(r => r.type === BIFF_RECORDS.BOUNDSHEET);

    boundsheets.forEach((bs) => {
        // BOUNDSHEET structure:
        // Offset 0-3: BOF position
        // Offset 4: Sheet state (0=visible, 1=hidden, 2=very hidden)
        // Offset 5: Sheet type
        // Offset 6+: Sheet name
        const sheetState = bs.data[4];
        const nameLen = bs.data[6];
        const flags = bs.data[7]; // String flags byte
        const isCompressed = (flags & 0x01) === 0;
        let name = '';
        let pos = 8; // Characters start after length and flags

        if (isCompressed) {
            // 8-bit characters
            for (let i = 0; i < nameLen && pos < bs.data.length; i++, pos++) {
                const byte = bs.data[pos];
                if (byte >= 0x20) {
                    name += String.fromCharCode(byte);
                }
            }
        } else {
            // 16-bit characters
            for (let i = 0; i < nameLen && pos + 1 < bs.data.length; i++, pos += 2) {
                const char = readUInt16LE(bs.data, pos);
                if (char >= 0x20) {
                    name += String.fromCharCode(char);
                }
            }
        }

        sheets.push({
            name,
            visibility: sheetState  // 0=visible, 1=hidden, 2=very hidden
        });
    });

    // Names, external sheets and add-in functions that formulas refer to
    const formulaContext = parseFormulaContext(records, sheets.map(s => s.name));

    // Parse cell data with all information
    const allCells = parseCellRecords(records, sst, xfStyles, workbookData);

    // Group cells by sheetIndex
    const cellsBySheet = {};
    allCells.forEach(cell => {
        if (!cellsBySheet[cell.sheetIndex]) {
            cellsBySheet[cell.sheetIndex] = [];
        }
        cellsBySheet[cell.sheetIndex].push(cell);
    });

    // Global style array to store unique CSS strings
    const globalStyles = [];
    const styleMap = new Map(); // CSS string -> index

    // Create worksheets - one for each sheet
    const worksheets = sheets.map((sheet, index) => {
        const sheetCells = cellsBySheet[index] || [];
        const sheetMergedCells = mergedCellsBySheet[index] || [];
        const sheetColumns = columnsBySheet[index] || [];
        const sheetRows = rowsBySheet[index] || {};
        const sheetData = cellsToJspreadsheet(sheetCells, sheetColumns, sheetRows, sheetMergedCells, formats, formulaContext);

        // Convert worksheet styles from CSS strings to global style indices
        const worksheetStyleIndices = {};
        if (sheetData.style) {
            Object.keys(sheetData.style).forEach(cellRef => {
                const cssString = sheetData.style[cellRef];

                // Get or create style index
                if (!styleMap.has(cssString)) {
                    styleMap.set(cssString, globalStyles.length);
                    globalStyles.push(cssString);
                }

                worksheetStyleIndices[cellRef] = styleMap.get(cssString);
            });
        }

        const worksheet = {
            worksheetName: sheet.name,
            ...sheetData,
            style: worksheetStyleIndices
        };

        // Add worksheet visibility state
        if (sheet.visibility === 1) {
            worksheet.worksheetState = 'hidden';
        } else if (sheet.visibility === 2) {
            worksheet.worksheetState = 'veryHidden';
        }
        // visibility === 0 means visible (default, no property needed)

        return worksheet;
    });

    // Return in Jspreadsheet Pro format
    // Include globalStyles array which contains CSS strings referenced by indices in worksheet.style
    return {
        worksheets,
        style: globalStyles
    };
}

/**
 * Parse XLS file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseXLS(input, options = {}) {
    return parse(async (inp) => {
        const buffer = await loadAsBuffer(inp);
        return parseXLSBuffer(buffer);
    }, input, options);
}

