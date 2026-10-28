import { loadAsBuffer, parse } from '../utils/loader.js';
import { readUInt8, readUInt16LE, readUInt32LE, readFloat64LE } from '../utils/helpers.js';

/**
 * Parse DBF (dBase) file format
 * Supports dBase III, IV, V, and FoxPro variants
 *
 * DBF is a database format, so it contains:
 * - Column definitions (name, type, width, decimals)
 * - Row data
 * - Metadata (version, record count, dates)
 *
 * DBF does NOT contain:
 * - Formulas (database stores calculated values)
 * - Fonts/styling (no visual formatting in binary format)
 * - Merged cells (database concept)
 */

// Language driver ID (header byte 29) -> codepage of the text fields, per the dBase / FoxPro tables
const LANGUAGE_DRIVERS = {
    0x01: 'cp437', 0x02: 'cp850', 0x03: 'cp1252', 0x04: 'macintosh', 0x08: 'cp865', 0x09: 'cp437',
    0x0A: 'cp850', 0x0B: 'cp437', 0x0D: 'cp437', 0x0E: 'cp850', 0x0F: 'cp437', 0x10: 'cp850',
    0x11: 'cp437', 0x12: 'cp850', 0x13: 'cp932', 0x14: 'cp850', 0x15: 'cp437', 0x16: 'cp850',
    0x17: 'cp865', 0x18: 'cp437', 0x19: 'cp437', 0x1A: 'cp850', 0x1B: 'cp437', 0x1C: 'cp863',
    0x1D: 'cp850', 0x1F: 'cp852', 0x22: 'cp852', 0x23: 'cp852', 0x24: 'cp860', 0x25: 'cp850',
    0x26: 'cp866', 0x37: 'cp850', 0x40: 'cp852', 0x4D: 'cp936', 0x4E: 'cp949', 0x4F: 'cp950',
    0x50: 'cp874', 0x57: 'cp1252', 0x58: 'cp1252', 0x59: 'cp1252', 0x64: 'cp852', 0x65: 'cp866',
    0x66: 'cp865', 0x67: 'cp861', 0x6A: 'cp737', 0x6B: 'cp857', 0x78: 'cp950', 0x79: 'cp949',
    0x7A: 'cp936', 0x7B: 'cp932', 0x7C: 'cp874', 0x86: 'cp737', 0x87: 'cp852', 0x88: 'cp857',
    0xC8: 'cp1250', 0xC9: 'cp1251', 0xCA: 'cp1254', 0xCB: 'cp1253', 0xCC: 'cp1257'
};

// Codepages that TextDecoder knows under a WHATWG label; the DOS ones (cp437, cp850...) need iconv-lite
const WHATWG_LABELS = {
    cp1250: 'windows-1250', cp1251: 'windows-1251', cp1252: 'windows-1252', cp1253: 'windows-1253',
    cp1254: 'windows-1254', cp1257: 'windows-1257', cp866: 'ibm866', cp874: 'windows-874',
    cp932: 'shift_jis', cp936: 'gbk', cp949: 'euc-kr', cp950: 'big5'
};

/**
 * Text decoder for an encoding name, or null when neither TextDecoder nor iconv-lite (Node.js) knows it
 */
async function getDecoder(encoding) {
    try {
        const decoder = new TextDecoder(WHATWG_LABELS[encoding] || encoding);
        return (bytes) => decoder.decode(bytes);
    } catch (e) {
        // Not a WHATWG label: try iconv-lite below
    }
    try {
        const module = await import('iconv-lite');
        const iconv = module.default || module;
        if (iconv.encodingExists(encoding)) {
            return (bytes) => iconv.decode(Buffer.from(bytes), encoding);
        }
    } catch (e) {
        // iconv-lite is optional (and absent from browser bundles)
    }
    return null;
}

// DBF version constants
const DBF_VERSION = {
    0x02: 'FoxBASE',
    0x03: 'dBASE III',
    0x04: 'dBASE IV',
    0x05: 'dBASE V',
    0x30: 'Visual FoxPro',
    0x31: 'Visual FoxPro with AutoIncrement',
    0x43: 'dBASE IV SQL table',
    0x63: 'dBASE IV SQL system',
    0x83: 'dBASE III+ with memo',
    0x8B: 'dBASE IV with memo',
    0x8E: 'dBASE IV SQL table with memo',
    0xF5: 'FoxPro 2.x with memo',
    0xFB: 'FoxPro without memo'
};

// Field type constants
const FIELD_TYPES = {
    'C': 'Character',
    'N': 'Numeric',
    'F': 'Float',
    'L': 'Logical',
    'D': 'Date',
    'M': 'Memo',
    'I': 'Integer',
    'B': 'Double',
    'G': 'General',
    'P': 'Picture',
    'T': 'DateTime',
    'Y': 'Currency',
    '@': 'Timestamp',
    'O': 'Double',
    '+': 'Autoincrement'
};

// Map DBF types to Jspreadsheet types
const DBF_TO_JS_TYPE = {
    'C': 'text',        // Character -> text
    'N': 'numeric',     // Numeric -> numeric
    'F': 'numeric',     // Float -> numeric
    'L': 'checkbox',    // Logical -> checkbox
    'D': 'calendar',    // Date -> calendar
    'M': 'text',        // Memo -> text
    'I': 'numeric',     // Integer -> numeric
    'B': 'numeric',     // Double -> numeric
    'T': 'calendar',    // DateTime -> calendar
    'Y': 'numeric',     // Currency -> numeric
    '@': 'calendar',    // Timestamp -> calendar
    'O': 'numeric',     // Double -> numeric
    '+': 'numeric'      // Autoincrement -> numeric
};

/**
 * Parse DBF header (32 bytes)
 */
function parseHeader(buffer) {
    const version = readUInt8(buffer, 0);

    // Last update date (YY MM DD)
    const year = 1900 + readUInt8(buffer, 1);
    const month = readUInt8(buffer, 2);
    const day = readUInt8(buffer, 3);
    const lastUpdate = new Date(year, month - 1, day);

    const recordCount = readUInt32LE(buffer, 4);
    const headerLength = readUInt16LE(buffer, 8);
    const recordLength = readUInt16LE(buffer, 10);

    // Reserved bytes
    const reserved = buffer.slice(12, 14);

    // Flags
    const incompleteTransaction = readUInt8(buffer, 14);
    const encryptionFlag = readUInt8(buffer, 15);

    // Multi-user processing (reserved)
    const multiUser = buffer.slice(16, 28);

    // MDX flag
    const mdxFlag = readUInt8(buffer, 28);

    // Language driver ID
    const languageDriver = readUInt8(buffer, 29);

    // Reserved
    const reserved2 = buffer.slice(30, 32);

    return {
        version,
        versionName: DBF_VERSION[version] || `Unknown (0x${version.toString(16)})`,
        lastUpdate,
        recordCount,
        headerLength,
        recordLength,
        incompleteTransaction: incompleteTransaction !== 0,
        encrypted: encryptionFlag !== 0,
        hasMDX: mdxFlag !== 0,
        languageDriver
    };
}

/**
 * Parse field descriptor (32 bytes)
 */
function parseFieldDescriptor(buffer, offset) {
    // Field name (11 bytes, null-terminated)
    let nameBytes = [];
    for (let i = 0; i < 11; i++) {
        const byte = buffer[offset + i];
        if (byte === 0) break;
        nameBytes.push(byte);
    }
    const name = String.fromCharCode(...nameBytes);

    // Field type (1 byte)
    const type = String.fromCharCode(buffer[offset + 11]);

    // Field data address (4 bytes) - in memory, not used in file
    const dataAddress = readUInt32LE(buffer, offset + 12);

    // Field length (1 byte)
    const length = readUInt8(buffer, offset + 16);

    // Decimal count (1 byte)
    const decimalCount = readUInt8(buffer, offset + 17);

    // Reserved (2 bytes)
    const reserved1 = buffer.slice(offset + 18, offset + 20);

    // Work area ID (1 byte)
    const workAreaId = readUInt8(buffer, offset + 20);

    // Reserved (10 bytes)
    const reserved2 = buffer.slice(offset + 21, offset + 31);

    // MDX flag (1 byte)
    const mdxFlag = readUInt8(buffer, offset + 31);

    return {
        name,
        type,
        typeName: FIELD_TYPES[type] || 'Unknown',
        length,
        decimalCount,
        workAreaId,
        hasMDX: mdxFlag !== 0
    };
}

/**
 * Parse a record value based on field type
 */
function parseFieldValue(buffer, offset, field, decode) {
    const bytes = buffer.slice(offset, offset + field.length);
    // Character fields carry text in the file's codepage; the other types are ASCII digits or binary
    const rawValue = (field.type === 'C' && decode ? decode(bytes) : String.fromCharCode(...bytes)).trim();

    if (!rawValue || rawValue === '') {
        return null;
    }

    switch (field.type) {
        case 'C': // Character
        case 'M': // Memo (returns memo block number, actual content in .dbt/.fpt file)
            return rawValue;

        case 'N': // Numeric
        case 'F': // Float
            const num = parseFloat(rawValue);
            return isNaN(num) ? null : num;

        case 'I': // Integer (4 bytes, little-endian)
            if (field.length === 4) {
                return readUInt32LE(bytes, 0);
            }
            const intNum = parseInt(rawValue, 10);
            return isNaN(intNum) ? null : intNum;

        case 'L': // Logical
            const logical = rawValue.toUpperCase();
            if (logical === 'T' || logical === 'Y') return true;
            if (logical === 'F' || logical === 'N') return false;
            return null;

        case 'D': // Date (YYYYMMDD)
            if (rawValue.length === 8) {
                const year = parseInt(rawValue.substring(0, 4), 10);
                const month = parseInt(rawValue.substring(4, 6), 10);
                const day = parseInt(rawValue.substring(6, 8), 10);
                if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
                    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                }
            }
            return null;

        case 'T': // DateTime (8 bytes: 4 for date, 4 for time)
            if (field.length === 8) {
                // Julian date (days since 4713 BC) + milliseconds since midnight
                const julianDate = readUInt32LE(bytes, 0);
                const milliseconds = readUInt32LE(bytes, 4);

                if (julianDate > 0) {
                    // Convert Julian to Gregorian (simplified)
                    const unixEpoch = 2440588; // Julian day for 1970-01-01
                    const days = julianDate - unixEpoch;
                    const date = new Date(days * 86400000 + milliseconds);
                    return date.toISOString();
                }
            }
            return null;

        case 'Y': // Currency (8 bytes, little-endian integer, 4 decimal places)
            if (field.length === 8) {
                // Read as 64-bit integer (approximate with two 32-bit)
                const low = readUInt32LE(bytes, 0);
                const high = readUInt32LE(bytes, 4);
                const value = high * 4294967296 + low;
                return value / 10000; // 4 decimal places
            }
            return parseFloat(rawValue) || null;

        case 'B': // Double (8 bytes)
        case 'O': // Double
            if (field.length === 8) {
                return readFloat64LE(bytes, 0);
            }
            return parseFloat(rawValue) || null;

        case '@': // Timestamp
            return rawValue; // Return as string for now

        case '+': // Autoincrement (4 bytes)
            if (field.length === 4) {
                return readUInt32LE(bytes, 0);
            }
            return parseInt(rawValue, 10) || null;

        case 'G': // General (OLE)
        case 'P': // Picture (binary)
            return `<binary:${rawValue.length}bytes>`;

        default:
            return rawValue;
    }
}

/**
 * Parse DBF from buffer
 * @param {Uint8Array} buffer - DBF file buffer
 * @returns {object} Jspreadsheet Pro format
 */
export function parseDBFBuffer(buffer, decode = null) {
    // Ensure buffer is Uint8Array
    if (!(buffer instanceof Uint8Array)) {
        buffer = new Uint8Array(buffer);
    }

    // Parse header
    const header = parseHeader(buffer);

    // Parse field descriptors
    const fields = [];
    let offset = 32; // Start after header

    while (offset < header.headerLength - 1) {
        const byte = buffer[offset];

        // 0x0D marks end of field descriptors
        if (byte === 0x0D) {
            break;
        }

        const field = parseFieldDescriptor(buffer, offset);
        fields.push(field);
        offset += 32;
    }

    // Skip to data start (should be at headerLength)
    offset = header.headerLength;

    // Parse records
    const data = [];
    const deletedRecords = [];

    // The record count comes from an untrusted header: never read more records than the file holds
    const available = header.recordLength > 0
        ? Math.max(0, Math.floor((buffer.length - header.headerLength) / header.recordLength))
        : 0;
    const recordCount = Math.min(header.recordCount, available);

    for (let i = 0; i < recordCount; i++) {
        // First byte of record: 0x20 = valid, 0x2A = deleted
        const deletionFlag = buffer[offset];
        offset++;

        const record = [];

        for (const field of fields) {
            const value = parseFieldValue(buffer, offset, field, decode);
            record.push(value);
            offset += field.length;
        }

        if (deletionFlag === 0x2A) {
            deletedRecords.push(i);
        } else {
            data.push(record);
        }
    }

    // Create column definitions with enhanced properties
    const columns = fields.map((field, index) => {
        const column = {
            title: field.name,
            name: field.name,
            width: Math.max(field.length * 10, 100), // Convert to pixels (approximate)
        };

        // Add Jspreadsheet type
        const jsType = DBF_TO_JS_TYPE[field.type];
        if (jsType) {
            column.type = jsType;
        }

        // Add decimal places for numeric fields
        if ((field.type === 'N' || field.type === 'F') && field.decimalCount > 0) {
            column.decimal = '.'.repeat(field.decimalCount);
            column.mask = '#,##0.' + '0'.repeat(field.decimalCount);
        }

        // Add currency mask
        if (field.type === 'Y') {
            column.mask = '$#,##0.00';
        }

        // Add date/datetime format
        if (field.type === 'D') {
            column.options = { format: 'YYYY-MM-DD' };
        } else if (field.type === 'T' || field.type === '@') {
            column.options = { format: 'YYYY-MM-DD HH:mm:ss' };
        }

        return column;
    });

    // Build comprehensive metadata
    const meta = {
        // DBF header information
        dbfVersion: header.versionName,
        dbfVersionCode: `0x${header.version.toString(16)}`,
        lastModified: header.lastUpdate.toISOString(),
        totalRecords: header.recordCount,
        activeRecords: data.length,
        deletedRecords: deletedRecords.length,
        recordLength: header.recordLength,
        headerLength: header.headerLength,

        // Flags
        encrypted: header.encrypted,
        hasMDX: header.hasMDX,
        incompleteTransaction: header.incompleteTransaction,
        languageDriver: header.languageDriver,

        // Field information
        fields: fields.map(field => ({
            name: field.name,
            type: field.type,
            typeName: field.typeName,
            length: field.length,
            decimals: field.decimalCount,
            jsType: DBF_TO_JS_TYPE[field.type] || 'text'
        })),

        // Deleted record indices (if any)
        deletedRecordIndices: deletedRecords.length > 0 ? deletedRecords : undefined
    };

    // Return in Jspreadsheet Pro format
    return {
        worksheets: [{
            worksheetName: 'DBF Data',
            data,
            columns,
            meta
        }]
    };
}

/**
 * Parse DBF file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {string} options.encoding - Encoding of the text fields (default: the codepage declared by the file's
 *     language driver byte, else latin1)
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseDBF(input, options = {}) {
    return parse(async (inp) => {
        const buffer = await loadAsBuffer(inp);
        const encoding = options.encoding || (buffer.length > 29 ? LANGUAGE_DRIVERS[buffer[29]] : undefined);
        const decode = encoding ? await getDecoder(encoding) : null;
        if (options.encoding && !decode) {
            throw new Error(`Unknown encoding '${options.encoding}'`);
        }
        return parseDBFBuffer(buffer, decode);
    }, input, options);
}
