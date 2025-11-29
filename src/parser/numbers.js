import JSZip from 'jszip';
import { loadAsBuffer, parse } from '../utils/loader.js';

/**
 * Parse Apple Numbers file (.numbers format)
 *
 * Numbers files are ZIP archives containing:
 * - Index.zip: Main data in IWA (iWork Archive) format with Protobuf messages
 * - preview.jpg: Preview image
 * - Metadata/: Various metadata files
 *
 * The IWA format uses Protocol Buffers to store structured data.
 * Each .iwa file contains serialized Protobuf messages.
 */

/**
 * Read varint (variable-length integer) from buffer
 * Used in Protobuf encoding
 */
function readVarint(buffer, offset) {
    let result = 0;
    let shift = 0;
    let byte;
    let pos = offset;

    do {
        if (pos >= buffer.length) {
            throw new Error('Unexpected end of buffer while reading varint');
        }
        byte = buffer[pos++];
        result |= (byte & 0x7F) << shift;
        shift += 7;
    } while (byte & 0x80);

    return { value: result, length: pos - offset };
}

/**
 * Parse IWA (iWork Archive) file
 * IWA files contain Protobuf messages with specific structure
 */
function parseIWA(buffer) {
    const messages = [];
    let offset = 0;

    try {
        while (offset < buffer.length) {
            // Read message length (varint)
            const lengthData = readVarint(buffer, offset);
            offset += lengthData.length;

            const messageLength = lengthData.value;
            if (messageLength === 0 || offset + messageLength > buffer.length) {
                break;
            }

            // Read message data
            const messageData = buffer.slice(offset, offset + messageLength);
            offset += messageLength;

            // Parse the Protobuf message
            const message = parseProtobufMessage(messageData);
            if (message) {
                messages.push(message);
            }
        }
    } catch (e) {
        // Parsing error - return what we have so far
        console.warn('IWA parsing stopped:', e.message);
    }

    return messages;
}

/**
 * Parse a Protobuf message
 * Simplified parser that extracts field values
 */
function parseProtobufMessage(buffer) {
    const fields = {};
    let offset = 0;

    try {
        while (offset < buffer.length) {
            // Read field tag (varint)
            const tagData = readVarint(buffer, offset);
            offset += tagData.length;

            const tag = tagData.value;
            const fieldNumber = tag >>> 3;
            const wireType = tag & 0x07;

            // Parse based on wire type
            switch (wireType) {
                case 0: // Varint
                    const varintData = readVarint(buffer, offset);
                    offset += varintData.length;
                    fields[fieldNumber] = varintData.value;
                    break;

                case 1: // 64-bit
                    if (offset + 8 <= buffer.length) {
                        const value = buffer.readDoubleLE(offset);
                        fields[fieldNumber] = value;
                        offset += 8;
                    } else {
                        return fields;
                    }
                    break;

                case 2: // Length-delimited (string, bytes, embedded message)
                    const lengthData = readVarint(buffer, offset);
                    offset += lengthData.length;
                    const length = lengthData.value;

                    if (offset + length <= buffer.length) {
                        const data = buffer.slice(offset, offset + length);
                        offset += length;

                        // Try to parse as UTF-8 string
                        try {
                            const str = data.toString('utf-8');
                            // Only use if it looks like valid text (no control chars except whitespace)
                            if (/^[\x20-\x7E\s\u0080-\uFFFF]*$/.test(str)) {
                                fields[fieldNumber] = str;
                            } else {
                                // Could be embedded message or binary data
                                fields[fieldNumber] = data;
                            }
                        } catch (e) {
                            fields[fieldNumber] = data;
                        }
                    } else {
                        return fields;
                    }
                    break;

                case 5: // 32-bit
                    if (offset + 4 <= buffer.length) {
                        const value = buffer.readFloatLE(offset);
                        fields[fieldNumber] = value;
                        offset += 4;
                    } else {
                        return fields;
                    }
                    break;

                default:
                    // Unknown wire type, skip
                    return fields;
            }
        }
    } catch (e) {
        // Parsing error
    }

    return fields;
}

/**
 * Extract cell data from IWA messages
 * This is a simplified extraction - full Numbers format is complex
 */
function extractCellData(messages) {
    const cells = [];
    const strings = [];
    const numbers = [];

    // First pass: collect strings and numbers
    for (const msg of messages) {
        // Look for string values (typically in field 3 or similar)
        if (typeof msg[3] === 'string') {
            strings.push(msg[3]);
        }
        // Look for numeric values
        if (typeof msg[2] === 'number' || typeof msg[4] === 'number') {
            numbers.push(msg[2] || msg[4]);
        }
    }

    return { strings, numbers };
}

/**
 * Build worksheet data from extracted information
 */
function buildWorksheetData(cellData, rowCount = 50, colCount = 26) {
    const data = [];

    // Create empty grid
    for (let r = 0; r < rowCount; r++) {
        const row = [];
        for (let c = 0; c < colCount; c++) {
            row.push('');
        }
        data.push(row);
    }

    // Fill in extracted data
    // This is a simplified approach - actual Numbers format has complex cell addressing
    const { strings, numbers } = cellData;

    let stringIdx = 0;
    let numberIdx = 0;

    // Distribute data across grid
    for (let r = 0; r < Math.min(rowCount, Math.max(strings.length, numbers.length)); r++) {
        if (stringIdx < strings.length) {
            data[r][0] = strings[stringIdx++];
        }
        if (numberIdx < numbers.length) {
            data[r][1] = numbers[numberIdx++];
        }
    }

    return data;
}

/**
 * Parse Numbers file string content
 */
async function parseNumbersFromZip(zip) {
    const result = {
        worksheets: []
    };

    // Extract Index.zip which contains the actual data
    const indexZipFile = zip.file('Index.zip');
    if (!indexZipFile) {
        throw new Error('Invalid .numbers file: Index.zip not found');
    }

    const indexZipBuffer = await indexZipFile.async('nodebuffer');
    const indexZip = await JSZip.loadAsync(indexZipBuffer);

    // Find .iwa files (iWork Archive files containing Protobuf data)
    const iwaFiles = [];
    indexZip.forEach((relativePath, file) => {
        if (relativePath.endsWith('.iwa')) {
            iwaFiles.push({ path: relativePath, file });
        }
    });

    console.log(`Numbers: Found ${iwaFiles.length} IWA files`);

    // Parse each IWA file
    const allMessages = [];
    for (const { path, file } of iwaFiles) {
        try {
            const buffer = await file.async('nodebuffer');
            const messages = parseIWA(buffer);
            allMessages.push(...messages);
            console.log(`Numbers: Parsed ${messages.length} messages from ${path}`);
        } catch (e) {
            console.warn(`Numbers: Failed to parse ${path}:`, e.message);
        }
    }

    console.log(`Numbers: Total messages parsed: ${allMessages.length}`);

    // Extract cell data
    const cellData = extractCellData(allMessages);
    console.log(`Numbers: Extracted ${cellData.strings.length} strings, ${cellData.numbers.length} numbers`);

    // Build worksheet
    const data = buildWorksheetData(cellData);

    // Create columns
    const columns = [];
    for (let i = 0; i < 26; i++) {
        columns.push({
            title: String.fromCharCode(65 + i),
            width: 100
        });
    }

    result.worksheets.push({
        worksheetName: 'Sheet1',
        data,
        columns
    });

    // Add warning about limited support
    result.warnings = [
        'Numbers format support is limited. This is a basic extraction.',
        'Complex features like formulas, formatting, and charts may not be preserved.',
        'For best results, export to XLSX or CSV from Numbers app.'
    ];

    return result;
}

/**
 * Parse Apple Numbers file (.numbers)
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseNumbers(input, options = {}) {
    return parse(async (inp) => {
        // Load as buffer
        const buffer = await loadAsBuffer(inp);

        // Parse as ZIP
        const zip = await JSZip.loadAsync(buffer);

        return await parseNumbersFromZip(zip);
    }, input, options);
}

/**
 * Check if a file is a Numbers file by checking for Index.zip
 */
export async function isNumbersFile(input) {
    try {
        const buffer = await loadAsBuffer(input);
        const zip = await JSZip.loadAsync(buffer);
        return zip.file('Index.zip') !== null;
    } catch (e) {
        return false;
    }
}
