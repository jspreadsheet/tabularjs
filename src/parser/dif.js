import { loadAsString, parse, detectEncoding } from '../utils/loader.js';
import { getColumnName } from '../utils/helpers.js';

/**
 * Parse DIF from string content
 */
export function parseDIFString(fileContent) {
    const lines = fileContent.split(/\r?\n/);

    let i = 0;
    let inDataSection = false;
    let inLabelSection = false;
    const data = [];
    const columnLabels = [];
    const rowLabels = [];
    const comments = [];
    let currentRow = null;
    let currentLabelRow = null;
    let numVectors = 0; // Number of columns
    let numTuples = 0;  // Number of rows
    let tableVersion = '';

    // Parse the file line by line
    while (i < lines.length) {
        const line = lines[i].trim();

        // Skip empty lines
        if (!line) {
            i++;
            continue;
        }

        // Check for section headers
        if (line === 'TABLE') {
            i++;
            // Parse TABLE metadata (version, etc.)
            if (i < lines.length) {
                const typeLine = lines[i].trim();
                i++;
                if (i < lines.length) {
                    tableVersion = lines[i].trim();
                    i++;
                }
            }
            // Skip rest of TABLE metadata until next section
            while (i < lines.length && !['VECTORS', 'TUPLES', 'DATA', 'LABEL', 'COMMENT'].includes(lines[i].trim())) {
                i++;
            }
            continue;
        }

        if (line === 'VECTORS') {
            i++;
            // Next line should be 0,numVectors
            if (i < lines.length) {
                const vectorLine = lines[i].trim();
                const match = vectorLine.match(/^0,(\d+)/);
                if (match) {
                    numVectors = parseInt(match[1]);
                }
                i += 2; // Skip value and empty line
            }
            continue;
        }

        if (line === 'TUPLES') {
            i++;
            // Next line should be 0,numTuples
            if (i < lines.length) {
                const tupleLine = lines[i].trim();
                const match = tupleLine.match(/^0,(\d+)/);
                if (match) {
                    numTuples = parseInt(match[1]);
                }
                i += 2; // Skip value and empty line
            }
            continue;
        }

        if (line === 'LABEL') {
            inLabelSection = true;
            inDataSection = false;
            currentLabelRow = null;
            i++;
            continue;
        }

        if (line === 'COMMENT') {
            i++;
            // Read comment lines until next section
            while (i < lines.length) {
                const commentLine = lines[i].trim();
                if (['TABLE', 'VECTORS', 'TUPLES', 'DATA', 'LABEL'].includes(commentLine)) {
                    break;
                }
                if (commentLine && !commentLine.match(/^[0-2],-?\d+$/)) {
                    comments.push(commentLine);
                }
                i++;
            }
            continue;
        }

        if (line === 'DATA') {
            inDataSection = true;
            inLabelSection = false;
            i++;
            continue;
        }

        // Process LABEL section (for column/row headers)
        if (inLabelSection) {
            const parts = line.split(',');
            if (parts.length < 2) {
                i++;
                continue;
            }

            const type = parseInt(parts[0]);
            const value = parts.slice(1).join(',');

            // -1,0 indicates BOT (Beginning of Tuple)
            if (type === -1) {
                if (currentLabelRow !== null) {
                    rowLabels.push(currentLabelRow);
                }
                currentLabelRow = [];
                i += 2;
                continue;
            }

            // String label
            if (type === 1) {
                i++;
                let labelValue = '';
                if (i < lines.length) {
                    labelValue = lines[i];
                    if (labelValue.startsWith('"') && labelValue.endsWith('"')) {
                        labelValue = labelValue.slice(1, -1);
                    }
                }
                if (currentLabelRow !== null) {
                    currentLabelRow.push(labelValue);
                } else {
                    columnLabels.push(labelValue);
                }
                i++;
                continue;
            }

            i += 2;
        }

        // Process DATA section
        if (inDataSection) {
            // Parse type,value line
            const parts = line.split(',');
            if (parts.length < 2) {
                i++;
                continue;
            }

            const type = parseInt(parts[0]);
            const value = parts.slice(1).join(','); // Join back in case value contains commas

            // -1,0 indicates BOT (Beginning of Tuple/Row) or EOD (End of Data)
            if (type === -1) {
                if (value === '0') {
                    // Check if this is EOD by looking at the next line
                    if (i + 1 < lines.length) {
                        const nextLine = lines[i + 1].trim();
                        if (nextLine === 'EOD' || nextLine === '') {
                            // End of data
                            if (currentRow !== null && currentRow.length > 0) {
                                data.push(currentRow);
                            }
                            break;
                        }
                    }

                    // BOT - start new row
                    if (currentRow !== null) {
                        data.push(currentRow);
                    }
                    currentRow = [];
                }
                i += 2; // Skip value line
                continue;
            }

            // 0,value indicates a numeric value
            if (type === 0) {
                const numValue = parseFloat(value);
                if (currentRow !== null) {
                    currentRow.push(isNaN(numValue) ? '' : numValue);
                }
                i += 2; // Skip next line (usually empty for numbers)
                continue;
            }

            // 1,0 indicates a string value (actual value on next line)
            if (type === 1) {
                i++; // Move to next line for string content
                let stringValue = '';
                if (i < lines.length) {
                    stringValue = lines[i];
                    // Remove quotes if present
                    if (stringValue.startsWith('"') && stringValue.endsWith('"')) {
                        stringValue = stringValue.slice(1, -1);
                    }
                }
                if (currentRow !== null) {
                    currentRow.push(stringValue);
                }
                i++;
                continue;
            }

            // 2,value indicates special values (TRUE=1, FALSE=0, ERROR=-1, NA=-2)
            if (type === 2) {
                const specialValue = parseInt(value);
                let cellValue;
                switch (specialValue) {
                    case 1:
                        cellValue = true; // TRUE
                        break;
                    case 0:
                        cellValue = false; // FALSE
                        break;
                    case -1:
                        cellValue = '#ERROR'; // ERROR
                        break;
                    case -2:
                        cellValue = '#N/A'; // NA
                        break;
                    default:
                        cellValue = '';
                }
                if (currentRow !== null) {
                    currentRow.push(cellValue);
                }
                i += 2; // Skip next line
                continue;
            }

            // Unknown type, skip
            i += 2;
        } else {
            i++;
        }
    }

    // Add the last row if it exists
    if (currentRow !== null && currentRow.length > 0) {
        data.push(currentRow);
    }

    // Determine the number of columns (use the max from rows or VECTORS header)
    let maxColumns = numVectors > 0 ? numVectors : 0;
    data.forEach(row => {
        if (row.length > maxColumns) {
            maxColumns = row.length;
        }
    });

    // Ensure all rows have the same number of columns
    data.forEach(row => {
        while (row.length < maxColumns) {
            row.push('');
        }
    });

    // Create column definitions with labels if available
    const columns = [];
    for (let i = 0; i < maxColumns; i++) {
        const columnLabel = columnLabels[i] || getColumnName(i);
        columns.push({
            title: columnLabel,
            width: 100
        });
    }

    // Build result
    const result = {
        worksheets: [
            {
                data: data,
                columns: columns
            }
        ]
    };

    // Add metadata if available (using Jspreadsheet's meta property)
    if (comments.length > 0 || tableVersion) {
        result.worksheets[0].meta = {};
        if (tableVersion) {
            result.worksheets[0].meta.version = tableVersion;
        }
        if (comments.length > 0) {
            result.worksheets[0].meta.comments = comments;
        }
    }

    return result;
}

/**
 * Parse DIF file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {string} options.encoding - Character encoding (default: auto-detect with fallbacks). Common values: 'utf-8', 'latin1', 'windows-1252'. If not specified, tries multiple encodings.
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseDIF(input, options = {}) {
    return parse(async (inp) => {
        let encoding = options.encoding;
        let content;

        if (!encoding) {
            // Try encodings in order of likelihood for DIF files
            const encodingsToTry = [
                'cp850',       // DOS/OEM encoding (common for older Excel exports)
                'cp437',       // DOS encoding (alternative)
                'latin1',      // Windows-1252/ISO-8859-1 (most common for older files)
                'utf-8',       // Modern files
                'utf-16le',    // Windows UTF-16
            ];

            // First, try auto-detection
            try {
                const detected = await detectEncoding(inp);
                // Put detected encoding first
                const idx = encodingsToTry.indexOf(detected);
                if (idx > 0) {
                    encodingsToTry.splice(idx, 1);
                    encodingsToTry.unshift(detected);
                }
            } catch (e) {
                // Detection failed, use default order
            }

            // Try each encoding and score based on content quality
            let bestEncoding = null;
            let bestContent = null;
            let bestScore = -1;

            for (const enc of encodingsToTry) {
                try {
                    const testContent = await loadAsString(inp, enc);

                    // Score this encoding
                    let score = 0;

                    // Penalize replacement characters heavily
                    const replacementCount = (testContent.match(/\uFFFD/g) || []).length;
                    score -= replacementCount * 1000;

                    // Reward common spreadsheet characters
                    if (testContent.includes('°')) score += 100; // Degree symbol
                    if (testContent.includes('µ')) score += 50;  // Micro symbol
                    if (testContent.includes('±')) score += 50;  // Plus-minus

                    // Penalize unusual control characters
                    const controlCharCount = (testContent.match(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g) || []).length;
                    score -= controlCharCount * 10;

                    console.log(`DIF: Tried encoding ${enc}, score: ${score}`);

                    if (score > bestScore) {
                        bestScore = score;
                        bestEncoding = enc;
                        bestContent = testContent;
                    }

                    // If we have a perfect score (no replacements, has special chars), use it immediately
                    if (replacementCount === 0 && score > 50) {
                        encoding = enc;
                        content = testContent;
                        console.log(`DIF: Successfully parsed with encoding: ${encoding} (score: ${score})`);
                        break;
                    }
                } catch (e) {
                    console.log(`DIF: Failed with encoding ${enc}:`, e.message);
                    continue;
                }
            }

            // Use best encoding found
            if (!content && bestContent) {
                encoding = bestEncoding;
                content = bestContent;
                console.log(`DIF: Using best encoding: ${encoding} (score: ${bestScore})`);
            }

            // If all failed, use first attempt (latin1)
            if (!content) {
                encoding = 'latin1';
                content = await loadAsString(inp, encoding);
                console.log(`DIF: Using fallback encoding: ${encoding}`);
            }
        } else {
            // Use specified encoding
            content = await loadAsString(inp, encoding);
            console.log(`DIF: Using specified encoding: ${encoding}`);
        }

        return parseDIFString(content);
    }, input, options);
}
