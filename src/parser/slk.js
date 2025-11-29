import { loadAsString, parse } from '../utils/loader.js';
import { convertWidthToPixels, convertR1C1toA1, getColumnName, getCellNameFromCoords } from '../utils/helpers.js';

/**
 * Excel ColorIndex palette (1-56)
 * Maps color indices to hex color values
 */
const EXCEL_COLOR_PALETTE = {
    1: '#000000', 2: '#FFFFFF', 3: '#FF0000', 4: '#00FF00', 5: '#0000FF', 6: '#FFFF00',
    7: '#FF00FF', 8: '#00FFFF', 9: '#800000', 10: '#008000', 11: '#000080', 12: '#808000',
    13: '#800080', 14: '#008080', 15: '#C0C0C0', 16: '#808080', 17: '#9999FF', 18: '#993366',
    19: '#FFFFCC', 20: '#CCFFFF', 21: '#660066', 22: '#FF8080', 23: '#0066CC', 24: '#CCCCFF',
    25: '#000080', 26: '#FF00FF', 27: '#FFFF00', 28: '#00FFFF', 29: '#800080', 30: '#800000',
    31: '#008080', 32: '#0000FF', 33: '#00CCFF', 34: '#CCFFFF', 35: '#CCFFCC', 36: '#FFFF99',
    37: '#99CCFF', 38: '#FF99CC', 39: '#CC99FF', 40: '#FFCC99', 41: '#3366FF', 42: '#33CCCC',
    43: '#99CC00', 44: '#FFCC00', 45: '#FF9900', 46: '#FF6600', 47: '#666699', 48: '#969696',
    49: '#003366', 50: '#339966', 51: '#003300', 52: '#333300', 53: '#993300', 54: '#993366',
    55: '#333399', 56: '#333333'
};

/**
 * Convert format properties to CSS string
 */
function formatPropsToCSS(formatProps) {
    const cssProps = [];

    // Font properties
    if (formatProps.fontFamily) cssProps.push(`font-family: ${formatProps.fontFamily}`);
    if (formatProps.fontSize) cssProps.push(`font-size: ${formatProps.fontSize}`);
    if (formatProps.fontWeight) cssProps.push(`font-weight: ${formatProps.fontWeight}`);
    if (formatProps.fontStyle) cssProps.push(`font-style: ${formatProps.fontStyle}`);
    if (formatProps.textDecoration) cssProps.push(`text-decoration: ${formatProps.textDecoration}`);
    if (formatProps.textAlign) cssProps.push(`text-align: ${formatProps.textAlign}`);

    // Color (convert color index to actual color)
    if (formatProps.colorIndex) {
        const color = EXCEL_COLOR_PALETTE[parseInt(formatProps.colorIndex)];
        if (color) cssProps.push(`color: ${color}`);
    }

    // Background color (from G parameter + color index)
    if (formatProps.gValue && formatProps.colorIndex) {
        const bgColor = EXCEL_COLOR_PALETTE[parseInt(formatProps.colorIndex)];
        if (bgColor) cssProps.push(`background-color: ${bgColor}`);
    }

    // Borders - individual properties
    if (formatProps.borderLeft) cssProps.push(`border-left: ${formatProps.borderLeft}`);
    if (formatProps.borderRight) cssProps.push(`border-right: ${formatProps.borderRight}`);
    if (formatProps.borderTop) cssProps.push(`border-top: ${formatProps.borderTop}`);
    if (formatProps.borderBottom) cssProps.push(`border-bottom: ${formatProps.borderBottom}`);

    return cssProps.length > 0 ? cssProps.join('; ') : null;
}

/**
 * Parse S (Style) parameter for borders
 * S parameter format: SDLRTBM<size>
 * D = has borders, L = left, R = right, T = top, B = bottom
 * M<num> = border size/thickness
 */
function parseBorderStyle(sValue) {
    const borders = {};

    if (!sValue || sValue.length < 2) return borders;

    // Check for border indicators
    const hasLeft = sValue.includes('L');
    const hasRight = sValue.includes('R');
    const hasTop = sValue.includes('T');
    const hasBottom = sValue.includes('B');

    // Extract size if present (M followed by digits)
    const sizeMatch = sValue.match(/M(\d+)/);
    const borderSize = sizeMatch ? parseInt(sizeMatch[1]) : 1;

    // Map size to CSS border width - use 1px for standard borders
    const borderWidth = '1px';
    const borderStyle = `${borderWidth} solid #000000`;

    if (hasLeft) borders.borderLeft = borderStyle;
    if (hasRight) borders.borderRight = borderStyle;
    if (hasTop) borders.borderTop = borderStyle;
    if (hasBottom) borders.borderBottom = borderStyle;

    return borders;
}

/**
 * Parse G (shading/background) parameter
 * G parameter seems to indicate background fill pattern
 * The exact encoding is not well documented
 */
function parseBackgroundStyle(gValue, colorIndex) {
    const style = {};

    // G parameter might indicate pattern type
    // For now, if G is present and we have a color, apply it as background
    if (gValue && colorIndex) {
        const color = EXCEL_COLOR_PALETTE[parseInt(colorIndex)];
        if (color) {
            style.backgroundColor = color;
        }
    }

    return style;
}

/**
 * Decode SLK escape sequences
 */
function decodeSLKString(str) {
    if (!str) return str;

    // Handle \x1B escape sequences (ESC + character code)
    // Common: \x1BN0C = degree symbol (°C), \x1BN0F = °F, \x1BN5L = µL
    return str.replace(/\x1BN([0-9A-F]{2})/g, (match, code) => {
        const charCode = parseInt(code, 16);
        // Map common character codes
        const charMap = {
            0x0C: '°', // Degree symbol
            0x5: 'µ'   // Micro symbol
        };
        return charMap[charCode] || String.fromCharCode(charCode);
    });
}

/**
 * Parse SYLK/SLK parameter
 */
function parseParam(param) {
    const code = param.charAt(0);
    const value = param.substring(1);
    return { code, value };
}

/**
 * Deduplicate borders between adjacent cells
 * If cell A has border-right and cell B (to the right) has border-left, remove one
 * If cell A has border-bottom and cell B (below) has border-top, remove one
 */
function deduplicateBorders(cellStyles, maxRow, maxCol) {
    for (let row = 0; row < maxRow; row++) {
        for (let col = 0; col < maxCol; col++) {
            const cellAddr = getCellNameFromCoords(col, row);
            const cellStyle = cellStyles[cellAddr];

            if (!cellStyle) continue;

            // Check right neighbor
            if (col < maxCol - 1) {
                const rightAddr = getCellNameFromCoords(col + 1, row);
                const rightStyle = cellStyles[rightAddr];

                if (cellStyle.includes('border-right:') && rightStyle && rightStyle.includes('border-left:')) {
                    // Remove border-left from right cell
                    cellStyles[rightAddr] = rightStyle.replace(/border-left:\s*[^;]+;\s*/g, '').trim();
                    if (cellStyles[rightAddr].endsWith(';')) {
                        cellStyles[rightAddr] = cellStyles[rightAddr].slice(0, -1).trim();
                    }
                    if (!cellStyles[rightAddr]) delete cellStyles[rightAddr];
                }
            }

            // Check bottom neighbor
            if (row < maxRow - 1) {
                const bottomAddr = getCellNameFromCoords(col, row + 1);
                const bottomStyle = cellStyles[bottomAddr];

                if (cellStyle.includes('border-bottom:') && bottomStyle && bottomStyle.includes('border-top:')) {
                    // Remove border-top from bottom cell
                    cellStyles[bottomAddr] = bottomStyle.replace(/border-top:\s*[^;]+;\s*/g, '').trim();
                    if (cellStyles[bottomAddr].endsWith(';')) {
                        cellStyles[bottomAddr] = cellStyles[bottomAddr].slice(0, -1).trim();
                    }
                    if (!cellStyles[bottomAddr]) delete cellStyles[bottomAddr];
                }
            }
        }
    }
}

/**
 * Parse SLK from string content
 */
export function parseSLKString(fileContent) {
    const lines = fileContent.split(/\r?\n/);

    const cells = {};
    const cellStyles = {}; // Cell address -> CSS style
    const cellMeta = {}; // Cell address -> metadata (formula, etc.)
    const formats = {}; // Format ID -> format definition
    const colWidths = {}; // Column index -> width
    const comments = {}; // Cell address -> comment
    let maxRow = 0;
    let maxCol = 0;
    let currentRow = 1;
    let currentCol = 1;

    // Parse each line
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Skip empty lines
        if (!line) continue;

        // Comments (lines starting with ;)
        if (line.startsWith(';')) {
            // Store as general comment (could be enhanced to associate with cells)
            continue;
        }

        // Get the record type (first character)
        const recordType = line.charAt(0);

        // Skip if not a recognized record type
        if (!/^[IDBCFPOEWR]/.test(recordType)) continue;

        // End of file
        if (recordType === 'E') break;

        // Parse Format record
        if (recordType === 'F') {
            const parts = line.substring(2).split(';');
            let formatId = null;
            const formatProps = {};

            for (const part of parts) {
                if (!part) continue;
                const param = parseParam(part);

                switch (param.code) {
                    case 'Y': // Row position
                        currentRow = parseInt(param.value);
                        break;
                    case 'X': // Column position
                        currentCol = parseInt(param.value);
                        break;
                    case 'P': // Format ID
                        formatId = param.value;
                        break;
                    case 'F': // Format code (number style + alignment), not font
                        // F parameter format: Fch1digs ch2
                        // ch1 = number style (G=general, D=default, etc.)
                        // digs = number of digits
                        // ch2 = alignment (L=left, R=right, C=center)
                        // We parse alignment from the last character if it's L, R, or C
                        const lastChar = param.value.charAt(param.value.length - 1);
                        if (lastChar === 'L') formatProps.textAlign = 'left';
                        else if (lastChar === 'R') formatProps.textAlign = 'right';
                        else if (lastChar === 'C') formatProps.textAlign = 'center';
                        break;
                    case 'S': // Style (font styles and borders)
                        // If S contains border indicators (D, L, R, T, B with M), parse as borders
                        if (param.value.match(/[DLRTB]/) && param.value.match(/M\d+/)) {
                            const borders = parseBorderStyle(param.value);
                            Object.assign(formatProps, borders);
                        } else {
                            // Parse as font styles
                            if (param.value.includes('B')) formatProps.fontWeight = 'bold';
                            if (param.value.includes('I')) formatProps.fontStyle = 'italic';
                            if (param.value.includes('U')) formatProps.textDecoration = 'underline';
                        }
                        break;
                    case 'G': // Background/shading pattern
                        formatProps.gValue = param.value;
                        break;
                    case 'H': // Font height (in points)
                        formatProps.fontSize = param.value + 'pt';
                        break;
                    case 'M': // Number format mask
                        formatProps.mask = param.value;
                        break;
                    case 'A': // Alignment (L=left, R=right, C=center, G=general, J=justify)
                        switch (param.value.toUpperCase()) {
                            case 'L': formatProps.textAlign = 'left'; break;
                            case 'R': formatProps.textAlign = 'right'; break;
                            case 'C': formatProps.textAlign = 'center'; break;
                            case 'J': formatProps.textAlign = 'justify'; break;
                            case 'G': formatProps.textAlign = 'left'; break; // General defaults to left
                        }
                        break;
                    case 'C': // Color (typically foreground color index)
                        formatProps.colorIndex = param.value;
                        break;
                    case 'W': // Column width (format: W<start> <end> <width>)
                        const widthParts = param.value.split(' ');
                        if (widthParts.length >= 3) {
                            const startCol = parseInt(widthParts[0]) - 1;
                            const endCol = parseInt(widthParts[1]) - 1;
                            const width = parseInt(widthParts[2]);
                            const widthPx = convertWidthToPixels(width, 'pt');
                            for (let c = startCol; c <= endCol; c++) {
                                colWidths[c] = widthPx;
                            }
                        }
                        break;
                }
            }

            // Store format by ID if P parameter is present
            if (formatId) {
                formats[formatId] = formatProps;
            }

            // If X and/or Y are present, apply format directly to that cell/row/column
            if (currentRow > 0 && currentCol > 0) {
                // Format applies to specific cell
                const rowIdx = currentRow - 1;
                const colIdx = currentCol - 1;
                const cellAddress = getCellNameFromCoords(colIdx, rowIdx);

                // Convert formatProps to CSS and apply
                const css = formatPropsToCSS(formatProps);
                if (css) {
                    cellStyles[cellAddress] = css;
                }
            }
        }

        // Parse column Width record
        if (recordType === 'W') {
            const parts = line.substring(2).split(';');
            let col = currentCol;
            let width = null;

            for (const part of parts) {
                if (!part) continue;
                const param = parseParam(part);

                if (param.code === 'X') {
                    col = parseInt(param.value);
                } else if (param.code === 'W') {
                    width = parseInt(param.value);
                }
            }

            if (width) {
                colWidths[col - 1] = convertWidthToPixels(width, 'pt');
            }
        }

        // Parse cell record
        if (recordType === 'C') {
            // Split by semicolon to get parameters
            const parts = line.substring(2).split(';');

            let row = currentRow;
            let col = currentCol;
            let value = null;
            let formula = null;
            let formatId = null;
            const inlineFormat = {}; // Store inline formatting for this cell

            // Parse each parameter
            for (const part of parts) {
                if (!part) continue;

                const param = parseParam(part);

                switch (param.code) {
                    case 'Y': // Row (1-based)
                        row = parseInt(param.value);
                        currentRow = row;
                        break;

                    case 'X': // Column (1-based)
                        col = parseInt(param.value);
                        currentCol = col;
                        break;

                    case 'K': // Constant value
                        value = param.value;
                        // Remove quotes if it's a string
                        if (value.startsWith('"') && value.endsWith('"')) {
                            value = value.slice(1, -1);
                            // Decode SLK escape sequences
                            value = decodeSLKString(value);
                        } else {
                            // Try to parse as number
                            const numValue = parseFloat(value);
                            if (!isNaN(numValue)) {
                                value = numValue;
                            }
                        }
                        break;

                    case 'E': // Expression/Formula
                        formula = param.value;
                        break;

                    case 'P': // Format ID reference
                        formatId = param.value;
                        break;

                    // Inline formatting (same as F record parameters)
                    case 'S': // Style (font styles and borders)
                        // If S contains border indicators, parse as borders
                        if (param.value.match(/[DLRTB]/) && param.value.match(/M\d+/)) {
                            const borders = parseBorderStyle(param.value);
                            Object.assign(inlineFormat, borders);
                        } else {
                            // Parse as font styles
                            if (param.value.includes('B')) inlineFormat.fontWeight = 'bold';
                            if (param.value.includes('I')) inlineFormat.fontStyle = 'italic';
                            if (param.value.includes('U')) inlineFormat.textDecoration = 'underline';
                        }
                        break;

                    case 'G': // Background/shading pattern
                        inlineFormat.gValue = param.value;
                        break;

                    case 'F': // Format code (number style + alignment), not font
                        // F parameter format: Fch1digs ch2
                        const lastCharInline = param.value.charAt(param.value.length - 1);
                        if (lastCharInline === 'L') inlineFormat.textAlign = 'left';
                        else if (lastCharInline === 'R') inlineFormat.textAlign = 'right';
                        else if (lastCharInline === 'C') inlineFormat.textAlign = 'center';
                        break;

                    case 'H': // Font height (in points)
                        inlineFormat.fontSize = param.value + 'pt';
                        break;

                    case 'M': // Number format mask
                        inlineFormat.mask = param.value;
                        break;

                    case 'A': // Alignment
                        switch (param.value.toUpperCase()) {
                            case 'L': inlineFormat.textAlign = 'left'; break;
                            case 'R': inlineFormat.textAlign = 'right'; break;
                            case 'C': inlineFormat.textAlign = 'center'; break;
                            case 'J': inlineFormat.textAlign = 'justify'; break;
                            case 'G': inlineFormat.textAlign = 'left'; break;
                        }
                        break;

                    case 'C': // Color
                        inlineFormat.colorIndex = param.value;
                        break;
                }
            }

            // Store cell data (convert to 0-based indices)
            const rowIdx = row - 1;
            const colIdx = col - 1;
            const cellKey = `${rowIdx},${colIdx}`;
            const cellAddress = getCellNameFromCoords(colIdx, rowIdx);

            // If there's a formula, store it with = prefix (Jspreadsheet format)
            if (formula) {
                // Formulas in SLK might not start with =, so add it if needed
                let formulaStr = formula.startsWith('=') ? formula : '=' + formula;
                // Convert RC notation to A1 notation
                formulaStr = convertR1C1toA1(formulaStr, rowIdx, colIdx);
                cells[cellKey] = formulaStr;

                // Store formula metadata
                if (!cellMeta[cellAddress]) cellMeta[cellAddress] = {};
                cellMeta[cellAddress].formula = formulaStr;
            } else if (value !== null) {
                cells[cellKey] = value;
            } else {
                cells[cellKey] = '';
            }

            // Merge format ID and inline formatting (inline takes precedence)
            const mergedFormat = {};
            if (formatId && formats[formatId]) {
                Object.assign(mergedFormat, formats[formatId]);
            }
            // Apply inline formatting (overrides format ID)
            Object.assign(mergedFormat, inlineFormat);

            // Apply merged formatting if present
            if (Object.keys(mergedFormat).length > 0) {
                // Convert to CSS using the helper function
                const css = formatPropsToCSS(mergedFormat);
                if (css) {
                    cellStyles[cellAddress] = css;
                }

                // Store non-CSS metadata (like number format mask)
                if (mergedFormat.mask) {
                    if (!cellMeta[cellAddress]) cellMeta[cellAddress] = {};
                    cellMeta[cellAddress].mask = mergedFormat.mask;
                }
            }

            // Update max dimensions
            if (row > maxRow) maxRow = row;
            if (col > maxCol) maxCol = col;

            // Move to next column for relative positioning
            currentCol++;
        }

        // Parse bounds record (optional, defines sheet dimensions)
        if (recordType === 'B') {
            const parts = line.substring(2).split(';');
            for (const part of parts) {
                if (!part) continue;
                const param = parseParam(part);

                if (param.code === 'Y') {
                    const boundRow = parseInt(param.value);
                    if (boundRow > maxRow) maxRow = boundRow;
                }
                if (param.code === 'X') {
                    const boundCol = parseInt(param.value);
                    if (boundCol > maxCol) maxCol = boundCol;
                }
            }
        }
    }

    // Build 2D array from cells
    const data = [];
    for (let r = 0; r < maxRow; r++) {
        const row = [];
        for (let c = 0; c < maxCol; c++) {
            const cellKey = `${r},${c}`;
            row.push(cells[cellKey] !== undefined ? cells[cellKey] : '');
        }
        data.push(row);
    }

    // Create column definitions
    const columns = [];
    for (let i = 0; i < maxCol; i++) {
        columns.push({
            title: getColumnName(i),
            width: colWidths[i] || 100  // Use parsed width or default to 100
        });
    }

    // Deduplicate borders (remove redundant borders between adjacent cells)
    deduplicateBorders(cellStyles, maxRow, maxCol);

    // Return in Jspreadsheet Pro format
    return {
        worksheets: [
            {
                data: data,
                columns: columns,
                rows: {},           // SLK doesn't have explicit row properties
                cells: cellMeta,    // Cell metadata
                style: cellStyles,  // Cell styles (CSS)
                mergeCells: {},     // SLK doesn't support merged cells
                comments: comments  // Comments
            }
        ]
    };
}

/**
 * Parse SLK file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseSLK(input, options = {}) {
    return parse(async (inp) => {
        const content = await loadAsString(inp);
        return parseSLKString(content);
    }, input, options);
}
