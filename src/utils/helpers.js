/**
 * Spreadsheet helper utilities for cell references, conversions, and formatting
 * @module helpers
 */

// Column Names
const columnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

const getColumnNameCache = {};

/**
 * Convert column index to letter notation (0=A, 1=B, 26=AA, etc.)
 *
 * @param {number|string} num - Zero-based column index
 * @returns {string} Column letter(s)
 * @example
 * getColumnName(0)   // 'A'
 * getColumnName(25)  // 'Z'
 * getColumnName(26)  // 'AA'
 * getColumnName(701) // 'ZZ'
 */
const getColumnName = function(num) {
    if (typeof(num) === 'string') {
        num = parseInt(num);
    }

    let result = getColumnNameCache[num];
    if (typeof(result) !== 'undefined') {
        return result;
    } else {
        let letter = ''
        let quotient = num + 1
        let remainder;
        while (quotient > 0) {
            remainder = (quotient - 1) % 26
            letter = columnNames[remainder] + letter
            quotient = ((quotient - remainder) / 26) | 0
        }
        return getColumnNameCache[num] = letter;
    }
};

/**
 * Convert zero-based coordinates to cell name (e.g., 0,0 -> A1)
 *
 * @param {number|string} x - Zero-based column index
 * @param {number|string} y - Zero-based row index
 * @returns {string} Cell name in A1 notation
 * @example
 * getCellNameFromCoords(0, 0)   // 'A1'
 * getCellNameFromCoords(5, 10)  // 'F11'
 * getCellNameFromCoords(26, 0)  // 'AA1'
 */
const getCellNameFromCoords = function(x, y) {
    if (typeof(x) === 'string') {
        x = parseInt(x);
    }
    if (typeof(y) === 'string') {
        y = parseInt(y);
    }
    return getColumnName(x) + (y + 1);
}

/**
 * Convert cell name to zero-based coordinates (e.g., A1 -> [0, 0])
 *
 * @param {string} name - Cell name in A1 notation
 * @returns {number[]} Array of [columnIndex, rowIndex] (zero-based)
 * @example
 * getCoordsFromCellName('A1')   // [0, 0]
 * getCoordsFromCellName('F11')  // [5, 10]
 * getCoordsFromCellName('AA1')  // [26, 0]
 */
const getCoordsFromCellName = function(name) {
    if (! name) {
        return [];
    }

    let letter = null;
    let number = [];
    let size = name.length;
    let c;

    // Convert column name to numeric code
    for (let i = 0; i < size; i++) {
        c = name.charCodeAt(i);
        if (c >= 65 && c <= 90) { // Only process uppercase letters
            letter = letter * 26 + (c - 64);
        } else if (c >= 48 && c <= 57) {
            number.push(c-48);
        }
    }

    // Convert row number to numeric value
    let total = null;
    size = number.length;
    if (size) {
        for (let i = 0; i < size; i++) {
            total = total * 10 + number[i];
        }
        if (total) {
            total = total - 1;
        }
    }

    if (letter === null) {
        return [letter, total];
    }

    return [letter - 1, total];
}

/**
 * Convert column letter to zero-based index (A=0, B=1, AA=26, etc.)
 *
 * @param {string} letter - Column letter(s)
 * @returns {number} Zero-based column index
 * @example
 * getColumnIndex('A')   // 0
 * getColumnIndex('Z')   // 25
 * getColumnIndex('AA')  // 26
 * getColumnIndex('ZZ')  // 701
 */
const getColumnIndex = function(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 65 + 1);
    }
    return index - 1;
};

/**
 * Get the coordinates from a range.
 * @param {string} range
 * @param {boolean?} adjust range to the worksheet dimension
 * @return {number[]}
 */
const getCoordsFromRange = function(range, adjust) {
    if (range.indexOf('!') !== -1) {
        range = range.split('!');
        range = range[1];
    }
    range = range.split(':');
    if (! range[1]) {
        range[1] = range[0];
    }

    let t1 = getCoordsFromCellName(range[0]);
    let t2 = getCoordsFromCellName(range[1]);
    // For A:A and 1:1 style
    if (adjust !== false) {
        if (this && typeof(this.cols) !== 'undefined') {
            if (t2[0] == null) {
                t2[0] = this.cols.length - 1;
            }
            if (t2[1] == null) {
                t2[1] = this.rows.length - 1;
            }
        }

        if (t1[0] == null) {
            t1[0] = 0;
        }
        if (t1[1] == null) {
            t1[1] = 0;
        }
        if (t2[0] == null) {
            t2[0] = 0;
        }
        if (t2[1] == null) {
            t2[1] = 0;
        }
    }

    return [ t1[0], t1[1], t2[0], t2[1] ];
}

/**
 * Get the tokens from the coordinates
 * @param {number} x1
 * @param {number} y1
 * @param {number} x2
 * @param {number} y2
 * @param {string} worksheetName
 * @return {string[]}
 */
const getTokensFromCoords = function(x1, y1, x2, y2, worksheetName) {
    let f = [];
    let t;
    for (let j = y1; j <= y2; j++) {
        for (let i = x1; i <= x2; i++) {
            t = getCellNameFromCoords(i, j);
            if (worksheetName) {
                t = worksheetName + '!' + t;
            }
            f.push(t);
        }
    }
    return f;
}

const tokenIdentifier = function (s) {
    if (! s || typeof s !== 'string') {
        return false;
    }

    // To be a valid spreadsheet like variable token, the first character of the string should be a letter or '
    let c = s.charCodeAt(0);

    // First valid character
    if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || c === 39 || c === 36 || c === 95) {
        let index = 0;
        let numOfChars = s.length;
        let worksheetName = null;

        if (c === 39) {
            // Check the string
            for (index = 1; index < numOfChars; index++) {
                // Get the next char
                c = s.charCodeAt(index);

                // Inside quotes
                if (c === 39) {
                    index++;

                    if (s[index] === '!') {
                        worksheetName = true;
                        break;
                    } else {
                        return false;
                    }
                } else if (c === 58) {
                    // Cannot have :
                    return false;
                }
            }

            if (! worksheetName) {
                return false;
            }

            index++;
        } else {
            let hasSpecial = c >= 48 && c <= 57;
            // Only scan if not starting with quote (quoted handling below)
            for (let i = 0; i < numOfChars; i++) {
                c = s.charCodeAt(i);
                if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || c === 95) {
                    // Do nothing
                } else if (c === 33) {
                    if (hasSpecial === true) {
                        return false;
                    } else {
                        worksheetName = true;
                        index = i + 1;
                        break;
                    }
                } else {
                    hasSpecial = true;
                }
            }
        }

        let x1 = '';
        let y1 = '';
        let x2 = '';
        let y2 = '';

        // Letter counters for performance (excluding dollar signs)
        let x1LetterCount = 0;
        let x2LetterCount = 0;

        // Not a range 0 or range 1 (alphanumeric), 2 (letter only), 3 (numeric only)
        let range = 0;

        // Check the string
        for (let i = index; i < numOfChars; i++) {
            // Get the next char
            c = s.charCodeAt(i);

            // Letters, underscore
            if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122)) {
                if (range) {
                    if (y2) {
                        return false;
                    }
                    x2 += s[i];
                    x2LetterCount++;
                } else {
                    if (y1) {
                        return false;
                    }
                    x1 += s[i];
                    x1LetterCount++;
                }
            } else if (c >= 48 && c <= 57) {
                if (range) {
                    y2 += s[i];
                } else {
                    y1 += s[i];
                }
            } else if (c === 58) {
                if (range) {
                    // String cannot contain two ':' in a valid string
                    return false;
                } else {
                    // Type of range
                    if (x1 && y1) {
                        range = 1;
                    } else if (x1 && !y1) {
                        range = 2;
                    } else if (!x1 && y1) {
                        range = 3;
                    } else {
                        // Not a valid token
                        return false;
                    }
                }
            } else if (c === 33) {
                // Exclamation mark
                if (range || worksheetName) {
                    return false;
                } else {
                    worksheetName = x1 + y1;
                    x1 = '';
                    y1 = '';
                    x1LetterCount = 0;
                }
            } else if (c === 36) {
                let next = s.charCodeAt(i+1);

                if (next >= 65 && next <= 90) {
                    if (range) {
                        if (x2) {
                            return false;
                        }
                        x2 += '$';
                    } else {
                        if (x1) {
                            return false;
                        }
                        x1 += '$';
                    }
                } else if (next >= 48 && next <= 57) {
                    if (range) {
                        if (y2) {
                            return false;
                        }
                        y2 += '$';
                    } else {
                        if (y1) {
                            return false;
                        }
                        y1 += '$';
                    }
                } else {
                    return false;
                }
            } else {
                return false;
            }
        }

        // Only allow up to 3 letters
        if (x2LetterCount > 3) {
            return false;
        }
        // Only allow up to 3 letters
        if (x1LetterCount > 3) {
            return false;
        }

        if (range) {
            // Type of range
            if (x2 && y2) {
                if (range === 1) {
                    return true;
                }
            } else if (x2 && !y2) {
                if (range === 2) {
                    return true;
                }
            } else if (!x2 && y2) {
                if (range === 3) {
                    return true;
                }
            }
        } else {
            if (x1 && y1) {
                return true;
            }
        }
    }

    return false;
}

// HTML/XML Helper Functions
/**
 * Get property value from node props (for HTML/XML nodes)
 * @param {object} node
 * @param {string} propName
 * @return {string|undefined}
 */
const getProp = function(node, propName) {
    if (!node.props) return undefined;
    const prop = node.props.find(p => p.name === propName || p.name.endsWith(':' + propName));
    return prop ? prop.value : undefined;
};

/**
 * Get all children of a specific type (for HTML/XML nodes)
 * @param {object} node
 * @param {string} type
 * @return {array}
 */
const getChildren = function(node, type) {
    if (!node.children) return [];
    return node.children.filter(child =>
        child.type === type || child.type.endsWith(':' + type)
    );
};

/**
 * Get text content from XML node recursively
 * @param {object|string|array} node
 * @return {string}
 */
const getTextContent = function(node) {
    if (!node) return '';
    if (typeof node === 'string') return node;
    if (Array.isArray(node)) {
        return node.map(n => getTextContent(n)).join('');
    }

    // Handle #text nodes with textContent prop
    if (node.type === '#text' && node.props) {
        const textProp = node.props.find(p => p.name === 'textContent');
        if (textProp) return textProp.value;
    }

    if (node.children) {
        return getTextContent(node.children);
    }
    return '';
};

/**
 * Parse XML attributes into object
 * @param {object} element
 * @return {object}
 */
const parseAttributes = function(element) {
    const attrs = {};
    if (element && element.props) {
        element.props.forEach(attr => {
            attrs[attr.name] = attr.value;
        });
    }
    // Fallback for old format
    if (element && element.attributes) {
        element.attributes.forEach(attr => {
            attrs[attr.name] = attr.value;
        });
    }
    return attrs;
};

/**
 * Find nodes by type recursively
 * @param {object|array} node
 * @param {string} type
 * @return {array}
 */
const findNodes = function(node, type) {
    const results = [];

    const search = (n) => {
        if (!n) return;

        if (Array.isArray(n)) {
            n.forEach(search);
            return;
        }

        if (typeof n === 'object') {
            if (n.type === type) {
                results.push(n);
            }
            if (n.children) {
                search(n.children);
            }
        }
    };

    search(node);
    return results;
};

// Binary Reading Helper Functions
/**
 * Read unsigned 8-bit integer
 * @param {Uint8Array} buffer
 * @param {number} offset
 * @return {number}
 */
const readUInt8 = function(buffer, offset) {
    return buffer[offset];
};

/**
 * Read unsigned 16-bit integer (little-endian)
 * @param {Uint8Array} buffer
 * @param {number} offset
 * @return {number}
 */
const readUInt16LE = function(buffer, offset) {
    return buffer[offset] | (buffer[offset + 1] << 8);
};

/**
 * Read unsigned 32-bit integer (little-endian)
 * @param {Uint8Array} buffer
 * @param {number} offset
 * @return {number}
 */
const readUInt32LE = function(buffer, offset) {
    return (buffer[offset] | (buffer[offset + 1] << 8) |
           (buffer[offset + 2] << 16) | (buffer[offset + 3] << 24)) >>> 0;
};

/**
 * Read signed 16-bit integer (little-endian)
 * @param {Uint8Array} buffer
 * @param {number} offset
 * @return {number}
 */
const readInt16LE = function(buffer, offset) {
    const val = readUInt16LE(buffer, offset);
    return val > 32767 ? val - 65536 : val;
};

/**
 * Read 64-bit float (little-endian)
 * @param {Uint8Array} buffer
 * @param {number} offset
 * @return {number}
 */
const readFloat64LE = function(buffer, offset) {
    const view = new DataView(buffer.buffer, buffer.byteOffset + offset, 8);
    return view.getFloat64(0, true);
};

// Export functions for use in other modules
/**
 * Decode HTML entities
 */
const decodeHTMLEntities = function(text) {
    return text
        .replace(/&quot;/gi, '"')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&apos;/gi, "'");
};

/**
 * Convert column width from various units to pixels
 * @param {number} width - Width value
 * @param {string} unit - Unit type: 'pt' (points), 'char' (character width), or 'px' (pixels)
 * @return {number} Width in pixels
 */
const convertWidthToPixels = function(width, unit = 'pt') {
    if (!width || isNaN(width)) return 100; // Default width

    switch (unit) {
        case 'pt':
            // Points to pixels: 1pt ≈ 7.15px (based on Excel's rendering)
            return Math.round(parseFloat(width) * 7.15);
        case 'char':
            // Character width to pixels: 1 char ≈ 8.43px (Excel default font)
            return Math.round(parseFloat(width) * 8.43);
        case 'px':
            // Already in pixels
            return Math.round(parseFloat(width));
        default:
            // Default to points conversion
            return Math.round(parseFloat(width) * 7.15);
    }
};

/**
 * Convert R1C1 notation to A1 notation
 * @param {string} formula - Formula in R1C1 format
 * @param {number} currentRow - Current row index (0-based)
 * @param {number} currentCol - Current column index (0-based)
 */
const convertR1C1toA1 = function(formula, currentRow, currentCol) {
    if (!formula) return formula;

    // First decode HTML entities
    formula = decodeHTMLEntities(formula);

    // Pattern to match R1C1 references
    // Matches: RC, RC[n], R[n]C, R[n]C[n], RnC, RnCn, RnC[n], R[n]Cn
    const r1c1Pattern = /R(\[?-?\d+\]?)?C(\[?-?\d+\]?)?/g;

    return formula.replace(r1c1Pattern, (match, rowPart, colPart) => {
        let row, col;
        let rowAbsolute = true;
        let colAbsolute = true;

        // Parse row part
        if (!rowPart) {
            // RC or C - relative, same row
            row = currentRow;
            rowAbsolute = false;
        } else if (rowPart.startsWith('[')) {
            // R[n] - relative reference
            const offset = parseInt(rowPart.slice(1, -1));
            row = currentRow + offset;
            rowAbsolute = false;
        } else {
            // Rn - absolute reference (1-based in R1C1)
            row = parseInt(rowPart) - 1;
            rowAbsolute = true;
        }

        // Parse column part
        if (!colPart) {
            // This shouldn't happen in valid R1C1, but handle it
            col = currentCol;
            colAbsolute = false;
        } else if (colPart.startsWith('[')) {
            // C[n] - relative reference
            const offset = parseInt(colPart.slice(1, -1));
            col = currentCol + offset;
            colAbsolute = false;
        } else {
            // Cn - absolute reference (1-based in R1C1)
            col = parseInt(colPart) - 1;
            colAbsolute = true;
        }

        // Convert to A1 notation
        const colLetter = getColumnName(col);
        const rowNumber = row + 1;

        // Add $ for absolute references
        const colPrefix = colAbsolute ? '$' : '';
        const rowPrefix = rowAbsolute ? '$' : '';

        return colPrefix + colLetter + rowPrefix + rowNumber;
    });
};

const getDefaultTheme = function() {
    return {
        arrayColors: [
            "FFFFFF",
            "000000",
            "E7E6E6",
            "44546A",
            "4472C4",
            "ED7D31",
            "A5A5A5",
            "FFC000",
            "5B9BD5",
            "70AD47",
            "0563C1",
            "954F72"
        ],
        objColors: {
            dk1: "000000",
            lt1: "FFFFFF",
            dk2: "44546A",
            lt2: "E7E6E6",
            accent1: "4472C4",
            accent2: "ED7D31",
            accent3: "A5A5A5",
            accent4: "FFC000",
            accent5: "5B9BD5",
            accent6: "70AD47",
            hlink: "0563C1",
            folHlink: "954F72",
            tx1: "000000",
            tx2: "44546A",
            bg1: "FFFFFF",
            bg2: "E7E6E6"
        },
    };
}

const numberFormats = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'm/d/yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
    56: '',
    65535: 'General',
};

const borderStyles = {
    thin: ['solid', '1px'],
    medium: ['solid', '2px'],
    thick: ['solid', '3px'],
    dotted: ['dotted', '1px'],
    dashed: ['dashed', '1px'],
    double: ['double', '3px'],
    mediumDashed: ['dashed', '2px'],
}

const excelValidationTypes = {
    whole: 'number',
    decimal: 'number',
    list: 'list',
    textLength: 'textLength',
    date: 'date',
    custom: 'formula',
    time: 'time',
}

const excelValidationOperations = {
    notBetween: 'not between',
    equal: '=',
    notEqual: '!=',
    greaterThan: '>',
    lessThan: '<',
    greaterThanOrEqual: '>=',
    lessThanOrEqual: '<=',
};

const exclusions = ['_xlfn.', '_xll.', '_xlws.', '_xlpm.', '_xleta.', '_xlnm.'];

const excelCFSimpleTypes = {
    containsBlanks: 'empty',
    notContainsBlanks: 'notEmpty'
}

const excelCFTextTypes = {
    containsText: 'contains',
    notContainsText: 'not contains',
    beginsWith: 'begins with',
    endsWith: 'ends with',
};

const excelCFNumericOperators = {
    equal: '=',
    notEqual: '!=',
    greaterThan: '>',
    lessThan: '<',
    greaterThanOrEqual: '>=',
    lessThanOrEqual: '<=',
    between: 'between',
    notBetween: 'not between',
};

const shapeMap = {
    rect: 'rectangle',
    roundRect: 'rounded-rectangle',
    triangle: 'triangle',
    rtTriangle: 'right-triangle',
    ellipse: 'ellipse',
    diamond: 'diamond',
    trapezoid: 'trapezium',
    pentagon: 'pentagon',
    parallelogram: 'parallelogram',
    hexagon: 'hexagon',
    heptagon: 'heptagon',
    octagon: 'octagon',
    decagon: 'decagon',
    snip1Rect: 'drawSingleCutCornerRectangle',
    snip2SameRect: 'drawDoubleCutCornerRectangle',
    snip2DiagRect: 'drawOppositeCutCornerRectangle',
    round1Rect: 'drawRoundedTopRightRectangle',
    pie: 'drawOvalInterfaceIcon',
    chord: 'drawCutCircle',
    frame: 'drawDoubleLineSquare',
    halfFrame: 'drawdoubledtoprighangle',
    corner: 'drawDoubleLineShape',
    diagStripe: 'drawMouldingCrown',
    plus: 'drawCrossIcon',
    plaque: 'drawPlaqueShape',
    can: 'drawCylinderShape',
    cube: 'drawCubeShape',
    bevel: 'drawBevelShape',
    donut: 'drawDonutShape',
    noSmoking: 'drawBlockShape',
    blockArc: 'drawBlockArk',
    smileyFace: 'drawSmileyFace',
    heart: 'drawHeart',
    lightningBolt: 'drawLightningBolt',
    sun: 'drawSun',
    moon: 'drawCrescentMoon',
    cloud: 'drawCloud',
    arc: 'drawArcCurve',
    bracePair: 'drawFlowerBrackets',
    leftBrace: 'drawLeftFlowerBracket',
    rightBrace: 'drawRightFlowerBracket',
    rightArrow: 'drawRightBlockArrow',
    leftArrow: 'drawLeftBlockArrow',
    upArrow: 'drawUpBlockArrow',
    downArrow: 'drawDownBlockArrow',
    leftRightArrow: 'drawLeftRightArrow',
    upDownArrow: 'drawUpDownArrow',
    quadArrow: 'drawQuadArrow',
    leftRightUpArrow: 'drawLeftRightUpArrow',
    bentArrow: 'drawBentArrow',
    uturnArrow: 'drawUTurnArrow',
    leftUpArrow: 'drawLeftUpArrow',
    bentUpArrow: 'drawBentUpArrow',
    curvedRightArrow: 'drawCurvedRightArrow',
    curvedLeftArrow: 'drawCurvedLeftArrow',
    curvedDownArrow: 'drawCurvedDownArrow',
    curvedUpArrow: 'drawCurvedUpArrow',
    stripedRightArrow: 'drawStripedRightArrow',
    notchedRightArrow: 'drawNotchedRightArrow',
    homePlate: 'drawPentagonArrow',
    chevron: 'drawChevron',
    rightArrowCallout: 'drawRightArrowCallout',
    downArrowCallout: 'drawDownArrowCallout',
    leftArrowCallout: 'drawLeftArrowCallout',
    upArrowCallout: 'drawUpArrowCallout',
    leftRightArrowCallout: 'drawLeftRightArrowCallout',
    quadArrowCallout: 'drawQuadArrowCallout',
    circularArrow: 'drawCircularArrow',
    mathPlus: 'drawPlus',
    mathMinus: 'drawMinus',
    mathMultiply: 'drawMultiplication',
    mathDivide: 'drawDivision',
    mathEqual: 'drawEqual',
    mathNotEqual: 'drawNotEqualTo',
    flowChartProcess: 'rectangle',
    flowChartAlternateProcess: 'rounded-rectangle',
    flowChartDecision: 'diamond',
    flowChartInputOutput: 'parallelogram',
    flowChartPredefinedProcess: 'drawPredefinedTask',
    flowChartInternalStorage: 'drawInternalStorage',
    flowChartDocument: 'drawDocument',
    flowChartMultidocument: 'drawMultitaskingDocuments',
    flowChartTerminator: 'drawTerminator',
    flowChartPreparation: 'drawPreparation',
    flowChartManualInput: 'drawManualInput',
    flowChartManualOperation: 'drawManualOperation',
    flowChartConnector: 'ellipse',
    flowChartOffpageConnector: 'drawOffpageConnector',
    flowChartPunchedCard: 'drawCard',
    flowChartPunchedTape: 'drawPunchedTape',
    flowChartSummingJunction: 'drawSwimmingJunction',
    flowChartOr: 'drawOrShape',
    flowChartCollate: 'drawCollateShape',
    flowChartSort: 'drawSortShape',
    flowChartExtract: 'drawExtractShape',
    flowChartMerge: 'drawMergeShape',
    flowChartDelay: 'drawDelayShape',
    flowChartMagneticTape: 'drawSequentialAccessStorageShape',
    flowChartMagneticDisk: 'drawCylinderShape',
    flowChartMagneticDrum: 'drawDirectAccessStorage',
    flowChartDisplay: 'drawDisplayShape',
};


const hex2RGB = function(h) {
    const o = h.slice(h[0] === "#" ? 1 : 0).slice(0, 6);

    return [parseInt(o.slice(0, 2), 16), parseInt(o.slice(2, 4), 16), parseInt(o.slice(4, 6), 16)];
}

const rgb2Hex = function (rgb) {
    let o = 1;

    for (let i = 0; i !== 3; ++i) {
        o = o * 256 + (rgb[i] > 255 ? 255 : rgb[i] < 0 ? 0 : rgb[i]);
    }

    return o.toString(16).toUpperCase().slice(1);
}

const rgb2HSL = function(rgb) {
    const R = rgb[0] / 255,
        G = rgb[1] / 255,
        B = rgb[2] / 255,
        M = Math.max(R, G, B),
        m = Math.min(R, G, B),
        C = M - m;

    if (C === 0) return [0, 0, R];

    const L2 = (M + m);
    const S = C / (L2 > 1 ? 2 - L2 : L2);

    let H6 = 0;

    switch (M) {
        case R:
            H6 = ((G - B) / C + 6) % 6;
            break;
        case G:
            H6 = ((B - R) / C + 2);
            break;
        case B:
            H6 = ((R - G) / C + 4);
            break;
    }

    return [H6 / 6, S, L2 / 2];
}

const hsl2RGB = function(hsl) {
    const H = hsl[0],
        S = hsl[1],
        L = hsl[2],
        C = S * 2 * (L < 0.5 ? L : 1 - L),
        m = L - C / 2,
        rgb = [m, m, m],
        h6 = 6 * H;

    let X;
    if (S !== 0) {
        switch (h6 | 0) {
            case 0:
            case 6:
                X = C * h6;
                rgb[0] += C;
                rgb[1] += X;
                break;
            case 1:
                X = C * (2 - h6);
                rgb[0] += X;
                rgb[1] += C;
                break;
            case 2:
                X = C * (h6 - 2);
                rgb[1] += C;
                rgb[2] += X;
                break;
            case 3:
                X = C * (4 - h6);
                rgb[1] += X;
                rgb[2] += C;
                break;
            case 4:
                X = C * (h6 - 4);
                rgb[2] += C;
                rgb[0] += X;
                break;
            case 5:
                X = C * (6 - h6);
                rgb[2] += X;
                rgb[0] += C;
                break;
        }
    }

    for (let i = 0; i != 3; ++i) {
        rgb[i] = Math.round(rgb[i] * 255);
    }

    return rgb;
}

const rgb_tint = function(hex, tint) {
    if (!tint) {
        return hex;
    }
    tint = parseFloat(tint);

    const hsl = rgb2HSL(hex2RGB(hex));
    if (tint < 0) {
        hsl[2] = hsl[2] * (1 + tint);
    } else {
        hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
    }
    return rgb2Hex(hsl2RGB(hsl));
}

export {
    getColumnName,
    getColumnIndex,
    getCellNameFromCoords,
    getCoordsFromCellName,
    getCoordsFromRange,
    getTokensFromCoords,
    tokenIdentifier,
    getProp,
    getChildren,
    getTextContent,
    parseAttributes,
    findNodes,
    readUInt8,
    readUInt16LE,
    readUInt32LE,
    readInt16LE,
    readFloat64LE,
    decodeHTMLEntities,
    convertR1C1toA1,
    convertWidthToPixels,
    getDefaultTheme,
    numberFormats,
    borderStyles,
    exclusions,
    excelValidationTypes,
    excelValidationOperations,
    excelCFSimpleTypes,
    excelCFTextTypes,
    excelCFNumericOperators,
    shapeMap,
    hex2RGB,
    rgb2Hex,
    rgb2HSL,
    hsl2RGB,
    rgb_tint
};
