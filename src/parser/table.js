import { loadAsString, parse } from '../utils/loader.js';
import { parser } from '@lemonadejs/html-to-json';
import { getColumnName, getCellNameFromCoords, getProp, getChildren, getTextContent } from '../utils/helpers.js';

/**
 * Whether a string is HTML markup rather than a file path
 */
export function isHTMLString(input) {
    return typeof input === 'string' && input.trimStart().startsWith('<');
}

/**
 * HTML source of the input: a DOM element, an HTML string, or a file (path, File, Blob, bytes)
 */
async function readHTML(input) {
    if (typeof HTMLElement !== 'undefined' && input instanceof HTMLElement) {
        return input.outerHTML;
    }
    if (isHTMLString(input)) {
        return input;
    }
    return loadAsString(input);
}

/**
 * Recursively find all table nodes in the tree
 */
function findAllTables(node) {
    const tables = [];

    if (node.type === 'table') {
        tables.push(node);
    }

    if (node.children) {
        for (const child of node.children) {
            tables.push(...findAllTables(child));
        }
    }

    return tables;
}

// Named HTML entities: the XML ones, all of Latin-1 (U+00A0..U+00FF), Greek letters and common typography/math.
// &nbsp; is decoded to a regular space (U+0020) so spreadsheet values never carry U+00A0.
const htmlEntities = (() => {
    const map = { quot: '"', amp: '&', lt: '<', gt: '>', apos: "'", nbsp: ' ' };
    const latin1 = ('iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn sup2 sup3 ' +
        'acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute Acirc Atilde Auml Aring ' +
        'AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times ' +
        'Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde auml aring aelig ccedil egrave ' +
        'eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde ouml divide oslash ugrave ' +
        'uacute ucirc uuml yacute thorn yuml').split(' ');
    latin1.forEach((name, i) => { map[name] = String.fromCharCode(0xA1 + i); });
    const greek = 'Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa Lambda Mu Nu Xi Omicron Pi Rho - Sigma Tau Upsilon Phi Chi Psi Omega'.split(' ');
    greek.forEach((name, i) => {
        if (name === '-') return;
        map[name] = String.fromCharCode(0x391 + i);
        map[name.toLowerCase()] = String.fromCharCode(0x3B1 + i);
    });
    map.sigmaf = 'ς';
    Object.assign(map, {
        OElig: 'Œ', oelig: 'œ', Scaron: 'Š', scaron: 'š', Yuml: 'Ÿ', fnof: 'ƒ',
        circ: 'ˆ', tilde: '˜', ensp: ' ', emsp: ' ', thinsp: ' ', zwnj: '‌',
        zwj: '‍', lrm: '‎', rlm: '‏', ndash: '–', mdash: '—', lsquo: '‘',
        rsquo: '’', sbquo: '‚', ldquo: '“', rdquo: '”', bdquo: '„', dagger: '†',
        Dagger: '‡', bull: '•', hellip: '…', permil: '‰', prime: '′', Prime: '″',
        lsaquo: '‹', rsaquo: '›', oline: '‾', frasl: '⁄', euro: '€', trade: '™',
        larr: '←', uarr: '↑', rarr: '→', darr: '↓', harr: '↔', minus: '−',
        lowast: '∗', radic: '√', prop: '∝', infin: '∞', asymp: '≈', ne: '≠',
        equiv: '≡', le: '≤', ge: '≥', sum: '∑', prod: '∏', part: '∂',
        forall: '∀', exist: '∃', isin: '∈', notin: '∉', cap: '∩', cup: '∪',
        int: '∫', loz: '◊', spades: '♠', clubs: '♣', hearts: '♥', diams: '♦'
    });
    return map;
})();

// Windows-1252 code points that HTML maps numeric references &#128;..&#159; to
const cp1252 = {
    0x80: 0x20AC, 0x82: 0x201A, 0x83: 0x0192, 0x84: 0x201E, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021,
    0x88: 0x02C6, 0x89: 0x2030, 0x8A: 0x0160, 0x8B: 0x2039, 0x8C: 0x0152, 0x8E: 0x017D, 0x91: 0x2018,
    0x92: 0x2019, 0x93: 0x201C, 0x94: 0x201D, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02DC,
    0x99: 0x2122, 0x9A: 0x0161, 0x9B: 0x203A, 0x9C: 0x0153, 0x9E: 0x017E, 0x9F: 0x0178
};

/**
 * Decode HTML character references (named, decimal and hex) in a single pass.
 * The html-to-json parser keeps text and attribute values exactly as written.
 * Unknown named entities are left untouched.
 */
function decodeEntities(text) {
    if (!text || text.indexOf('&') === -1) return text;

    return text.replace(/&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|([A-Za-z][A-Za-z0-9]*));/g, (match, dec, hex, name) => {
        if (name !== undefined) {
            return Object.prototype.hasOwnProperty.call(htmlEntities, name) ? htmlEntities[name] : match;
        }
        let code = dec !== undefined ? parseInt(dec, 10) : parseInt(hex, 16);
        if (cp1252[code]) {
            code = cp1252[code];
        }
        if (code === 0 || code > 0x10FFFF || (code >= 0xD800 && code <= 0xDFFF) || isNaN(code)) {
            return '�';
        }
        return String.fromCodePoint(code);
    });
}

/**
 * Parse a single table node from HTML tree
 */
function parseTableNode(tableNode, options) {
    const { firstRowAsHeader = true } = options;

    // Collect all rows in display order: thead, then every tbody and any <tr>
    // placed directly under <table> (in document order), then tfoot
    const headRows = [];
    const bodyRows = [];
    const footRows = [];

    if (tableNode.children) {
        for (const child of tableNode.children) {
            if (child.type === 'thead') {
                headRows.push(...getChildren(child, 'tr'));
            } else if (child.type === 'tbody') {
                bodyRows.push(...getChildren(child, 'tr'));
            } else if (child.type === 'tfoot') {
                footRows.push(...getChildren(child, 'tr'));
            } else if (child.type === 'tr') {
                bodyRows.push(child);
            }
        }
    }

    const allRows = [...headRows, ...bodyRows, ...footRows];

    if (allRows.length === 0) {
        return {
            data: [],
            columns: []
        };
    }

    // Parse rows and cells
    const grid = []; // Temporary grid to handle rowspan/colspan
    const mergeCells = {};
    const styles = {};
    const comments = {};
    const nestedHeaders = [];
    let maxCols = 0;

    // Track which cells are occupied by rowspan/colspan
    const occupied = new Map(); // Map<"row,col", true>

    allRows.forEach((rowNode, rowIndex) => {
        // th and td cells in document order
        const cells = (rowNode.children || []).filter(child => child.type === 'th' || child.type === 'td');

        let colIndex = 0;

        cells.forEach(cellNode => {
            // Skip occupied cells
            while (occupied.has(`${rowIndex},${colIndex}`)) {
                colIndex++;
            }

            // Get cell content
            let cellValue = '';

            // Check for formula (data-formula attribute)
            const formula = getProp(cellNode, 'data-formula');
            if (formula) {
                cellValue = decodeEntities(formula);
                cellValue = cellValue.startsWith('=') ? cellValue : '=' + cellValue;
            } else {
                cellValue = decodeEntities(getTextContent(cellNode)).trim();
            }

            // Get rowspan and colspan
            const rowspan = parseInt(getProp(cellNode, 'rowspan')) || 1;
            const colspan = parseInt(getProp(cellNode, 'colspan')) || 1;

            // Initialize grid row if needed
            if (!grid[rowIndex]) {
                grid[rowIndex] = [];
            }

            // Store cell value
            grid[rowIndex][colIndex] = cellValue;

            // Mark merged cells
            if (rowspan > 1 || colspan > 1) {
                const cellAddress = getCellNameFromCoords(colIndex, rowIndex);
                mergeCells[cellAddress] = [colspan, rowspan];
            }

            // Mark occupied cells
            for (let r = 0; r < rowspan; r++) {
                for (let c = 0; c < colspan; c++) {
                    if (r === 0 && c === 0) continue; // Skip the main cell
                    occupied.set(`${rowIndex + r},${colIndex + c}`, true);

                    // Initialize grid rows for spanned cells
                    if (!grid[rowIndex + r]) {
                        grid[rowIndex + r] = [];
                    }
                    grid[rowIndex + r][colIndex + c] = '';
                }
            }

            // Parse inline styles
            const style = getProp(cellNode, 'style');
            if (style) {
                const cellAddress = getCellNameFromCoords(colIndex, rowIndex);
                styles[cellAddress] = style;
            }

            // Check for title attribute (can be used as comment)
            const title = getProp(cellNode, 'title');
            if (title) {
                const cellAddress = getCellNameFromCoords(colIndex, rowIndex);
                comments[cellAddress] = decodeEntities(title);
            }

            // Update max columns
            if (colIndex + colspan > maxCols) {
                maxCols = colIndex + colspan;
            }

            colIndex += colspan;
        });
    });

    // Convert grid to 2D array
    const data = [];
    for (let r = 0; r < grid.length; r++) {
        const row = [];
        for (let c = 0; c < maxCols; c++) {
            row.push(grid[r] && grid[r][c] !== undefined ? grid[r][c] : '');
        }
        data.push(row);
    }

    // Create column definitions
    const columns = [];

    // If first row is header, use those values as column titles
    if (firstRowAsHeader && data.length > 0) {
        for (let i = 0; i < maxCols; i++) {
            columns.push({
                title: data[0][i] || getColumnName(i),
                width: 100
            });
        }
        // Remove the header row from data
        data.shift();

        // Adjust mergeCells and other references after removing first row
        const adjustedMergeCells = {};
        const adjustedStyles = {};
        const adjustedComments = {};

        for (const [addr, value] of Object.entries(mergeCells)) {
            const match = addr.match(/^([A-Z]+)(\d+)$/);
            if (match) {
                const col = match[1];
                const row = parseInt(match[2]);
                if (row > 1) {
                    adjustedMergeCells[`${col}${row - 1}`] = value;
                }
            }
        }

        for (const [addr, value] of Object.entries(styles)) {
            const match = addr.match(/^([A-Z]+)(\d+)$/);
            if (match) {
                const col = match[1];
                const row = parseInt(match[2]);
                if (row > 1) {
                    adjustedStyles[`${col}${row - 1}`] = value;
                }
            }
        }

        for (const [addr, value] of Object.entries(comments)) {
            const match = addr.match(/^([A-Z]+)(\d+)$/);
            if (match) {
                const col = match[1];
                const row = parseInt(match[2]);
                if (row > 1) {
                    adjustedComments[`${col}${row - 1}`] = value;
                }
            }
        }

        Object.keys(mergeCells).forEach(key => delete mergeCells[key]);
        Object.keys(styles).forEach(key => delete styles[key]);
        Object.keys(comments).forEach(key => delete comments[key]);
        Object.assign(mergeCells, adjustedMergeCells);
        Object.assign(styles, adjustedStyles);
        Object.assign(comments, adjustedComments);
    } else {
        // Use default column letters
        for (let i = 0; i < maxCols; i++) {
            columns.push({
                title: getColumnName(i),
                width: 100
            });
        }
    }

    // Build worksheet result
    const worksheet = {
        data,
        columns
    };

    // Add mergeCells if any
    if (Object.keys(mergeCells).length > 0) {
        worksheet.mergeCells = mergeCells;
    }

    // Add styles if any
    if (Object.keys(styles).length > 0) {
        worksheet.style = styles;
    }

    // Add comments if any
    if (Object.keys(comments).length > 0) {
        worksheet.comments = comments;
    }

    return worksheet;
}

/**
 * Parse HTML table - works in both Browser and Node.js
 * @param {string|HTMLElement|File|Uint8Array} input - HTML string, DOM element, File path (Node.js), File object (Browser), or buffer
 * @param {object} options - Parsing options
 * @param {number} options.tableIndex - Which table to parse (0-based), default: 0
 * @param {boolean} options.firstRowAsHeader - Treat first row as column headers, default: true
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseHTMLTable(input, options = {}) {
    const {
        tableIndex = 0,
        firstRowAsHeader = true
    } = options;

    return parse(async (inp) => {
        const fileContent = await readHTML(inp);

        // Parse HTML to JSON
        const tree = parser(fileContent, { ignore: ['script', 'style'] });

        if (!tree) {
            throw new Error('Failed to parse HTML');
        }

        // Find all table nodes
        const tables = findAllTables(tree);

        if (tables.length === 0) {
            throw new Error('No table elements found in HTML');
        }

        if (tableIndex >= tables.length) {
            throw new Error(`Table index ${tableIndex} out of range. Found ${tables.length} tables.`);
        }

        // Parse the selected table
        const table = tables[tableIndex];
        const worksheet = parseTableNode(table, { firstRowAsHeader });

        return {
            worksheets: [{
                worksheetName: `Table${tableIndex + 1}`,
                ...worksheet
            }]
        };
    }, input, options);
}

/**
 * Parse all tables from an HTML file - works in both Browser and Node.js
 * @param {string|HTMLElement|File|Uint8Array} input - HTML string, DOM element, File path (Node.js), File object (Browser), or buffer
 * @param {object} options - Parsing options
 * @param {boolean} options.firstRowAsHeader - Treat first row as column headers, default: true
 * @returns {Promise<object>} Jspreadsheet Pro format with multiple worksheets
 */
export async function parseAllHTMLTables(input, options = {}) {
    return parse(async (inp) => {
        const fileContent = await readHTML(inp);

        // Parse HTML to JSON
        const tree = parser(fileContent, { ignore: ['script', 'style'] });

        if (!tree) {
            throw new Error('Failed to parse HTML');
        }

        // Find all table nodes
        const tables = findAllTables(tree);

        if (tables.length === 0) {
            throw new Error('No table elements found in HTML');
        }

        // Parse all tables
        const worksheets = tables.map((table, index) => {
            const worksheet = parseTableNode(table, options);
            return {
                worksheetName: `Table${index + 1}`,
                ...worksheet
            };
        });

        return { worksheets };
    }, input, options);
}
