import { loadAsString, parse } from '../utils/loader.js';
import { parser } from '@lemonadejs/html-to-json';
import { getColumnName, getCellNameFromCoords, getProp, getChildren, getTextContent } from '../utils/helpers.js';

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

/**
 * Parse a single table node from HTML tree
 */
function parseTableNode(tableNode, options) {
    const { firstRowAsHeader = true } = options;

    // Get thead, tbody, tfoot sections
    const thead = getChildren(tableNode, 'thead')[0];
    const tbody = getChildren(tableNode, 'tbody')[0];
    const tfoot = getChildren(tableNode, 'tfoot')[0];

    // Collect all rows in order: thead, tbody, tfoot
    let allRows = [];

    if (thead) {
        allRows.push(...getChildren(thead, 'tr'));
    }

    if (tbody) {
        allRows.push(...getChildren(tbody, 'tr'));
    } else {
        // If no tbody, get tr directly from table
        allRows.push(...getChildren(tableNode, 'tr'));
    }

    if (tfoot) {
        allRows.push(...getChildren(tfoot, 'tr'));
    }

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
        const cells = [
            ...getChildren(rowNode, 'th'),
            ...getChildren(rowNode, 'td')
        ];

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
                cellValue = formula.startsWith('=') ? formula : '=' + formula;
            } else {
                cellValue = getTextContent(cellNode).trim();
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
                comments[cellAddress] = title;
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
        let fileContent;

        // Check if input is a DOM element
        if (typeof HTMLElement !== 'undefined' && inp instanceof HTMLElement) {
            fileContent = inp.outerHTML;
        }
        // Check if input is an HTML table string
        else if (typeof inp === 'string' && inp.trim().toLowerCase().startsWith('<table')) {
            fileContent = inp;
        }
        // Otherwise, load as file
        else {
            fileContent = await loadAsString(inp);
        }

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
            worksheets: [worksheet]
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
        let fileContent;

        // Check if input is a DOM element
        if (typeof HTMLElement !== 'undefined' && inp instanceof HTMLElement) {
            fileContent = inp.outerHTML;
        }
        // Check if input is an HTML table string
        else if (typeof inp === 'string' && inp.trim().toLowerCase().startsWith('<table')) {
            fileContent = inp;
        }
        // Otherwise, load as file
        else {
            fileContent = await loadAsString(inp);
        }

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
