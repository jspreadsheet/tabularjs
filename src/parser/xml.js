import { loadAsString, parse } from '../utils/loader.js';
import { parser } from '@lemonadejs/html-to-json';
import { getColumnName, getCellNameFromCoords, convertR1C1toA1, getProp, getChildren, getTextContent } from '../utils/helpers.js';

/**
 * Recursively find all nodes of a specific type in the tree
 */
function findAllNodes(node, type, results = []) {
    if (!node) return results;

    // Check if this node matches
    if (node.type === type || node.type.endsWith(':' + type)) {
        results.push(node);
    }

    // Recursively search children
    if (node.children) {
        node.children.forEach(child => findAllNodes(child, type, results));
    }

    return results;
}

/**
 * Get first child of a specific type
 */
function getChild(node, type) {
    const children = getChildren(node, type);
    return children.length > 0 ? children[0] : null;
}

/**
 * Parse styles from the Styles section
 */
function parseStyles(tree) {
    const styles = {};
    const styleNodes = findAllNodes(tree, 'Style');

    styleNodes.forEach(styleNode => {
        const styleId = getProp(styleNode, 'ID');
        if (!styleId) return;

        const style = {};

        // Parse Font
        const fontNode = getChild(styleNode, 'Font');
        if (fontNode) {
            const fontName = getProp(fontNode, 'FontName');
            const fontSize = getProp(fontNode, 'Size');
            const fontColor = getProp(fontNode, 'Color');
            const bold = getProp(fontNode, 'Bold');
            const italic = getProp(fontNode, 'Italic');
            const underline = getProp(fontNode, 'Underline');

            style.font = {
                name: fontName,
                size: fontSize,
                color: fontColor,
                bold: bold === '1',
                italic: italic === '1',
                underline: underline !== undefined
            };
        }

        // Parse Interior (background)
        const interiorNode = getChild(styleNode, 'Interior');
        if (interiorNode) {
            const color = getProp(interiorNode, 'Color');
            const pattern = getProp(interiorNode, 'Pattern');
            if (color) {
                style.interior = { color, pattern };
            }
        }

        // Parse Borders
        const bordersNode = getChild(styleNode, 'Borders');
        if (bordersNode) {
            const borderChildren = getChildren(bordersNode, 'Border');
            if (borderChildren.length > 0) {
                style.borders = {};
                borderChildren.forEach(border => {
                    const position = getProp(border, 'Position');
                    const lineStyle = getProp(border, 'LineStyle');
                    const weight = getProp(border, 'Weight');
                    const color = getProp(border, 'Color');

                    if (position) {
                        style.borders[position.toLowerCase()] = {
                            style: lineStyle || 'Continuous', // Default to Continuous if not specified
                            weight: weight,
                            color: color
                        };
                    }
                });
            }
        }

        // Parse Alignment
        const alignmentNode = getChild(styleNode, 'Alignment');
        if (alignmentNode) {
            const horizontal = getProp(alignmentNode, 'Horizontal');
            const vertical = getProp(alignmentNode, 'Vertical');
            const wrapText = getProp(alignmentNode, 'WrapText');

            style.alignment = {
                horizontal: horizontal,
                vertical: vertical,
                wrapText: wrapText === '1'
            };
        }

        // Parse NumberFormat
        const numberFormatNode = getChild(styleNode, 'NumberFormat');
        if (numberFormatNode) {
            const format = getProp(numberFormatNode, 'Format');
            if (format) {
                style.numberFormat = format;
            }
        }

        styles[styleId] = style;
    });

    return styles;
}

/**
 * Convert Excel number format to jspreadsheet mask
 * This is a basic converter - some formats may need manual adjustment
 */
function convertExcelFormatToMask(format) {
    if (!format) return null;

    // Remove Excel-specific codes and convert to simpler format
    let mask = format;

    // Common Excel format conversions
    const conversions = {
        // Date formats
        'yyyy-mm-dd': 'YYYY-MM-DD',
        'dd/mm/yyyy': 'DD/MM/YYYY',
        'mm/dd/yyyy': 'MM/DD/YYYY',
        'dd-mmm-yy': 'DD-MMM-YY',
        'dddd, mmmm dd, yyyy': 'dddd, MMMM DD, YYYY',

        // Currency
        '[$€-407]#,##0.00': '#,##0.00',
        '[$£-809]#,##0.00': '#,##0.00',
        '[$¥-411]#,##0': '#,##0',

        // Percentage
        '0.00%': '0.00%',
        '0%': '0%',

        // Number
        '#,##0.00': '#,##0.00',
        '#,##0': '#,##0',
        '0.00': '0.00',
        '0': '0',

        // Scientific notation
        '0.00E+00': '0.00E+00',
        '0.000E+00': '0.000E+00',
        '0.0000E+00': '0.0000E+00',
        '0.00000E+00': '0.00000E+00'
    };

    // Try direct conversion
    if (conversions[format]) {
        return conversions[format];
    }

    // Check for scientific notation pattern (e.g., 0.000000000000000000000E+00)
    if (/^0\.0+E\+00$/i.test(format)) {
        return format.toUpperCase(); // Normalize to uppercase E
    }

    // Clean up Excel-specific syntax
    mask = mask
        .replace(/\[\$[^\]]+\]/g, '') // Remove currency codes like [$€-407]
        .replace(/_;/g, '') // Remove underscores
        .replace(/\\/g, '') // Remove escape characters
        .replace(/"/g, '') // Remove quotes
        .trim();

    // Check if cleaned mask is scientific notation
    if (/^0\.0+E\+00$/i.test(mask)) {
        return mask.toUpperCase();
    }

    // If it's a simple format, return it
    if (mask) {
        return mask;
    }

    return null;
}

/**
 * Convert style object to CSS string
 */
function styleToCSS(style) {
    if (!style) return '';

    const css = [];

    // Font
    if (style.font) {
        if (style.font.name) css.push(`font-family: ${style.font.name}`);
        if (style.font.size) css.push(`font-size: ${style.font.size}pt`);
        if (style.font.color) css.push(`color: ${style.font.color}`);
        if (style.font.bold) css.push('font-weight: bold');
        if (style.font.italic) css.push('font-style: italic');
        if (style.font.underline) css.push('text-decoration: underline');
    }

    // Interior (background)
    if (style.interior && style.interior.color) {
        css.push(`background-color: ${style.interior.color}`);
    }

    // Borders
    if (style.borders) {
        // Map SpreadsheetML border styles to CSS
        const styleMap = {
            'None': null,           // No border
            'Continuous': 'solid',
            'Dash': 'dashed',
            'Dot': 'dotted',
            'DashDot': 'dashed',
            'DashDotDot': 'dashed',
            'SlantDashDot': 'dashed',
            'Double': 'double'
        };

        ['left', 'right', 'top', 'bottom', 'diagonal'].forEach(position => {
            const border = style.borders[position];
            if (border) {
                // Skip if style is 'None'
                const cssStyle = styleMap[border.style];
                if (cssStyle === null) return;

                // Default to solid if unknown style
                const borderStyle = cssStyle || 'solid';

                // Weight: 0=hairline, 1=thin, 2=medium, 3=thick
                const weight = border.weight ? parseInt(border.weight) : 1;
                const width = weight === 0 ? '1' : weight;

                const color = border.color || '#000000';
                css.push(`border-${position}: ${width}px ${borderStyle} ${color}`);
            }
        });
    }

    // Alignment
    if (style.alignment) {
        if (style.alignment.horizontal) {
            const align = style.alignment.horizontal.toLowerCase();
            css.push(`text-align: ${align}`);
        }
        if (style.alignment.vertical) {
            const vAlign = {
                'top': 'top',
                'center': 'middle',
                'bottom': 'bottom'
            };
            const align = vAlign[style.alignment.vertical.toLowerCase()] || 'middle';
            css.push(`vertical-align: ${align}`);
        }
        if (style.alignment.wrapText) {
            css.push('white-space: pre-wrap');
        }
    }

    return css.join('; ');
}

/**
 * Parse XML Spreadsheet from string content
 */
export function parseXMLSpreadsheetString(fileContent) {
    // Parse XML to JSON using html-to-json parser
    const tree = parser(fileContent, { ignore: [] });

    if (!tree) {
        throw new Error('Failed to parse XML file');
    }

    // Parse styles
    const styles = parseStyles(tree);

    // Find all Worksheet nodes recursively in the entire tree
    // This handles cases where the XML structure may have unexpected nesting
    const worksheetNodes = findAllNodes(tree, 'Worksheet');

    if (worksheetNodes.length === 0) {
        throw new Error('No Worksheet elements found in XML');
    }

    // Parse each worksheet
    const worksheets = worksheetNodes.map(worksheetNode => {
        const worksheetName = getProp(worksheetNode, 'Name') || 'Sheet1';

        // Get the Table node
        const tableNode = getChild(worksheetNode, 'Table');

        if (!tableNode) {
            return {
                name: worksheetName,
                data: [],
                columns: []
            };
        }

        // Parse column definitions
        const columnNodes = getChildren(tableNode, 'Column');
        const columns = [];
        let columnIndex = 0;

        columnNodes.forEach(colNode => {
            const width = getProp(colNode, 'Width');
            const span = getProp(colNode, 'Span');
            const index = getProp(colNode, 'Index');
            const hidden = getProp(colNode, 'Hidden');
            const styleId = getProp(colNode, 'StyleID');

            // If Index is specified, fill gaps with default columns
            if (index) {
                const targetIndex = parseInt(index) - 1; // Convert to 0-based
                while (columnIndex < targetIndex) {
                    columns.push({
                        title: getColumnName(columnIndex),
                        width: 100
                    });
                    columnIndex++;
                }
            }

            // Add the column
            const column = {
                title: getColumnName(columnIndex)
            };

            if (width) {
                // XML Width is already in points, use directly as jspreadsheet expects similar values
                column.width = parseFloat(width);
            }

            if (hidden === '1') {
                column.visible = false;
            }

            // Store styleId for later application to cells
            if (styleId) {
                column.styleId = styleId;
            }

            columns.push(column);
            columnIndex++;

            // Handle span (repeat column definition)
            if (span) {
                const spanCount = parseInt(span);
                for (let i = 0; i < spanCount; i++) {
                    const spanColumn = {
                        title: getColumnName(columnIndex),
                        width: width ? parseFloat(width) : 100
                    };
                    if (hidden === '1') {
                        spanColumn.visible = false;
                    }
                    // Apply same styleId to spanned columns
                    if (styleId) {
                        spanColumn.styleId = styleId;
                    }
                    columns.push(spanColumn);
                    columnIndex++;
                }
            }
        });

        // Parse rows
        const rowNodes = getChildren(tableNode, 'Row');
        const data = [];
        const rows = []; // Row properties (height, visibility, etc.)
        const mergeCells = {};
        const comments = {};
        const cellStyles = {}; // Cell address -> CSS string
        const cells = {}; // Cell metadata (formula, mask, etc.)
        let rowIndex = 0;
        let maxColumns = columns.length;

        rowNodes.forEach(rowNode => {
            const index = getProp(rowNode, 'Index');

            // If Index is specified, fill gaps with empty rows
            if (index) {
                const targetIndex = parseInt(index) - 1; // Convert to 0-based
                while (rowIndex < targetIndex) {
                    data.push([]);
                    rows.push({}); // Add empty row properties
                    rowIndex++;
                }
            }

            // Parse row properties
            const rowProps = {};
            const height = getProp(rowNode, 'Height');
            const hidden = getProp(rowNode, 'Hidden');

            if (height) {
                rowProps.height = parseFloat(height);
            }
            if (hidden === '1') {
                rowProps.visible = false;
            }

            rows.push(rowProps);

            // Parse cells in this row
            const cellNodes = getChildren(rowNode, 'Cell');
            const row = [];
            let cellIndex = 0;

            cellNodes.forEach(cellNode => {
                const index = getProp(cellNode, 'Index');

                // If Index is specified, fill gaps with empty cells
                if (index) {
                    const targetIndex = parseInt(index) - 1; // Convert to 0-based
                    while (cellIndex < targetIndex) {
                        row.push('');
                        cellIndex++;
                    }
                }

                // Get cell formula
                let formula = getProp(cellNode, 'Formula');

                // Get cell data
                const dataNode = getChild(cellNode, 'Data');
                let value = '';

                if (dataNode) {
                    const dataType = getProp(dataNode, 'Type');

                    // Get text content from Data node
                    if (dataNode.children && dataNode.children.length > 0) {
                        const textNode = dataNode.children.find(child => child.type === '#text');
                        if (textNode) {
                            value = getTextContent(textNode);
                        }
                    }

                    // Convert value based on type
                    if (dataType === 'Number') {
                        const numValue = parseFloat(value);
                        value = isNaN(numValue) ? value : numValue;
                    } else if (dataType === 'Boolean') {
                        value = value === '1' || value.toLowerCase() === 'true';
                    }
                }

                // If there's a formula, convert from R1C1 to A1 notation
                // Formula takes precedence over value
                if (formula) {
                    // Convert R1C1 notation to A1
                    formula = convertR1C1toA1(formula, rowIndex, cellIndex);
                    // Add = prefix if not present (Jspreadsheet format)
                    value = formula.startsWith('=') ? formula : '=' + formula;
                }

                row.push(value);

                // Get current cell address BEFORE any index manipulation
                const currentCellAddress = getCellNameFromCoords(cellIndex, rowIndex);

                // Handle MergeAcross (colspan) and MergeDown (rowspan)
                const mergeAcross = getProp(cellNode, 'MergeAcross');
                const mergeDown = getProp(cellNode, 'MergeDown');

                if (mergeAcross || mergeDown) {
                    const colspan = mergeAcross ? parseInt(mergeAcross) + 1 : 1;
                    const rowspan = mergeDown ? parseInt(mergeDown) + 1 : 1;
                    mergeCells[currentCellAddress] = [colspan, rowspan];
                }

                // Check for comment (use current cell address)
                const commentNode = getChild(cellNode, 'Comment');
                if (commentNode) {
                    const commentData = getChild(commentNode, 'Data');
                    if (commentData && commentData.children) {
                        const textNode = commentData.children.find(child => child.type === '#text');
                        if (textNode) {
                            const commentText = getTextContent(textNode);
                            comments[currentCellAddress] = commentText;
                        }
                    }
                }

                // Apply cell style if StyleID is present (use current cell address)
                const styleId = getProp(cellNode, 'StyleID');
                if (styleId && styles[styleId]) {
                    const style = styles[styleId];

                    // Add CSS styles
                    const cssStyle = styleToCSS(style);
                    if (cssStyle) {
                        cellStyles[currentCellAddress] = cssStyle;
                    }

                    // Add number format mask if present
                    if (style.numberFormat) {
                        const mask = convertExcelFormatToMask(style.numberFormat);
                        if (mask) {
                            if (!cells[currentCellAddress]) {
                                cells[currentCellAddress] = {};
                            }
                            cells[currentCellAddress].mask = mask;
                        }
                    }
                }

                // Increment cell index
                cellIndex++;

                // Handle MergeAcross by adding empty cells AFTER style is applied
                if (mergeAcross) {
                    const mergeCount = parseInt(mergeAcross);
                    for (let i = 0; i < mergeCount; i++) {
                        row.push('');
                        cellIndex++;
                    }
                }
            });

            // Update max columns if this row is wider
            if (row.length > maxColumns) {
                maxColumns = row.length;
            }

            data.push(row);
            rowIndex++;
        });

        // Ensure all rows have the same number of columns
        data.forEach(row => {
            while (row.length < maxColumns) {
                row.push('');
            }
        });

        // Ensure we have enough column definitions
        while (columns.length < maxColumns) {
            columns.push({
                title: getColumnName(columns.length),
                width: 100
            });
        }

        // Parse worksheet-level properties
        const worksheetVisible = getProp(worksheetNode, 'Visible');

        // Parse default dimensions from Table node
        const defaultColumnWidth = getProp(tableNode, 'DefaultColumnWidth');
        const defaultRowHeight = getProp(tableNode, 'DefaultRowHeight');

        // Parse WorksheetOptions for various sheet settings
        const worksheetOptionsNode = getChild(worksheetNode, 'WorksheetOptions');
        let freezeColumns = null;
        let freezeRows = null;
        let protectObjects = null;
        let protectScenarios = null;
        let pageSetup = null;
        let visibleState = null;
        let displayGridlines = true; // default is true
        let panes = null;
        let enableSelection = null;

        if (worksheetOptionsNode) {
            // Parse freeze panes
            const splitNode = getChild(worksheetOptionsNode, 'SplitHorizontal');
            const splitVertical = getChild(worksheetOptionsNode, 'SplitVertical');
            const freezePanesNode = getChild(worksheetOptionsNode, 'FreezePanes');

            // If FreezePanes is present, parse split positions
            if (freezePanesNode) {
                if (splitNode) {
                    const splitValue = getTextContent(splitNode.children?.[0]);
                    if (splitValue) {
                        freezeRows = parseInt(splitValue);
                    }
                }
                if (splitVertical) {
                    const splitValue = getTextContent(splitVertical.children?.[0]);
                    if (splitValue) {
                        freezeColumns = parseInt(splitValue);
                    }
                }
            }

            // Parse protection settings
            const protectObjectsNode = getChild(worksheetOptionsNode, 'ProtectObjects');
            if (protectObjectsNode && protectObjectsNode.children?.[0]) {
                const value = getTextContent(protectObjectsNode.children[0]);
                protectObjects = value.toLowerCase() === 'true';
            }

            const protectScenariosNode = getChild(worksheetOptionsNode, 'ProtectScenarios');
            if (protectScenariosNode && protectScenariosNode.children?.[0]) {
                const value = getTextContent(protectScenariosNode.children[0]);
                protectScenarios = value.toLowerCase() === 'true';
            }

            // Parse visibility state from WorksheetOptions
            const visibleNode = getChild(worksheetOptionsNode, 'Visible');
            if (visibleNode && visibleNode.children?.[0]) {
                visibleState = getTextContent(visibleNode.children[0]);
            }

            // Parse gridlines display
            const noGridlinesNode = getChild(worksheetOptionsNode, 'DoNotDisplayGridlines');
            if (noGridlinesNode) {
                displayGridlines = false;
            }

            // Parse EnableSelection
            const enableSelectionNode = getChild(worksheetOptionsNode, 'EnableSelection');
            if (enableSelectionNode && enableSelectionNode.children?.[0]) {
                enableSelection = getTextContent(enableSelectionNode.children[0]);
            }

            // Parse PageSetup
            const pageSetupNode = getChild(worksheetOptionsNode, 'PageSetup');
            if (pageSetupNode) {
                pageSetup = {};

                const headerNode = getChild(pageSetupNode, 'Header');
                if (headerNode) {
                    const margin = getProp(headerNode, 'Margin');
                    if (margin) pageSetup.headerMargin = parseFloat(margin);
                }

                const footerNode = getChild(pageSetupNode, 'Footer');
                if (footerNode) {
                    const margin = getProp(footerNode, 'Margin');
                    if (margin) pageSetup.footerMargin = parseFloat(margin);
                }

                const marginsNode = getChild(pageSetupNode, 'PageMargins');
                if (marginsNode) {
                    pageSetup.margins = {};
                    const bottom = getProp(marginsNode, 'Bottom');
                    const left = getProp(marginsNode, 'Left');
                    const right = getProp(marginsNode, 'Right');
                    const top = getProp(marginsNode, 'Top');

                    if (bottom) pageSetup.margins.bottom = parseFloat(bottom);
                    if (left) pageSetup.margins.left = parseFloat(left);
                    if (right) pageSetup.margins.right = parseFloat(right);
                    if (top) pageSetup.margins.top = parseFloat(top);
                }
            }

            // Parse Panes (for split window info)
            // We'll use this to extract freeze position if not already set
            const panesNode = getChild(worksheetOptionsNode, 'Panes');
            if (panesNode) {
                const paneNodes = getChildren(panesNode, 'Pane');

                // In Excel, pane numbers indicate split/freeze state:
                // 0 = top-left, 1 = top-right, 2 = bottom-left, 3 = bottom-right
                // Multiple panes typically means there's a freeze or split
                if (paneNodes.length > 1 && (freezeRows === null || freezeColumns === null)) {
                    // Try to infer freeze position from active pane
                    const activePane = paneNodes.find(p => {
                        const numNode = getChild(p, 'ActivePane');
                        return numNode !== null;
                    }) || paneNodes[0];

                    const numberNode = getChild(activePane, 'Number');
                    if (numberNode && numberNode.children?.[0]) {
                        const paneNumber = parseInt(getTextContent(numberNode.children[0]));

                        // Pane 3 (bottom-right) suggests both row and column freeze
                        // Pane 1 (top-right) suggests column freeze only
                        // Pane 2 (bottom-left) suggests row freeze only
                        if (paneNumber === 3 || paneNumber === 1) {
                            // Has vertical split (frozen columns)
                            if (freezeColumns === null) {
                                // We can't determine exact count without more info,
                                // but we know there's a freeze
                                // Keep it null and let explicit split values handle it
                            }
                        }
                        if (paneNumber === 3 || paneNumber === 2) {
                            // Has horizontal split (frozen rows)
                            if (freezeRows === null) {
                                // Keep it null and let explicit split values handle it
                            }
                        }
                    }
                }
            }
        }

        // Build worksheet result
        const worksheet = {
            worksheetName: worksheetName,
            data,
            columns
        };

        // Add rows if any have properties
        if (rows.length > 0 && rows.some(r => Object.keys(r).length > 0)) {
            worksheet.rows = rows;
        }

        // Add cells if any have metadata (formulas, masks, etc.)
        if (Object.keys(cells).length > 0) {
            worksheet.cells = cells;
        }

        // Add mergeCells if any
        if (Object.keys(mergeCells).length > 0) {
            worksheet.mergeCells = mergeCells;
        }

        // Add comments if any
        if (Object.keys(comments).length > 0) {
            worksheet.comments = comments;
        }

        // Add cell styles if any
        if (Object.keys(cellStyles).length > 0) {
            worksheet.style = cellStyles;
        }

        // Add worksheet state (hidden)
        // Check both ss:Visible attribute and WorksheetOptions/Visible element
        if (worksheetVisible === '0' || visibleState === 'SheetHidden' || visibleState === 'SheetVeryHidden') {
            worksheet.worksheetState = visibleState === 'SheetVeryHidden' ? 'veryHidden' : 'hidden';
        }

        // Add default dimensions
        if (defaultColumnWidth) {
            worksheet.defaultColWidth = parseFloat(defaultColumnWidth);
        }
        if (defaultRowHeight) {
            worksheet.defaultRowHeight = parseFloat(defaultRowHeight);
        }

        // Add freeze panes
        if (freezeColumns !== null && freezeColumns > 0) {
            worksheet.freezeColumns = freezeColumns;
        }
        if (freezeRows !== null && freezeRows > 0) {
            worksheet.freezeRows = freezeRows;
        }

        // Add protection settings
        if (protectObjects !== null || protectScenarios !== null || enableSelection !== null) {
            worksheet.protection = {};
            if (protectObjects !== null) {
                worksheet.protection.protectObjects = protectObjects;
            }
            if (protectScenarios !== null) {
                worksheet.protection.protectScenarios = protectScenarios;
            }
            if (enableSelection !== null) {
                worksheet.protection.enableSelection = enableSelection;
            }
        }

        // Add display settings
        if (!displayGridlines) {
            worksheet.displayGridlines = false;
        }

        // Note: pageSetup and panes are parsed but not added to output
        // pageSetup: primarily for printing, not relevant for Jspreadsheet Pro
        // panes: used internally to extract freeze info, but not needed in output

        return worksheet;
    });

    // Parse defined names (named ranges) at workbook level
    const definedNames = {};
    const namesNodes = findAllNodes(tree, 'Names');

    namesNodes.forEach(namesNode => {
        const namedRangeNodes = getChildren(namesNode, 'NamedRange');
        namedRangeNodes.forEach(namedRange => {
            let name = getProp(namedRange, 'Name');
            const refersTo = getProp(namedRange, 'RefersTo');

            if (name && refersTo) {
                // Remove all leading underscores and numeric prefixes
                // Handles: _1_, _2_, __, ___, _1____, etc.
                name = name.replace(/^_+\d*_*/g, '');

                // Remove Excel internal prefixes (with or without leading underscore)
                // Handles: _XLNM., XLNM., _xlnm., xlnm., etc.
                const prefixPatterns = [
                    /^_*xlnm[._]/i,    // _XLNM., XLNM., _xlnm_, xlnm_, etc.
                    /^_*xleta[._]/i,   // _XLETA., XLETA., etc.
                    /^_*xlfn[._]/i,    // _xlfn., xlfn., etc.
                    /^_*xll[._]/i,     // _xll., xll., etc.
                    /^_*xlws[._]/i,    // _xlws., xlws., etc.
                    /^_*xlpm[._]/i     // _xlpm., xlpm., etc.
                ];

                prefixPatterns.forEach(pattern => {
                    name = name.replace(pattern, '');
                });

                // Convert R1C1 reference to A1 notation if needed
                // RefersTo typically looks like "=Sheet1!R1C1:R10C5" or "=Sheet1!A1:E10"
                let cleanedRef = refersTo;

                // Remove leading '=' if present
                if (cleanedRef.startsWith('=')) {
                    cleanedRef = cleanedRef.substring(1);
                }

                // Remove Excel internal prefixes from the reference formula as well
                prefixPatterns.forEach(pattern => {
                    cleanedRef = cleanedRef.replace(new RegExp(pattern.source, 'gi'), '');
                });

                definedNames[name] = cleanedRef;
            }
        });
    });

    // Return in Jspreadsheet Pro format
    const result = {
        worksheets: worksheets
    };

    // Add defined names (named ranges)
    if (Object.keys(definedNames).length > 0) {
        result.definedNames = definedNames;
    }

    return result;
}

/**
 * Parse XML Spreadsheet file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseXMLSpreadsheet(input, options = {}) {
    return parse(async (inp) => {
        const content = await loadAsString(inp);
        return parseXMLSpreadsheetString(content);
    }, input, options);
}
