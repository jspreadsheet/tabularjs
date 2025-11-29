import { loadAsBuffer, parse } from '../utils/loader.js';
import JSZip from 'jszip';
import { parser } from '@lemonadejs/html-to-json';
import {
    getColumnName,
    getCellNameFromCoords,
    getCoordsFromCellName,
    getTextContent,
    parseAttributes,
    convertWidthToPixels
} from '../utils/helpers.js';

// Parse number format from number style definition
function parseNumberFormat(node) {
    if (!node || !node.children) return null;

    const attrs = parseAttributes(node);
    const styleName = attrs['style:name'] || attrs['number:style-name'];

    // Look for number:number, number:text, etc.
    let decimalPlaces = 0;
    let minDecimalPlaces = 0;
    let minIntegerDigits = 1;
    let isPercentage = node.type === 'number:percentage-style';
    let isCurrency = node.type === 'number:currency-style';
    let currencySymbol = '';

    const findNumberProps = (n) => {
        if (!n) return;
        if (Array.isArray(n)) {
            n.forEach(findNumberProps);
            return;
        }
        if (typeof n === 'object') {
            if (n.type === 'number:number') {
                const nattrs = parseAttributes(n);
                if (nattrs['number:decimal-places']) {
                    decimalPlaces = parseInt(nattrs['number:decimal-places']) || 0;
                }
                if (nattrs['number:min-decimal-places']) {
                    minDecimalPlaces = parseInt(nattrs['number:min-decimal-places']) || 0;
                }
                if (nattrs['number:min-integer-digits']) {
                    minIntegerDigits = parseInt(nattrs['number:min-integer-digits']) || 1;
                }
            } else if (n.type === 'number:currency-symbol') {
                currencySymbol = getTextContent(n) || '$';
            }
            if (n.children) findNumberProps(n.children);
        }
    };

    findNumberProps(node);

    // Build Excel-compatible format code
    let mask = '';

    if (isCurrency) {
        mask = currencySymbol + '#,##0';
        if (decimalPlaces > 0) {
            mask += '.' + '0'.repeat(decimalPlaces);
        }
    } else if (isPercentage) {
        mask = '0';
        if (decimalPlaces > 0) {
            mask += '.' + '0'.repeat(decimalPlaces);
        }
        mask += '%';
    } else {
        // Regular number format
        if (decimalPlaces > 0 || minDecimalPlaces > 0) {
            const decimals = Math.max(decimalPlaces, minDecimalPlaces);
            mask = '0.' + '0'.repeat(decimals);
        } else {
            mask = '0';
        }
    }

    return mask;
}

// Parse number format definitions from XML
function parseNumberFormats(xml) {
    const formats = {};

    if (!xml) return formats;

    const parsed = parser(xml);

    const findFormats = (node) => {
        if (!node) return;

        if (Array.isArray(node)) {
            node.forEach(findFormats);
            return;
        }

        if (typeof node === 'object') {
            // Check for number format style types
            if (node.type === 'number:number-style' ||
                node.type === 'number:percentage-style' ||
                node.type === 'number:currency-style' ||
                node.type === 'number:date-style' ||
                node.type === 'number:time-style') {

                const attrs = parseAttributes(node);
                const name = attrs['style:name'];

                if (name) {
                    const mask = parseNumberFormat(node);
                    if (mask) {
                        formats[name] = mask;
                    }
                }
            }

            if (node.children) {
                findFormats(node.children);
            }
        }
    };

    findFormats(parsed);
    return formats;
}

// Parse style definitions from styles.xml
function parseStyles(stylesXml) {
    const styles = {};

    if (!stylesXml) return styles;

    const parsed = parser(stylesXml);

    // Find style:style elements
    const findStyles = (node) => {
        if (!node) return;

        if (Array.isArray(node)) {
            node.forEach(findStyles);
            return;
        }

        if (typeof node === 'object') {
            if (node.type === 'style:style') {
                const attrs = parseAttributes(node);
                const styleName = attrs['style:name'];

                if (styleName) {
                    styles[styleName] = {
                        family: attrs['style:family'],
                        parentStyleName: attrs['style:parent-style-name'],
                        dataStyleName: attrs['style:data-style-name'] // Store number format reference
                    };

                    // Parse style properties
                    if (node.children) {
                        node.children.forEach(child => {
                            if (child.type === 'style:text-properties') {
                                styles[styleName].text = parseAttributes(child);
                            } else if (child.type === 'style:paragraph-properties') {
                                styles[styleName].paragraph = parseAttributes(child);
                            } else if (child.type === 'style:table-cell-properties') {
                                styles[styleName].cell = parseAttributes(child);
                            } else if (child.type === 'style:table-column-properties') {
                                styles[styleName].column = parseAttributes(child);
                            } else if (child.type === 'style:table-row-properties') {
                                styles[styleName].row = parseAttributes(child);
                            }
                        });
                    }
                }
            }

            if (node.children) {
                findStyles(node.children);
            }
        }
    };

    findStyles(parsed);
    return styles;
}

// Parse table style definitions (for worksheet visibility)
function parseTableStyles(xml) {
    const tableStyles = {};

    if (!xml) return tableStyles;

    const parsed = parser(xml);

    const findTableStyles = (node) => {
        if (!node) return;

        if (Array.isArray(node)) {
            node.forEach(findTableStyles);
            return;
        }

        if (typeof node === 'object') {
            if (node.type === 'style:style') {
                const attrs = parseAttributes(node);
                if (attrs['style:family'] === 'table') {
                    const styleName = attrs['style:name'];
                    tableStyles[styleName] = { name: styleName };

                    // Look for table-properties child
                    if (node.children) {
                        node.children.forEach(child => {
                            if (child.type === 'style:table-properties') {
                                const props = parseAttributes(child);
                                tableStyles[styleName].display = props['table:display'];
                            }
                        });
                    }
                }
            }

            if (node.children) {
                findTableStyles(node.children);
            }
        }
    };

    findTableStyles(parsed);
    return tableStyles;
}

// Convert ODS style to CSS
function styleToCSS(styleName, styles) {
    if (!styleName || !styles[styleName]) return '';

    const style = styles[styleName];
    const css = [];

    // Text properties
    if (style.text) {
        const t = style.text;
        if (t['fo:font-family']) css.push(`font-family: ${t['fo:font-family']}`);
        if (t['fo:font-size']) css.push(`font-size: ${t['fo:font-size']}`);
        if (t['fo:font-weight']) css.push(`font-weight: ${t['fo:font-weight']}`);
        if (t['fo:font-style']) css.push(`font-style: ${t['fo:font-style']}`);
        if (t['style:text-underline-style']) css.push('text-decoration: underline');
        if (t['fo:color']) css.push(`color: ${t['fo:color']}`);
    }

    // Helper to normalize border style (convert ODS format to CSS format)
    const normalizeBorder = (border) => {
        if (!border) return border;
        // Convert "thin" to "1px", "medium" to "2px", "thick" to "3px"
        return border.replace(/\bthin\b/g, '1px')
                     .replace(/\bmedium\b/g, '2px')
                     .replace(/\bthick\b/g, '3px');
    };

    // Cell properties
    if (style.cell) {
        const c = style.cell;
        if (c['fo:background-color']) css.push(`background-color: ${c['fo:background-color']}`);

        // Expand shorthand fo:border into individual borders for deduplication to work
        if (c['fo:border']) {
            const normalizedBorder = normalizeBorder(c['fo:border']);
            // Apply to all four sides individually instead of using shorthand
            css.push(`border-left: ${normalizedBorder}`);
            css.push(`border-right: ${normalizedBorder}`);
            css.push(`border-top: ${normalizedBorder}`);
            css.push(`border-bottom: ${normalizedBorder}`);
        }

        // Individual borders override the shorthand if present
        // Skip borders set to "none" as they don't need to be explicitly stated in CSS
        if (c['fo:border-left'] && c['fo:border-left'] !== 'none') {
            css.push(`border-left: ${normalizeBorder(c['fo:border-left'])}`);
        }
        if (c['fo:border-right'] && c['fo:border-right'] !== 'none') {
            css.push(`border-right: ${normalizeBorder(c['fo:border-right'])}`);
        }
        if (c['fo:border-top'] && c['fo:border-top'] !== 'none') {
            css.push(`border-top: ${normalizeBorder(c['fo:border-top'])}`);
        }
        if (c['fo:border-bottom'] && c['fo:border-bottom'] !== 'none') {
            css.push(`border-bottom: ${normalizeBorder(c['fo:border-bottom'])}`);
        }
    }

    // Paragraph properties
    if (style.paragraph) {
        const p = style.paragraph;
        if (p['fo:text-align']) css.push(`text-align: ${p['fo:text-align']}`);
    }

    return css.join('; ');
}

// Helper: Parse rich text from cell node
function parseRichText(node) {
    if (!node) return null;

    const segments = [];
    let hasFormatting = false;

    const processTextNode = (n, inFormattedNode = false) => {
        if (!n) return;

        if (Array.isArray(n)) {
            n.forEach(child => processTextNode(child, inFormattedNode));
            return;
        }

        if (typeof n === 'object') {
            // Handle text:span (formatted text) - these have style
            if (n.type === 'text:span') {
                hasFormatting = true;
                const attrs = parseAttributes(n);
                const styleName = attrs['text:style-name'];
                const text = getTextContent(n);

                if (text) {
                    segments.push({
                        text,
                        style: styleName || null
                    });
                }
                // Don't recurse - getTextContent already handled children
                return;
            }
            // Handle text:a (hyperlink) - these have links
            else if (n.type === 'text:a') {
                hasFormatting = true;
                const attrs = parseAttributes(n);
                const text = getTextContent(n);

                if (text) {
                    segments.push({
                        text,
                        link: attrs['xlink:href'] || null
                    });
                }
                // Don't recurse - getTextContent already handled children
                return;
            }
            // Handle text:p (paragraph) - just container
            else if (n.type === 'text:p') {
                if (n.children) {
                    processTextNode(n.children, inFormattedNode);
                }
                return;
            }
            // Handle #text nodes only if not inside formatted node
            else if (n.type === '#text' && !inFormattedNode) {
                const text = getTextContent(n);
                if (text) {
                    segments.push({ text });
                }
            }

            if (n.children && n.type !== 'text:span' && n.type !== 'text:a') {
                processTextNode(n.children, inFormattedNode);
            }
        }
    };

    processTextNode(node);

    // Only return segments if there's actual formatting (span/link tags)
    // Plain text cells don't need rich text structure
    return hasFormatting && segments.length > 0 ? segments : null;
}

// Helper: Parse image from draw:frame
function parseImage(frameNode, rowIndex, colIndex) {
    const attrs = parseAttributes(frameNode);
    const image = {
        type: 'image',
        anchorType: attrs['text:anchor-type'] || 'cell',
        x: attrs['svg:x'] || '0',
        y: attrs['svg:y'] || '0',
        width: attrs['svg:width'] || '0',
        height: attrs['svg:height'] || '0',
        row: rowIndex,
        col: colIndex,
        zIndex: parseInt(attrs['draw:z-index']) || 0
    };

    // Find draw:image child
    const findImage = (node) => {
        if (!node) return null;
        if (Array.isArray(node)) {
            for (let n of node) {
                const result = findImage(n);
                if (result) return result;
            }
            return null;
        }
        if (typeof node === 'object') {
            if (node.type === 'draw:image') {
                const imgAttrs = parseAttributes(node);
                return {
                    href: imgAttrs['xlink:href'],
                    type: imgAttrs['xlink:type']
                };
            }
            if (node.children) {
                return findImage(node.children);
            }
        }
        return null;
    };

    const imageData = findImage(frameNode);
    if (imageData) {
        image.href = imageData.href;
    }

    return image;
}

// Helper: Parse chart from draw:object
function parseChart(objectNode, rowIndex, colIndex) {
    const attrs = parseAttributes(objectNode);
    return {
        type: 'chart',
        href: attrs['xlink:href'],
        x: attrs['svg:x'] || '0',
        y: attrs['svg:y'] || '0',
        width: attrs['svg:width'] || '0',
        height: attrs['svg:height'] || '0',
        row: rowIndex,
        col: colIndex,
        zIndex: parseInt(attrs['draw:z-index']) || 0
    };
}

// Helper: Parse comment/annotation
function parseComment(annotationNode, rowIndex, colIndex) {
    const attrs = parseAttributes(annotationNode);
    const comment = {
        author: attrs['office:author'] || '',
        date: attrs['office:date'] || '',
        text: getTextContent(annotationNode)
    };

    // Try to find display attribute
    if (attrs['office:display'] === 'true') {
        comment.visible = true;
    }

    return comment;
}

// Helper: Parse data validation
function parseDataValidation(validationNode) {
    const attrs = parseAttributes(validationNode);
    const validation = {
        name: attrs['table:name'],
        condition: attrs['table:condition'],
        baseCellAddress: attrs['table:base-cell-address'],
        allowEmpty: attrs['table:allow-empty-cell'] === 'true',
        displayList: attrs['table:display-list']
    };

    // Parse error message
    const findErrorMessage = (node) => {
        if (!node) return;
        if (Array.isArray(node)) {
            node.forEach(findErrorMessage);
            return;
        }
        if (typeof node === 'object') {
            if (node.type === 'table:error-message') {
                const msgAttrs = parseAttributes(node);
                validation.errorMessage = {
                    title: msgAttrs['table:title'],
                    display: msgAttrs['table:display'] === 'true',
                    text: getTextContent(node)
                };
            }
            if (node.children) findErrorMessage(node.children);
        }
    };

    if (validationNode.children) {
        findErrorMessage(validationNode.children);
    }

    return validation;
}

// Helper: Parse conditional format
function parseConditionalFormat(formatNode) {
    const attrs = parseAttributes(formatNode);
    const format = {
        targetRange: attrs['table:target-range-address']
    };

    // Parse conditions
    const conditions = [];

    const findConditions = (node) => {
        if (!node) return;
        if (Array.isArray(node)) {
            node.forEach(findConditions);
            return;
        }
        if (typeof node === 'object') {
            if (node.type === 'table:condition') {
                const condAttrs = parseAttributes(node);
                conditions.push({
                    condition: condAttrs['table:condition'],
                    applyStyleName: condAttrs['table:apply-style-name'],
                    baseCellAddress: condAttrs['table:base-cell-address']
                });
            }
            if (node.children) findConditions(node.children);
        }
    };

    if (formatNode.children) {
        findConditions(formatNode.children);
    }

    format.conditions = conditions;
    return format;
}

// Parse table from content.xml
function parseTable(tableNode, styles, numberFormats, tableStyles = {}) {
    const result = {
        name: '',
        data: [],
        columns: [],
        rows: {},
        cells: {},
        style: {},
        mergeCells: {},
        media: [], // Unified array for images, charts, and shapes
        comments: {}
    };


    if (!tableNode) return result;

    const attrs = parseAttributes(tableNode);
    result.name = attrs['table:name'] || 'Sheet1';

    // Parse table display settings (grid lines from table:print="false" attribute)
    // ODS doesn't have direct gridline control like XLSX, but we can check table:display
    if (attrs['table:display'] === 'false') {
        result.gridline = false;
    }

    // Parse worksheet visibility state (only set if hidden, like XLSX)
    // Check both direct attribute and table style
    const tableStyleName = attrs['table:style-name'];
    const tableStyle = tableStyles[tableStyleName];

    if (attrs['table:display'] === 'false' || tableStyle?.display === 'false') {
        result.worksheetState = 'hidden';
    }

    // Parse table protection
    if (attrs['table:protected'] === 'true') {
        result.locked = true;
        if (attrs['table:protection-key']) {
            // Table is password protected
            result.locked = { password: true };
        }
    }

    // Parse default column width from table style (reuse tableStyleName from above)
    if (tableStyleName && styles[tableStyleName] && styles[tableStyleName].table) {
        const tableProps = styles[tableStyleName].table;
        const defaultWidth = tableProps['table:default-column-width'];
        if (defaultWidth) {
            // Convert ODS units to pixels
            const match = defaultWidth.match(/^([\d.]+)(cm|in|pt|px)?$/);
            if (match) {
                const value = parseFloat(match[1]);
                const unit = match[2] || 'cm';
                let widthPx;
                if (unit === 'cm') {
                    widthPx = Math.round(value * 37.795);
                } else if (unit === 'in') {
                    widthPx = Math.round(value * 96);
                } else if (unit === 'pt') {
                    widthPx = convertWidthToPixels(value, 'pt');
                } else if (unit === 'px') {
                    widthPx = Math.round(value);
                }
                result.defaultColWidth = widthPx + 'px';
            }
        }
    }

    let rowIndex = 0;
    let maxCol = 0;

    const processNode = (node) => {
        if (!node) return;

        if (Array.isArray(node)) {
            node.forEach(processNode);
            return;
        }

        // Parse columns
        if (node.type === 'table:table-column') {
            const colAttrs = parseAttributes(node);
            let repeat = parseInt(colAttrs['table:number-columns-repeated']) || 1;
            const styleName = colAttrs['table:style-name'];
            const visibility = colAttrs['table:visibility'];

            // Limit repeated columns to avoid creating too many empty columns
            if (!styleName && repeat > 100) {
                repeat = 0; // Skip large blocks of empty columns
            }

            for (let i = 0; i < repeat; i++) {
                // Get width from style if available
                let width = 100; // Default width
                if (styleName && styles[styleName] && styles[styleName].column) {
                    const colProps = styles[styleName].column;
                    const styleWidth = colProps['style:column-width'];
                    if (styleWidth) {
                        // ODS widths can be in various units: cm, in, pt, px
                        // Extract number and unit
                        const match = styleWidth.match(/^([\d.]+)(cm|in|pt|px)?$/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            const unit = match[2] || 'cm'; // Default to cm

                            // Convert to pixels based on unit
                            if (unit === 'cm') {
                                width = Math.round(value * 37.795); // 1cm ≈ 37.795px
                            } else if (unit === 'in') {
                                width = Math.round(value * 96); // 1in = 96px
                            } else if (unit === 'pt') {
                                width = convertWidthToPixels(value, 'pt');
                            } else if (unit === 'px') {
                                width = Math.round(value);
                            }
                        }
                    }
                }

                const col = {
                    width: width,
                    title: getColumnName(result.columns.length),
                    type: 'text'
                };

                // Add style if present
                if (styleName) {
                    col.style = styleName;
                }

                // Check if column is hidden (visible: false for hidden)
                if (visibility === 'collapse' || visibility === 'filter') {
                    col.visible = false;
                }

                result.columns.push(col);
            }
        }

        // Parse rows
        if (node.type === 'table:table-row') {
            const rowAttrs = parseAttributes(node);
            let rowRepeat = parseInt(rowAttrs['table:number-rows-repeated']) || 1;
            const visibility = rowAttrs['table:visibility'];

            // Check if row is empty (only has empty cells with no styles)
            let isEmptyRow = true;
            if (node.children) {
                for (let cellNode of node.children) {
                    if (cellNode.type === 'table:table-cell' || cellNode.type === 'table:covered-table-cell') {
                        const cellAttrs = parseAttributes(cellNode);
                        const valueType = cellAttrs['office:value-type'];
                        const textContent = getTextContent(cellNode);
                        if (valueType || textContent || cellAttrs['table:formula']) {
                            isEmptyRow = false;
                            break;
                        }
                    }
                }
            }

            // Limit repeated empty rows to avoid memory issues
            if (isEmptyRow && rowRepeat > 100) {
                rowRepeat = 0; // Skip large blocks of empty rows
            }

            for (let r = 0; r < rowRepeat; r++) {
                const currentRow = rowIndex + r;
                result.data[currentRow] = [];

                // Store row properties (height, visible status)
                if (visibility === 'collapse' || visibility === 'filter') {
                    if (!result.rows[currentRow]) result.rows[currentRow] = {};
                    result.rows[currentRow].visible = false;
                }

                let colIndex = 0;

                // Parse cells in row
                if (node.children) {
                    node.children.forEach(cellNode => {
                        if (cellNode.type === 'table:table-cell' || cellNode.type === 'table:covered-table-cell') {
                            const cellAttrs = parseAttributes(cellNode);
                            let colRepeat = parseInt(cellAttrs['table:number-columns-repeated']) || 1;
                            const colspan = parseInt(cellAttrs['table:number-columns-spanned']) || 1;
                            const rowspan = parseInt(cellAttrs['table:number-rows-spanned']) || 1;
                            const valueType = cellAttrs['office:value-type'];
                            const cellStyleName = cellAttrs['table:style-name'];

                            let value = '';

                            // Extract cell value
                            if (valueType === 'string') {
                                value = getTextContent(cellNode);
                            } else if (valueType === 'float') {
                                value = parseFloat(cellAttrs['office:value']) || 0;
                            } else if (valueType === 'percentage') {
                                value = parseFloat(cellAttrs['office:value']) || 0;
                            } else if (valueType === 'currency') {
                                value = parseFloat(cellAttrs['office:value']) || 0;
                            } else if (valueType === 'date') {
                                value = cellAttrs['office:date-value'] || '';
                            } else if (valueType === 'time') {
                                value = cellAttrs['office:time-value'] || '';
                            } else if (valueType === 'boolean') {
                                value = cellAttrs['office:boolean-value'] === 'true';
                            } else {
                                // Default to text content
                                value = getTextContent(cellNode);
                            }

                            // Handle formulas
                            const formula = cellAttrs['table:formula'];

                            // Check for cell protection
                            const cellProtected = cellAttrs['table:protected'] === 'true';

                            // Parse rich text / hyperlinks from cell children
                            let richText = null;
                            let hyperlink = null;

                            if (cellNode.children) {
                                // Check for hyperlinks first
                                const findHyperlink = (n) => {
                                    if (!n) return null;
                                    if (Array.isArray(n)) {
                                        for (let item of n) {
                                            const result = findHyperlink(item);
                                            if (result) return result;
                                        }
                                        return null;
                                    }
                                    if (typeof n === 'object') {
                                        if (n.type === 'text:a') {
                                            const attrs = parseAttributes(n);
                                            return attrs['xlink:href'];
                                        }
                                        if (n.children) return findHyperlink(n.children);
                                    }
                                    return null;
                                };

                                hyperlink = findHyperlink(cellNode.children);

                                // Parse rich text (text:span elements)
                                richText = parseRichText(cellNode);

                                // Check for images in cell
                                const findDrawElements = (n) => {
                                    if (!n) return;
                                    if (Array.isArray(n)) {
                                        n.forEach(findDrawElements);
                                        return;
                                    }
                                    if (typeof n === 'object') {
                                        // Images
                                        if (n.type === 'draw:frame') {
                                            const image = parseImage(n, currentRow, colIndex);
                                            if (image) {
                                                image.type = 'image';
                                                result.media.push(image);
                                            }
                                        }
                                        // Charts
                                        else if (n.type === 'draw:object') {
                                            const chart = parseChart(n, currentRow, colIndex);
                                            if (chart) {
                                                chart.type = 'chart';
                                                result.media.push(chart);
                                            }
                                        }
                                        // Comments
                                        else if (n.type === 'office:annotation') {
                                            const cellAddr = getCellNameFromCoords(colIndex, currentRow);
                                            result.comments[cellAddr] = parseComment(n, currentRow, colIndex);
                                        }

                                        if (n.children) findDrawElements(n.children);
                                    }
                                };

                                findDrawElements(cellNode.children);
                            }

                            // Limit repeated empty cells
                            if (!valueType && !value && !formula && colRepeat > 100) {
                                colRepeat = 0; // Skip large blocks of empty cells
                            }

                            // Place value in cells
                            for (let c = 0; c < colRepeat; c++) {
                                const col = colIndex + c;
                                // If cell has formula, export formula string instead of calculated value
                                if (formula && c === 0) {
                                    // Clean ODS formula format: remove "of:" prefix and convert to Excel format
                                    let cleanFormula = formula.replace(/^of:/, '');

                                    // Decode HTML entities
                                    cleanFormula = cleanFormula
                                        .replace(/&quot;/g, '"')
                                        .replace(/&lt;/g, '<')
                                        .replace(/&gt;/g, '>')
                                        .replace(/&amp;/g, '&');

                                    // Convert ODS range references [.A1:.B10] or [.$A$1:.$B$10] to Excel format A1:B10
                                    cleanFormula = cleanFormula.replace(/\[\.(\$?)([A-Z]+)(\$?)(\d+):\.(\$?)([A-Z]+)(\$?)(\d+)\]/g, '$1$2$3$4:$5$6$7$8');

                                    // Convert ODS cell references [.A1] or [.$A$1] to Excel format A1 or $A$1
                                    cleanFormula = cleanFormula.replace(/\[\.(\$?)([A-Z]+)(\$?)(\d+)\]/g, '$1$2$3$4');

                                    // Convert ODS semicolons to Excel commas in function arguments
                                    cleanFormula = cleanFormula.replace(/;/g, ',');

                                    result.data[currentRow][col] = cleanFormula;
                                } else {
                                    result.data[currentRow][col] = value;
                                }

                                if (col > maxCol) maxCol = col;

                                // Add style if present
                                if (cellStyleName) {
                                    const cellAddr = getCellNameFromCoords(col, currentRow);
                                    const cssStyle = styleToCSS(cellStyleName, styles);
                                    if (cssStyle) {
                                        result.style[cellAddr] = cssStyle;
                                    }

                                    // Add number format if present (only for numeric value types)
                                    // Don't apply number formats to string cells
                                    const isNumericType = valueType === 'float' || valueType === 'percentage' ||
                                                         valueType === 'currency' || valueType === 'date' || valueType === 'time';

                                    if (isNumericType && styles[cellStyleName] && styles[cellStyleName].dataStyleName) {
                                        const dataStyleName = styles[cellStyleName].dataStyleName;
                                        if (numberFormats && numberFormats[dataStyleName]) {
                                            const format = numberFormats[dataStyleName];

                                            // Skip "default" formats that don't add meaningful information
                                            // Don't apply integer format "0" to decimal values
                                            const isDefaultFormat = format === '0' || format === 'General';
                                            const hasDecimal = value && typeof value === 'number' && value % 1 !== 0;

                                            if (!isDefaultFormat || !hasDecimal) {
                                                if (!result.cells[cellAddr]) result.cells[cellAddr] = {};
                                                // Only set format if not already set (preserve first occurrence)
                                                if (!result.cells[cellAddr].format) {
                                                    result.cells[cellAddr].format = format;
                                                }
                                            }
                                        }
                                    }
                                }

                                // Add hyperlink if present (store in cells as url property)
                                if (hyperlink && c === 0) {
                                    const cellAddr = getCellNameFromCoords(col, currentRow);
                                    if (!result.cells[cellAddr]) result.cells[cellAddr] = {};
                                    result.cells[cellAddr].url = hyperlink;
                                }

                                // Add merged cells
                                if ((colspan > 1 || rowspan > 1) && c === 0) {
                                    const cellAddr = getCellNameFromCoords(col, currentRow);
                                    result.mergeCells[cellAddr] = [colspan, rowspan];
                                }
                            }

                            colIndex += colRepeat;
                        }
                    });
                }
            }

            rowIndex += rowRepeat;
        }

        // Recurse into children
        if (node.children) {
            processNode(node.children);
        }
    };

    if (tableNode.children) {
        processNode(tableNode.children);
    }

    // Fill empty cells
    for (let r = 0; r < result.data.length; r++) {
        if (!result.data[r]) result.data[r] = [];
        for (let c = 0; c <= maxCol; c++) {
            if (result.data[r][c] === undefined) {
                result.data[r][c] = '';
            }
        }
    }

    // Trim columns array to match actual data width
    if (result.columns.length > maxCol + 1) {
        result.columns = result.columns.slice(0, maxCol + 1);
    }

    // Ensure we have at least maxCol+1 columns
    while (result.columns.length <= maxCol) {
        result.columns.push({
            width: 100,
            title: getColumnName(result.columns.length),
            type: 'text'
        });
    }

    // Add minDimensions [columns, rows]
    result.minDimensions = [maxCol + 1, result.data.length];

    return result;
}

/**
 * Parse ODS file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseODS(input, options = {}) {
    return parse(async (inp) => {
        // JSZip can handle File objects directly in browser, or buffers
        let zipInput;
        if (typeof File !== 'undefined' && inp instanceof File) {
            zipInput = inp; // JSZip handles File objects directly
        } else {
            zipInput = await loadAsBuffer(inp);
        }

        const zip = await JSZip.loadAsync(zipInput);

        // Extract content.xml
        const contentFile = zip.file('content.xml');
        if (!contentFile) {
            throw new Error('content.xml not found in ODS file');
        }

        const contentXml = await contentFile.async('string');
        const content = parser(contentXml);

        // Extract styles.xml
        let styles = {};
        const stylesFile = zip.file('styles.xml');
        if (stylesFile) {
            const stylesXml = await stylesFile.async('string');
            styles = parseStyles(stylesXml);
        }

        // Also parse automatic styles from content.xml
        const contentStyles = parseStyles(contentXml);
        styles = { ...styles, ...contentStyles };

        // Parse table styles for worksheet visibility
        let tableStyles = {};
        if (stylesFile) {
            const stylesXml = await stylesFile.async('string');
            tableStyles = parseTableStyles(stylesXml);
        }
        const contentTableStyles = parseTableStyles(contentXml);
        tableStyles = { ...tableStyles, ...contentTableStyles };

        // Parse number formats from both styles.xml and content.xml
        let numberFormats = {};
        if (stylesFile) {
            const stylesXml = await stylesFile.async('string');
            numberFormats = parseNumberFormats(stylesXml);
        }
        const contentFormats = parseNumberFormats(contentXml);
        numberFormats = { ...numberFormats, ...contentFormats };


        // Parse document-level features
        const definedNames = {};

        const parseDocumentFeatures = (node) => {
            if (!node) return;

            if (Array.isArray(node)) {
                node.forEach(parseDocumentFeatures);
                return;
            }

            if (typeof node === 'object') {
                // Named ranges
                if (node.type === 'table:named-range') {
                    const attrs = parseAttributes(node);
                    const name = attrs['table:name'];
                    const range = attrs['table:cell-range-address'];
                    if (name && range) {
                        // Store as simple string to match XLSX format
                        definedNames[name] = range;
                    }
                }

                if (node.children) {
                    parseDocumentFeatures(node.children);
                }
            }
        };

        parseDocumentFeatures(content);

        // Parse frozen panes from settings.xml
        let frozenPanes = {};
        const settingsFile = zip.file('settings.xml');
        if (settingsFile) {
            const settingsXml = await settingsFile.async('string');
            const settings = parser(settingsXml);

            const findConfig = (node, sheetName) => {
                if (!node) return null;
                if (Array.isArray(node)) {
                    for (let n of node) {
                        const result = findConfig(n, sheetName);
                        if (result) return result;
                    }
                    return null;
                }
                if (typeof node === 'object') {
                    if (node.type === 'config:config-item-map-entry') {
                        const attrs = parseAttributes(node);
                        const name = attrs['config:name'];

                        if (name && node.children) {
                            const panes = {};
                            const findPaneSettings = (n) => {
                                if (!n) return;
                                if (Array.isArray(n)) {
                                    n.forEach(findPaneSettings);
                                    return;
                                }
                                if (typeof n === 'object' && n.type === 'config:config-item') {
                                    const itemAttrs = parseAttributes(n);
                                    const itemName = itemAttrs['config:name'];
                                    const value = getTextContent(n);

                                    if (itemName === 'HorizontalSplitMode' && value === '2') {
                                        panes.horizontalSplit = true;
                                    } else if (itemName === 'VerticalSplitMode' && value === '2') {
                                        panes.verticalSplit = true;
                                    } else if (itemName === 'PositionRight') {
                                        panes.frozenColumns = parseInt(value) || 0;
                                    } else if (itemName === 'PositionBottom') {
                                        panes.frozenRows = parseInt(value) || 0;
                                    }

                                    if (n.children) findPaneSettings(n.children);
                                } else if (n.children) {
                                    findPaneSettings(n.children);
                                }
                            };
                            findPaneSettings(node.children);

                            if (panes.frozenRows || panes.frozenColumns) {
                                frozenPanes[name] = panes;
                            }
                        }
                    }
                    if (node.children) {
                        const result = findConfig(node.children, sheetName);
                        if (result) return result;
                    }
                }
                return null;
            };

            findConfig(settings);
        }

        // Find all tables (worksheets)
        const worksheets = [];
        const allValidations = []; // Collect validations (includes conditional formatting with action:'format')

        const findTables = (node) => {
            if (!node) return;

            if (Array.isArray(node)) {
                node.forEach(findTables);
                return;
            }

            if (typeof node === 'object') {
                if (node.type === 'table:table') {
                    const tableData = parseTable(node, styles, numberFormats, tableStyles);

                    // Add frozen panes if available for this sheet
                    const sheetFrozenPanes = frozenPanes[tableData.name] || {};

                    const worksheet = {
                        worksheetName: tableData.name,
                        data: tableData.data,
                        columns: tableData.columns,
                        rows: tableData.rows,
                        cells: tableData.cells,
                        style: tableData.style,
                        mergeCells: tableData.mergeCells,
                        media: tableData.media,
                        comments: tableData.comments,
                        minDimensions: tableData.minDimensions,
                        freezeRows: sheetFrozenPanes.frozenRows || 0,
                        freezeColumns: sheetFrozenPanes.frozenColumns || 0
                    };

                    // Add optional properties if present
                    if (tableData.gridline !== undefined) {
                        worksheet.gridline = tableData.gridline;
                    }
                    if (tableData.locked !== undefined) {
                        worksheet.locked = tableData.locked;
                    }
                    if (tableData.defaultColWidth !== undefined) {
                        worksheet.defaultColWidth = tableData.defaultColWidth;
                    }
                    if (tableData.worksheetState !== undefined) {
                        worksheet.worksheetState = tableData.worksheetState;
                    }

                    // Collect validations from this worksheet (includes conditional formatting)
                    if (tableData.validations) {
                        // Add worksheet name to each validation/conditional formatting
                        tableData.validations.forEach(validation => {
                            // Prefix the range with worksheet name if not already qualified
                            if (!validation.range.includes('!')) {
                                validation.range = `${tableData.name}!${validation.range}`;
                            }
                            allValidations.push(validation);
                        });
                    }

                    worksheets.push(worksheet);
                    // Don't recurse into table children - parseTable handles that
                    return;
                }

                if (node.children) {
                    findTables(node.children);
                }
            }
        };

        findTables(content);

        // Build global styles array as CSS strings and convert style references to indices
        const styleNames = Object.keys(styles);
        const globalStyles = styleNames.map(name => styleToCSS(name, styles));

        // Create style name to index mapping
        const styleIndexMap = {};
        styleNames.forEach((name, index) => {
            styleIndexMap[name] = index;
        });

        // Convert all style references in worksheets to indices
        worksheets.forEach(worksheet => {
            // Convert column style references
            worksheet.columns.forEach(col => {
                if (col.style && styleIndexMap[col.style] !== undefined) {
                    col.style = styleIndexMap[col.style];
                } else if (col.style) {
                    // Style not found, remove it
                    delete col.style;
                }
            });

            // Apply border deduplication to prevent double borders between adjacent cells
            // Track cells that should NOT have top/left borders
            const cellsWithoutTopBorder = new Set();
            const cellsWithoutLeftBorder = new Set();

            // First pass: identify which cells should not have top/left borders
            Object.keys(worksheet.style).forEach(cellAddr => {
                const cssString = worksheet.style[cellAddr];

                // Check if this cell has border-right or border-bottom
                if (cssString.includes('border-right:')) {
                    // Mark cells to the right to not have border-left
                    // Account for merged cells - if this cell is merged, mark cells to the right of the merge
                    const coords = getCoordsFromCellName(cellAddr);
                    if (coords && coords[0] !== null && coords[1] !== null) {
                        const mergeInfo = worksheet.mergeCells?.[cellAddr];
                        if (mergeInfo && Array.isArray(mergeInfo)) {
                            // Cell is merged - mark all cells to the right of the merged range
                            const [colSpan, rowSpan] = mergeInfo;
                            for (let rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
                                const rightCell = getCellNameFromCoords(coords[0] + colSpan, coords[1] + rowOffset);
                                cellsWithoutLeftBorder.add(rightCell);
                            }
                        } else {
                            // Single cell - mark only the cell directly to the right
                            const rightCell = getCellNameFromCoords(coords[0] + 1, coords[1]);
                            cellsWithoutLeftBorder.add(rightCell);
                        }
                    }
                }
                if (cssString.includes('border-bottom:')) {
                    // Mark cells below to not have border-top
                    // Account for merged cells - if this cell is merged, mark all cells below the merge
                    const coords = getCoordsFromCellName(cellAddr);
                    if (coords && coords[0] !== null && coords[1] !== null) {
                        const mergeInfo = worksheet.mergeCells?.[cellAddr];
                        if (mergeInfo && Array.isArray(mergeInfo)) {
                            // Cell is merged - mark all cells below the merged range
                            const [colSpan, rowSpan] = mergeInfo;
                            for (let colOffset = 0; colOffset < colSpan; colOffset++) {
                                const belowCell = getCellNameFromCoords(coords[0] + colOffset, coords[1] + rowSpan);
                                cellsWithoutTopBorder.add(belowCell);
                            }
                        } else {
                            // Single cell - mark only the cell directly below
                            const belowCell = getCellNameFromCoords(coords[0], coords[1] + 1);
                            cellsWithoutTopBorder.add(belowCell);
                        }
                    }
                }
            });

            // Second pass: remove duplicate borders
            Object.keys(worksheet.style).forEach(cellAddr => {
                let cssString = worksheet.style[cellAddr];

                // Remove border-top if this cell should not have one
                if (cellsWithoutTopBorder.has(cellAddr)) {
                    cssString = cssString.replace(/border-top:\s*[^;]+;?\s*/g, '');
                }

                // Remove border-left if this cell should not have one
                if (cellsWithoutLeftBorder.has(cellAddr)) {
                    cssString = cssString.replace(/border-left:\s*[^;]+;?\s*/g, '');
                }

                worksheet.style[cellAddr] = cssString;
            });

            // Convert cell style references to indices
            const newStyleObj = {};
            Object.keys(worksheet.style).forEach(cellAddr => {
                const cssString = worksheet.style[cellAddr];
                // Try to find matching style index by CSS string
                const matchingIndex = globalStyles.findIndex(css => css === cssString);
                if (matchingIndex !== -1) {
                    newStyleObj[cellAddr] = matchingIndex;
                } else {
                    // Add new style to global array and use that index
                    globalStyles.push(cssString);
                    newStyleObj[cellAddr] = globalStyles.length - 1;
                }
            });
            worksheet.style = newStyleObj;
        });

        const result = { worksheets };

        // Add global style array if any styles exist
        if (globalStyles.length > 0) {
            result.style = globalStyles;
        }

        // Add defined names if any
        if (Object.keys(definedNames).length > 0) {
            result.definedNames = definedNames;
        }

        // Add validations at spreadsheet level if any (includes conditional formatting with action:'format')
        if (allValidations.length > 0) {
            result.validations = allValidations;
        }

        return result;
    }, input, options);
}
