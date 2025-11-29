import { loadAsBuffer, parse } from '../utils/loader.js';
import JSZip from 'jszip';
import { parser } from '@lemonadejs/html-to-json';
import {
    getColumnName,
    getColumnIndex,
    getCellNameFromCoords,
    getCoordsFromCellName,
    getTextContent,
    parseAttributes,
    findNodes,
    convertWidthToPixels,
    borderStyles,
    getDefaultTheme,
    exclusions,
    excelValidationTypes,
    excelValidationOperations
} from '../utils/helpers.js';

// Convert cell reference to coordinates (A1 -> {row: 0, col: 0})
function cellRefToCoords(ref) {
    const coords = getCoordsFromCellName(ref);
    if (!coords || coords[0] === null || coords[1] === null) return null;
    return {
        col: coords[0],
        row: coords[1]
    };
}

// Clean up Excel internal references from formulas
function cleanFormula(formula) {
    if (!formula) return formula;

    // Decode HTML entities
    formula = formula
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");

    // Remove Excel internal prefixes (case insensitive)
    // Includes: _xlnm., _xlfn., _xll., _xlws.
    exclusions.forEach(prefix => {
        const regex = new RegExp(prefix.replace('.', '\\.'), 'gi');
        formula = formula.replace(regex, '');
    });

    // Also remove _xlnm. specifically (backward compatibility)
    formula = formula.replace(/_xlnm\./gi, '');

    return formula;
}

// Parse shared strings
async function parseSharedStrings(zip) {
    const sharedStrings = [];
    const file = zip.file('xl/sharedStrings.xml');

    if (!file) return sharedStrings;

    const xml = await file.async('string');
    const parsed = parser(xml);

    // Find all <si> (string item) elements
    const stringItems = findNodes(parsed, 'si');

    stringItems.forEach(si => {
        // Get text from <t> elements
        const textNodes = findNodes(si, 't');
        const text = textNodes.map(t => getTextContent(t)).join('');
        sharedStrings.push(text);
    });

    return sharedStrings;
}

// Parse styles
async function parseStyles(zip) {
    const styles = {
        fonts: [],
        fills: [],
        borders: [],
        cellXfs: [],
        numFmts: {}
    };

    const file = zip.file('xl/styles.xml');
    if (!file) return styles;

    const xml = await file.async('string');
    const parsed = parser(xml);

    // Get default theme colors
    const theme = getDefaultTheme();

    // Helper function to resolve color (handles both RGB and theme colors)
    const resolveColor = (colorAttrs) => {
        if (!colorAttrs) return null;

        if (colorAttrs.rgb) {
            // RGB color (already in hex format, usually with FF prefix for alpha)
            return colorAttrs.rgb;
        } else if (colorAttrs.theme !== undefined) {
            // Theme color - resolve using theme array
            const themeIndex = parseInt(colorAttrs.theme);
            if (themeIndex >= 0 && themeIndex < theme.arrayColors.length) {
                return theme.arrayColors[themeIndex];
            }
        }

        return null;
    };

    // Parse number formats
    const numFmts = findNodes(parsed, 'numFmt');
    numFmts.forEach(fmt => {
        const attrs = parseAttributes(fmt);
        styles.numFmts[attrs.numFmtId] = attrs.formatCode;
    });

    // Parse fonts
    const fonts = findNodes(parsed, 'font');
    fonts.forEach(font => {
        const fontData = {};

        const name = findNodes(font, 'name')[0];
        if (name) fontData.name = parseAttributes(name).val;

        const size = findNodes(font, 'sz')[0];
        if (size) fontData.size = parseAttributes(size).val;

        const color = findNodes(font, 'color')[0];
        if (color) fontData.color = resolveColor(parseAttributes(color));

        const bold = findNodes(font, 'b')[0];
        if (bold) fontData.bold = true;

        const italic = findNodes(font, 'i')[0];
        if (italic) fontData.italic = true;

        const underline = findNodes(font, 'u')[0];
        if (underline) fontData.underline = true;

        const strike = findNodes(font, 'strike')[0];
        if (strike) fontData.strike = true;

        styles.fonts.push(fontData);
    });

    // Parse fills
    const fills = findNodes(parsed, 'fill');
    fills.forEach(fill => {
        const fillData = {};
        const patternFill = findNodes(fill, 'patternFill')[0];

        if (patternFill) {
            const attrs = parseAttributes(patternFill);
            fillData.patternType = attrs.patternType;

            const fgColor = findNodes(patternFill, 'fgColor')[0];
            if (fgColor) {
                fillData.fgColor = resolveColor(parseAttributes(fgColor));
            }

            const bgColor = findNodes(patternFill, 'bgColor')[0];
            if (bgColor) {
                fillData.bgColor = resolveColor(parseAttributes(bgColor));
            }
        }

        styles.fills.push(fillData);
    });

    // Parse borders
    const borders = findNodes(parsed, 'border');
    borders.forEach(border => {
        const borderData = {};

        ['left', 'right', 'top', 'bottom', 'diagonal'].forEach(side => {
            const sideNode = findNodes(border, side)[0];
            if (sideNode) {
                const attrs = parseAttributes(sideNode);
                borderData[side] = {
                    style: attrs.style
                };
                const color = findNodes(sideNode, 'color')[0];
                if (color) {
                    borderData[side].color = resolveColor(parseAttributes(color));
                }
            }
        });

        styles.borders.push(borderData);
    });

    // Parse cell formats (xf elements in cellXfs)
    const cellXfs = findNodes(parsed, 'cellXfs')[0];
    if (cellXfs) {
        const xfs = findNodes(cellXfs, 'xf');
        xfs.forEach(xf => {
            const attrs = parseAttributes(xf);
            const xfData = {
                numFmtId: attrs.numFmtId,
                fontId: attrs.fontId,
                fillId: attrs.fillId,
                borderId: attrs.borderId,
                xfId: attrs.xfId
            };

            // Parse alignment
            const alignment = findNodes(xf, 'alignment')[0];
            if (alignment) {
                const alignAttrs = parseAttributes(alignment);
                xfData.alignment = {
                    horizontal: alignAttrs.horizontal,
                    vertical: alignAttrs.vertical,
                    wrapText: alignAttrs.wrapText === '1'
                };
            }

            styles.cellXfs.push(xfData);
        });
    }

    // Parse dxf (differential formats) used for conditional formatting
    const dxfs = findNodes(parsed, 'dxfs')[0];
    if (dxfs) {
        const dxfNodes = findNodes(dxfs, 'dxf');
        styles.dxfs = [];

        dxfNodes.forEach(dxf => {
            const dxfData = {};

            // Parse font
            const font = findNodes(dxf, 'font')[0];
            if (font) {
                const fontColor = findNodes(font, 'color')[0];
                if (fontColor) {
                    dxfData.fontColor = resolveColor(parseAttributes(fontColor));
                }

                const bold = findNodes(font, 'b')[0];
                if (bold) dxfData.bold = true;

                const italic = findNodes(font, 'i')[0];
                if (italic) dxfData.italic = true;
            }

            // Parse fill
            const fill = findNodes(dxf, 'fill')[0];
            if (fill) {
                const patternFill = findNodes(fill, 'patternFill')[0];
                if (patternFill) {
                    const bgColor = findNodes(patternFill, 'bgColor')[0];
                    if (bgColor) {
                        dxfData.bgColor = resolveColor(parseAttributes(bgColor));
                    }
                }
            }

            // Parse border
            const border = findNodes(dxf, 'border')[0];
            if (border) {
                dxfData.border = {};
                ['left', 'right', 'top', 'bottom'].forEach(side => {
                    const sideNode = findNodes(border, side)[0];
                    if (sideNode) {
                        const attrs = parseAttributes(sideNode);
                        if (attrs.style) {
                            dxfData.border[side] = {
                                style: attrs.style
                            };
                            const color = findNodes(sideNode, 'color')[0];
                            if (color) {
                                dxfData.border[side].color = resolveColor(parseAttributes(color));
                            }
                        }
                    }
                });
            }

            styles.dxfs.push(dxfData);
        });
    }

    return styles;
}

// Convert XLSX style to CSS
function styleToCSS(styleIndex, styles) {
    if (!styleIndex || !styles.cellXfs[styleIndex]) return '';

    const xf = styles.cellXfs[styleIndex];
    const css = [];

    // Font
    if (xf.fontId && styles.fonts[xf.fontId]) {
        const font = styles.fonts[xf.fontId];
        if (font.name) css.push(`font-family: ${font.name}`);
        if (font.size) css.push(`font-size: ${font.size}pt`);
        if (font.bold) css.push('font-weight: bold');
        if (font.italic) css.push('font-style: italic');
        if (font.underline) css.push('text-decoration: underline');
        if (font.strike) css.push('text-decoration: line-through');
        if (font.color && font.color.startsWith('FF')) {
            css.push(`color: #${font.color.substring(2)}`);
        }
    }

    // Fill
    if (xf.fillId && styles.fills[xf.fillId]) {
        const fill = styles.fills[xf.fillId];
        if (fill.fgColor && fill.fgColor.startsWith('FF')) {
            css.push(`background-color: #${fill.fgColor.substring(2)}`);
        }
    }

    // Borders
    if (xf.borderId && styles.borders[xf.borderId]) {
        const border = styles.borders[xf.borderId];

        // Helper to get CSS border style from Excel border style
        const getBorderStyle = (side) => {
            if (!side || !side.style) return null;

            const color = side.color ? `#${side.color.substring(2)}` : '#000';
            const excelStyle = side.style;

            // Map Excel border style to CSS
            const cssStyle = borderStyles[excelStyle] || borderStyles.thin;
            const [lineStyle, width] = cssStyle;

            return `${width} ${lineStyle} ${color}`;
        };

        if (border.left && border.left.style) {
            css.push(`border-left: ${getBorderStyle(border.left)}`);
        }
        if (border.right && border.right.style) {
            css.push(`border-right: ${getBorderStyle(border.right)}`);
        }
        if (border.top && border.top.style) {
            css.push(`border-top: ${getBorderStyle(border.top)}`);
        }
        if (border.bottom && border.bottom.style) {
            css.push(`border-bottom: ${getBorderStyle(border.bottom)}`);
        }
    }

    // Alignment
    if (xf.alignment) {
        if (xf.alignment.horizontal) {
            css.push(`text-align: ${xf.alignment.horizontal}`);
        }
        if (xf.alignment.vertical) {
            const vAlign = { top: 'top', center: 'middle', bottom: 'bottom' };
            css.push(`vertical-align: ${vAlign[xf.alignment.vertical] || 'middle'}`);
        }
        if (xf.alignment.wrapText) {
            css.push('white-space: pre-wrap');
        }
    }

    return css.join('; ');
}

// Excel shape type to Jspreadsheet shape type mapping
const shapeMap = {
    'rect': 'rectangle',
    'roundRect': 'rounded-rectangle',
    'triangle': 'triangle',
    'rtTriangle': 'right-triangle',
    'ellipse': 'ellipse',
    'diamond': 'diamond',
    'trapezoid': 'trapezium',
    'pentagon': 'pentagon',
    'parallelogram': 'parallelogram',
    'hexagon': 'hexagon',
    'octagon': 'octagon',
    'plus': 'drawCrossIcon',
    'star5': 'star',
    'rightArrow': 'drawRightBlockArrow',
    'leftArrow': 'drawLeftBlockArrow',
    'upArrow': 'drawUpBlockArrow',
    'downArrow': 'drawDownBlockArrow',
    'leftRightArrow': 'drawLeftRightArrow',
    'upDownArrow': 'drawUpDownArrow',
    'heart': 'drawHeart',
    'cloud': 'drawCloud',
    'smileyFace': 'drawSmileyFace'
    // Add more mappings as needed
};

// Helper to extract color from drawing color tags
function getDrawingColor(colorNode) {
    if (!colorNode) return null;

    const tagName = colorNode.tagName || colorNode.type;
    const attrs = parseAttributes(colorNode);

    if (tagName === 'srgbClr' && attrs.val) {
        return '#' + attrs.val;
    } else if (tagName === 'schemeClr') {
        // Theme colors - simplified mapping
        const themeColors = {
            'lt1': '#FFFFFF',
            'dk1': '#000000',
            'accent1': '#4472C4',
            'accent2': '#ED7D31',
            'accent3': '#A5A5A5',
            'accent4': '#FFC000',
            'accent5': '#5B9BD5',
            'accent6': '#70AD47'
        };
        return themeColors[attrs.val] || '#000000';
    }

    return null;
}

// Parse drawings/images/shapes for a worksheet
async function parseDrawings(zip, sheetIndex) {
    const drawingsFile = zip.file(`xl/drawings/drawing${sheetIndex}.xml`);
    if (!drawingsFile) return [];

    const xml = await drawingsFile.async('string');
    const parsed = parser(xml);
    const media = [];

    // Find all anchor types (with namespace prefix xdr:)
    const anchors = [
        ...findNodes(parsed, 'xdr:twoCellAnchor'),
        ...findNodes(parsed, 'xdr:oneCellAnchor'),
        ...findNodes(parsed, 'xdr:absoluteAnchor')
    ];

    for (const anchor of anchors) {
        // Check if it's a picture (xdr:pic)
        const picNode = findNodes(anchor, 'xdr:pic')[0];
        if (picNode) {
            const mediaObj = await parseImage(zip, sheetIndex, anchor, picNode);
            if (mediaObj) media.push(mediaObj);
            continue;
        }

        // Check if it's a shape (xdr:sp)
        const spNode = findNodes(anchor, 'xdr:sp')[0];
        if (spNode) {
            const shapeObj = parseShape(anchor, spNode);
            if (shapeObj) media.push(shapeObj);
            continue;
        }

        // Check if it's a chart (xdr:graphicFrame)
        const graphicFrameNode = findNodes(anchor, 'xdr:graphicFrame')[0];
        if (graphicFrameNode) {
            const chartObj = await parseChart(zip, sheetIndex, anchor, graphicFrameNode);
            if (chartObj) media.push(chartObj);
            continue;
        }
    }

    return media;
}

// Parse image from anchor
async function parseImage(zip, sheetIndex, anchor, pic) {
    const mediaObj = {};

    // Get the image relationship ID
    const blipNode = findNodes(pic, 'blip')[0];
    if (!blipNode) return null;

    const blipAttrs = parseAttributes(blipNode);
    const rId = blipAttrs['r:embed'];

    if (rId) {
        // Get the drawing relationships to find the actual image file
        const relsFile = zip.file(`xl/drawings/_rels/drawing${sheetIndex}.xml.rels`);
        if (relsFile) {
            const relsXml = await relsFile.async('string');
            const relsParsed = parser(relsXml);
            const relationships = findNodes(relsParsed, 'Relationship');

            for (const rel of relationships) {
                const relAttrs = parseAttributes(rel);
                if (relAttrs.Id === rId) {
                    const imagePath = `xl/${relAttrs.Target.replace('../', '')}`;
                    const imageFile = zip.file(imagePath);

                    if (imageFile) {
                        // Get image as base64
                        const imageData = await imageFile.async('base64');
                        const extension = imagePath.split('.').pop().toLowerCase();
                        const mimeType = {
                            'png': 'image/png',
                            'jpg': 'image/jpeg',
                            'jpeg': 'image/jpeg',
                            'gif': 'image/gif',
                            'bmp': 'image/bmp',
                            'svg': 'image/svg+xml'
                        }[extension] || 'image/png';

                        mediaObj.src = `data:${mimeType};base64,${imageData}`;
                    }
                    break;
                }
            }
        }
    }

    // Get anchor position
    const anchorData = getAnchorData(anchor);
    Object.assign(mediaObj, anchorData);

    return mediaObj.src ? mediaObj : null;
}

// Parse shape from anchor
function parseShape(anchor, spNode) {
    const shapeObj = {
        type: 'shape'
    };

    // Get shape properties
    const spPrNode = findNodes(spNode, 'spPr')[0];
    if (!spPrNode) return null;

    // Get shape type
    const prstGeomNode = findNodes(spPrNode, 'prstGeom')[0];
    if (prstGeomNode) {
        const attrs = parseAttributes(prstGeomNode);
        const excelShape = attrs.prst;
        shapeObj.options = {
            type: shapeMap[excelShape] || 'rectangle'
        };
    } else {
        shapeObj.options = { type: 'rectangle' };
    }

    // Get background color
    const solidFillNode = findNodes(spPrNode, 'solidFill')[0];
    if (solidFillNode) {
        const colorNodes = solidFillNode.children || [];
        if (colorNodes.length > 0) {
            const color = getDrawingColor(colorNodes[0]);
            if (color) shapeObj.options.backgroundColor = color;
        }
    } else if (findNodes(spPrNode, 'noFill')[0]) {
        shapeObj.options.backgroundColor = 'transparent';
    }

    // Get border
    const lnNode = findNodes(spPrNode, 'ln')[0];
    if (lnNode) {
        const lnAttrs = parseAttributes(lnNode);
        if (lnAttrs.w) {
            // Convert EMU to pixels (EMU / 9525 = points, points * 1.333 = pixels)
            const emu = parseInt(lnAttrs.w);
            shapeObj.options.borderWidth = Math.round(emu / 9525 * 1.333);
        }

        const lnSolidFillNode = findNodes(lnNode, 'solidFill')[0];
        if (lnSolidFillNode) {
            const colorNodes = lnSolidFillNode.children || [];
            if (colorNodes.length > 0) {
                const color = getDrawingColor(colorNodes[0]);
                if (color) shapeObj.options.borderColor = color;
            }
        } else if (findNodes(lnNode, 'noFill')[0]) {
            shapeObj.options.borderWidth = 0;
        }
    }

    // Get text content
    const txBodyNode = findNodes(spNode, 'txBody')[0];
    if (txBodyNode) {
        const textParts = [];
        const pNodes = findNodes(txBodyNode, 'p');

        for (const pNode of pNodes) {
            const tNodes = findNodes(pNode, 't');
            const lineText = tNodes.map(t => getTextContent(t)).join('');
            if (lineText) textParts.push(lineText);
        }

        if (textParts.length > 0) {
            shapeObj.options.text = textParts.join('\n');
        }

        // Try to get text color
        const rPrNode = findNodes(txBodyNode, 'rPr')[0];
        if (rPrNode) {
            const solidFillNode = findNodes(rPrNode, 'solidFill')[0];
            if (solidFillNode) {
                const colorNodes = solidFillNode.children || [];
                if (colorNodes.length > 0) {
                    const color = getDrawingColor(colorNodes[0]);
                    if (color) shapeObj.options.fontColor = color;
                }
            }
        }
    }

    // Get anchor position
    const anchorData = getAnchorData(anchor);
    Object.assign(shapeObj, anchorData);

    return shapeObj;
}

// Join multiple Excel ranges into a single bounding box range
// e.g., ["Sheet1!$B$2:$B$10", "Sheet1!$C$2:$C$10"] => "Sheet1!$B$2:$C$10"
function joinExcelRanges(ranges) {
    if (!ranges || ranges.length === 0) return null;
    if (ranges.length === 1) return ranges[0];

    // Extract sheet name from first range
    const firstRange = ranges[0];
    const exclamationIndex = firstRange.lastIndexOf('!');
    const sheetPrefix = exclamationIndex > -1 ? firstRange.substring(0, exclamationIndex + 1) : '';

    // Initialize bounding box coordinates
    let minCol = Infinity, minRow = Infinity, maxCol = -Infinity, maxRow = -Infinity;

    for (const range of ranges) {
        // Remove sheet prefix if present
        let cellRange = range;
        if (exclamationIndex > -1) {
            cellRange = range.substring(range.lastIndexOf('!') + 1);
        }

        // Parse range (e.g., "$B$2:$B$10" or "$B$2")
        const rangeParts = cellRange.split(':');
        const startCell = rangeParts[0];
        const endCell = rangeParts[1] || startCell;

        // Parse start cell
        const startCoords = cellRefToCoords(startCell);
        if (startCoords) {
            minCol = Math.min(minCol, startCoords.col);
            minRow = Math.min(minRow, startCoords.row);
            maxCol = Math.max(maxCol, startCoords.col);
            maxRow = Math.max(maxRow, startCoords.row);
        }

        // Parse end cell
        if (rangeParts[1]) {
            const endCoords = cellRefToCoords(endCell);
            if (endCoords) {
                minCol = Math.min(minCol, endCoords.col);
                minRow = Math.min(minRow, endCoords.row);
                maxCol = Math.max(maxCol, endCoords.col);
                maxRow = Math.max(maxRow, endCoords.row);
            }
        }
    }

    // Build combined range
    const startCellName = getCellNameFromCoords(minCol, minRow);
    const endCellName = getCellNameFromCoords(maxCol, maxRow);

    return sheetPrefix + startCellName + ':' + endCellName;
}

// Generate a simple GUID/UUID
function generateGUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        const r = Math.random() * 16 | 0;
        const v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// Parse Excel range like "Sheet1!$C$2:$E$13" into coordinates
function parseExcelRange(range) {
    if (!range) return null;

    // Remove sheet name if present (Sheet1!$C$2:$E$13 -> $C$2:$E$13)
    const parts = range.split('!');
    const rangeStr = parts.length > 1 ? parts[1] : parts[0];

    // Parse range like $C$2:$E$13 or C2:E13
    const match = rangeStr.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
    if (!match) return null;

    const startCol = getColumnIndex(match[1]);
    const startRow = parseInt(match[2]) - 1;
    const endCol = getColumnIndex(match[3]);
    const endRow = parseInt(match[4]) - 1;

    return { startCol, startRow, endCol, endRow };
}

// Parse generic axis properties (for category axis)
function parseGenericAxis(axNode) {
    const axisConfig = {};

    // Check if axis is deleted/hidden
    const deleteNode = findNodes(axNode, 'c:delete')[0];
    if (deleteNode) {
        const deleteAttrs = parseAttributes(deleteNode);
        if (deleteAttrs.val === '1') {
            axisConfig.display = false;
        }
    }

    // Check if axis is reversed
    const scalingNode = findNodes(axNode, 'c:scaling')[0];
    if (scalingNode) {
        const orientationNode = findNodes(scalingNode, 'c:orientation')[0];
        if (orientationNode) {
            const orientationAttrs = parseAttributes(orientationNode);
            if (orientationAttrs.val === 'maxMin') {
                axisConfig.reverse = true;
            }
        }
    }

    // Parse axis title
    const titleNode = findNodes(axNode, 'c:title')[0];
    if (titleNode) {
        const txNode = findNodes(titleNode, 'c:tx')[0];
        if (txNode) {
            const richNode = findNodes(txNode, 'c:rich')[0];
            if (richNode) {
                const tNodes = findNodes(richNode, 'a:t');
                const titleText = tNodes.map(t => getTextContent(t)).join('');
                if (titleText) {
                    axisConfig.title = { text: titleText };
                }
            }
        }
    }

    return axisConfig;
}

// Parse linear axis properties (for value axis)
function parseLinearAxis(axNode) {
    const axisConfig = parseGenericAxis(axNode);

    // Parse min/max scaling
    const scalingNode = findNodes(axNode, 'c:scaling')[0];
    if (scalingNode) {
        const minNode = findNodes(scalingNode, 'c:min')[0];
        if (minNode) {
            const minAttrs = parseAttributes(minNode);
            if (minAttrs.val) {
                axisConfig.min = parseFloat(minAttrs.val);
            }
        }

        const maxNode = findNodes(scalingNode, 'c:max')[0];
        if (maxNode) {
            const maxAttrs = parseAttributes(maxNode);
            if (maxAttrs.val) {
                axisConfig.max = parseFloat(maxAttrs.val);
            }
        }

        if (axisConfig.min !== undefined || axisConfig.max !== undefined) {
            axisConfig.forceTheLimits = true;
        }
    }

    return axisConfig;
}

// Parse chart from anchor
async function parseChart(zip, sheetIndex, anchor, graphicFrameNode) {
    const chartObj = {
        id: generateGUID(),
        type: 'chart'
    };

    // Get chart relationship ID (c: namespace)
    const chartNode = findNodes(graphicFrameNode, 'c:chart')[0];
    if (!chartNode) return null;

    const chartAttrs = parseAttributes(chartNode);
    const rId = chartAttrs['r:id'];

    if (!rId) return null;

    // Load drawing relationships to find chart file
    const relsFile = zip.file(`xl/drawings/_rels/drawing${sheetIndex}.xml.rels`);
    if (!relsFile) return null;

    const relsXml = await relsFile.async('string');
    const relsParsed = parser(relsXml);
    const relationships = findNodes(relsParsed, 'Relationship');

    let chartPath = null;
    for (const rel of relationships) {
        const relAttrs = parseAttributes(rel);
        if (relAttrs.Id === rId) {
            chartPath = `xl/${relAttrs.Target.replace('../', '')}`;
            break;
        }
    }

    if (!chartPath) return null;

    // Load the chart XML file
    const chartFile = zip.file(chartPath);
    if (!chartFile) return null;

    const chartXml = await chartFile.async('string');
    const chartParsed = parser(chartXml);

    // Determine chart type (c: namespace for chart elements)
    const plotAreaNode = findNodes(chartParsed, 'c:plotArea')[0];
    if (!plotAreaNode) return null;

    let chartType = 'bar'; // default
    const chartOptions = {
        type: chartType
    };

    // Check for different chart types (c: namespace)
    if (findNodes(plotAreaNode, 'c:barChart')[0]) {
        chartType = 'bar';
        chartOptions.type = 'bar';
    } else if (findNodes(plotAreaNode, 'c:lineChart')[0]) {
        chartType = 'line';
        chartOptions.type = 'line';
    } else if (findNodes(plotAreaNode, 'c:pieChart')[0]) {
        chartType = 'pie';
        chartOptions.type = 'pie';
    } else if (findNodes(plotAreaNode, 'c:doughnutChart')[0]) {
        chartType = 'doughnut';
        chartOptions.type = 'doughnut';
    } else if (findNodes(plotAreaNode, 'c:areaChart')[0]) {
        chartType = 'area';
        chartOptions.type = 'area';
    } else if (findNodes(plotAreaNode, 'c:scatterChart')[0]) {
        chartType = 'scatter';
        chartOptions.type = 'scatter';
    } else if (findNodes(plotAreaNode, 'c:radarChart')[0]) {
        chartType = 'radar';
        chartOptions.type = 'radar';
    }

    // Get chart title (c: namespace)
    const titleNode = findNodes(chartParsed, 'c:title')[0];
    if (titleNode) {
        const txNode = findNodes(titleNode, 'tx')[0];
        if (txNode) {
            const richNode = findNodes(txNode, 'rich')[0];
            if (richNode) {
                const tNodes = findNodes(richNode, 't');
                const titleText = tNodes.map(t => getTextContent(t)).join('');
                if (titleText) {
                    chartOptions.title = {
                        display: true,
                        text: titleText
                    };
                }
            }
        }
    }

    // Get legend (c: namespace)
    const legendNode = findNodes(chartParsed, 'c:legend')[0];
    if (legendNode) {
        chartOptions.legend = {
            display: true
        };

        const legendPosNode = findNodes(legendNode, 'legendPos')[0];
        if (legendPosNode) {
            const posAttrs = parseAttributes(legendPosNode);
            const pos = posAttrs.val;
            // Map Excel legend positions to common positions
            const posMap = {
                'r': 'right',
                'l': 'left',
                't': 'top',
                'b': 'bottom',
                'tr': 'top'
            };
            chartOptions.legend.position = posMap[pos] || 'right';
        }
    }

    // Parse axis configuration (c: namespace)
    const catAxNode = findNodes(plotAreaNode, 'c:catAx')[0];
    const valAxNode = findNodes(plotAreaNode, 'c:valAx')[0];

    if (catAxNode && valAxNode) {
        const axis = {
            base: parseGenericAxis(catAxNode),
            side: parseLinearAxis(valAxNode)
        };
        chartOptions.axis = axis;
    }

    // Get series data and build combined range (c: namespace)
    // Following Jspreadsheet Pro format from parser.js
    const serNodes = findNodes(plotAreaNode, 'c:ser');
    const datasetRanges = [];
    let categoryRange = null;

    if (serNodes.length > 0) {
        const seriesStyles = [];

        for (let i = 0; i < serNodes.length; i++) {
            const serNode = serNodes[i];

            // Get data range reference for this series (c: namespace)
            const valNode = findNodes(serNode, 'c:val')[0];
            if (valNode) {
                const numRefNode = findNodes(valNode, 'c:numRef')[0];
                if (numRefNode) {
                    const fNode = findNodes(numRefNode, 'c:f')[0];
                    if (fNode) {
                        const dataRange = getTextContent(fNode);
                        if (dataRange) {
                            datasetRanges.push(dataRange);
                        }
                    }
                }
            }

            // Get category range (labels) - usually same for all series (c: namespace)
            if (!categoryRange) {
                const catNode = findNodes(serNode, 'c:cat')[0];
                if (catNode) {
                    const strRefNode = findNodes(catNode, 'c:strRef')[0] || findNodes(catNode, 'c:numRef')[0];
                    if (strRefNode) {
                        const fNode = findNodes(strRefNode, 'c:f')[0];
                        if (fNode) {
                            categoryRange = getTextContent(fNode);
                        }
                    }
                }
            }

            // Build series styling object (following Jspreadsheet format)
            const seriesStyle = {};

            // Get series color from spPr (shape properties)
            const spPrNode = findNodes(serNode, 'c:spPr')[0];
            if (spPrNode) {
                // Try to get fill color
                const solidFillNode = findNodes(spPrNode, 'a:solidFill')[0];
                if (solidFillNode) {
                    const colorNodes = solidFillNode.children || [];
                    if (colorNodes.length > 0) {
                        const color = getDrawingColor(colorNodes[0]);
                        if (color) seriesStyle.color = color;
                    }
                }

                // Try to get border color
                const lnNode = findNodes(spPrNode, 'a:ln')[0];
                if (lnNode) {
                    const lnSolidFillNode = findNodes(lnNode, 'a:solidFill')[0];
                    if (lnSolidFillNode) {
                        const colorNodes = lnSolidFillNode.children || [];
                        if (colorNodes.length > 0) {
                            const borderColor = getDrawingColor(colorNodes[0]);
                            if (borderColor) seriesStyle.borderColor = borderColor;
                        }
                    }
                }
            }

            seriesStyles.push(seriesStyle);
        }

        // Build the combined range following Jspreadsheet Pro format
        // The range should encompass: categories + all data series
        const allRanges = [...datasetRanges];
        if (categoryRange) {
            allRanges.push(categoryRange);
        }

        if (allRanges.length > 0) {
            // Calculate the combined bounding box range
            chartOptions.range = joinExcelRanges(allRanges);
            chartOptions.headers = true; // Excel charts typically have headers

            // Calculate orientation: true if column-based (vertical), false if row-based (horizontal)
            // Parse first dataset range to determine orientation
            if (datasetRanges.length > 0) {
                const firstDatasetCoords = parseExcelRange(datasetRanges[0]);
                if (firstDatasetCoords) {
                    // If start column === end column, it's a vertical range (column-based) = true
                    // If start row === end row, it's a horizontal range (row-based) = false
                    chartOptions.orientation = firstDatasetCoords.startCol === firstDatasetCoords.endCol;
                }
            }

            // Calculate dataset indices and labels index (offsets within the combined range)
            const totalRangeCoords = parseExcelRange(chartOptions.range);
            if (totalRangeCoords) {
                const coordIndex = chartOptions.orientation ? 0 : 1; // column index if vertical, row index if horizontal

                // Calculate dataset indices
                chartOptions.datasets = datasetRanges.map(range => {
                    const rangeCoords = parseExcelRange(range);
                    if (!rangeCoords) return 0;
                    const datasetPos = chartOptions.orientation ? rangeCoords.startCol : rangeCoords.startRow;
                    const totalStart = chartOptions.orientation ? totalRangeCoords.startCol : totalRangeCoords.startRow;
                    return datasetPos - totalStart;
                });

                // Calculate labels index
                if (categoryRange) {
                    const labelCoords = parseExcelRange(categoryRange);
                    if (labelCoords) {
                        const labelPos = chartOptions.orientation ? labelCoords.startCol : labelCoords.startRow;
                        const totalStart = chartOptions.orientation ? totalRangeCoords.startCol : totalRangeCoords.startRow;
                        chartOptions.labels = labelPos - totalStart;
                    }
                }
            }

            // Store series styling
            if (seriesStyles.length > 0) {
                chartOptions.series = seriesStyles;
            }
        }
    }

    chartObj.options = chartOptions;

    // Get anchor position
    const anchorData = getAnchorData(anchor);
    Object.assign(chartObj, anchorData);

    return chartObj;
}

// Convert EMU (English Metric Units) to pixels
// 1 EMU = 1/914400 inch, 1 pixel = 1/96 inch (standard DPI)
// So: pixels = EMU / (914400 / 96) = EMU / 9525
function emuToPixels(emu) {
    return Math.round(emu / 9525);
}

// Extract position data from anchor
function getAnchorData(anchor) {
    const data = {};

    // Get position/anchor information from xfrm (transform) (xdr: namespace)
    const xfrmNode = findNodes(anchor, 'xdr:xfrm')[0];
    if (xfrmNode) {
        const offNode = findNodes(xfrmNode, 'off')[0];
        const extNode = findNodes(xfrmNode, 'ext')[0];

        if (offNode) {
            const offAttrs = parseAttributes(offNode);
            // Convert EMU to pixels and use 'left' and 'top' properties
            const xEmu = parseInt(offAttrs.x) || 0;
            const yEmu = parseInt(offAttrs.y) || 0;
            data.left = emuToPixels(xEmu);
            data.top = emuToPixels(yEmu);
        }

        if (extNode) {
            const extAttrs = parseAttributes(extNode);
            // Convert EMU to pixels
            const widthEmu = parseInt(extAttrs.cx) || 0;
            const heightEmu = parseInt(extAttrs.cy) || 0;
            data.width = emuToPixels(widthEmu);
            data.height = emuToPixels(heightEmu);
        }
    }

    // If xfrm is not present, try to calculate from twoCellAnchor (xdr: namespace)
    if (!data.left && !data.top) {
        const fromNode = findNodes(anchor, 'xdr:from')[0];
        const toNode = findNodes(anchor, 'xdr:to')[0];

        if (fromNode && toNode) {
            // Extract from position (xdr: namespace)
            const fromCol = parseInt(getTextContent(findNodes(fromNode, 'xdr:col')[0] || {}) || 0);
            const fromRow = parseInt(getTextContent(findNodes(fromNode, 'xdr:row')[0] || {}) || 0);
            const fromColOff = parseInt(getTextContent(findNodes(fromNode, 'xdr:colOff')[0] || {}) || 0);
            const fromRowOff = parseInt(getTextContent(findNodes(fromNode, 'xdr:rowOff')[0] || {}) || 0);

            // Extract to position (xdr: namespace)
            const toCol = parseInt(getTextContent(findNodes(toNode, 'xdr:col')[0] || {}) || 0);
            const toRow = parseInt(getTextContent(findNodes(toNode, 'xdr:row')[0] || {}) || 0);
            const toColOff = parseInt(getTextContent(findNodes(toNode, 'xdr:colOff')[0] || {}) || 0);
            const toRowOff = parseInt(getTextContent(findNodes(toNode, 'xdr:rowOff')[0] || {}) || 0);

            // Approximate pixel position (simplified - uses average column/row size)
            // This is a rough estimate; real calculation would need actual column widths and row heights
            const avgColWidth = 64; // Average column width in pixels
            const avgRowHeight = 20; // Average row height in pixels

            data.left = Math.round(fromCol * avgColWidth + emuToPixels(fromColOff));
            data.top = Math.round(fromRow * avgRowHeight + emuToPixels(fromRowOff));
            data.width = Math.round((toCol - fromCol) * avgColWidth + emuToPixels(toColOff - fromColOff));
            data.height = Math.round((toRow - fromRow) * avgRowHeight + emuToPixels(toRowOff - fromRowOff));
        }
    }

    // Try to get cell anchor (from/to positions) for cellAnchor property (xdr: namespace)
    const fromNode2 = findNodes(anchor, 'xdr:from')[0];
    if (fromNode2) {
        const colNode = findNodes(fromNode2, 'xdr:col')[0];
        const rowNode = findNodes(fromNode2, 'xdr:row')[0];

        if (colNode && rowNode) {
            const col = parseInt(getTextContent(colNode));
            const row = parseInt(getTextContent(rowNode));
            data.cellAnchor = getCellNameFromCoords(col, row);
        }
    }

    return data;
}

// Parse comments for a worksheet
async function parseComments(zip, sheetIndex) {
    const commentsFile = zip.file(`xl/comments${sheetIndex}.xml`);
    if (!commentsFile) return {};

    const xml = await commentsFile.async('string');
    const parsed = parser(xml);
    const comments = {};

    const commentList = findNodes(parsed, 'comment');
    commentList.forEach(comment => {
        const attrs = parseAttributes(comment);
        const ref = attrs.ref;

        // Get comment text from <text> node
        const textNode = findNodes(comment, 'text')[0];
        if (textNode) {
            const tNodes = findNodes(textNode, 't');
            const commentText = tNodes.map(t => getTextContent(t)).join('');
            if (commentText) {
                comments[ref] = commentText;
            }
        }
    });

    return comments;
}

// Parse hyperlinks for a worksheet
async function parseHyperlinks(zip, sheetIndex, sheetParsed) {
    const hyperlinksNode = findNodes(sheetParsed, 'hyperlinks')[0];
    if (!hyperlinksNode) return {};

    const hyperlinks = {};

    // Load worksheet relationships to resolve hyperlink targets
    const relsFile = zip.file(`xl/worksheets/_rels/sheet${sheetIndex}.xml.rels`);
    if (!relsFile) return {};

    const relsXml = await relsFile.async('string');
    const relsParsed = parser(relsXml);

    // Get all hyperlink elements
    const hyperlinkNodes = findNodes(hyperlinksNode, 'hyperlink');

    hyperlinkNodes.forEach(hyperlink => {
        const attrs = parseAttributes(hyperlink);
        const ref = attrs.ref; // Cell reference like "A1"
        const rId = attrs['r:id']; // Relationship ID

        if (ref && rId) {
            // Find the relationship to get the actual URL
            const relationships = findNodes(relsParsed, 'Relationship');
            for (const rel of relationships) {
                const relAttrs = parseAttributes(rel);
                if (relAttrs.Id === rId) {
                    const target = relAttrs.Target;
                    if (target) {
                        hyperlinks[ref] = {
                            url: target
                        };
                    }
                    break;
                }
            }
        } else if (ref && attrs.location) {
            // Internal link (location attribute instead of r:id)
            hyperlinks[ref] = {
                url: attrs.location
            };
        }
    });

    return hyperlinks;
}

// Parse worksheet
async function parseWorksheet(zip, sheetPath, sheetIndex, sharedStrings, styles) {
    const file = zip.file(sheetPath);
    if (!file) return null;

    const xml = await file.async('string');
    const parsed = parser(xml);

    const result = {
        data: [],
        columns: [],
        rows: {},
        cells: {},
        style: {},
        mergeCells: {},
        comments: {},
        media: []
    };

    // Parse dimension to get range
    const dimension = findNodes(parsed, 'dimension')[0];
    let maxRow = 0;
    let maxCol = 0;

    if (dimension) {
        const ref = parseAttributes(dimension).ref;
        if (ref && ref.includes(':')) {
            const [, end] = ref.split(':');
            const coords = cellRefToCoords(end);
            if (coords) {
                maxRow = coords.row;
                maxCol = coords.col;
            }
        }
    }

    // Parse default column width from sheetFormatPr
    const sheetFormatPr = findNodes(parsed, 'sheetFormatPr')[0];
    if (sheetFormatPr) {
        const formatAttrs = parseAttributes(sheetFormatPr);
        if (formatAttrs.defaultColWidth) {
            const defaultWidth = parseFloat(formatAttrs.defaultColWidth);
            result.defaultColWidth = convertWidthToPixels(defaultWidth, 'char') + 'px';
        }
    }

    // Parse column definitions
    const cols = findNodes(parsed, 'col');
    cols.forEach(col => {
        const attrs = parseAttributes(col);
        const min = parseInt(attrs.min) - 1;
        let max = parseInt(attrs.max) - 1;
        const width = parseFloat(attrs.width) || 10;
        const hidden = attrs.hidden === '1';

        // Don't create columns beyond what's actually used - cap at maxCol
        // Excel often defines columns up to 16384 for default styling
        if (max > maxCol) {
            max = maxCol;
        }

        for (let c = min; c <= max; c++) {
            const column = {
                width: convertWidthToPixels(width, 'char'),
                title: getColumnName(c),
                type: 'text'
            };
            if (hidden) {
                column.visible = false;
            }
            result.columns[c] = column;
        }
    });

    // Parse rows
    const rows = findNodes(parsed, 'row');
    rows.forEach(row => {
        const attrs = parseAttributes(row);
        const rowNum = parseInt(attrs.r) - 1;

        const rowProps = {
            height: parseFloat(attrs.ht) || 21
        };
        if (attrs.hidden === '1') {
            rowProps.visible = false;
        }
        result.rows[rowNum] = rowProps;

        if (rowNum > maxRow) maxRow = rowNum;

        // Parse cells in row
        const cells = findNodes(row, 'c');
        cells.forEach(cell => {
            const cellAttrs = parseAttributes(cell);
            const ref = cellAttrs.r;
            const coords = cellRefToCoords(ref);

            if (!coords) return;

            if (coords.col > maxCol) maxCol = coords.col;

            const cellType = cellAttrs.t;
            const styleIdx = cellAttrs.s;

            let value = '';

            // Get cell value
            const vNode = findNodes(cell, 'v')[0];
            if (vNode) {
                const rawValue = getTextContent(vNode);

                if (cellType === 's') {
                    // Shared string
                    const idx = parseInt(rawValue);
                    value = sharedStrings[idx] || '';
                } else if (cellType === 'b') {
                    // Boolean
                    value = rawValue === '1';
                } else if (cellType === 'e') {
                    // Error
                    value = `#${rawValue}`;
                } else if (cellType === 'str') {
                    // String formula result
                    value = rawValue;
                } else {
                    // Number
                    value = parseFloat(rawValue) || rawValue;
                }
            }

            // Check for formula - if formula exists, use it in data instead of calculated value
            const fNode = findNodes(cell, 'f')[0];
            let finalValue = value;
            if (fNode) {
                const formula = getTextContent(fNode);
                if (formula) {
                    // Clean up Excel internal references and store formula in data with = prefix
                    const cleanedFormula = cleanFormula(formula);
                    finalValue = '=' + cleanedFormula;
                }
            }

            // Store value (formula if exists, otherwise raw value)
            if (!result.data[coords.row]) result.data[coords.row] = [];
            result.data[coords.row][coords.col] = finalValue;

            // Store style
            if (styleIdx) {
                const cssStyle = styleToCSS(parseInt(styleIdx), styles);
                if (cssStyle) {
                    result.style[ref] = cssStyle;
                }

                // Extract cell-level properties from style
                const xf = styles.cellXfs[parseInt(styleIdx)];
                if (xf) {
                    const cellProps = {};

                    // Extract number format
                    if (xf.numFmtId) {
                        const numFmtId = parseInt(xf.numFmtId);
                        let formatCode = null;

                        // Check for custom format in styles.numFmts
                        if (styles.numFmts[numFmtId]) {
                            formatCode = styles.numFmts[numFmtId];
                        } else {
                            // Built-in Excel number formats
                            const builtInFormats = {
                                1: '0',
                                2: '0.00',
                                3: '#,##0',
                                4: '#,##0.00',
                                9: '0%',
                                10: '0.00%',
                                11: '0.00E+00',
                                12: '# ?/?',
                                13: '# ??/??',
                                14: 'mm-dd-yy',
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
                            };
                            formatCode = builtInFormats[numFmtId];
                        }

                        if (formatCode) {
                            // Decode HTML entities in format code
                            formatCode = formatCode
                                .replace(/&amp;/g, '&')
                                .replace(/&lt;/g, '<')
                                .replace(/&gt;/g, '>')
                                .replace(/&quot;/g, '"')
                                .replace(/&apos;/g, "'");
                            cellProps.format = formatCode;
                        }
                    }

                    // Extract alignment
                    if (xf.alignment) {
                        if (xf.alignment.horizontal) {
                            cellProps.align = xf.alignment.horizontal;
                        }
                        if (xf.alignment.wrapText) {
                            cellProps.wrap = true;
                        }
                    }

                    // Store cell properties if any exist
                    if (Object.keys(cellProps).length > 0) {
                        result.cells[ref] = cellProps;
                    }
                }
            }
        });
    });

    // Parse merged cells
    const mergeCells = findNodes(parsed, 'mergeCell');
    mergeCells.forEach(merge => {
        const ref = parseAttributes(merge).ref;
        if (ref && ref.includes(':')) {
            const [start, end] = ref.split(':');
            const startCoords = cellRefToCoords(start);
            const endCoords = cellRefToCoords(end);

            if (startCoords && endCoords) {
                const colspan = endCoords.col - startCoords.col + 1;
                const rowspan = endCoords.row - startCoords.row + 1;
                result.mergeCells[start] = [colspan, rowspan];

                // Update maxRow and maxCol if merged cell extends beyond current max
                if (endCoords.row > maxRow) maxRow = endCoords.row;
                if (endCoords.col > maxCol) maxCol = endCoords.col;
            }
        }
    });

    // Parse comments for this worksheet
    const comments = await parseComments(zip, sheetIndex);
    if (Object.keys(comments).length > 0) {
        result.comments = comments;
    }

    // Parse hyperlinks for this worksheet
    const hyperlinks = await parseHyperlinks(zip, sheetIndex, parsed);
    if (Object.keys(hyperlinks).length > 0) {
        // Merge hyperlinks into cells object
        Object.keys(hyperlinks).forEach(cellRef => {
            if (!result.cells[cellRef]) {
                result.cells[cellRef] = {};
            }
            // Add url via options object (consistent with parser.js line 2399)
            if (!result.cells[cellRef].options) {
                result.cells[cellRef].options = {};
            }
            result.cells[cellRef].options.url = hyperlinks[cellRef].url;
            // Set type to text if not already set
            if (!result.cells[cellRef].type) {
                result.cells[cellRef].type = 'text';
            }
        });
    }

    // Parse grid lines visibility
    const sheetViews = findNodes(parsed, 'sheetView');
    for (const sheetView of sheetViews) {
        const viewAttrs = parseAttributes(sheetView);
        if (viewAttrs.showGridLines === '0') {
            result.gridline = false;
            break;
        }
    }

    // Parse sheet protection
    const sheetProtectionNode = findNodes(parsed, 'sheetProtection')[0];
    if (sheetProtectionNode) {
        const protAttrs = parseAttributes(sheetProtectionNode);

        // Only process if sheet protection is enabled
        if (protAttrs.sheet === '1') {
            const configObj = {};

            // Check what operations are allowed (0 means restricted, so we set allow to true)
            if (protAttrs.insertColumns === '0') {
                configObj.allowInsertColumn = true;
            }
            if (protAttrs.deleteColumns === '0') {
                configObj.allowDeleteColumn = true;
            }
            if (protAttrs.insertRows === '0') {
                configObj.allowInsertRow = true;
            }
            if (protAttrs.deleteRows === '0') {
                configObj.allowDeleteRow = true;
            }

            // Set locked property
            result.locked = Object.keys(configObj).length > 0 ? configObj : true;

            // Handle selection restrictions
            if (protAttrs.selectUnlockedCells === '1') {
                result.selectUnLockedCells = false;
            }
            if (protAttrs.selectLockedCells === '1') {
                result.selectLockedCells = false;
            }

            // Parse protected ranges (cell-level protection)
            const protectedRangesNode = findNodes(parsed, 'protectedRanges')[0];
            if (protectedRangesNode) {
                const protectedRangeNodes = findNodes(protectedRangesNode, 'protectedRange');

                protectedRangeNodes.forEach(rangeNode => {
                    const rangeAttrs = parseAttributes(rangeNode);
                    const sqref = rangeAttrs.sqref;

                    if (sqref) {
                        // Parse the range (e.g., "A1:B10" or "A1")
                        const ranges = sqref.split(' ');
                        ranges.forEach(range => {
                            if (range.includes(':')) {
                                // Range like A1:B10
                                const [start, end] = range.split(':');
                                const startCoords = cellRefToCoords(start);
                                const endCoords = cellRefToCoords(end);

                                if (startCoords && endCoords) {
                                    // Mark all cells in range as locked
                                    for (let r = startCoords.row; r <= endCoords.row; r++) {
                                        for (let c = startCoords.col; c <= endCoords.col; c++) {
                                            const cellRef = getCellNameFromCoords(c, r);
                                            if (!result.cells[cellRef]) {
                                                result.cells[cellRef] = {};
                                            }
                                            result.cells[cellRef].locked = true;
                                        }
                                    }
                                }
                            } else {
                                // Single cell like A1
                                if (!result.cells[range]) {
                                    result.cells[range] = {};
                                }
                                result.cells[range].locked = true;
                            }
                        });
                    }
                });
            }
        }
    }

    // Parse frozen panes
    const paneNode = findNodes(parsed, 'pane').find(pane => {
        const attrs = parseAttributes(pane);
        return attrs.state === 'frozen' || attrs.state === 'frozenSplit';
    });

    if (paneNode) {
        const paneAttrs = parseAttributes(paneNode);

        // Get the first visible cell in the unfrozen area (topLeftCell from parent sheetView)
        const sheetViews = findNodes(parsed, 'sheetView');
        let topLeftCell = null;

        for (const sheetView of sheetViews) {
            const viewAttrs = parseAttributes(sheetView);
            if (viewAttrs.topLeftCell) {
                topLeftCell = viewAttrs.topLeftCell;
                break;
            }
        }

        let firstFrozenColumn = 0;
        let firstFrozenRow = 0;

        if (topLeftCell) {
            const coords = cellRefToCoords(topLeftCell);
            if (coords) {
                firstFrozenColumn = coords.col;
                firstFrozenRow = coords.row;
            }
        }

        // Parse frozen columns (xSplit)
        const xSplit = paneAttrs.xSplit;
        if (xSplit) {
            const xCoord = parseInt(xSplit);
            const freezeColumns = [];
            for (let x = 0; x < xCoord; x++) {
                freezeColumns.push(x + firstFrozenColumn);
            }
            result.freezeColumns = freezeColumns;
        }

        // Parse frozen rows (ySplit)
        const ySplit = paneAttrs.ySplit;
        if (ySplit) {
            const yCoord = parseInt(ySplit);
            const freezeRows = [];
            for (let y = 0; y < yCoord; y++) {
                freezeRows.push(y + firstFrozenRow);
            }
            result.freezeRows = freezeRows;
        }
    }

    // Parse data validations - will be collected and returned separately
    const validations = [];
    const dataValidations = findNodes(parsed, 'dataValidation');
    dataValidations.forEach(validation => {
        const attrs = parseAttributes(validation);
        const sqref = attrs.sqref; // Cell range like "A1:A10"

        if (sqref) {
            const validationObj = {
                range: sqref
            };

            // Map Excel type to Jspreadsheet type using helper mappings
            const excelType = attrs.type;
            if (excelType && excelValidationTypes[excelType]) {
                validationObj.type = excelValidationTypes[excelType];
            } else if (excelType === 'custom') {
                validationObj.type = 'formula';
            } else {
                validationObj.type = 'text'; // Default
            }

            // Map Excel error style to Jspreadsheet action
            // stop -> reject, warning -> warning, information -> warning
            const errorStyle = attrs.errorStyle || 'stop';
            validationObj.action = errorStyle === 'stop' ? 'reject' : 'warning';

            // Map Excel operator to Jspreadsheet criteria using helper mappings
            const operator = attrs.operator;
            if (operator && excelValidationOperations[operator]) {
                validationObj.criteria = excelValidationOperations[operator];
            }

            // Error message as text
            if (attrs.error) {
                validationObj.text = attrs.error;
            } else if (attrs.errorTitle) {
                validationObj.text = attrs.errorTitle;
            }

            // Allow blank
            if (attrs.allowBlank === '1') {
                validationObj.allowBlank = true;
            }

            // Parse formula nodes
            const formula1Node = findNodes(validation, 'formula1')[0];
            const formula2Node = findNodes(validation, 'formula2')[0];

            if (formula1Node) {
                const formula1 = cleanFormula(getTextContent(formula1Node));

                if (excelType === 'list') {
                    // For list type, convert to dropdown with value array
                    validationObj.dropdown = true;
                    // If formula is a simple list (e.g., "Item1,Item2,Item3")
                    if (formula1 && !formula1.includes('!') && !formula1.includes(':')) {
                        // Split by comma and clean up quotes
                        validationObj.value = formula1.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
                    } else {
                        // It's a range reference, keep as formula
                        validationObj.value = [formula1];
                    }
                } else {
                    // For other types, store as value array
                    const values = [formula1];
                    if (formula2Node) {
                        const formula2 = cleanFormula(getTextContent(formula2Node));
                        values.push(formula2);
                    }
                    validationObj.value = values;
                }
            }

            validations.push(validationObj);
        }
    });

    // Store validations in result for collection at spreadsheet level
    if (validations.length > 0) {
        result.validations = validations;
    }

    // Parse conditional formatting (stored as validations with action:'format')
    const conditionalFormattingNodes = findNodes(parsed, 'conditionalFormatting');

    conditionalFormattingNodes.forEach(cfNode => {
        const cfAttrs = parseAttributes(cfNode);
        const sqref = cfAttrs.sqref; // Cell range like "A1:A10"

        if (!sqref) return;

        // Get all cfRule elements within this conditional formatting
        const cfRuleNodes = findNodes(cfNode, 'cfRule');

        cfRuleNodes.forEach(cfRule => {
            const ruleAttrs = parseAttributes(cfRule);
            const ruleType = ruleAttrs.type;

            if (!ruleType) return;

            const cfObj = {
                range: sqref,
                action: 'format' // Conditional formatting uses action:'format'
            };

            // Map rule type to Jspreadsheet format
            const excelCFSimpleTypes = {
                'containsBlanks': 'empty',
                'notContainsBlanks': 'notEmpty'
            };

            const excelCFTextTypes = {
                'containsText': 'contains',
                'notContainsText': 'not contains',
                'beginsWith': 'begins with',
                'endsWith': 'ends with'
            };

            const excelCFNumericOperators = {
                'equal': '=',
                'notEqual': '!=',
                'greaterThan': '>',
                'lessThan': '<',
                'greaterThanOrEqual': '>=',
                'lessThanOrEqual': '<=',
                'between': 'between',
                'notBetween': 'not between'
            };

            // Handle simple types (empty/notEmpty)
            if (excelCFSimpleTypes[ruleType]) {
                cfObj.type = excelCFSimpleTypes[ruleType];
            }
            // Handle text types (contains, begins with, etc.)
            else if (excelCFTextTypes[ruleType]) {
                cfObj.type = 'text';
                cfObj.criteria = excelCFTextTypes[ruleType];
                if (ruleAttrs.text) {
                    cfObj.value = [ruleAttrs.text];
                }
            }
            // Handle cellIs (numeric comparisons)
            else if (ruleType === 'cellIs') {
                const operator = ruleAttrs.operator;
                const criteria = excelCFNumericOperators[operator];

                if (criteria) {
                    const formulaNodes = findNodes(cfRule, 'formula');
                    const values = [];

                    formulaNodes.forEach(fNode => {
                        const formulaText = getTextContent(fNode);
                        if (formulaText) {
                            // Try to parse as number
                            const numValue = parseFloat(formulaText);
                            if (!isNaN(numValue)) {
                                values.push(numValue);
                            } else if (formulaText.startsWith('"') && formulaText.endsWith('"')) {
                                // Text value in quotes
                                cfObj.type = 'text';
                                values.push(formulaText.slice(1, -1));
                            }
                        }
                    });

                    cfObj.type = cfObj.type || 'number';
                    cfObj.criteria = criteria;
                    if (values.length > 0) {
                        cfObj.value = values;
                    }
                }
            }
            // Handle expression (formula-based)
            else if (ruleType === 'expression') {
                const formulaNode = findNodes(cfRule, 'formula')[0];
                if (formulaNode) {
                    const formula = getTextContent(formulaNode);
                    if (formula) {
                        cfObj.type = 'formula';
                        cfObj.value = ['=' + cleanFormula(formula)];
                    }
                }
            }
            // Handle data bars, color scales, icon sets
            else if (ruleType === 'dataBar') {
                cfObj.type = 'dataBar';
                // Additional data bar properties could be parsed here
            }
            else if (ruleType === 'colorScale') {
                cfObj.type = 'colorScale';
                // Additional color scale properties could be parsed here
            }
            else if (ruleType === 'iconSet') {
                cfObj.type = 'iconSet';
                // Additional icon set properties could be parsed here
            }

            // Get style from dxfId
            const dxfId = ruleAttrs.dxfId;
            if (dxfId && styles.dxfs && styles.dxfs[parseInt(dxfId)]) {
                const dxf = styles.dxfs[parseInt(dxfId)];
                const cssProps = [];

                if (dxf.fontColor && dxf.fontColor.startsWith('FF')) {
                    cssProps.push(`color: #${dxf.fontColor.substring(2)}`);
                }
                if (dxf.bgColor && dxf.bgColor.startsWith('FF')) {
                    cssProps.push(`background-color: #${dxf.bgColor.substring(2)}`);
                }
                if (dxf.bold) {
                    cssProps.push('font-weight: bold');
                }
                if (dxf.italic) {
                    cssProps.push('font-style: italic');
                }

                // Handle borders
                if (dxf.border) {
                    ['left', 'right', 'top', 'bottom'].forEach(side => {
                        if (dxf.border[side]) {
                            const borderInfo = dxf.border[side];
                            const color = borderInfo.color && borderInfo.color.startsWith('FF')
                                ? `#${borderInfo.color.substring(2)}`
                                : '#000';
                            const cssStyle = borderStyles[borderInfo.style] || borderStyles.thin;
                            cssProps.push(`border-${side}: ${cssStyle[1]} ${cssStyle[0]} ${color}`);
                        }
                    });
                }

                if (cssProps.length > 0) {
                    cfObj.format = {};
                    cssProps.forEach(prop => {
                        const [key, value] = prop.split(':').map(s => s.trim());
                        cfObj.format[key] = value;
                    });
                }
            }

            // Only add if we have a valid type - push to validations array with action:'format'
            if (cfObj.type) {
                validations.push(cfObj);
            }
        });
    });

    // Parse drawings/images for this worksheet
    const media = await parseDrawings(zip, sheetIndex);
    if (media.length > 0) {
        result.media = media;
    }

    // Fill empty cells
    for (let r = 0; r <= maxRow; r++) {
        if (!result.data[r]) result.data[r] = [];
        for (let c = 0; c <= maxCol; c++) {
            if (result.data[r][c] === undefined) {
                result.data[r][c] = '';
            }
        }
    }

    // Fill missing columns
    for (let c = 0; c <= maxCol; c++) {
        if (!result.columns[c]) {
            result.columns[c] = {
                width: 100,
                title: getColumnName(c),
                type: 'text'
            };
        }
    }

    // Add minDimensions [columns, rows]
    result.minDimensions = [maxCol + 1, maxRow + 1];

    return result;
}

// Parse workbook
async function parseWorkbook(zip) {
    const file = zip.file('xl/workbook.xml');
    if (!file) throw new Error('workbook.xml not found');

    const xml = await file.async('string');
    const parsed = parser(xml);

    const sheets = [];
    const sheetNodes = findNodes(parsed, 'sheet');

    sheetNodes.forEach(sheet => {
        const attrs = parseAttributes(sheet);
        sheets.push({
            name: attrs.name,
            sheetId: attrs.sheetId,
            id: attrs['r:id'],
            state: attrs.state // hidden, veryHidden, or undefined for visible
        });
    });

    // Parse defined names (named ranges) - include ALL names
    const definedNames = {};
    const definedNameNodes = findNodes(parsed, 'definedName');
    definedNameNodes.forEach(node => {
        const attrs = parseAttributes(node);
        let name = attrs.name;
        const content = getTextContent(node);

        if (name && content) {
            definedNames[name] = cleanFormula(content);
        }
    });

    return { sheets, definedNames };
}

// Main XLSX parser
/**
 * Parse XLSX file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseXLSX(input, options = {}) {
    return parse(async (inp) => {
        // JSZip can handle File objects directly in browser, or buffers
        let zipInput;
        if (typeof File !== 'undefined' && inp instanceof File) {
            zipInput = inp; // JSZip handles File objects directly
        } else {
            zipInput = await loadAsBuffer(inp);
        }

        const zip = await JSZip.loadAsync(zipInput);

    // Parse shared strings
    const sharedStrings = await parseSharedStrings(zip);

    // Parse styles
    const styles = await parseStyles(zip);

    // Parse workbook to get sheet names and defined names
    const { sheets, definedNames } = await parseWorkbook(zip);

    // Parse each worksheet
    const worksheets = [];
    const allValidations = []; // Collect validations (includes conditional formatting with action:'format')
    const globalStyles = []; // Global style array
    const styleMap = new Map(); // CSS string -> index

    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        const sheetPath = `xl/worksheets/sheet${i + 1}.xml`;
        const sheetData = await parseWorksheet(zip, sheetPath, i + 1, sharedStrings, styles);

        if (sheetData) {
            const worksheet = {
                worksheetName: sheet.name,
                ...sheetData
            };

            // Apply border deduplication to prevent double borders between adjacent cells
            // Track cells that should NOT have top/left borders
            const cellsWithoutTopBorder = new Set();
            const cellsWithoutLeftBorder = new Set();

            if (worksheet.style) {
                Object.keys(worksheet.style).forEach(cellRef => {
                    const cssString = worksheet.style[cellRef];

                    // Check if this cell has border-right or border-bottom
                    if (cssString.includes('border-right:')) {
                        // Mark the cell to the right to not have border-left
                        const coords = cellRefToCoords(cellRef);
                        if (coords) {
                            const rightCell = getCellNameFromCoords(coords.col + 1, coords.row);
                            cellsWithoutLeftBorder.add(rightCell);
                        }
                    }
                    if (cssString.includes('border-bottom:')) {
                        // Mark the cell below to not have border-top
                        const coords = cellRefToCoords(cellRef);
                        if (coords) {
                            const belowCell = getCellNameFromCoords(coords.col, coords.row + 1);
                            cellsWithoutTopBorder.add(belowCell);
                        }
                    }
                });

                // Now remove duplicate borders
                Object.keys(worksheet.style).forEach(cellRef => {
                    let cssString = worksheet.style[cellRef];

                    // Remove border-top if this cell should not have one
                    if (cellsWithoutTopBorder.has(cellRef)) {
                        cssString = cssString.replace(/border-top:\s*[^;]+;?\s*/g, '');
                    }

                    // Remove border-left if this cell should not have one
                    if (cellsWithoutLeftBorder.has(cellRef)) {
                        cssString = cssString.replace(/border-left:\s*[^;]+;?\s*/g, '');
                    }

                    worksheet.style[cellRef] = cssString;
                });
            }

            // Convert worksheet styles from CSS strings to global style indices
            const worksheetStyleIndices = {};
            if (worksheet.style) {
                Object.keys(worksheet.style).forEach(cellRef => {
                    const cssString = worksheet.style[cellRef];

                    // Get or create style index
                    if (!styleMap.has(cssString)) {
                        styleMap.set(cssString, globalStyles.length);
                        globalStyles.push(cssString);
                    }

                    worksheetStyleIndices[cellRef] = styleMap.get(cssString);
                });
                worksheet.style = worksheetStyleIndices;
            }

            // Collect validations from this worksheet (includes conditional formatting)
            if (worksheet.validations) {
                // Add worksheet name to each validation/conditional formatting
                worksheet.validations.forEach(validation => {
                    // Prefix the range with worksheet name if not already qualified
                    if (!validation.range.includes('!')) {
                        validation.range = `${sheet.name}!${validation.range}`;
                    }
                    allValidations.push(validation);
                });
                delete worksheet.validations; // Remove from worksheet level
            }

            // Add worksheetState if sheet is hidden
            if (sheet.state === 'hidden' || sheet.state === 'veryHidden') {
                worksheet.worksheetState = 'hidden';
            }

            worksheets.push(worksheet);
        }
    }

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
