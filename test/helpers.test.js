import { expect } from 'chai';
import {
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
} from '../src/utils/helpers.js';

describe('Column and Cell Helpers', () => {
    describe('getColumnName', () => {
        it('should convert 0 to A', () => {
            expect(getColumnName(0)).to.equal('A');
        });

        it('should convert 25 to Z', () => {
            expect(getColumnName(25)).to.equal('Z');
        });

        it('should convert 26 to AA', () => {
            expect(getColumnName(26)).to.equal('AA');
        });

        it('should convert 701 to ZZ', () => {
            expect(getColumnName(701)).to.equal('ZZ');
        });

        it('should convert 702 to AAA', () => {
            expect(getColumnName(702)).to.equal('AAA');
        });

        it('should handle string input', () => {
            expect(getColumnName('5')).to.equal('F');
        });

        it('should use cache for repeated calls', () => {
            const result1 = getColumnName(10);
            const result2 = getColumnName(10);
            expect(result1).to.equal(result2);
            expect(result1).to.equal('K');
        });
    });

    describe('getColumnIndex', () => {
        it('should convert A to 0', () => {
            expect(getColumnIndex('A')).to.equal(0);
        });

        it('should convert Z to 25', () => {
            expect(getColumnIndex('Z')).to.equal(25);
        });

        it('should convert AA to 26', () => {
            expect(getColumnIndex('AA')).to.equal(26);
        });

        it('should convert ZZ to 701', () => {
            expect(getColumnIndex('ZZ')).to.equal(701);
        });

        it('should convert AAA to 702', () => {
            expect(getColumnIndex('AAA')).to.equal(702);
        });
    });

    describe('getCellNameFromCoords', () => {
        it('should convert (0, 0) to A1', () => {
            expect(getCellNameFromCoords(0, 0)).to.equal('A1');
        });

        it('should convert (5, 10) to F11', () => {
            expect(getCellNameFromCoords(5, 10)).to.equal('F11');
        });

        it('should convert (26, 0) to AA1', () => {
            expect(getCellNameFromCoords(26, 0)).to.equal('AA1');
        });

        it('should handle string inputs', () => {
            expect(getCellNameFromCoords('5', '10')).to.equal('F11');
        });
    });

    describe('getCoordsFromCellName', () => {
        it('should convert A1 to [0, 0]', () => {
            expect(getCoordsFromCellName('A1')).to.deep.equal([0, 0]);
        });

        it('should convert F11 to [5, 10]', () => {
            expect(getCoordsFromCellName('F11')).to.deep.equal([5, 10]);
        });

        it('should convert AA1 to [26, 0]', () => {
            expect(getCoordsFromCellName('AA1')).to.deep.equal([26, 0]);
        });

        it('should convert Z100 to [25, 99]', () => {
            expect(getCoordsFromCellName('Z100')).to.deep.equal([25, 99]);
        });

        it('should return empty array for empty input', () => {
            expect(getCoordsFromCellName('')).to.deep.equal([]);
        });

        it('should return empty array for null input', () => {
            expect(getCoordsFromCellName(null)).to.deep.equal([]);
        });
    });

    describe('getCoordsFromRange', () => {
        it('should convert A1:B2 to [0, 0, 1, 1]', () => {
            expect(getCoordsFromRange('A1:B2')).to.deep.equal([0, 0, 1, 1]);
        });

        it('should convert single cell A1 to [0, 0, 0, 0]', () => {
            expect(getCoordsFromRange('A1')).to.deep.equal([0, 0, 0, 0]);
        });

        it('should handle sheet names with ! separator', () => {
            expect(getCoordsFromRange('Sheet1!A1:B2')).to.deep.equal([0, 0, 1, 1]);
        });

        it('should handle column-only ranges A:A', () => {
            expect(getCoordsFromRange('A:A')).to.deep.equal([0, 0, 0, 0]);
        });

        it('should handle row-only ranges 1:1', () => {
            expect(getCoordsFromRange('1:1')).to.deep.equal([0, 0, 0, 0]);
        });

        it('should handle large ranges', () => {
            expect(getCoordsFromRange('AA10:ZZ100')).to.deep.equal([26, 9, 701, 99]);
        });
    });

    describe('getTokensFromCoords', () => {
        it('should return single cell for single coord', () => {
            expect(getTokensFromCoords(0, 0, 0, 0)).to.deep.equal(['A1']);
        });

        it('should return array of cells for range', () => {
            const tokens = getTokensFromCoords(0, 0, 1, 1);
            expect(tokens).to.deep.equal(['A1', 'B1', 'A2', 'B2']);
        });

        it('should include worksheet name when provided', () => {
            const tokens = getTokensFromCoords(0, 0, 1, 0, 'Sheet1');
            expect(tokens).to.deep.equal(['Sheet1!A1', 'Sheet1!B1']);
        });
    });

    describe('tokenIdentifier', () => {
        it('should return true for valid cell reference A1', () => {
            expect(tokenIdentifier('A1')).to.be.true;
        });

        it('should return true for valid range A1:B2', () => {
            expect(tokenIdentifier('A1:B2')).to.be.true;
        });

        it('should return true for column range A:A', () => {
            expect(tokenIdentifier('A:A')).to.be.true;
        });

        it('should return true for row range 1:1', () => {
            expect(tokenIdentifier('1:1')).to.be.true;
        });

        it('should return true for absolute reference $A$1', () => {
            expect(tokenIdentifier('$A$1')).to.be.true;
        });

        it('should return true for mixed reference $A1', () => {
            expect(tokenIdentifier('$A1')).to.be.true;
        });

        it('should return true for mixed reference A$1', () => {
            expect(tokenIdentifier('A$1')).to.be.true;
        });

        it('should return true for sheet reference Sheet1!A1', () => {
            expect(tokenIdentifier('Sheet1!A1')).to.be.true;
        });

        it('should return true for quoted sheet reference \'My Sheet\'!A1', () => {
            expect(tokenIdentifier('\'My Sheet\'!A1')).to.be.true;
        });

        it('should return false for empty string', () => {
            expect(tokenIdentifier('')).to.be.false;
        });

        it('should return false for null', () => {
            expect(tokenIdentifier(null)).to.be.false;
        });

        it('should return false for number', () => {
            expect(tokenIdentifier(123)).to.be.false;
        });

        it('should return false for invalid reference with colon in sheet name', () => {
            expect(tokenIdentifier('\'Sheet:1\'!A1')).to.be.false;
        });

        it('should return false for more than 3 letters in column', () => {
            expect(tokenIdentifier('AAAA1')).to.be.false;
        });
    });
});

describe('HTML/XML Helper Functions', () => {
    describe('getProp', () => {
        it('should return prop value when found', () => {
            const node = {
                props: [
                    { name: 'id', value: '123' },
                    { name: 'class', value: 'test' }
                ]
            };
            expect(getProp(node, 'id')).to.equal('123');
        });

        it('should return undefined when prop not found', () => {
            const node = { props: [{ name: 'id', value: '123' }] };
            expect(getProp(node, 'class')).to.be.undefined;
        });

        it('should return undefined when node has no props', () => {
            const node = {};
            expect(getProp(node, 'id')).to.be.undefined;
        });

        it('should handle namespaced props', () => {
            const node = {
                props: [{ name: 'xml:id', value: '123' }]
            };
            expect(getProp(node, 'id')).to.equal('123');
        });
    });

    describe('getChildren', () => {
        it('should return children of specified type', () => {
            const node = {
                children: [
                    { type: 'div', value: '1' },
                    { type: 'span', value: '2' },
                    { type: 'div', value: '3' }
                ]
            };
            const divs = getChildren(node, 'div');
            expect(divs).to.have.lengthOf(2);
            expect(divs[0].value).to.equal('1');
            expect(divs[1].value).to.equal('3');
        });

        it('should return empty array when no children match', () => {
            const node = {
                children: [{ type: 'div', value: '1' }]
            };
            expect(getChildren(node, 'span')).to.deep.equal([]);
        });

        it('should return empty array when node has no children', () => {
            const node = {};
            expect(getChildren(node, 'div')).to.deep.equal([]);
        });

        it('should handle namespaced types', () => {
            const node = {
                children: [{ type: 'xml:div', value: '1' }]
            };
            expect(getChildren(node, 'div')).to.have.lengthOf(1);
        });
    });

    describe('getTextContent', () => {
        it('should return empty string for null', () => {
            expect(getTextContent(null)).to.equal('');
        });

        it('should return string as is', () => {
            expect(getTextContent('hello')).to.equal('hello');
        });

        it('should join array of strings', () => {
            expect(getTextContent(['hello', ' ', 'world'])).to.equal('hello world');
        });

        it('should extract text from #text node', () => {
            const node = {
                type: '#text',
                props: [{ name: 'textContent', value: 'hello' }]
            };
            expect(getTextContent(node)).to.equal('hello');
        });

        it('should recursively extract text from children', () => {
            const node = {
                children: [
                    { type: '#text', props: [{ name: 'textContent', value: 'hello' }] },
                    { type: '#text', props: [{ name: 'textContent', value: ' world' }] }
                ]
            };
            expect(getTextContent(node)).to.equal('hello world');
        });

        it('should handle nested structures', () => {
            const node = {
                children: [
                    'hello',
                    { children: [' ', 'world'] }
                ]
            };
            expect(getTextContent(node)).to.equal('hello world');
        });
    });

    describe('parseAttributes', () => {
        it('should parse props into object', () => {
            const element = {
                props: [
                    { name: 'id', value: '123' },
                    { name: 'class', value: 'test' }
                ]
            };
            expect(parseAttributes(element)).to.deep.equal({
                id: '123',
                class: 'test'
            });
        });

        it('should handle attributes fallback', () => {
            const element = {
                attributes: [
                    { name: 'id', value: '456' }
                ]
            };
            expect(parseAttributes(element)).to.deep.equal({
                id: '456'
            });
        });

        it('should return empty object for null', () => {
            expect(parseAttributes(null)).to.deep.equal({});
        });

        it('should return empty object for element without props', () => {
            expect(parseAttributes({})).to.deep.equal({});
        });
    });

    describe('findNodes', () => {
        it('should find nodes by type', () => {
            const tree = {
                type: 'root',
                children: [
                    { type: 'div', value: '1' },
                    {
                        type: 'span',
                        children: [
                            { type: 'div', value: '2' }
                        ]
                    }
                ]
            };
            const divs = findNodes(tree, 'div');
            expect(divs).to.have.lengthOf(2);
            expect(divs[0].value).to.equal('1');
            expect(divs[1].value).to.equal('2');
        });

        it('should handle array input', () => {
            const nodes = [
                { type: 'div', value: '1' },
                { type: 'span', value: '2' }
            ];
            const divs = findNodes(nodes, 'div');
            expect(divs).to.have.lengthOf(1);
        });

        it('should return empty array when no matches', () => {
            const tree = { type: 'root' };
            expect(findNodes(tree, 'div')).to.deep.equal([]);
        });

        it('should handle null input', () => {
            expect(findNodes(null, 'div')).to.deep.equal([]);
        });
    });
});

describe('Binary Reading Helper Functions', () => {
    describe('readUInt8', () => {
        it('should read unsigned 8-bit integer', () => {
            const buffer = new Uint8Array([255, 128, 0]);
            expect(readUInt8(buffer, 0)).to.equal(255);
            expect(readUInt8(buffer, 1)).to.equal(128);
            expect(readUInt8(buffer, 2)).to.equal(0);
        });
    });

    describe('readUInt16LE', () => {
        it('should read unsigned 16-bit integer little-endian', () => {
            const buffer = new Uint8Array([0xFF, 0xFF, 0x00, 0x01]);
            expect(readUInt16LE(buffer, 0)).to.equal(65535);
            expect(readUInt16LE(buffer, 2)).to.equal(256);
        });

        it('should handle zero', () => {
            const buffer = new Uint8Array([0x00, 0x00]);
            expect(readUInt16LE(buffer, 0)).to.equal(0);
        });
    });

    describe('readUInt32LE', () => {
        it('should read unsigned 32-bit integer little-endian', () => {
            const buffer = new Uint8Array([0xFF, 0xFF, 0xFF, 0xFF]);
            expect(readUInt32LE(buffer, 0)).to.equal(4294967295);
        });

        it('should handle smaller values', () => {
            const buffer = new Uint8Array([0x01, 0x00, 0x00, 0x00]);
            expect(readUInt32LE(buffer, 0)).to.equal(1);
        });
    });

    describe('readInt16LE', () => {
        it('should read signed 16-bit integer little-endian', () => {
            const buffer = new Uint8Array([0xFF, 0xFF, 0x00, 0x80]);
            expect(readInt16LE(buffer, 0)).to.equal(-1);
            expect(readInt16LE(buffer, 2)).to.equal(-32768);
        });

        it('should handle positive values', () => {
            const buffer = new Uint8Array([0x00, 0x7F]);
            expect(readInt16LE(buffer, 0)).to.equal(32512);
        });
    });

    describe('readFloat64LE', () => {
        it('should read 64-bit float little-endian', () => {
            const buffer = new Uint8Array(8);
            const view = new DataView(buffer.buffer);
            view.setFloat64(0, 3.14159, true);
            expect(readFloat64LE(buffer, 0)).to.be.closeTo(3.14159, 0.00001);
        });

        it('should handle zero', () => {
            const buffer = new Uint8Array(8);
            expect(readFloat64LE(buffer, 0)).to.equal(0);
        });

        it('should handle negative values', () => {
            const buffer = new Uint8Array(8);
            const view = new DataView(buffer.buffer);
            view.setFloat64(0, -123.456, true);
            expect(readFloat64LE(buffer, 0)).to.be.closeTo(-123.456, 0.001);
        });
    });
});

describe('String Utility Functions', () => {
    describe('decodeHTMLEntities', () => {
        it('should decode &quot; to "', () => {
            expect(decodeHTMLEntities('&quot;hello&quot;')).to.equal('"hello"');
        });

        it('should decode &amp; to &', () => {
            expect(decodeHTMLEntities('A &amp; B')).to.equal('A & B');
        });

        it('should decode &lt; and &gt;', () => {
            expect(decodeHTMLEntities('&lt;div&gt;')).to.equal('<div>');
        });

        it('should decode &apos; to \'', () => {
            expect(decodeHTMLEntities('it&apos;s')).to.equal('it\'s');
        });

        it('should handle multiple entities', () => {
            expect(decodeHTMLEntities('&lt;tag&gt; &amp; &quot;text&quot;'))
                .to.equal('<tag> & "text"');
        });

        it('should handle text without entities', () => {
            expect(decodeHTMLEntities('plain text')).to.equal('plain text');
        });
    });

    describe('convertR1C1toA1', () => {
        it('should convert absolute reference R1C1 to $A$1', () => {
            expect(convertR1C1toA1('R1C1', 0, 0)).to.equal('$A$1');
        });

        it('should convert relative reference RC to current cell', () => {
            expect(convertR1C1toA1('RC', 5, 3)).to.equal('D6');
        });

        it('should convert R[1]C[1] to relative offset', () => {
            expect(convertR1C1toA1('R[1]C[1]', 0, 0)).to.equal('B2');
        });

        it('should convert R[-1]C[-1] to negative offset', () => {
            expect(convertR1C1toA1('R[-1]C[-1]', 5, 5)).to.equal('E5');
        });

        it('should handle mixed absolute and relative R1C[1]', () => {
            expect(convertR1C1toA1('R1C[1]', 0, 0)).to.equal('B$1');
        });

        it('should handle HTML entities in formula', () => {
            expect(convertR1C1toA1('=R1C1&amp;R2C2', 0, 0)).to.equal('=$A$1&$B$2');
        });

        it('should return input if no R1C1 references', () => {
            expect(convertR1C1toA1('=SUM(A1:A10)', 0, 0)).to.equal('=SUM(A1:A10)');
        });

        it('should handle null input', () => {
            expect(convertR1C1toA1(null, 0, 0)).to.be.null;
        });
    });
});

describe('Width Conversion Functions', () => {
    describe('convertWidthToPixels', () => {
        it('should convert points to pixels', () => {
            expect(convertWidthToPixels(10, 'pt')).to.equal(72);
        });

        it('should convert character width to pixels', () => {
            expect(convertWidthToPixels(10, 'char')).to.equal(84);
        });

        it('should convert pixels to pixels (rounded)', () => {
            expect(convertWidthToPixels(10.7, 'px')).to.equal(11);
        });

        it('should default to points conversion', () => {
            expect(convertWidthToPixels(10, 'unknown')).to.equal(72);
        });

        it('should return default for invalid width', () => {
            expect(convertWidthToPixels(null, 'pt')).to.equal(100);
            expect(convertWidthToPixels(NaN, 'pt')).to.equal(100);
        });

        it('should default unit to pt', () => {
            expect(convertWidthToPixels(10)).to.equal(72);
        });
    });
});

describe('Theme and Constants', () => {
    describe('getDefaultTheme', () => {
        it('should return theme with arrayColors', () => {
            const theme = getDefaultTheme();
            expect(theme.arrayColors).to.be.an('array');
            expect(theme.arrayColors).to.have.lengthOf(12);
            expect(theme.arrayColors[0]).to.equal('FFFFFF');
        });

        it('should return theme with objColors', () => {
            const theme = getDefaultTheme();
            expect(theme.objColors).to.be.an('object');
            expect(theme.objColors.dk1).to.equal('000000');
            expect(theme.objColors.lt1).to.equal('FFFFFF');
        });
    });

    describe('numberFormats', () => {
        it('should contain standard format codes', () => {
            expect(numberFormats[0]).to.equal('General');
            expect(numberFormats[1]).to.equal('0');
            expect(numberFormats[2]).to.equal('0.00');
            expect(numberFormats[14]).to.equal('m/d/yy');
        });
    });

    describe('borderStyles', () => {
        it('should contain border style definitions', () => {
            expect(borderStyles.thin).to.deep.equal(['solid', '1px']);
            expect(borderStyles.medium).to.deep.equal(['solid', '2px']);
            expect(borderStyles.thick).to.deep.equal(['solid', '3px']);
        });
    });

    describe('excelValidationTypes', () => {
        it('should map Excel validation types', () => {
            expect(excelValidationTypes.whole).to.equal('number');
            expect(excelValidationTypes.list).to.equal('list');
            expect(excelValidationTypes.date).to.equal('date');
        });
    });

    describe('excelValidationOperations', () => {
        it('should map Excel validation operations', () => {
            expect(excelValidationOperations.equal).to.equal('=');
            expect(excelValidationOperations.greaterThan).to.equal('>');
            expect(excelValidationOperations.notBetween).to.equal('not between');
        });
    });

    describe('excelCFSimpleTypes', () => {
        it('should map conditional format simple types', () => {
            expect(excelCFSimpleTypes.containsBlanks).to.equal('empty');
            expect(excelCFSimpleTypes.notContainsBlanks).to.equal('notEmpty');
        });
    });

    describe('excelCFTextTypes', () => {
        it('should map conditional format text types', () => {
            expect(excelCFTextTypes.containsText).to.equal('contains');
            expect(excelCFTextTypes.beginsWith).to.equal('begins with');
        });
    });

    describe('excelCFNumericOperators', () => {
        it('should map conditional format numeric operators', () => {
            expect(excelCFNumericOperators.equal).to.equal('=');
            expect(excelCFNumericOperators.between).to.equal('between');
        });
    });

    describe('shapeMap', () => {
        it('should map Excel shape names', () => {
            expect(shapeMap.rect).to.equal('rectangle');
            expect(shapeMap.ellipse).to.equal('ellipse');
            expect(shapeMap.triangle).to.equal('triangle');
        });
    });
});

describe('Color Conversion Functions', () => {
    describe('hex2RGB', () => {
        it('should convert hex to RGB', () => {
            expect(hex2RGB('#FF0000')).to.deep.equal([255, 0, 0]);
            expect(hex2RGB('00FF00')).to.deep.equal([0, 255, 0]);
            expect(hex2RGB('#0000FF')).to.deep.equal([0, 0, 255]);
        });

        it('should handle white and black', () => {
            expect(hex2RGB('#FFFFFF')).to.deep.equal([255, 255, 255]);
            expect(hex2RGB('#000000')).to.deep.equal([0, 0, 0]);
        });

        it('should handle gray scale', () => {
            expect(hex2RGB('#808080')).to.deep.equal([128, 128, 128]);
        });
    });

    describe('rgb2Hex', () => {
        it('should convert RGB to hex', () => {
            expect(rgb2Hex([255, 0, 0])).to.equal('FF0000');
            expect(rgb2Hex([0, 255, 0])).to.equal('00FF00');
            expect(rgb2Hex([0, 0, 255])).to.equal('0000FF');
        });

        it('should handle white and black', () => {
            expect(rgb2Hex([255, 255, 255])).to.equal('FFFFFF');
            expect(rgb2Hex([0, 0, 0])).to.equal('000000');
        });

        it('should clamp values above 255', () => {
            expect(rgb2Hex([300, 0, 0])).to.equal('FF0000');
        });

        it('should clamp values below 0', () => {
            expect(rgb2Hex([-10, 0, 0])).to.equal('000000');
        });
    });

    describe('rgb2HSL', () => {
        it('should convert RGB to HSL', () => {
            const hsl = rgb2HSL([255, 0, 0]); // Red
            expect(hsl[0]).to.be.closeTo(0, 0.01);
            expect(hsl[1]).to.equal(1);
            expect(hsl[2]).to.be.closeTo(0.5, 0.01);
        });

        it('should handle gray (no saturation)', () => {
            const hsl = rgb2HSL([128, 128, 128]);
            expect(hsl[0]).to.equal(0);
            expect(hsl[1]).to.equal(0);
            expect(hsl[2]).to.be.closeTo(0.502, 0.01);
        });

        it('should handle white', () => {
            const hsl = rgb2HSL([255, 255, 255]);
            expect(hsl[0]).to.equal(0);
            expect(hsl[1]).to.equal(0);
            expect(hsl[2]).to.equal(1);
        });

        it('should handle black', () => {
            const hsl = rgb2HSL([0, 0, 0]);
            expect(hsl[0]).to.equal(0);
            expect(hsl[1]).to.equal(0);
            expect(hsl[2]).to.equal(0);
        });
    });

    describe('hsl2RGB', () => {
        it('should convert HSL to RGB', () => {
            const rgb = hsl2RGB([0, 1, 0.5]); // Red
            expect(rgb[0]).to.be.closeTo(255, 1);
            expect(rgb[1]).to.be.closeTo(0, 1);
            expect(rgb[2]).to.be.closeTo(0, 1);
        });

        it('should handle no saturation (gray)', () => {
            const rgb = hsl2RGB([0, 0, 0.5]);
            expect(rgb[0]).to.be.closeTo(128, 1);
            expect(rgb[1]).to.be.closeTo(128, 1);
            expect(rgb[2]).to.be.closeTo(128, 1);
        });

        it('should handle white', () => {
            const rgb = hsl2RGB([0, 0, 1]);
            expect(rgb).to.deep.equal([255, 255, 255]);
        });

        it('should handle different hues', () => {
            const green = hsl2RGB([1/3, 1, 0.5]); // Green
            expect(green[0]).to.be.closeTo(0, 1);
            expect(green[1]).to.be.closeTo(255, 1);
            expect(green[2]).to.be.closeTo(0, 1);
        });
    });

    describe('rgb_tint', () => {
        it('should return hex unchanged when tint is 0', () => {
            expect(rgb_tint('FF0000', 0)).to.equal('FF0000');
        });

        it('should return hex unchanged when tint is falsy', () => {
            expect(rgb_tint('FF0000', null)).to.equal('FF0000');
            expect(rgb_tint('FF0000', undefined)).to.equal('FF0000');
        });

        it('should lighten color with positive tint', () => {
            const result = rgb_tint('808080', 0.5);
            // Should be lighter than original
            const original = hex2RGB('808080');
            const tinted = hex2RGB(result);
            expect(tinted[0]).to.be.greaterThan(original[0]);
        });

        it('should darken color with negative tint', () => {
            const result = rgb_tint('808080', -0.5);
            // Should be darker than original
            const original = hex2RGB('808080');
            const tinted = hex2RGB(result);
            expect(tinted[0]).to.be.lessThan(original[0]);
        });

        it('should handle string tint values', () => {
            const result = rgb_tint('808080', '0.5');
            expect(result).to.be.a('string');
            expect(result.length).to.equal(6);
        });
    });
});

describe('Round-trip Conversions', () => {
    it('should maintain value through getColumnName and getColumnIndex', () => {
        for (let i = 0; i < 1000; i++) {
            const name = getColumnName(i);
            const index = getColumnIndex(name);
            expect(index).to.equal(i);
        }
    });

    it('should maintain cell through getCellNameFromCoords and getCoordsFromCellName', () => {
        const testCases = [
            [0, 0], [5, 10], [26, 0], [100, 200]
        ];
        testCases.forEach(([x, y]) => {
            const cellName = getCellNameFromCoords(x, y);
            const coords = getCoordsFromCellName(cellName);
            expect(coords).to.deep.equal([x, y]);
        });
    });

    it('should maintain color through RGB and Hex conversions', () => {
        const testColors = [
            [255, 0, 0],
            [0, 255, 0],
            [0, 0, 255],
            [128, 128, 128],
            [255, 255, 255],
            [0, 0, 0]
        ];
        testColors.forEach(rgb => {
            const hex = rgb2Hex(rgb);
            const rgbBack = hex2RGB(hex);
            expect(rgbBack).to.deep.equal(rgb);
        });
    });

    it('should maintain color through RGB and HSL conversions', () => {
        const testColors = [
            [255, 0, 0],
            [0, 255, 0],
            [0, 0, 255],
            [128, 128, 128]
        ];
        testColors.forEach(rgb => {
            const hsl = rgb2HSL(rgb);
            const rgbBack = hsl2RGB(hsl);
            expect(rgbBack[0]).to.be.closeTo(rgb[0], 1);
            expect(rgbBack[1]).to.be.closeTo(rgb[1], 1);
            expect(rgbBack[2]).to.be.closeTo(rgb[2], 1);
        });
    });
});
