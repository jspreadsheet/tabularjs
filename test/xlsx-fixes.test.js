import { expect } from 'chai';
import JSZip from 'jszip';
import tabularjs from '../src/index.js';
import * as xlsx from '../src/parser/xlsx.js';

// Build a minimal single-sheet .xlsx from the <sheetData> inner XML
async function buildXlsx(sheetData, sharedStrings) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        (sharedStrings ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : '') +
        '</Types>');
    zip.file('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>');
    zip.file('xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>');
    zip.file('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        (sharedStrings ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : '') +
        '</Relationships>');
    zip.file('xl/worksheets/sheet1.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetData>' + sheetData + '</sheetData>' +
        '</worksheet>');
    if (sharedStrings) {
        zip.file('xl/sharedStrings.xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sharedStrings.length + '" uniqueCount="' + sharedStrings.length + '">' +
            sharedStrings.map(s => typeof s === 'object' ? s.xml : '<si><t>' + s + '</t></si>').join('') +
            '</sst>');
    }
    return zip.generateAsync({ type: 'uint8array' });
}

async function parseSheet(sheetData, sharedStrings) {
    const result = await tabularjs(await buildXlsx(sheetData, sharedStrings), { extension: 'xlsx' });
    return result.worksheets[0].data;
}

describe('XLSX parser fixes', () => {
    describe('Inline strings (t="inlineStr")', () => {
        it('should read plain inline strings', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>' +
                '<row r="2"><c r="A2" t="inlineStr"><is><t>Apple</t></is></c><c r="B2"><v>3</v></c></row>'
            );
            expect(data[0]).to.deep.equal(['Name', 'Qty']);
            expect(data[1]).to.deep.equal(['Apple', 3]);
        });

        it('should join rich text runs and skip phonetic runs', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1" t="inlineStr"><is>' +
                '<r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> and plain</t></r>' +
                '<rPh sb="0" eb="1"><t>PHONETIC</t></rPh>' +
                '</is></c></row>'
            );
            expect(data[0][0]).to.equal('Bold and plain');
        });

        it('should decode XML entities in inline strings', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1" t="inlineStr"><is><t>R&amp;D &lt;2024&gt; &amp;lt;</t></is></c></row>'
            );
            expect(data[0][0]).to.equal('R&D <2024> &lt;');
        });

        it('should keep an empty inline string as an empty string', async () => {
            const data = await parseSheet('<row r="1"><c r="A1" t="inlineStr"><is><t></t></is></c><c r="B1"><v>1</v></c></row>');
            expect(data[0]).to.deep.equal(['', 1]);
        });
    });

    describe('Numeric cells', () => {
        it('should return 0 as a number', async () => {
            const data = await parseSheet('<row r="1"><c r="A1"><v>0</v></c><c r="B1" t="n"><v>0</v></c></row>');
            expect(data[0][0]).to.equal(0);
            expect(data[0][1]).to.equal(0);
        });

        it('should return 0.0 and negative zero as the number 0', async () => {
            const data = await parseSheet('<row r="1"><c r="A1"><v>0.0</v></c><c r="B1"><v>-0</v></c><c r="C1"><v>0E+0</v></c></row>');
            expect(data[0]).to.deep.equal([0, 0, 0]);
            expect(Object.is(data[0][1], -0)).to.equal(false);
        });

        it('should still parse other numbers', async () => {
            const data = await parseSheet('<row r="1"><c r="A1"><v>-12.5</v></c><c r="B1"><v>1.5E-3</v></c><c r="C1"><v>42</v></c></row>');
            expect(data[0]).to.deep.equal([-12.5, 0.0015, 42]);
        });

        it('should leave genuinely non-numeric values and strings untouched', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1"><v>abc</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="str"><v>0</v></c><c r="D1" t="d"><v>2024-01-15</v></c></row>',
                ['zero']
            );
            expect(data[0]).to.deep.equal(['abc', 'zero', '0', '2024-01-15']);
        });
    });

    describe('Shared formulas', () => {
        it('should translate a shared formula down a column', async () => {
            let rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>n</t></is></c><c r="B1" t="inlineStr"><is><t>double</t></is></c></row>';
            for (let r = 2; r <= 10; r++) {
                const f = r === 2 ? '<f t="shared" ref="B2:B10" si="0">A2*2</f>' : '<f t="shared" si="0"/>';
                rows += `<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}">${f}<v>${r * 2}</v></c></row>`;
            }
            const data = await parseSheet(rows);
            expect(data[1][1]).to.equal('=A2*2');
            expect(data[2][1]).to.equal('=A3*2');
            expect(data[9][1]).to.equal('=A10*2');
        });

        it('should translate across columns and keep absolute parts', async () => {
            const data = await parseSheet(
                '<row r="1">' +
                '<c r="B1"><f t="shared" ref="B1:D1" si="0">SUM(B2:B5)/$A$1+B$2+$A2</f><v>1</v></c>' +
                '<c r="C1"><f t="shared" si="0"/><v>2</v></c>' +
                '<c r="D1"><f t="shared" si="0"/><v>3</v></c>' +
                '</row>'
            );
            expect(data[0]).to.deep.equal([
                '',
                '=SUM(B2:B5)/$A$1+B$2+$A2',
                '=SUM(C2:C5)/$A$1+C$2+$A2',
                '=SUM(D2:D5)/$A$1+D$2+$A2'
            ]);
        });

        it('should keep several shared groups apart', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">C1+1</f><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="1">C1*2</f><v>2</v></c></row>' +
                '<row r="2"><c r="A2"><f t="shared" si="0"/><v>1</v></c><c r="B2"><f t="shared" si="1"/><v>2</v></c></row>'
            );
            expect(data[1]).to.deep.equal(['=C2+1', '=C2*2']);
        });

        it('should keep the cached value when the master formula is missing', async () => {
            const data = await parseSheet('<row r="2"><c r="B2"><f t="shared" si="7"/><v>5</v></c></row>');
            expect(data[1][1]).to.equal(5);
        });
    });

    describe('translateFormula', () => {
        const t = (formula, rows, cols) => xlsx.translateFormula(formula, rows, cols);

        it('should shift relative references and keep absolute ones', () => {
            expect(t('A2*2', 1, 0)).to.equal('A3*2');
            expect(t('$A$2+A$2+$A2', 3, 2)).to.equal('$A$2+C$2+$A5');
            expect(t('SUM(A2:B3)', 3, 2)).to.equal('SUM(C5:D6)');
        });

        it('should handle sheet-qualified references', () => {
            expect(t("Sheet2!A1+'My Sheet'!B$2+'it''s'!C3", 3, 2)).to.equal("Sheet2!C4+'My Sheet'!D$2+'it''s'!E6");
            expect(t('VLOOKUP(A2,Sheet2!$A:$C,2,FALSE)', 1, 1)).to.equal('VLOOKUP(B3,Sheet2!$A:$C,2,FALSE)');
        });

        it('should handle whole-column and whole-row references', () => {
            expect(t('SUM(A:A)+SUM($B:B)', 3, 2)).to.equal('SUM(C:C)+SUM($B:D)');
            expect(t('SUM(1:1)+SUM($2:3)', 3, 2)).to.equal('SUM(4:4)+SUM($2:6)');
        });

        it('should leave strings, function names, names and numbers alone', () => {
            expect(t('IF(A2>0,"A1 ""B2""",C3)', 1, 1)).to.equal('IF(B3>0,"A1 ""B2""",D4)');
            expect(t('LOG10(A2)+ATAN2(B2,C2)', 1, 0)).to.equal('LOG10(A3)+ATAN2(B3,C3)');
            expect(t('1.5E+3*A2+TaxRate', 1, 0)).to.equal('1.5E+3*A3+TaxRate');
            expect(t('Table1[[#This Row],[Col1]]*A2', 1, 0)).to.equal('Table1[[#This Row],[Col1]]*A3');
        });

        it('should return #REF! for references moved off the sheet', () => {
            expect(t('A1', -1, 0)).to.equal('#REF!');
            expect(t('XFD1+A1', 0, 1)).to.equal('#REF!+B1');
        });

        it('should return the formula unchanged for a zero offset', () => {
            expect(t('A1+B2', 0, 0)).to.equal('A1+B2');
        });
    });

    describe('Error cells', () => {
        it('should read errors exactly as Excel shows them', async () => {
            const errors = ['#VALUE!', '#DIV/0!', '#N/A', '#REF!', '#NAME?', '#NUM!', '#NULL!'];
            const cells = errors.map((e, i) => `<c r="${String.fromCharCode(65 + i)}1" t="e"><v>${e}</v></c>`).join('');
            const data = await parseSheet('<row r="1">' + cells + '</row>');
            expect(data[0]).to.deep.equal(errors);
        });

        it('should keep the formula of an error cell', async () => {
            const data = await parseSheet('<row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="B1" t="e"><v>#DIV/0!</v></c></row>');
            expect(data[0]).to.deep.equal(['=1/0', '#DIV/0!']);
        });
    });

    describe('Shared strings', () => {
        it('should decode XML entities exactly once', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>',
                ['R&amp;D &lt;b&gt; &quot;q&quot; &apos;a&apos;', '&amp;lt; &amp;amp; &amp;#65;', '&#65;&#x42;&#X43; &#233; &#x1F600;']
            );
            expect(data[0]).to.deep.equal(['R&D <b> "q" \'a\'', '&lt; &amp; &#65;', 'ABC é 😀']);
        });

        it('should concatenate rich text runs and skip phonetic runs', async () => {
            const data = await parseSheet(
                '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>',
                [
                    { xml: '<si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> &amp; plain</t></r></si>' },
                    { xml: '<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>' }
                ]
            );
            expect(data[0]).to.deep.equal(['Bold & plain', '東京']);
        });

        it('should decode numeric entities in inline strings too', async () => {
            const data = await parseSheet('<row r="1"><c r="A1" t="inlineStr"><is><t>&#65;&#x42; &amp;#65;</t></is></c></row>');
            expect(data[0][0]).to.equal('AB &#65;');
        });
    });
});
