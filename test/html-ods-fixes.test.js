import { expect } from 'chai';
import JSZip from 'jszip';
import tabularjs from '../src/index.js';
import { parseHTMLTable, parseAllHTMLTables } from '../src/parser/table.js';
import { parseODS } from '../src/parser/ods.js';

const ODS_NS = [
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
    'xmlns:dc="http://purl.org/dc/elements/1.1/"',
    'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'
].join(' ');

// Build a minimal ODS file from <table:table> markup
async function buildODS(tablesXml) {
    const content = `<?xml version="1.0" encoding="UTF-8"?>` +
        `<office:document-content ${ODS_NS} office:version="1.2"><office:body><office:spreadsheet>` +
        tablesXml +
        `</office:spreadsheet></office:body></office:document-content>`;

    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    zip.file('content.xml', content);
    zip.file('META-INF/manifest.xml',
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">' +
        '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
        '</manifest:manifest>');
    return zip.generateAsync({ type: 'uint8array' });
}

// One-row sheet whose cells carry the given formulas
function formulaSheet(name, formulas) {
    const cells = formulas.map(f =>
        `<table:table-cell table:formula="${f}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>`
    ).join('');
    return `<table:table table:name="${name}"><table:table-row>${cells}</table:table-row></table:table>`;
}

async function parseFormulas(formulas) {
    const bytes = await buildODS(formulaSheet('Main', formulas) + formulaSheet('Rates', []));
    const result = await parseODS(bytes);
    return result.worksheets[0].data[0];
}

describe('HTML table parser fixes', () => {
    describe('multiple tbody elements', () => {
        it('reads every tbody in document order, between thead and tfoot', async () => {
            const html = `<table>
                <tfoot><tr><td>Total</td><td>10</td></tr></tfoot>
                <thead><tr><th>Name</th><th>Value</th></tr></thead>
                <tbody><tr><td>a</td><td>1</td></tr></tbody>
                <tbody><tr><td>b</td><td>2</td></tr><tr><td>c</td><td>3</td></tr></tbody>
                <tbody><tr><td>d</td><td>4</td></tr></tbody>
            </table>`;

            const ws = (await tabularjs(html)).worksheets[0];

            expect(ws.columns.map(c => c.title)).to.deep.equal(['Name', 'Value']);
            expect(ws.data).to.deep.equal([
                ['a', '1'], ['b', '2'], ['c', '3'], ['d', '4'], ['Total', '10']
            ]);
        });

        it('keeps <tr> placed directly under <table> next to tbody rows', async () => {
            const html = `<table>
                <tbody><tr><td>a</td></tr></tbody>
                <tr><td>b</td></tr>
                <tbody><tr><td>c</td></tr></tbody>
            </table>`;

            const ws = (await parseHTMLTable(html, { firstRowAsHeader: false })).worksheets[0];

            expect(ws.data).to.deep.equal([['a'], ['b'], ['c']]);
        });
    });

    describe('mixed th/td rows', () => {
        it('keeps a row header <th> before its <td> cells', async () => {
            const html = `<table>
                <tr><td>Quarter</td><td>Sales</td></tr>
                <tr><th>Q1</th><td>10</td></tr>
                <tr><td>20</td><th>Q2</th></tr>
            </table>`;

            const ws = (await tabularjs(html)).worksheets[0];

            expect(ws.data).to.deep.equal([['Q1', '10'], ['20', 'Q2']]);
        });

        it('keeps document order with colspan and rowspan', async () => {
            const html = `<table>
                <tr><td>a</td><th rowspan="2">H</th><td>b</td></tr>
                <tr><td>c</td><th>d</th></tr>
                <tr><th colspan="2">wide</th><td>e</td></tr>
            </table>`;

            const ws = (await parseHTMLTable(html, { firstRowAsHeader: false })).worksheets[0];

            expect(ws.data).to.deep.equal([
                ['a', 'H', 'b'],
                ['c', '', 'd'],
                ['wide', '', 'e']
            ]);
            expect(ws.mergeCells).to.deep.equal({ B1: [1, 2], A3: [2, 1] });
        });
    });

    describe('worksheet name with tableIndex', () => {
        it('names the selected table Table{index+1}, like the all-tables path', async () => {
            const html = '<table><tr><td>one</td></tr></table><table><tr><td>two</td></tr></table>';

            const first = await tabularjs(html, { tableIndex: 0 });
            const second = await tabularjs(html, { tableIndex: 1 });
            const all = await parseAllHTMLTables(html);

            expect(first.worksheets[0].worksheetName).to.equal('Table1');
            expect(second.worksheets[0].worksheetName).to.equal('Table2');
            expect(second.worksheets[0].worksheetName).to.equal(all.worksheets[1].worksheetName);
        });
    });

    describe('HTML entities', () => {
        it('decodes named, decimal and hex entities in cell text', async () => {
            const html = '<table><tr>' +
                '<td>Fish &amp; Chips</td>' +
                '<td>&lt;b&gt; &quot;x&quot; &#39;y&#39; &apos;z&apos;</td>' +
                '<td>caf&eacute; &copy; &euro;5 &mdash; &hellip; &Omega;</td>' +
                '<td>&#x2014;&#8212;&#X1F600;&#128;</td>' +
                '<td>&amp;lt; &unknown; AT&T</td>' +
                '</tr></table>';

            const ws = (await parseHTMLTable(html, { firstRowAsHeader: false })).worksheets[0];

            expect(ws.data[0]).to.deep.equal([
                'Fish & Chips',
                '<b> "x" \'y\' \'z\'',
                'café © €5 — … Ω',
                '——\u{1F600}€',
                '&lt; &unknown; AT&T'
            ]);
        });

        it('decodes &nbsp; to a regular space', async () => {
            const html = '<table><tr><td>1&nbsp;000</td><td>&nbsp;</td><td>a&#160;b</td></tr></table>';

            const ws = (await parseHTMLTable(html, { firstRowAsHeader: false })).worksheets[0];

            expect(ws.data[0][0]).to.equal('1 000');
            expect(ws.data[0][1]).to.equal('');
            // A numeric reference to U+00A0 is kept as written
            expect(ws.data[0][2]).to.equal('a b');
        });

        it('decodes entities in headers, titles and data-formula', async () => {
            const html = '<table>' +
                '<tr><th>R&amp;D</th></tr>' +
                '<tr><td title="Tom &amp; Jerry" data-formula="=IF(A1&gt;0,&quot;y&quot;,&quot;n&quot;)">y</td></tr>' +
                '</table>';

            const ws = (await tabularjs(html)).worksheets[0];

            expect(ws.columns[0].title).to.equal('R&D');
            expect(ws.data[0][0]).to.equal('=IF(A1>0,"y","n")');
            expect(ws.comments.A1).to.equal('Tom & Jerry');
        });

        it('decodes entities in bytes passed with the html extension', async () => {
            const bytes = new TextEncoder().encode('<table><tr><td>&eacute;t&eacute;</td></tr></table>');

            const ws = (await tabularjs(bytes, { extension: 'html', firstRowAsHeader: false })).worksheets[0];

            expect(ws.data[0][0]).to.equal('été');
        });
    });
});

describe('ODS parser fixes', () => {
    describe('cell comments', () => {
        const sheet = '<table:table table:name="Sheet1"><table:table-row>' +
            '<table:table-cell office:value-type="string">' +
            '<office:annotation office:display="true">' +
            '<dc:creator>Ann Author</dc:creator><dc:date>2024-01-02T03:04:05</dc:date>' +
            '<text:p>First line</text:p><text:p>Second line</text:p>' +
            '</office:annotation>' +
            '<text:p>Hello</text:p></table:table-cell>' +
            '<table:table-cell><office:annotation><dc:creator>Bob</dc:creator><text:p>Only a note</text:p></office:annotation></table:table-cell>' +
            '<table:table-cell office:value-type="float" office:value="42"><office:annotation><text:p>number</text:p></office:annotation><text:p>42</text:p></table:table-cell>' +
            '</table:table-row></table:table>';

        it('keeps the annotation out of the cell value', async () => {
            const ws = (await parseODS(await buildODS(sheet))).worksheets[0];

            expect(ws.data[0][0]).to.equal('Hello');
            expect(ws.data[0][1]).to.equal('');
            expect(ws.data[0][2]).to.equal(42);
        });

        it('stores the comment text, author and date in worksheet.comments', async () => {
            const ws = (await tabularjs(await buildODS(sheet), { extension: 'ods' })).worksheets[0];

            expect(ws.comments.A1).to.deep.equal({
                author: 'Ann Author',
                date: '2024-01-02T03:04:05',
                text: 'First line\nSecond line',
                visible: true
            });
            expect(ws.comments.B1).to.deep.equal({ author: 'Bob', date: '', text: 'Only a note' });
            expect(ws.comments.C1.text).to.equal('number');
        });
    });

    describe('multi-line cells', () => {
        it('separates the paragraphs of a cell with line breaks', async () => {
            const sheet = '<table:table table:name="Sheet1"><table:table-row>' +
                '<table:table-cell office:value-type="string"><text:p>Line one</text:p><text:p>Line two</text:p></table:table-cell>' +
                '</table:table-row></table:table>';
            const ws = (await parseODS(await buildODS(sheet))).worksheets[0];
            expect(ws.data[0][0]).to.equal('Line one\nLine two');
        });
    });

    describe('sheet-qualified formula references', () => {
        it('converts single-cell references to other sheets', async () => {
            const data = await parseFormulas([
                'of:=[$Rates.A1]*2',
                'of:=[Rates.$B$3]+[.C1]'
            ]);

            expect(data[0]).to.equal('=Rates!A1*2');
            expect(data[1]).to.equal('=Rates!$B$3+C1');
        });

        it('converts ranges on another sheet, keeping absolute parts', async () => {
            const data = await parseFormulas([
                'of:=SUM([$Rates.A1:.B2])',
                'of:=SUM([$Rates.$A$1:$Rates.B$2])',
                'of:=SUM([$Rates.A:.A])'
            ]);

            expect(data[0]).to.equal('=SUM(Rates!A1:B2)');
            expect(data[1]).to.equal('=SUM(Rates!$A$1:B$2)');
            expect(data[2]).to.equal('=SUM(Rates!A:A)');
        });

        it('quotes sheet names that need it, Excel style', async () => {
            const data = await parseFormulas([
                "of:=[$'My Sheet'.A1]",
                "of:=SUM([$'My Sheet'.A1:.B2])",
                "of:=[$'It''s'.A1]",
                "of:=[$'1Q'.A1]&amp;[$'A1'.B1]&amp;[$'Sales-2024'.C1]"
            ]);

            expect(data[0]).to.equal("='My Sheet'!A1");
            expect(data[1]).to.equal("=SUM('My Sheet'!A1:B2)");
            expect(data[2]).to.equal("='It''s'!A1");
            expect(data[3]).to.equal("='1Q'!A1&'A1'!B1&'Sales-2024'!C1");
        });

        it('converts ranges spanning several sheets to 3D references', async () => {
            const data = await parseFormulas([
                'of:=SUM([$Jan.A1:$Mar.B2])',
                "of:=SUM([$'Jan 1'.A1:$Mar.B2])"
            ]);

            expect(data[0]).to.equal('=SUM(Jan:Mar!A1:B2)');
            expect(data[1]).to.equal("=SUM('Jan 1:Mar'!A1:B2)");
        });

        it('leaves string literals untouched while converting separators', async () => {
            const data = await parseFormulas([
                'of:=IF([$Rates.A1]&gt;0;&quot;[$Rates.A1]; ok&quot;;[.B1])'
            ]);

            expect(data[0]).to.equal('=IF(Rates!A1>0,"[$Rates.A1]; ok",B1)');
        });
    });
});
