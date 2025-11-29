import { expect } from 'chai';
import { parseHTMLTable, parseAllHTMLTables } from '../src/parser/table.js';

describe('HTML Table Parser', () => {
    describe('parseHTMLTable - Basic functionality', () => {
        it('should parse simple table', async () => {
            const html = `
                <table>
                    <tr><th>Name</th><th>Age</th></tr>
                    <tr><td>John</td><td>25</td></tr>
                    <tr><td>Jane</td><td>30</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            expect(result).to.have.property('worksheets');
            expect(result.worksheets).to.have.lengthOf(1);

            const ws = result.worksheets[0];
            expect(ws.data).to.deep.equal([
                ['John', '25'],
                ['Jane', '30']
            ]);
            expect(ws.columns).to.have.lengthOf(2);
            expect(ws.columns[0].title).to.equal('Name');
            expect(ws.columns[1].title).to.equal('Age');
        });

        it('should parse table without header', async () => {
            const html = `
                <table>
                    <tr><td>John</td><td>25</td></tr>
                    <tr><td>Jane</td><td>30</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html, { firstRowAsHeader: false });

            const ws = result.worksheets[0];
            expect(ws.data).to.have.lengthOf(2);
            expect(ws.columns[0].title).to.equal('A');
            expect(ws.columns[1].title).to.equal('B');
        });

        it('should handle tbody, thead, tfoot', async () => {
            const html = `
                <table>
                    <thead>
                        <tr><th>Name</th><th>Age</th></tr>
                    </thead>
                    <tbody>
                        <tr><td>John</td><td>25</td></tr>
                        <tr><td>Jane</td><td>30</td></tr>
                    </tbody>
                    <tfoot>
                        <tr><td>Total</td><td>2</td></tr>
                    </tfoot>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data).to.have.lengthOf(3);
            expect(ws.data[0]).to.deep.equal(['John', '25']);
            expect(ws.data[2]).to.deep.equal(['Total', '2']);
        });
    });

    describe('parseHTMLTable - Merged cells', () => {
        it('should handle colspan', async () => {
            const html = `
                <table>
                    <tr><td>Row1A</td><td>Row1B</td><td>Row1C</td></tr>
                    <tr><td>John</td><td colspan="2">Engineer, 25 years old</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html, { firstRowAsHeader: false });

            const ws = result.worksheets[0];
            expect(ws.mergeCells).to.exist;
            // Cell B2 has colspan 2
            expect(ws.mergeCells['B2']).to.deep.equal([2, 1]);
        });

        it('should handle rowspan', async () => {
            const html = `
                <table>
                    <tr><td>Name</td><td>Age</td></tr>
                    <tr><td rowspan="2">John</td><td>25</td></tr>
                    <tr><td>26</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html, { firstRowAsHeader: false });

            const ws = result.worksheets[0];
            expect(ws.mergeCells).to.exist;
            // Rowspan starts in row 2 (A2)
            expect(ws.mergeCells['A2']).to.deep.equal([1, 2]);
        });

        it('should handle both colspan and rowspan', async () => {
            const html = `
                <table>
                    <tr><td>A</td><td>B</td><td>C</td></tr>
                    <tr><td colspan="2" rowspan="2">Merged</td><td>1</td></tr>
                    <tr><td>2</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html, { firstRowAsHeader: false });

            const ws = result.worksheets[0];
            // Merge starts in row 2 (A2)
            expect(ws.mergeCells['A2']).to.deep.equal([2, 2]);
        });
    });

    describe('parseHTMLTable - Formulas and attributes', () => {
        it('should parse data-formula attribute', async () => {
            const html = `
                <table>
                    <tr><th>A</th><th>B</th><th>Sum</th></tr>
                    <tr><td>10</td><td>20</td><td data-formula="A1+B1">30</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data[0][2]).to.equal('=A1+B1');
        });

        it('should add = prefix if missing in formula', async () => {
            const html = `
                <table>
                    <tr><th>Result</th></tr>
                    <tr><td data-formula="SUM(A1:A10)">55</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.equal('=SUM(A1:A10)');
        });

        it('should parse inline styles', async () => {
            const html = `
                <table>
                    <tr><th>Name</th></tr>
                    <tr><td style="color: red; font-weight: bold;">John</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.style).to.exist;
            expect(ws.style['A1']).to.equal('color: red; font-weight: bold;');
        });

        it('should parse title attribute as comment', async () => {
            const html = `
                <table>
                    <tr><th>Name</th></tr>
                    <tr><td title="This is a comment">John</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.comments).to.exist;
            expect(ws.comments['A1']).to.equal('This is a comment');
        });
    });

    describe('parseHTMLTable - Edge cases', () => {
        it('should handle empty table', async () => {
            const html = '<table></table>';

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data).to.deep.equal([]);
            expect(ws.columns).to.deep.equal([]);
        });

        it('should handle table with only header', async () => {
            const html = `
                <table>
                    <tr><th>Name</th><th>Age</th></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data).to.deep.equal([]);
            expect(ws.columns).to.have.lengthOf(2);
        });

        it('should handle irregular rows', async () => {
            const html = `
                <table>
                    <tr><th>A</th><th>B</th><th>C</th></tr>
                    <tr><td>1</td><td>2</td></tr>
                    <tr><td>3</td><td>4</td><td>5</td><td>6</td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            // Should pad rows to match max columns
            expect(ws.data[0]).to.have.lengthOf(4);
            expect(ws.data[0][2]).to.equal('');
        });

        it('should trim cell content', async () => {
            const html = `
                <table>
                    <tr><th>Name</th></tr>
                    <tr><td>  John  </td></tr>
                </table>
            `;

            const result = await parseHTMLTable(html);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.equal('John');
        });

        it('should throw error if no table found', async () => {
            const html = '<div>No table here</div>';

            try {
                await parseHTMLTable(html);
                expect.fail('Should have thrown error');
            } catch (error) {
                // Error could be file not found (in Node.js) or no table found
                expect(error.message).to.satisfy(msg =>
                    msg.includes('No table elements found') || msg.includes('ENOENT')
                );
            }
        });

        it('should throw error if table index out of range', async () => {
            const html = '<table><tr><td>1</td></tr></table>';

            try {
                await parseHTMLTable(html, { tableIndex: 5 });
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('out of range');
            }
        });
    });

    describe('parseAllHTMLTables', () => {
        it('should parse multiple tables', async () => {
            const html = '<table><tr><th>Table 1</th></tr><tr><td>Data 1</td></tr></table>' +
                '<table><tr><th>Table 2</th></tr><tr><td>Data 2</td></tr></table>';

            const result = await parseAllHTMLTables(html);

            expect(result.worksheets).to.have.lengthOf(2);
            expect(result.worksheets[0].worksheetName).to.equal('Table1');
            expect(result.worksheets[1].worksheetName).to.equal('Table2');
        });

        it('should parse nested tables', async () => {
            const html = `
                <table>
                    <tr><th>Outer</th></tr>
                    <tr><td>
                        <table>
                            <tr><th>Inner</th></tr>
                            <tr><td>Data</td></tr>
                        </table>
                    </td></tr>
                </table>
            `;

            const result = await parseAllHTMLTables(html);

            // Should find both outer and inner table
            expect(result.worksheets).to.have.lengthOf(2);
        });
    });

    describe('parseHTMLTable - tableIndex option', () => {
        it('should parse second table when tableIndex is 1', async () => {
            const html = `
                <table><tr><th>First</th></tr></table>
                <table><tr><th>Second</th></tr></table>
            `;

            const result = await parseHTMLTable(html, { tableIndex: 1 });

            const ws = result.worksheets[0];
            expect(ws.columns[0].title).to.equal('Second');
        });
    });
});
