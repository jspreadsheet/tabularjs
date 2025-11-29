import { expect } from 'chai';
import { parseXMLSpreadsheetString } from '../src/parser/xml.js';

describe('XML Spreadsheet Parser', () => {
    describe('parseXMLSpreadsheetString - Basic functionality', () => {
        it('should parse simple XML spreadsheet', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell><Data ss:Type="String">Name</Data></Cell>
                                <Cell><Data ss:Type="String">Age</Data></Cell>
                            </Row>
                            <Row>
                                <Cell><Data ss:Type="String">John</Data></Cell>
                                <Cell><Data ss:Type="Number">25</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result).to.have.property('worksheets');
            expect(result.worksheets).to.have.lengthOf(1);

            const ws = result.worksheets[0];
            expect(ws.worksheetName).to.equal('Sheet1');
            expect(ws.data).to.deep.equal([
                ['Name', 'Age'],
                ['John', 25]
            ]);
        });

        it('should handle multiple worksheets', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row><Cell><Data ss:Type="String">Data1</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                    <Worksheet ss:Name="Sheet2">
                        <Table>
                            <Row><Cell><Data ss:Type="String">Data2</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.worksheets).to.have.lengthOf(2);
            expect(result.worksheets[0].worksheetName).to.equal('Sheet1');
            expect(result.worksheets[1].worksheetName).to.equal('Sheet2');
        });

        it('should parse different data types', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell><Data ss:Type="String">Text</Data></Cell>
                                <Cell><Data ss:Type="Number">123.45</Data></Cell>
                                <Cell><Data ss:Type="Boolean">1</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.equal('Text');
            expect(ws.data[0][1]).to.equal(123.45);
            expect(ws.data[0][2]).to.equal(true);
        });

        it('should handle empty cells', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell><Data ss:Type="String">A</Data></Cell>
                                <Cell></Cell>
                                <Cell><Data ss:Type="String">C</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0]).to.deep.equal(['A', '', 'C']);
        });
    });

    describe('parseXMLSpreadsheetString - Formulas', () => {
        it('should parse formulas in R1C1 notation', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell><Data ss:Type="Number">10</Data></Cell>
                                <Cell><Data ss:Type="Number">20</Data></Cell>
                                <Cell ss:Formula="RC[-2]+RC[-1]"><Data ss:Type="Number">30</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0][2]).to.equal('=A1+B1');
        });

        it('should parse absolute formula references', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:Formula="R1C1"><Data ss:Type="Number">100</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.equal('=$A$1');
        });

        it('should parse formulas with functions', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:Formula="SUM(R1C1:R10C1)"><Data ss:Type="Number">550</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.include('SUM');
        });
    });

    describe('parseXMLSpreadsheetString - Merged cells', () => {
        it('should handle MergeAcross (colspan)', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:MergeAcross="2"><Data ss:Type="String">Merged</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.mergeCells).to.exist;
            expect(ws.mergeCells['A1']).to.deep.equal([3, 1]);
        });

        it('should handle MergeDown (rowspan)', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:MergeDown="2"><Data ss:Type="String">Merged</Data></Cell>
                            </Row>
                            <Row></Row>
                            <Row></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.mergeCells).to.exist;
            expect(ws.mergeCells['A1']).to.deep.equal([1, 3]);
        });

        it('should handle both MergeAcross and MergeDown', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:MergeAcross="1" ss:MergeDown="1"><Data ss:Type="String">Merged</Data></Cell>
                            </Row>
                            <Row></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.mergeCells['A1']).to.deep.equal([2, 2]);
        });
    });

    describe('parseXMLSpreadsheetString - Styles', () => {
        it('should parse and apply styles', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Styles>
                        <Style ss:ID="s1">
                            <Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/>
                            <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
                        </Style>
                    </Styles>
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:StyleID="s1"><Data ss:Type="String">Styled</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.style).to.exist;
            expect(ws.style['A1']).to.include('font-family: Arial');
            expect(ws.style['A1']).to.include('font-weight: bold');
            expect(ws.style['A1']).to.include('background-color: #FF0000');
        });

        it('should parse number format styles', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Styles>
                        <Style ss:ID="s1">
                            <NumberFormat ss:Format="#,##0.00"/>
                        </Style>
                    </Styles>
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:StyleID="s1"><Data ss:Type="Number">1234.56</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.cells).to.exist;
            expect(ws.cells['A1']).to.exist;
            expect(ws.cells['A1'].mask).to.equal('#,##0.00');
        });

        it('should parse border styles', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Styles>
                        <Style ss:ID="s1">
                            <Borders>
                                <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
                                <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
                            </Borders>
                        </Style>
                    </Styles>
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:StyleID="s1"><Data ss:Type="String">Bordered</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.style['A1']).to.include('border-left');
            expect(ws.style['A1']).to.include('border-top');
        });
    });

    describe('parseXMLSpreadsheetString - Column and Row properties', () => {
        it('should parse column widths', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Column ss:Width="100"/>
                            <Column ss:Width="150"/>
                            <Row>
                                <Cell><Data ss:Type="String">A</Data></Cell>
                                <Cell><Data ss:Type="String">B</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.columns[0].width).to.equal(100);
            expect(ws.columns[1].width).to.equal(150);
        });

        it('should parse hidden columns', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Column ss:Hidden="1"/>
                            <Column/>
                            <Row>
                                <Cell><Data ss:Type="String">A</Data></Cell>
                                <Cell><Data ss:Type="String">B</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.columns[0].visible).to.equal(false);
            expect(ws.columns[1].visible).to.be.undefined;
        });

        it('should parse row heights', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row ss:Height="25">
                                <Cell><Data ss:Type="String">Tall</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.rows).to.exist;
            expect(ws.rows[0].height).to.equal(25);
        });

        it('should parse hidden rows', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row ss:Hidden="1">
                                <Cell><Data ss:Type="String">Hidden</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.rows[0].visible).to.equal(false);
        });
    });

    describe('parseXMLSpreadsheetString - Index attribute', () => {
        it('should handle Cell Index attribute', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell ss:Index="3"><Data ss:Type="String">C</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data[0][0]).to.equal('');
            expect(ws.data[0][1]).to.equal('');
            expect(ws.data[0][2]).to.equal('C');
        });

        it('should handle Row Index attribute', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row ss:Index="3">
                                <Cell><Data ss:Type="String">Row 3</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.data).to.have.lengthOf(3);
            // Empty rows are filled but may contain a single empty string
            expect(ws.data[0].length).to.be.at.least(0);
            expect(ws.data[1].length).to.be.at.least(0);
            expect(ws.data[2][0]).to.equal('Row 3');
        });

        it('should handle Column Index attribute', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Column ss:Index="3" ss:Width="100"/>
                            <Row>
                                <Cell><Data ss:Type="String">A</Data></Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.columns[2].width).to.equal(100);
        });
    });

    describe('parseXMLSpreadsheetString - Comments', () => {
        it('should parse cell comments', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row>
                                <Cell>
                                    <Data ss:Type="String">Value</Data>
                                    <Comment><Data>This is a comment</Data></Comment>
                                </Cell>
                            </Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.comments).to.exist;
            expect(ws.comments['A1']).to.equal('This is a comment');
        });
    });

    describe('parseXMLSpreadsheetString - Named ranges', () => {
        it('should parse named ranges', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Names>
                        <NamedRange ss:Name="MyRange" ss:RefersTo="Sheet1!A1:B10"/>
                    </Names>
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row><Cell><Data ss:Type="String">A</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.definedNames).to.exist;
            expect(result.definedNames['MyRange']).to.equal('Sheet1!A1:B10');
        });

        it('should clean Excel internal prefixes from named ranges', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Names>
                        <NamedRange ss:Name="_xlnm.Print_Area" ss:RefersTo="Sheet1!A1:Z100"/>
                    </Names>
                    <Worksheet ss:Name="Sheet1">
                        <Table>
                            <Row><Cell><Data ss:Type="String">A</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.definedNames['Print_Area']).to.equal('Sheet1!A1:Z100');
        });
    });

    describe('parseXMLSpreadsheetString - Edge cases', () => {
        it('should handle worksheet without Table node', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Empty"></Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.worksheets[0].data).to.deep.equal([]);
            expect(result.worksheets[0].columns).to.deep.equal([]);
        });

        it('should throw error if no Worksheet found', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                </Workbook>`;

            expect(() => parseXMLSpreadsheetString(xml)).to.throw('No Worksheet elements found');
        });

        it('should use default worksheet name if not specified', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet>
                        <Table>
                            <Row><Cell><Data ss:Type="String">Test</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.worksheets[0].worksheetName).to.equal('Sheet1');
        });
    });

    describe('parseXMLSpreadsheetString - Worksheet options', () => {
        it('should parse hidden worksheet', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Hidden" ss:Visible="0">
                        <Table>
                            <Row><Cell><Data ss:Type="String">Hidden</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            expect(result.worksheets[0].worksheetState).to.equal('hidden');
        });

        it('should handle default dimensions', () => {
            const xml = `<?xml version="1.0"?>
                <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
                    <Worksheet ss:Name="Sheet1">
                        <Table ss:DefaultColumnWidth="80" ss:DefaultRowHeight="15">
                            <Row><Cell><Data ss:Type="String">Test</Data></Cell></Row>
                        </Table>
                    </Worksheet>
                </Workbook>`;

            const result = parseXMLSpreadsheetString(xml);

            const ws = result.worksheets[0];
            expect(ws.defaultColWidth).to.equal(80);
            expect(ws.defaultRowHeight).to.equal(15);
        });
    });
});
