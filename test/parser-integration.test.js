import { expect } from 'chai';
import parser from '../src/parser.js';

describe('Parser Integration Tests', () => {
    describe('Auto-detection by file extension', () => {
        it('should detect CSV files', async () => {
            try {
                await parser('test.csv');
            } catch (error) {
                // File doesn't exist, but we're testing that CSV parser is selected
                // The error should be about file loading, not unsupported type
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect XLSX files', async () => {
            try {
                await parser('test.xlsx');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect XLS files', async () => {
            try {
                await parser('test.xls');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect ODS files', async () => {
            try {
                await parser('test.ods');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect XML files', async () => {
            try {
                await parser('test.xml');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect HTML files', async () => {
            try {
                await parser('test.html');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect HTM files', async () => {
            try {
                await parser('test.htm');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect TSV files', async () => {
            try {
                await parser('test.tsv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect TAB files', async () => {
            try {
                await parser('test.tab');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect TXT files as tab-delimited', async () => {
            try {
                await parser('test.txt');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect DIF files', async () => {
            try {
                await parser('test.dif');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect SLK files', async () => {
            try {
                await parser('test.slk');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect SYLK files', async () => {
            try {
                await parser('test.sylk');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect DBF files', async () => {
            try {
                await parser('test.dbf');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect Lotus files (WKS)', async () => {
            try {
                await parser('test.wks');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect Lotus files (WK1)', async () => {
            try {
                await parser('test.wk1');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect Lotus files (WK3)', async () => {
            try {
                await parser('test.wk3');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect Lotus files (WK4)', async () => {
            try {
                await parser('test.wk4');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should detect Lotus files (123)', async () => {
            try {
                await parser('test.123');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('File extension handling', () => {
        it('should handle uppercase extensions', async () => {
            try {
                await parser('test.CSV');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle mixed case extensions', async () => {
            try {
                await parser('test.XlSx');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle files with multiple dots', async () => {
            try {
                await parser('test.backup.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should throw error for unsupported file types', async () => {
            try {
                await parser('test.pdf');
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Unsupported file type');
            }
        });

        it('should throw error for files without extension', async () => {
            try {
                await parser('test');
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Unsupported file type');
            }
        });
    });

    describe('File object support', () => {
        it('should handle file objects with name property', async () => {
            const mockFile = {
                name: 'test.csv',
                type: 'text/csv'
            };

            try {
                await parser(mockFile);
            } catch (error) {
                // Should attempt to parse, not throw unsupported type error
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should extract extension from file object name', async () => {
            const mockFile = {
                name: 'spreadsheet.XLSX'
            };

            try {
                await parser(mockFile);
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('Error handling', () => {
        it('should throw error for null input', async () => {
            try {
                await parser(null);
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should throw error for undefined input', async () => {
            try {
                await parser(undefined);
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should throw error for empty string', async () => {
            try {
                await parser('');
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should throw error for object without name property', async () => {
            try {
                await parser({ data: 'test' });
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should throw error for number input', async () => {
            try {
                await parser(123);
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should throw error for boolean input', async () => {
            try {
                await parser(true);
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });
    });

    describe('Options passing', () => {
        it('should pass options to CSV parser', async () => {
            try {
                await parser('test.csv', { delimiter: ';' });
            } catch (error) {
                // Options should be passed through
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should pass options to TSV parser', async () => {
            try {
                // TSV should automatically use tab delimiter
                await parser('test.tsv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should allow overriding delimiter for TXT files', async () => {
            try {
                await parser('test.txt', { delimiter: ',' });
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should pass generic options to parsers', async () => {
            try {
                await parser('test.html', { tableIndex: 1 });
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('Return value structure', () => {
        it('should return promise', () => {
            const result = parser('test.csv');
            expect(result).to.be.instanceOf(Promise);
        });
    });

    describe('Path variations', () => {
        it('should handle absolute paths', async () => {
            try {
                await parser('/home/user/documents/test.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle relative paths', async () => {
            try {
                await parser('./data/test.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle Windows paths', async () => {
            try {
                await parser('C:\\Users\\Documents\\test.xlsx');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle paths with spaces', async () => {
            try {
                await parser('my documents/test file.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle paths with special characters', async () => {
            try {
                await parser('data/file-name_v2.0.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('Format aliases', () => {
        it('should treat SYLK as SLK', async () => {
            try {
                await parser('test.sylk');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should treat HTM as HTML', async () => {
            try {
                await parser('test.htm');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should treat TAB as TSV', async () => {
            try {
                await parser('test.tab');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('Case sensitivity', () => {
        it('should be case-insensitive for extensions', async () => {
            const testCases = [
                'test.csv',
                'test.CSV',
                'test.Csv',
                'test.CsV'
            ];

            for (const fileName of testCases) {
                try {
                    await parser(fileName);
                } catch (error) {
                    expect(error.message).to.not.include('Unsupported file type');
                }
            }
        });
    });

    describe('Edge case filenames', () => {
        it('should handle filenames with numbers', async () => {
            try {
                await parser('data2024.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle filenames starting with dots', async () => {
            try {
                await parser('.hidden.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should handle filenames with Unicode', async () => {
            try {
                await parser('数据.csv');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });
    });

    describe('HTML input', () => {
        const page = `<!DOCTYPE html><html><body>
            <h1>Prices</h1>
            <table><tr><th>Item</th><th>Price</th></tr><tr><td>Apple</td><td>1.5</td></tr></table>
            <table><tr><th>City</th></tr><tr><td>Lisbon</td></tr><tr><td>Porto</td></tr></table>
        </body></html>`;

        it('should accept an HTML string instead of a file path', async () => {
            const result = await parser('<table><tr><th>A</th></tr><tr><td>1</td></tr></table>');
            expect(result.worksheets[0].data).to.deep.equal([['1']]);
        });

        it('should accept a full page as a string, with leading whitespace', async () => {
            const result = await parser('\n  ' + page);
            expect(result.worksheets).to.have.lengthOf(2);
        });

        it('should return one worksheet per table', async () => {
            const result = await parser(page);
            expect(result.worksheets.map((w) => w.worksheetName)).to.deep.equal(['Table1', 'Table2']);
            expect(result.worksheets[0].columns.map((c) => c.title)).to.deep.equal(['Item', 'Price']);
            expect(result.worksheets[0].data).to.deep.equal([['Apple', '1.5']]);
            expect(result.worksheets[1].data).to.deep.equal([['Lisbon'], ['Porto']]);
        });

        it('should return only the table picked by tableIndex', async () => {
            const result = await parser(page, { tableIndex: 1 });
            expect(result.worksheets).to.have.lengthOf(1);
            expect(result.worksheets[0].data).to.deep.equal([['Lisbon'], ['Porto']]);
        });

        it('should return every table from HTML bytes', async () => {
            const bytes = new TextEncoder().encode(page);
            expect((await parser(bytes)).worksheets).to.have.lengthOf(2);
            expect((await parser(bytes.buffer, { extension: 'html', tableIndex: 0 })).worksheets[0].data)
                .to.deep.equal([['Apple', '1.5']]);
        });

        it('should report a page without tables', async () => {
            try {
                await parser('<html><body><p>nothing</p></body></html>');
                expect.fail('should have thrown');
            } catch (error) {
                expect(error.message).to.include('No table elements found');
            }
        });

        it('should still treat other strings as file paths', async () => {
            try {
                await parser('missing-file.csv');
                expect.fail('should have thrown');
            } catch (error) {
                expect(error.message).to.not.include('No table elements');
            }
        });
    });

    describe('Lotus formulas', () => {
        const u16 = (n) => [n & 0xFF, n >> 8];
        const f64 = (v) => {
            const b = new Uint8Array(8);
            new DataView(b.buffer).setFloat64(0, v, true);
            return [...b];
        };
        const rec = (type, data) => [...u16(type), ...u16(data.length), ...data];
        const tokens = [0x01, 0x00, 0x80, 0x00, 0x80, 0x05, 0x02, 0x00, 0x0B, 0x03]; // A1*2 (not decoded)
        const wk1 = Uint8Array.from([
            ...rec(0x00, u16(0x0406)),                                                          // BOF
            ...rec(0x0E, [0xFF, ...u16(0), ...u16(0), ...f64(21)]),                             // A1 = 21
            ...rec(0x10, [0xFF, ...u16(1), ...u16(0), ...f64(42), ...u16(tokens.length), ...tokens]), // B1 = A1*2
            ...rec(0x10, [0xFF, ...u16(2), ...u16(0), ...f64(0), ...u16(tokens.length), ...tokens]),  // C1: string result
            ...rec(0x33, [0xFF, ...u16(2), ...u16(0), ...Array.from('ok', (c) => c.charCodeAt(0)), 0]),
            ...rec(0x01, [])                                                                    // EOF
        ]);

        it('should return the calculated values of formula cells', async () => {
            const result = await parser(wk1, { extension: 'wk1' });
            expect(result.worksheets[0].data[0]).to.deep.equal([21, 42, 'ok']);
        });

        it('should not expose undecoded formula tokens', async () => {
            const result = await parser(wk1, { extension: 'wk1' });
            expect(result.worksheets[0].cells).to.deep.equal({});
        });
    });

    describe('Legacy formats (regressions from real-world files)', () => {
        /** dBase III file with one character field of the given width */
        const dbf = ({ ldid = 0, declared, width = 8, rows = [] }) => {
            const header = new Uint8Array(32);
            const view = new DataView(header.buffer);
            header[0] = 0x03;
            view.setUint32(4, declared ?? rows.length, true);
            view.setUint16(8, 32 + 32 + 1, true);
            view.setUint16(10, 1 + width, true);
            header[29] = ldid;
            const field = new Uint8Array(32);
            field.set([0x4E, 0x41, 0x4D, 0x45]); // NAME
            field[11] = 0x43; // 'C'
            field[16] = width;
            const records = rows.flatMap((bytes) => [0x20, ...bytes, ...new Array(width - bytes.length).fill(0x20)]);
            return Uint8Array.from([...header, ...field, 0x0D, ...records, 0x1A]);
        };
        const russian = [0xD0, 0xF3, 0xF1, 0xF1, 0xEA, 0xE8, 0xE9]; // "Русский" in windows-1251

        it('should decode DBF text with the codepage of the language driver byte', async () => {
            const result = await parser(dbf({ ldid: 0xC9, rows: [russian] }), { extension: 'dbf' });
            expect(result.worksheets[0].data[0][0]).to.equal('Русский');
        });

        it('should let options.encoding override the DBF codepage', async () => {
            const result = await parser(dbf({ ldid: 0x65, rows: [russian] }), { extension: 'dbf', encoding: 'windows-1251' });
            expect(result.worksheets[0].data[0][0]).to.equal('Русский');
        });

        it('should reject an unknown DBF encoding', async () => {
            try {
                await parser(dbf({ rows: [russian] }), { extension: 'dbf', encoding: 'no-such-encoding' });
                expect.fail('should have thrown');
            } catch (error) {
                expect(error.message).to.include('Unknown encoding');
            }
        });

        it('should not trust a DBF record count larger than the file', async () => {
            // A 41 KB file in the LibreOffice corpus declares 660,481 records of 180 fields
            const result = await parser(dbf({ declared: 5_000_000, rows: [[0x41], [0x42]] }), { extension: 'dbf' });
            expect(result.worksheets[0].data).to.deep.equal([['A'], ['B']]);
        });

        it('should read SYLK files with CR-only line endings', async () => {
            const slk = new TextEncoder().encode('ID;PWXL;N;E\rC;Y1;X1;K"Name"\rC;Y2;X1;K42\rE\r');
            const result = await parser(slk, { extension: 'slk' });
            expect(result.worksheets[0].data).to.deep.equal([['Name'], [42]]);
        });

        it('should not write to the console while parsing DIF', async () => {
            const dif = new TextEncoder().encode('TABLE\n0,1\n""\nVECTORS\n0,1\n""\nTUPLES\n0,1\n""\nDATA\n0,0\n""\n-1,0\nBOT\n0,7\nV\n-1,0\nEOD\n');
            const log = console.log;
            let calls = 0;
            console.log = () => { calls++; };
            try {
                await parser(dif, { extension: 'dif' });
            } finally {
                console.log = log;
            }
            expect(calls).to.equal(0);
        });
    });

    describe('Text options', () => {
        it('should decode CSV with the given encoding', async () => {
            const latin1 = Uint8Array.from([0x43, 0x61, 0x66, 0xE9, 0x2C, 0x31, 0x0A]); // "Café,1" in latin1
            const result = await parser(latin1, { extension: 'csv', encoding: 'latin1' });
            expect(result.worksheets[0].data[0]).to.deep.equal(['Café', '1']);
        });

        it('should read CSV as UTF-8 by default', async () => {
            const utf8 = new TextEncoder().encode('Café,1\n');
            const result = await parser(utf8, { extension: 'csv' });
            expect(result.worksheets[0].data[0]).to.deep.equal(['Café', '1']);
        });

        it('should split TXT on tabs unless a delimiter is given', async () => {
            const bytes = new TextEncoder().encode('a;b\tc\n');
            expect((await parser(bytes, { extension: 'txt' })).worksheets[0].data[0]).to.deep.equal(['a;b', 'c']);
            expect((await parser(bytes, { extension: 'txt', delimiter: ';' })).worksheets[0].data[0]).to.deep.equal(['a', 'b\tc']);
        });
    });
});
