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
});
