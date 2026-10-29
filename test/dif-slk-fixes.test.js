import { expect } from 'chai';
import { parseDIF, parseDIFString } from '../src/parser/dif.js';
import { parseSLK, parseSLKString } from '../src/parser/slk.js';

const DIF_HEADER = 'TABLE\n0,1\n""\nVECTORS\n0,2\n""\nTUPLES\n0,2\n""\nDATA\n0,0\n""\n';

describe('DIF and SYLK fixes', () => {
    describe('DIF end of data', () => {
        it('should not duplicate the last row when the data ends with -1,0 / EOD', () => {
            const dif = DIF_HEADER +
                '-1,0\nBOT\n0,1\nV\n0,2\nV\n' +
                '-1,0\nBOT\n0,3\nV\n0,4\nV\n' +
                '-1,0\nEOD\n';
            const result = parseDIFString(dif);
            expect(result.worksheets[0].data).to.deep.equal([[1, 2], [3, 4]]);
        });

        it('should not duplicate the last row with CRLF line endings', () => {
            const dif = (DIF_HEADER + '-1,0\nBOT\n1,0\n"a"\n-1,0\nBOT\n1,0\n"b"\n-1,0\nEOD\n').replace(/\n/g, '\r\n');
            const result = parseDIFString(dif);
            expect(result.worksheets[0].data.map(r => r[0])).to.deep.equal(['a', 'b']);
        });

        it('should not duplicate the last row when an empty line precedes EOD', () => {
            const dif = DIF_HEADER + '-1,0\nBOT\n0,1\nV\n0,2\nV\n-1,0\n\nEOD\n';
            const result = parseDIFString(dif);
            expect(result.worksheets[0].data).to.deep.equal([[1, 2]]);
        });

        it('should return 52 rows for the Excel sample test1.dif', async () => {
            const result = await parseDIF('samples/test1.dif');
            expect(result.worksheets[0].data).to.have.lengthOf(52);
        });
    });

    describe('DIF value indicators', () => {
        it('should return booleans for TRUE and FALSE', () => {
            const dif = DIF_HEADER + '-1,0\nBOT\n0,1\nTRUE\n0,0\nFALSE\n-1,0\nEOD\n';
            const data = parseDIFString(dif).worksheets[0].data;
            expect(data).to.deep.equal([[true, false]]);
        });

        it('should return #N/A for NA and #VALUE! for ERROR', () => {
            const dif = DIF_HEADER + '-1,0\nBOT\n0,0\nNA\n0,0\nERROR\n-1,0\nEOD\n';
            const data = parseDIFString(dif).worksheets[0].data;
            expect(data).to.deep.equal([['#N/A', '#VALUE!']]);
        });

        it('should keep numbers for the V indicator', () => {
            const dif = DIF_HEADER + '-1,0\nBOT\n0,1.5\nV\n0,-2\nV\n-1,0\nEOD\n';
            const data = parseDIFString(dif).worksheets[0].data;
            expect(data).to.deep.equal([[1.5, -2]]);
        });

        it('should return #VALUE! for the ERROR cells in test1.dif', async () => {
            const result = await parseDIF('samples/test1.dif');
            const lastRow = result.worksheets[0].data[51];
            expect(lastRow.slice(13, 16)).to.deep.equal(['#VALUE!', '#VALUE!', '#VALUE!']);
        });
    });

    describe('SYLK escaped semicolons', () => {
        it('should unescape ;; inside a string constant', () => {
            const result = parseSLKString('ID;PWXL;N;E\nC;Y1;X1;K"Tea;; green"\nE\n');
            expect(result.worksheets[0].data).to.deep.equal([['Tea; green']]);
        });

        it('should keep parsing fields after an escaped semicolon', () => {
            const result = parseSLKString('ID;PWXL;N;E\nC;Y1;X1;K"a;;b;;c";\nC;X2;K"d"\nE\n');
            expect(result.worksheets[0].data).to.deep.equal([['a;b;c', 'd']]);
        });

        it('should unescape ;; inside a formula', () => {
            const result = parseSLKString('ID;PWXL;N;E\nC;Y1;X1;K"x;;y";EIF(1,"x;;y","")\nE\n');
            expect(result.worksheets[0].cells.A1.formula).to.equal('=IF(1,"x;y","")');
        });

        it('should honour the escape in F records', () => {
            const result = parseSLKString('ID;PWXL;N;E\nF;P0;M0;;0;Y1;X1\nC;Y1;X1;K1;P0\nE\n');
            expect(result.worksheets[0].cells.A1.mask).to.equal('0;0');
        });
    });

    describe('SYLK encoding option', () => {
        // "Café" with é as 0xE9 (windows-1252 / latin1)
        const bytes = new Uint8Array([
            ...new TextEncoder().encode('ID;PWXL;N;E\nC;Y1;X1;K"Caf'),
            0xE9,
            ...new TextEncoder().encode('"\nE\n')
        ]);

        it('should decode using options.encoding', async () => {
            const result = await parseSLK(bytes, { encoding: 'windows-1252' });
            expect(result.worksheets[0].data[0][0]).to.equal('Café');
        });

        it('should keep UTF-8 as the default', async () => {
            const utf8 = new TextEncoder().encode('ID;PWXL;N;E\nC;Y1;X1;K"Café"\nE\n');
            const result = await parseSLK(utf8);
            expect(result.worksheets[0].data[0][0]).to.equal('Café');
        });
    });

    describe('DIF column count', () => {
        it('should size rows to the widest row, not the VECTORS header', () => {
            // Excel writes VECTORS = number of rows
            const dif = 'TABLE\n0,1\n"EXCEL"\nVECTORS\n0,3\n""\nTUPLES\n0,2\n""\nDATA\n0,0\n""\n' +
                '-1,0\nBOT\n0,1\nV\n0,2\nV\n' +
                '-1,0\nBOT\n0,3\nV\n' +
                '-1,0\nBOT\n0,4\nV\n0,5\nV\n' +
                '-1,0\nEOD\n';
            const sheet = parseDIFString(dif).worksheets[0];
            expect(sheet.data).to.deep.equal([[1, 2], [3, ''], [4, 5]]);
            expect(sheet.columns).to.have.lengthOf(2);
        });

        it('should return 52 rows x 17 columns for the Excel sample test1.dif', async () => {
            const sheet = (await parseDIF('samples/test1.dif')).worksheets[0];
            expect(sheet.data).to.have.lengthOf(52);
            expect(sheet.columns).to.have.lengthOf(17);
            sheet.data.forEach(row => expect(row).to.have.lengthOf(17));
        });
    });

    describe('DIF LABEL section', () => {
        it('should not skip an extra line after a numeric label entry', () => {
            const dif = 'TABLE\n0,1\n""\nLABEL\n0,5\nV\n1,0\n"Name"\n' +
                'DATA\n0,0\n""\n-1,0\nBOT\n1,0\n"x"\n-1,0\nEOD\n';
            const sheet = parseDIFString(dif).worksheets[0];
            expect(sheet.columns[0].title).to.equal('Name');
            expect(sheet.data).to.deep.equal([['x']]);
        });
    });
});
