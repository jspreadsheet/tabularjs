import { expect } from 'chai';
import fs from 'fs';
import JSZip from 'jszip';
import parser, { detectFormat } from '../src/parser.js';

const encode = (text) => new TextEncoder().encode(text);

describe('Format Detection (magic bytes)', () => {
    describe('detectFormat', () => {
        it('should detect OLE2 compound files as xls', async () => {
            const buffer = new Uint8Array(16);
            buffer.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
            expect(await detectFormat(buffer)).to.equal('xls');
        });

        it('should detect zip archives with [Content_Types].xml as xlsx', async () => {
            const zip = new JSZip();
            zip.file('[Content_Types].xml', '<Types/>');
            zip.file('xl/workbook.xml', '<workbook/>');
            const buffer = await zip.generateAsync({ type: 'uint8array' });
            expect(await detectFormat(buffer)).to.equal('xlsx');
        });

        it('should detect zip archives with an OpenDocument mimetype as ods', async () => {
            const zip = new JSZip();
            zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
            zip.file('content.xml', '<office:document-content/>');
            const buffer = await zip.generateAsync({ type: 'uint8array' });
            expect(await detectFormat(buffer)).to.equal('ods');
        });

        it('should detect zip archives with an Index.zip as numbers', async () => {
            const zip = new JSZip();
            zip.file('Index.zip', new Uint8Array([0x50, 0x4B, 0x05, 0x06]));
            zip.file('Metadata/Properties.plist', 'x');
            const buffer = await zip.generateAsync({ type: 'uint8array' });
            expect(await detectFormat(buffer)).to.equal('numbers');
        });

        it('should detect the Lotus 1-2-3 BOF record', async () => {
            const buffer = new Uint8Array([0x00, 0x00, 0x02, 0x00, 0x06, 0x04, 0x06, 0x00]);
            expect(await detectFormat(buffer)).to.equal('wk1');
        });

        it('should detect a dBase header', async () => {
            const buffer = new Uint8Array(65);
            buffer[0] = 0x03; // dBase III
            buffer[1] = 99;   // year
            buffer[2] = 6;    // month
            buffer[3] = 15;   // day
            buffer[8] = 65;   // header length (LE)
            buffer[10] = 32;  // record length
            expect(await detectFormat(buffer)).to.equal('dbf');
        });

        it('should detect SYLK content', async () => {
            expect(await detectFormat(encode('ID;PWXL;N;E\r\nC;Y1;X1;K"a"\r\nE'))).to.equal('slk');
        });

        it('should detect DIF content', async () => {
            expect(await detectFormat(encode('TABLE\n0,1\n"EXCEL"\nVECTORS\n0,2\n""\n'))).to.equal('dif');
        });

        it('should detect XML Spreadsheet 2003 content', async () => {
            const xml = '<?xml version="1.0"?><Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"></Workbook>';
            expect(await detectFormat(encode(xml))).to.equal('xml');
        });

        it('should detect HTML tables', async () => {
            expect(await detectFormat(encode('<!DOCTYPE html><html><body><table></table></body></html>'))).to.equal('html');
        });

        it('should detect CSV content', async () => {
            expect(await detectFormat(encode('name,age\nAlice,30\nBob,25\n'))).to.equal('csv');
        });

        it('should detect tab-delimited content as tsv', async () => {
            expect(await detectFormat(encode('name\tage\nAlice\t30\n'))).to.equal('tsv');
        });

        it('should detect CSV with a UTF-8 BOM', async () => {
            const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
            const body = encode('a,b\n1,2\n');
            const buffer = new Uint8Array(bom.length + body.length);
            buffer.set(bom);
            buffer.set(body, bom.length);
            expect(await detectFormat(buffer)).to.equal('csv');
        });

        it('should return null for unrecognizable binary content', async () => {
            const buffer = new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x00, 0x1A, 0x00, 0x0A]); // PNG-ish
            expect(await detectFormat(buffer)).to.equal(null);
        });

        it('should return null for empty buffers', async () => {
            expect(await detectFormat(new Uint8Array(0))).to.equal(null);
        });
    });

    describe('parser with buffer input', () => {
        // samples/ is not tracked in git; skip fixture-based tests when absent
        const requireSample = function (ctx, path) {
            if (!fs.existsSync(path)) {
                ctx.skip();
            }
        };

        it('should parse an xlsx sample passed as Uint8Array without a name', async function () {
            requireSample(this, 'samples/test1.xlsx');
            const buffer = new Uint8Array(fs.readFileSync('samples/test1.xlsx'));
            const result = await parser(buffer);
            expect(result.worksheets).to.be.an('array');
            expect(result.worksheets.length).to.be.greaterThan(0);
            expect(result.worksheets[0].data).to.be.an('array');
        });

        it('should parse an xls sample passed as Uint8Array without a name', async function () {
            requireSample(this, 'samples/test.xls');
            const buffer = new Uint8Array(fs.readFileSync('samples/test.xls'));
            const result = await parser(buffer);
            expect(result.worksheets).to.be.an('array');
            expect(result.worksheets.length).to.be.greaterThan(0);
        });

        it('should parse a csv sample passed as Uint8Array without a name', async function () {
            requireSample(this, 'samples/test1.csv');
            const buffer = new Uint8Array(fs.readFileSync('samples/test1.csv'));
            const result = await parser(buffer);
            expect(result.worksheets).to.be.an('array');
            expect(result.worksheets[0].data.length).to.be.greaterThan(0);
        });

        it('should accept ArrayBuffer input', async () => {
            const bytes = encode('a,b\n1,2\n');
            const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
            const result = await parser(arrayBuffer);
            expect(result.worksheets[0].data).to.deep.equal([['a', 'b'], ['1', '2']]);
        });

        it('should honor the extension option over content sniffing', async () => {
            // Semicolon-separated content forced through the CSV parser
            const buffer = encode('a;b\n1;2\n');
            const result = await parser(buffer, { extension: 'csv', delimiter: ';' });
            expect(result.worksheets[0].data).to.deep.equal([['a', 'b'], ['1', '2']]);
        });

        it('should honor the extension option with a leading dot', async () => {
            const buffer = encode('a,b\n1,2\n');
            const result = await parser(buffer, { extension: '.csv' });
            expect(result.worksheets[0].data).to.deep.equal([['a', 'b'], ['1', '2']]);
        });

        it('should sniff the content when the extension is unrecognized', async function () {
            requireSample(this, 'samples/test1.xlsx');
            // xlsx content behind a misleading extension falls back to sniffing
            const buffer = new Uint8Array(fs.readFileSync('samples/test1.xlsx'));
            const result = await parser(buffer, { extension: 'dat' });
            expect(result.worksheets).to.be.an('array');
            expect(result.worksheets.length).to.be.greaterThan(0);
        });

        it('should throw for unrecognizable buffers', async () => {
            try {
                await parser(new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x00, 0x1A]));
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Unsupported file type');
            }
        });
    });

    describe('parser dispatch (regression)', () => {
        it('should route .numbers files to the Numbers parser', async () => {
            try {
                await parser('missing-file.numbers');
            } catch (error) {
                expect(error.message).to.not.include('Unsupported file type');
            }
        });

        it('should still reject plain objects with Invalid file', async () => {
            try {
                await parser({ data: 'test' });
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Invalid file');
            }
        });

        it('should still reject unknown extensions for missing paths', async () => {
            try {
                await parser('test.pdf');
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.include('Unsupported file type');
            }
        });
    });
});
