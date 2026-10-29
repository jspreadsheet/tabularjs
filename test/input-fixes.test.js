import { expect } from 'chai';
import fs from 'fs';
import os from 'os';
import path from 'path';
import JSZip from 'jszip';
import parser from '../src/parser.js';
import { loadAsString } from '../src/utils/loader.js';
import { parseCSVString } from '../src/parser/csv.js';

const encode = (text) => new TextEncoder().encode(text);
const concat = (...parts) => {
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of parts) {
        out.set(p, at);
        at += p.length;
    }
    return out;
};
const utf16le = (text) => {
    const out = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) {
        out[i * 2] = text.charCodeAt(i) & 0xFF;
        out[i * 2 + 1] = text.charCodeAt(i) >> 8;
    }
    return out;
};
const utf16be = (text) => {
    const le = utf16le(text);
    for (let i = 0; i < le.length; i += 2) {
        [le[i], le[i + 1]] = [le[i + 1], le[i]];
    }
    return le;
};

const UTF8_BOM = new Uint8Array([0xEF, 0xBB, 0xBF]);
const UTF16LE_BOM = new Uint8Array([0xFF, 0xFE]);
const UTF16BE_BOM = new Uint8Array([0xFE, 0xFF]);

// --- Minimal BIFF8 workbook inside a CFB container ---------------------------

const u16 = (n) => [n & 0xFF, (n >> 8) & 0xFF];
const u32 = (n) => [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >>> 24) & 0xFF];
const ascii = (s) => Array.from(s, (c) => c.charCodeAt(0));
const rec = (type, data) => [...u16(type), ...u16(data.length), ...data];
const bof = (dt) => rec(0x0809, [...u16(0x0600), ...u16(dt), ...new Array(12).fill(0)]);
const eof = () => rec(0x000A, []);
const label = (row, col, s) => rec(0x0204, [...u16(row), ...u16(col), ...u16(0), ...u16(s.length), 0, ...ascii(s)]);
const boundsheet = (offset, name) => rec(0x0085, [...u32(offset), 0, 0, name.length, 0, ...ascii(name)]);

/** Workbook stream with one sheet "S" holding a single label in A1 */
function workbookStream(text, globalsExtra = []) {
    const globals = (offset) => [...bof(0x0005), ...globalsExtra, ...boundsheet(offset, 'S'), ...eof()];
    const sheet = [...bof(0x0010), ...label(0, 0, text), ...eof()];
    return new Uint8Array([...globals(globals(0).length), ...sheet]);
}

/** Minimal v3 compound file holding one stream (stored in regular sectors) */
function compoundFile(streamName, streamBytes) {
    const SECTOR = 512;
    const put32 = (b, at, v) => new DataView(b.buffer).setUint32(at, v >>> 0, true);
    const nStream = Math.max(1, Math.ceil(streamBytes.length / SECTOR));
    const file = new Uint8Array(SECTOR * (3 + nStream));
    file.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    file.set([0x3E, 0x00, 0x03, 0x00, 0xFE, 0xFF, 0x09, 0x00, 0x06, 0x00], 0x18);
    put32(file, 0x2C, 1);          // FAT sectors
    put32(file, 0x30, 1);          // first directory sector
    put32(file, 0x38, 0);          // mini stream cutoff: every stream in regular sectors
    put32(file, 0x3C, 0xFFFFFFFE); // no MiniFAT
    put32(file, 0x44, 0xFFFFFFFE); // no DIFAT sectors
    for (let i = 0; i < 109; i++) put32(file, 0x4C + i * 4, i === 0 ? 0 : 0xFFFFFFFF);

    const fat = SECTOR;
    for (let i = 0; i < 128; i++) put32(file, fat + i * 4, 0xFFFFFFFF);
    put32(file, fat, 0xFFFFFFFD);     // sector 0: FAT
    put32(file, fat + 4, 0xFFFFFFFE); // sector 1: directory
    for (let i = 0; i < nStream; i++) put32(file, fat + (2 + i) * 4, i === nStream - 1 ? 0xFFFFFFFE : 3 + i);

    const dir = SECTOR * 2;
    const entry = (at, name, type, child, start, size) => {
        for (let i = 0; i < name.length; i++) file[at + i * 2] = name.charCodeAt(i);
        file[at + 64] = (name.length + 1) * 2;
        file[at + 66] = type;
        put32(file, at + 68, 0xFFFFFFFF);
        put32(file, at + 72, 0xFFFFFFFF);
        put32(file, at + 76, child);
        put32(file, at + 116, start);
        put32(file, at + 120, size);
    };
    entry(dir, 'Root Entry', 5, 1, 0xFFFFFFFE, 0);
    entry(dir + 128, streamName, 2, 0xFFFFFFFF, 2, streamBytes.length);
    file.set(streamBytes, SECTOR * 3);
    return file;
}

// RC4 FILEPASS record ([MS-OFFCRYPTO] verifier for "password1", not the default password)
const hex = (s) => Array.from(s.match(/../g), (h) => parseInt(h, 16));
const filepass = rec(0x002F, [
    ...u16(1), ...u16(1), ...u16(1),
    ...hex('e8772c1d91c56a37964761b280183217'),
    ...hex('c9e997d454973d310bb1ba701426837e'),
    ...hex('b1de178f07e989c44dae5e4cf96ac407')
]);

// --- Minimal zip based workbooks ---------------------------------------------

async function buildXLSX(text) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${text}</t></is></c></row></sheetData></worksheet>`);
    return zip.generateAsync({ type: 'uint8array' });
}

async function buildODS(text) {
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    zip.file('content.xml', '<?xml version="1.0" encoding="UTF-8"?>'
        + '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
        + '<office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>'
        + text
        + '</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>');
    return zip.generateAsync({ type: 'uint8array' });
}

const XML_2003 = '<?xml version="1.0"?>\n'
    + '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
    + '<Worksheet ss:Name="S"><Table><Row><Cell><Data ss:Type="String">hello</Data></Cell></Row></Table></Worksheet></Workbook>';

const HTML_TABLE = '<html><body><table><tr><th>title</th></tr><tr><td>hello</td></tr></table></body></html>';

describe('Input handling fixes', () => {
    let dir;
    const write = (name, bytes) => {
        const file = path.join(dir, name);
        fs.writeFileSync(file, bytes);
        return file;
    };

    before(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tabularjs-input-'));
    });

    after(() => {
        fs.rmSync(dir, { recursive: true, force: true });
    });

    describe('Byte order marks', () => {
        const csv = encode('name,age\nAna,30');
        const expected = [['name', 'age'], ['Ana', '30']];

        it('should strip a UTF-8 BOM from a CSV given as a file path (regression)', async () => {
            const result = await parser(write('bom.csv', concat(UTF8_BOM, csv)));
            expect(result.worksheets[0].data).to.deep.equal(expected);
        });

        it('should strip a UTF-8 BOM from a TSV and a TXT given as a file path (regression)', async () => {
            const tsv = concat(UTF8_BOM, encode('name\tage\nAna\t30'));
            expect((await parser(write('bom.tsv', tsv))).worksheets[0].data[0]).to.deep.equal(['name', 'age']);
            expect((await parser(write('bom.txt', tsv))).worksheets[0].data[0]).to.deep.equal(['name', 'age']);
        });

        it('should give the same result for a path, a Buffer, a Uint8Array and a File (regression)', async () => {
            const bytes = concat(UTF8_BOM, csv);
            const file = write('same.csv', bytes);
            const results = [
                await parser(file),
                await parser(Buffer.from(bytes), { extension: 'csv' }),
                await parser(bytes, { extension: 'csv' }),
                await parser(new File([bytes], 'same.csv'))
            ];
            for (const result of results) {
                expect(result.worksheets[0].data).to.deep.equal(expected);
            }
        });

        it('should strip the BOM in loadAsString for every input kind (regression)', async () => {
            const bytes = concat(UTF8_BOM, encode('héllo'));
            const file = write('bom.txt', bytes);
            expect(await loadAsString(file)).to.equal('héllo');
            expect(await loadAsString(bytes)).to.equal('héllo');
            expect(await loadAsString(new Blob([bytes]))).to.equal('héllo');
        });

        it('should decode a UTF-16LE file announced by its BOM (regression)', async () => {
            const file = write('utf16le.csv', concat(UTF16LE_BOM, utf16le('name,city\nZoë,Köln\n')));
            const result = await parser(file);
            expect(result.worksheets[0].data[0]).to.deep.equal(['name', 'city']);
            expect(result.worksheets[0].data[1]).to.deep.equal(['Zoë', 'Köln']);
        });

        it('should decode a UTF-16BE File announced by its BOM (regression)', async () => {
            const bytes = concat(UTF16BE_BOM, utf16be('name,city\nZoë,Köln\n'));
            const result = await parser(new File([bytes], 'utf16be.csv'));
            expect(result.worksheets[0].data[1]).to.deep.equal(['Zoë', 'Köln']);
        });

        it('should let a UTF-8 BOM win over a different requested encoding (regression)', async () => {
            const file = write('bom-1252.csv', concat(UTF8_BOM, encode('café,1\n')));
            const result = await parser(file, { encoding: 'windows-1252' });
            expect(result.worksheets[0].data[0]).to.deep.equal(['café', '1']);
        });

        it('should keep decoding BOM-less files with the requested encoding', async () => {
            const file = write('plain-1252.csv', new Uint8Array([0x63, 0x61, 0x66, 0xE9, 0x2C, 0x31]));
            const result = await parser(file, { encoding: 'windows-1252' });
            expect(result.worksheets[0].data[0]).to.deep.equal(['café', '1']);
        });

        it('should strip a BOM from string content in parseCSVString (regression)', () => {
            const result = parseCSVString('﻿name,age\nAna,30');
            expect(result.worksheets[0].data[0]).to.deep.equal(['name', 'age']);
        });
    });

    describe('File and Blob objects in Node.js', () => {
        it('should parse an xlsx File (regression)', async () => {
            const result = await parser(new File([await buildXLSX('hello')], 'a.xlsx'));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an ods File (regression)', async () => {
            const result = await parser(new File([await buildODS('hello')], 'a.ods'));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an xls File', async () => {
            const bytes = compoundFile('Workbook', workbookStream('hello'));
            const result = await parser(new File([bytes], 'a.xls'));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse a csv File', async () => {
            const result = await parser(new File([encode('a,b\n1,2')], 'a.csv'));
            expect(result.worksheets[0].data).to.deep.equal([['a', 'b'], ['1', '2']]);
        });

        it('should parse an extensionless File with an explicit extension (regression)', async () => {
            const result = await parser(new File([await buildODS('hello')], 'upload'), { extension: 'ods' });
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse a nameless Blob with an explicit extension', async () => {
            const result = await parser(new Blob([await buildXLSX('hello')]), { extension: 'xlsx' });
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse a nameless Blob by sniffing its content', async () => {
            const result = await parser(new Blob([await buildODS('hello')]));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });
    });

    describe('Content detection versus file extension', () => {
        it('should parse an xlsx renamed to .xls (regression)', async () => {
            const result = await parser(write('renamed-xlsx.xls', await buildXLSX('hello')));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an xls renamed to .xlsx (regression)', async () => {
            const file = write('renamed-xls.xlsx', compoundFile('Workbook', workbookStream('hello')));
            const result = await parser(file);
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an ods renamed to .xlsx and an xlsx renamed to .ods (regression)', async () => {
            const ods = await parser(write('renamed-ods.xlsx', await buildODS('hello')));
            expect(ods.worksheets[0].data[0][0]).to.equal('hello');
            const xlsx = await parser(new File([await buildXLSX('hello')], 'renamed-xlsx.ods'));
            expect(xlsx.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse a CSV saved as .xls (regression)', async () => {
            const result = await parser(write('csv.xls', encode('a,b\n1,2')));
            expect(result.worksheets[0].data).to.deep.equal([['a', 'b'], ['1', '2']]);
        });

        it('should parse an XML Spreadsheet 2003 saved as .xls (regression)', async () => {
            const result = await parser(write('xml2003.xls', encode(XML_2003)));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an HTML table saved as .xls (regression)', async () => {
            const result = await parser(new File([encode(HTML_TABLE)], 'export.xls'));
            expect(result.worksheets[0].columns[0].title).to.equal('title');
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should parse an xlsx saved as .csv (regression)', async () => {
            const result = await parser(write('xlsx.csv', await buildXLSX('hello')));
            expect(result.worksheets[0].data[0][0]).to.equal('hello');
        });

        it('should never override an explicit options.extension', async () => {
            const bytes = await buildXLSX('hello');
            try {
                await parser(bytes, { extension: 'xls' });
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.match(/CFB signature/);
            }
        });

        it('should still report a password-protected .xls', async () => {
            const bytes = compoundFile('Workbook', workbookStream('hello', filepass));
            try {
                await parser(write('encrypted.xls', bytes));
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.equal('Workbook is encrypted with a password');
            }
        });

        it('should report a password-protected .xls renamed to .xlsx (regression)', async () => {
            const bytes = compoundFile('Workbook', workbookStream('hello', filepass));
            try {
                await parser(write('encrypted-xls.xlsx', bytes));
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.equal('Workbook is encrypted with a password');
            }
        });

        it('should report a password-protected .xlsx (OLE2 EncryptedPackage) (regression)', async () => {
            const bytes = compoundFile('EncryptedPackage', new Uint8Array(600).fill(7));
            try {
                await parser(write('encrypted.xlsx', bytes));
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.equal('Workbook is encrypted with a password');
            }
        });

        it('should throw the extension parser error when the fallback fails too', async () => {
            // A zip that is neither xlsx nor ods: the ods error is the meaningful one
            const zip = new JSZip();
            zip.file('readme.txt', 'nothing here');
            const bytes = await zip.generateAsync({ type: 'uint8array' });
            try {
                await parser(write('unknown.ods', bytes));
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message).to.match(/content\.xml not found/);
            }
        });

        it('should fire onload once and not onerror when the fallback succeeds (regression)', async () => {
            const calls = [];
            await parser(write('callbacks.xls', encode('a,b\n1,2')), {
                onload: () => calls.push('load'),
                onerror: () => calls.push('error')
            });
            expect(calls).to.deep.equal(['load']);
        });

        it('should fire onerror once when parsing fails', async () => {
            const calls = [];
            try {
                await parser(write('broken.xlsx', new Uint8Array([0x50, 0x4B, 0x03, 0x04, 0, 0, 0, 0])), {
                    onload: () => calls.push('load'),
                    onerror: () => calls.push('error')
                });
            } catch (e) {
                // expected
            }
            expect(calls).to.deep.equal(['error']);
        });
    });
});
