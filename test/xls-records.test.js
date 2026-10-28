import { expect } from 'chai';
import { parseWorkbookStream } from '../src/parser/xls.js';
import { md5, decryptWorkbook } from '../src/utils/xls-crypto.js';

// ------------------------------------------------------------------------- BIFF8 stream builders

const u16 = (n) => [n & 0xFF, (n >> 8) & 0xFF];
const u32 = (n) => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF];
const ascii = (s) => Array.from(s, (c) => c.charCodeAt(0));
const utf16 = (s) => Array.from(s).flatMap((c) => u16(c.charCodeAt(0)));
const rec = (type, data) => [...u16(type), ...u16(data.length), ...data];
const bof = (dt) => rec(0x0809, [...u16(0x0600), ...u16(dt), ...new Array(12).fill(0)]);
const eof = () => rec(0x000A, []);
const labelsst = (row, col, isst) => rec(0x00FD, [...u16(row), ...u16(col), ...u16(0), ...u32(isst)]);
const label = (row, col, s) => rec(0x0204, [...u16(row), ...u16(col), ...u16(0), ...u16(s.length), 0, ...ascii(s)]);
const number = (row, col, v) => {
    const b = new Uint8Array(8);
    new DataView(b.buffer).setFloat64(0, v, true);
    return rec(0x0203, [...u16(row), ...u16(col), ...u16(0), ...b]);
};
const boundsheet = (offset, type, name) => rec(0x0085, [...u32(offset), 0, type, name.length, 0, ...ascii(name)]);
const xlstr = (s) => [...u16(s.length), 0, ...ascii(s)];
const supbookSelf = (nsheets) => rec(0x01AE, [...u16(nsheets), 0x01, 0x04]);
const supbookExternal = (path, sheetNames) => rec(0x01AE, [...u16(sheetNames.length), ...xlstr(path), ...sheetNames.flatMap(xlstr)]);
const externname = (name) => rec(0x0023, [...u16(0), ...u32(0), name.length, 0, ...ascii(name)]);
const externsheet = (entries) => rec(0x0017, [...u16(entries.length), ...entries.flatMap((e) => [...u16(e[0]), ...u16(e[1]), ...u16(e[2])])]);
const definedName = (name, flags = 0) => rec(0x0018, [...u16(flags), 0, name.length, ...u16(0), ...u16(0), ...u16(0), 0, 0, 0, 0, 0, ...name]);
/** FORMULA record: result is the 8-byte FormulaValue, rgce the tokens, extra the RgbExtra */
const formula = (row, col, result, rgce, extra = []) =>
    rec(0x0006, [...u16(row), ...u16(col), ...u16(0), ...result, ...u16(0), ...u32(0), ...u16(rgce.length), ...rgce, ...extra]);
const numResult = (v) => {
    const b = new Uint8Array(8);
    new DataView(b.buffer).setFloat64(0, v, true);
    return [...b];
};
const special = (kind, v = 0) => [kind, 0, v, 0, 0, 0, 0xFF, 0xFF];
const tExp = (row, col) => [0x01, ...u16(row), ...u16(col)];

/** globals (with the given SST records) + sheet substreams; BOUNDSHEET offsets computed from the layout */
function workbook(sstRecords, sheets) {
    const globalsWithout = (offsets) => [
        ...bof(0x0005),
        ...sheets.flatMap((sh, i) => boundsheet(offsets[i], sh.type || 0, sh.name)),
        ...sstRecords,
        ...eof()
    ];
    let offsets = sheets.map(() => 0);
    let pos = globalsWithout(offsets).length;
    offsets = sheets.map((sh) => {
        const at = pos;
        pos += sh.body.length;
        return at;
    });
    return new Uint8Array([...globalsWithout(offsets), ...sheets.flatMap((sh) => sh.body)]);
}

const worksheet = (cells) => [...bof(0x0010), ...cells, ...eof()];

describe('XLS records', () => {
    describe('Shared String Table', () => {
        // 0 plain, 1 rich text (fRichSt: cRun + runs after the characters), 2 phonetic (fExtSt: cbExtRst + data),
        // 3 characters split across a CONTINUE record that switches to 16-bit, 4 a string after all of them
        const strings = [
            [...u16(5), 0x00, ...ascii('plain')],
            [...u16(4), 0x08, ...u16(2), ...ascii('Bold'), ...new Array(8).fill(0x07)],
            [...u16(3), 0x04, ...u32(6), ...ascii('Ext'), ...new Array(6).fill(0x01)],
            [...u16(5), 0x00, ...ascii('He')]
        ];
        const sst = [
            ...rec(0x00FC, [...u32(5), ...u32(5), ...strings.flat()]),
            ...rec(0x003C, [0x01, ...utf16('llo'), ...u16(5), 0x00, ...ascii('after')])
        ];
        const body = worksheet([0, 1, 2, 3, 4].map((i) => labelsst(0, i, i)).flat());

        it('should skip rich-text runs and phonetic data, and follow CONTINUE records', () => {
            const r = parseWorkbookStream(workbook(sst, [{ name: 'S', body }]));
            expect(r.worksheets[0].data[0]).to.deep.equal(['plain', 'Bold', 'Ext', 'Hello', 'after']);
        });
    });

    describe('Sheet substreams', () => {
        it('should keep cells after a chart embedded in a worksheet', () => {
            const body = worksheet([
                ...number(0, 0, 1),
                ...bof(0x0020), ...number(0, 0, 99), ...eof(), // chart substream: not cells of the sheet
                ...number(1, 0, 42)
            ]);
            const r = parseWorkbookStream(workbook([], [{ name: 'S', body }]));
            expect(r.worksheets[0].data[0][0]).to.equal(1);
            expect(r.worksheets[0].data[1][0]).to.equal(42);
        });

        it('should align sheets after a chart sheet', () => {
            const r = parseWorkbookStream(workbook([], [
                { name: 'A', body: worksheet(number(0, 0, 1)) },
                { name: 'Chart1', type: 2, body: [...bof(0x0020), ...eof()] },
                { name: 'B', body: worksheet([...number(0, 0, 7), ...label(0, 1, 'inline')]) }
            ]));
            expect(r.worksheets.map((w) => w.worksheetName)).to.deep.equal(['A', 'Chart1', 'B']);
            expect(r.worksheets[0].data[0][0]).to.equal(1);
            expect(r.worksheets[2].data[0]).to.deep.equal([7, 'inline']);
        });
    });

    describe('Formula context (SUPBOOK, EXTERNSHEET, EXTERNNAME, NAME)', () => {
        const globals = [
            ...supbookSelf(2),
            ...supbookExternal('\x01C:\x03data\x03Book2.xls', ['Prices']),
            ...externname('Rate'),
            ...externsheet([[0, 1, 1], [1, 0, 0], [0, 0, 1]]),
            ...definedName(ascii('Total')),
            ...definedName([0x06], 0x0020) // built-in name 6: Print_Area
        ];
        const cells = [
            ...formula(0, 0, numResult(5), [0x5A, ...u16(0), ...u16(0), ...u16(0xC000)]),            // =Other!A1
            ...formula(0, 1, numResult(10), [0x43, ...u16(1), ...u16(0), 0x1E, ...u16(2), 0x05]),     // =Total*2
            ...formula(0, 2, numResult(3), [0x5A, ...u16(1), ...u16(0), ...u16(0xC000)]),             // external sheet
            ...formula(0, 3, numResult(1), [0x59, ...u16(1), ...u16(1), ...u16(0)]),                  // external name
            ...formula(0, 4, numResult(0), [0x3B, ...u16(2), ...u16(0), ...u16(0), ...u16(0xC000), ...u16(0xC000),
                0x22, 0x01, ...u16(4)]),                                                               // SUM(Main:Other!A1:A1)
            ...formula(0, 5, numResult(0), [0x23, ...u16(2), ...u16(0)])                               // =Print_Area
        ];
        const r = parseWorkbookStream(workbook(globals, [
            { name: 'Main', body: worksheet(cells) },
            { name: 'Other', body: worksheet(number(0, 0, 5)) }
        ]));
        const row = r.worksheets[0].data[0];

        it('should decode references to other sheets', () => {
            expect(row[0]).to.equal('=Other!A1');
            expect(row[4]).to.equal("=SUM('Main:Other'!A1:A1)");
        });

        it('should decode defined names, including built-in names', () => {
            expect(row[1]).to.equal('=Total*2');
            expect(row[5]).to.equal('=Print_Area');
        });

        it('should decode external workbooks from the SUPBOOK path', () => {
            expect(row[2]).to.equal("='[Book2.xls]Prices'!A1");
            expect(row[3]).to.equal('=[Book2.xls]!Rate');
        });

        it('should keep the formula text in the cell metadata', () => {
            expect(r.worksheets[0].cells.B1.formula).to.equal('Total*2');
        });

        it('should split EXTERNSHEET across CONTINUE records', () => {
            const body = [...u16(1), ...u16(0), ...u16(1), ...u16(1)];
            const continued = [...rec(0x0017, body.slice(0, 5)), ...rec(0x003C, body.slice(5))]; // entry split in half
            const out = parseWorkbookStream(workbook([...supbookSelf(2), ...continued], [
                { name: 'Main', body: worksheet(formula(0, 0, numResult(5), [0x5A, ...u16(0), ...u16(0), ...u16(0xC000)])) },
                { name: 'Other', body: worksheet([]) }
            ]));
            expect(out.worksheets[0].data[0][0]).to.equal('=Other!A1');
        });
    });

    describe('Formula results', () => {
        // tExp without a SHRFMLA/ARRAY record cannot be decoded: the cell falls back to the cached result
        const orphan = tExp(9, 9);
        const r = parseWorkbookStream(workbook([], [{
            name: 'S', body: worksheet([
                ...formula(0, 0, special(1, 1), orphan),
                ...formula(0, 1, special(2, 0x07), orphan),
                ...formula(0, 2, special(3), orphan),
                ...formula(0, 3, numResult(2.5), orphan),
                ...formula(0, 4, special(1, 0), [0x1D, 0x00]), // =FALSE decodes normally
                ...rec(0x0205, [...u16(0), ...u16(5), ...u16(0), 0x2A, 0x01]), // BOOLERR #N/A
                ...rec(0x0205, [...u16(0), ...u16(6), ...u16(0), 0x01, 0x00])  // BOOLERR TRUE
            ])
        }]));
        const row = r.worksheets[0].data[0];

        it('should fall back to cached boolean, error, empty-string and number results', () => {
            expect(row.slice(0, 4)).to.deep.equal([true, '#DIV/0!', '', 2.5]);
        });

        it('should prefer the formula when it decodes', () => {
            expect(row[4]).to.equal('=FALSE');
        });

        it('should read BOOLERR errors and booleans', () => {
            expect(row[5]).to.equal('#N/A');
            expect(row[6]).to.equal(true);
        });
    });

    describe('Array formulas (ARRAY record)', () => {
        it('should give every cell of the range the array formula, with its constants', () => {
            // {=A3:B3*{2,3}} over A1:B1
            const rgce = [0x65, ...u16(2), ...u16(2), ...u16(0xC000), ...u16(0xC001), 0x60, ...new Array(7).fill(0), 0x05];
            const extra = [0x01, ...u16(0), 0x01, ...numResult(2), 0x01, ...numResult(3)];
            const array = rec(0x0221, [...u16(0), ...u16(0), 0, 1, ...u16(0), ...u32(0), ...u16(rgce.length), ...rgce, ...extra]);
            const r = parseWorkbookStream(workbook([], [{
                name: 'S', body: worksheet([
                    ...formula(0, 0, numResult(2), tExp(0, 0)),
                    ...array,
                    ...formula(0, 1, numResult(6), tExp(0, 0))
                ])
            }]));
            expect(r.worksheets[0].data[0]).to.deep.equal(['=A3:B3*{2,3}', '=A3:B3*{2,3}']);
        });
    });

    describe('RC4 encryption', () => {
        const hex = (s) => Uint8Array.from(s.match(/../g), (h) => parseInt(h, 16));

        it('should compute MD5', () => {
            const digest = (s) => Buffer.from(md5(Uint8Array.from(ascii(s)))).toString('hex');
            expect(digest('')).to.equal('d41d8cd98f00b204e9800998ecf8427e');
            expect(digest('The quick brown fox jumps over the lazy dog')).to.equal('9e107d9d372bb6826bd81d3542a419d6');
            expect(digest('a'.repeat(200))).to.equal('887f30b43b2867f4a9accceee7d16e6c'); // several 64-byte blocks
        });

        // [MS-OFFCRYPTO] RC4 verifier for password "password1" (vector from msoffcrypto-tool)
        const filepass = rec(0x002F, [
            ...u16(1), ...u16(1), ...u16(1),
            ...hex('e8772c1d91c56a37964761b280183217'),
            ...hex('c9e997d454973d310bb1ba701426837e'),
            ...hex('b1de178f07e989c44dae5e4cf96ac407')
        ]);

        it('should accept the right password and reject others', () => {
            const stream = new Uint8Array([...bof(0x0005), ...filepass, ...eof()]);
            expect(() => decryptWorkbook(stream, 'password1')).to.not.throw();
            expect(() => decryptWorkbook(stream)).to.throw(/encrypted with a password/);
        });

        it('should leave headers, BOF and the BOUNDSHEET offset plain, and round-trip the rest', () => {
            const plain = new Uint8Array([
                ...bof(0x0005), ...filepass, ...boundsheet(0x1234, 0, 'Sheet'), ...number(3, 4, 2.5), ...eof()
            ]);
            const enc = decryptWorkbook(plain, 'password1'); // RC4 is symmetric
            const fpAt = bof(0x0005).length;
            expect(Array.from(enc.slice(0, fpAt))).to.deep.equal(Array.from(plain.slice(0, fpAt)));
            const bsAt = fpAt + filepass.length;
            expect(Array.from(enc.slice(bsAt, bsAt + 8))).to.deep.equal(Array.from(plain.slice(bsAt, bsAt + 8)));
            expect(Array.from(enc.slice(bsAt + 8))).to.not.deep.equal(Array.from(plain.slice(bsAt + 8)));
            // decrypting again restores the data (FILEPASS was turned into record type 0)
            enc.set(u16(0x002F), fpAt);
            const back = decryptWorkbook(enc, 'password1');
            expect(Array.from(back.slice(bsAt))).to.deep.equal(Array.from(plain.slice(bsAt)));
        });

        it('should report unsupported encryption', () => {
            const xor = new Uint8Array([...bof(0x0005), ...rec(0x002F, [...u16(0), ...u16(0x1234), ...u16(0x5678)]), ...eof()]);
            expect(() => decryptWorkbook(xor)).to.throw(/XOR obfuscation/);
        });
    });
});
