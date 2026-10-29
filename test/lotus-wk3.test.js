import { expect } from 'chai';
import parser from '../src/parser.js';
import { parseLotusBuffer } from '../src/parser/lotus.js';

// Byte builders for synthetic Lotus 1-2-3 release 3+ files
const u16 = (n) => [n & 0xFF, (n >> 8) & 0xFF];
const u32 = (n) => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF];
const f64 = (v) => {
    const b = new Uint8Array(8);
    new DataView(b.buffer).setFloat64(0, v, true);
    return [...b];
};
// 80-bit x87 extended float (WK3/WK4 NUMBER and FORMULA)
const f80 = (v) => {
    if (v === 0) return new Array(10).fill(0);
    const view = new DataView(new ArrayBuffer(8));
    view.setFloat64(0, v);
    const bits = view.getBigUint64(0);
    const sign = Number(bits >> 63n);
    const exp = Number((bits >> 52n) & 0x7FFn) - 1023 + 16383;
    const mantissa = (1n << 63n) | ((bits & ((1n << 52n) - 1n)) << 11n);
    const out = [];
    for (let i = 0n; i < 8n; i++) out.push(Number((mantissa >> (8n * i)) & 0xFFn));
    return [...out, ...u16((sign << 15) | exp)];
};
const f80NaN = [0, 0, 0, 0, 0, 0, 0, 0xC0, 0xFF, 0xFF];
const text = (s) => [...Array.from(s, (c) => c.charCodeAt(0)), 0];
const rec = (type, data) => [...u16(type), ...u16(data.length), ...data];
// Cell address of release 3+ files: row u16, sheet u8, column u8
const addr = (row, sheet, col) => [...u16(row), sheet, col];
const bof = (version) => rec(0x00, [...u16(version), ...u16(0x0004), ...new Array(22).fill(0)]);
const EOF = rec(0x01, []);
const file = (...records) => Uint8Array.from(records.flat());

// 16-bit packed number: integer form, or (integer * factor[index])
const smallInt = (n) => (n << 1) & 0xFFFF;
const smallScaled = (n, index) => ((n << 4) | (index << 1) | 1) & 0xFFFF;
// 32-bit packed number: mantissa << 6 | negative << 5 | divide << 4 | power of ten
const num32 = (mantissa, power = 0, divide = false, negative = false) =>
    ((mantissa << 6) | (negative ? 0x20 : 0) | (divide ? 0x10 : 0) | power) >>> 0;

describe('Lotus 1-2-3 WK3 / WK4 / 123', () => {
    describe('WK3/WK4 records', () => {
        const formulaTokens = [0x01, 0x02, 0x03, 0x04, 0x05, 0x03]; // never decoded

        const wk3 = file(
            bof(0x1000),
            rec(0x1B, [...u16(14000), ...u16(0), ...text('Prices')]),            // named sheet 0
            rec(0x1B, [...u16(14000), ...u16(1), ...text('Totals')]),            // named sheet 1
            rec(0x1B, [...u16(2007), 0, 0, 0, 0]),                               // other extended record: ignored
            rec(0x07, [0, 0, 0, 0, 0, 12, 2, 20]),                               // sheet 0: A = 12 chars, C = 20 chars
            rec(0x07, [0, 1, 0, 0, 1, 50]),                                      // second window: ignored
            rec(0x08, [0, 0, 0, 0, 3]),                                          // sheet 0: D hidden
            rec(0x16, [...addr(0, 0, 0), ...text("'Left")]),
            rec(0x16, [...addr(0, 0, 1), ...text('"Right')]),
            rec(0x16, [...addr(0, 0, 2), ...text('^Center')]),
            rec(0x16, [...addr(0, 0, 3), ...text('|printer command')]),          // not cell content
            rec(0x17, [...addr(1, 0, 0), ...f80(3.25)]),
            rec(0x17, [...addr(1, 0, 1), ...f80(-1234.5)]),
            rec(0x17, [...addr(1, 0, 2), ...f80(1e-5)]),
            rec(0x17, [...addr(1, 0, 3), ...f80(0)]),
            rec(0x18, [...addr(2, 0, 0), ...u16(smallInt(7))]),
            rec(0x18, [...addr(2, 0, 1), ...u16(smallInt(-3))]),
            rec(0x18, [...addr(2, 0, 2), ...u16(smallScaled(3, 2))]),            // 3 * 0.05
            rec(0x18, [...addr(2, 0, 3), ...u16(smallScaled(-2, 0))]),           // -2 * 5000
            rec(0x19, [...addr(3, 0, 0), ...f80(42), ...formulaTokens]),
            rec(0x19, [...addr(3, 0, 1), ...f80NaN, ...formulaTokens]),          // string result follows
            rec(0x1A, [...addr(3, 0, 1), ...text('text result')]),
            rec(0x19, [...addr(3, 0, 2), ...f80NaN, ...formulaTokens]),          // no string record
            rec(0x14, addr(4, 0, 0)),
            rec(0x15, addr(4, 0, 1)),
            rec(0x16, [...addr(0, 1, 1), ...text("'Second sheet")]),
            rec(0x17, [...addr(1, 1, 1), ...f80(99)]),
            EOF,
            rec(0x16, [...addr(0, 0, 0), ...text("'after EOF")])                // ignored
        );

        it('should read labels, numbers, small numbers, formula results and error cells', () => {
            const [sheet] = parseLotusBuffer(wk3).worksheets;
            expect(sheet.data).to.deep.equal([
                ['Left', 'Right', 'Center', ''],
                [3.25, -1234.5, 1e-5, 0],
                [7, -3, 0.15000000000000002, -10000],
                [42, 'text result', '', ''],
                ['#ERR!', '#N/A', '', '']
            ]);
        });

        it('should apply the label alignment prefixes', () => {
            const [sheet] = parseLotusBuffer(wk3).worksheets;
            expect(sheet.style.A1).to.equal('text-align: left');
            expect(sheet.style.B1).to.equal('text-align: right');
            expect(sheet.style.C1).to.equal('text-align: center');
        });

        it('should read every sheet with its name', () => {
            const { worksheets } = parseLotusBuffer(wk3);
            expect(worksheets.map((w) => w.worksheetName)).to.deep.equal(['Prices', 'Totals']);
            expect(worksheets[1].data).to.deep.equal([['', 'Second sheet'], ['', 99]]);
        });

        it('should read column widths and hidden columns per sheet', () => {
            const { worksheets } = parseLotusBuffer(wk3);
            const cols = worksheets[0].columns;
            expect(cols[0].width).to.be.closeTo(12 * 8.43, 1e-9);
            expect(cols[1].width).to.equal(100);
            expect(cols[2].width).to.be.closeTo(20 * 8.43, 1e-9);
            expect(cols[3].hidden).to.equal(true);
            expect(worksheets[1].columns[1].width).to.equal(100);
        });

        it('should not expose undecoded formula tokens', () => {
            const [sheet] = parseLotusBuffer(wk3).worksheets;
            expect(sheet.cells).to.deep.equal({});
        });

        it('should read WK4 files (version 0x1002) with the same records', async () => {
            const wk4 = file(
                bof(0x1002),
                rec(0x16, [...addr(0, 0, 0), ...text("'wk4")]),
                rec(0x17, [...addr(0, 0, 1), ...f80(2.5)]),
                EOF
            );
            const result = await parser(wk4, { extension: 'wk4' });
            expect(result.worksheets[0].data).to.deep.equal([['wk4', 2.5]]);
        });

        it('should keep a named sheet that has no cells', () => {
            const result = parseLotusBuffer(file(
                bof(0x1000),
                rec(0x1B, [...u16(14000), ...u16(0), ...text('Data')]),
                rec(0x1B, [...u16(14000), ...u16(1), ...text('Empty')]),
                rec(0x17, [...addr(0, 0, 0), ...f80(1)]),
                EOF
            ));
            expect(result.worksheets.map((w) => w.worksheetName)).to.deep.equal(['Data', 'Empty']);
            expect(result.worksheets[1].data).to.deep.equal([]);
        });

        it('should reject password-protected files', () => {
            expect(() => parseLotusBuffer(file(bof(0x1000), rec(0x02, [1, 2, 3, 4]), EOF)))
                .to.throw('Password-protected');
        });
    });

    describe('1-2-3 97 / Millennium (.123) records', () => {
        const tokens = [0x01, 0x02, 0x03];

        const wk123 = file(
            bof(0x1003),
            rec(0x23, [...u16(14000), ...u16(0), ...text('Main')]),
            rec(0x23, [...u16(14000), ...u16(1), ...text('Other')]),
            rec(0x16, [...addr(0, 0, 0), ...text("'Label")]),
            rec(0x25, [...addr(1, 0, 0), ...u32(num32(1))]),                     // 1
            rec(0x25, [...addr(1, 0, 1), ...u32(num32(25, 1, true, true))]),     // -2.5
            rec(0x25, [...addr(1, 0, 2), ...u32(num32(12, 3))]),                 // 12000
            rec(0x27, [...addr(2, 0, 0), ...f64(Math.PI)]),
            rec(0x28, [...addr(2, 0, 1), ...f64(6), ...tokens]),
            rec(0x28, [...addr(2, 0, 2), ...f64(NaN), ...tokens]),
            rec(0x1A, [...addr(2, 0, 2), ...text('abc')]),
            rec(0x26, [...addr(3, 0, 4), ...text('A note')]),
            rec(0x27, [...addr(0, 1, 0), ...f64(5)]),
            EOF
        );

        it('should read labels, packed and IEEE numbers and formula results', () => {
            const [sheet] = parseLotusBuffer(wk123).worksheets;
            expect(sheet.data.slice(0, 3)).to.deep.equal([
                ['Label', '', '', '', ''],
                [1, -2.5, 12000, '', ''],
                [Math.PI, 6, 'abc', '', '']
            ]);
        });

        it('should read sheet names', () => {
            const { worksheets } = parseLotusBuffer(wk123);
            expect(worksheets.map((w) => w.worksheetName)).to.deep.equal(['Main', 'Other']);
            expect(worksheets[1].data).to.deep.equal([[5]]);
        });

        it('should return notes as comments', () => {
            const [sheet] = parseLotusBuffer(wk123).worksheets;
            expect(sheet.comments).to.deep.equal({ E4: 'A note' });
            expect(sheet.data).to.have.length(4);
        });

        it('should decode LMBCS text', () => {
            const result = parseLotusBuffer(file(
                bof(0x1005),
                // CR LF as group 0x0F controls, group 1 and plain code page 850 accents, a UTF-16 euro sign
                rec(0x16, [...addr(0, 0, 0), 0x27, 0x61, 0x0F, 0x2D, 0x0F, 0x2A, 0x62, 0x01, 0x82, 0x81, 0x14, 0x20, 0xAC, 0]),
                EOF
            ));
            expect(result.worksheets[0].data).to.deep.equal([['a\r\nbéü€']]);
        });
    });

    describe('Format detection by BOF version', () => {
        it('should route a .123 layout by its BOF, whatever the extension', async () => {
            const buffer = file(
                bof(0x1005),
                rec(0x25, [...addr(0, 0, 0), ...u32(num32(25000))]),
                EOF
            );
            const result = await parser(buffer, { extension: 'wks' });
            expect(result.worksheets[0].data).to.deep.equal([[25000]]);
        });

        it('should keep the WK1 layout for WKS/WK1 BOF versions', () => {
            // 0x16/0x17 mean something else in WK1; the legacy NUMBER record (0x0E) must still work
            const wk1 = file(
                rec(0x00, u16(0x0406)),
                rec(0x0E, [0xFF, ...u16(1), ...u16(0), ...f64(21)]),
                EOF
            );
            expect(parseLotusBuffer(wk1).worksheets[0].data).to.deep.equal([['', 21]]);
        });
    });

    describe('Damaged files', () => {
        it('should return no sheet for a file with only a BOF', () => {
            expect(parseLotusBuffer(file(bof(0x1000))).worksheets).to.deep.equal([]);
        });

        it('should keep the cells read before a truncated record', () => {
            const buffer = file(
                bof(0x1003),
                rec(0x27, [...addr(0, 0, 0), ...f64(1)]),
                [...u16(0x16), ...u16(8000), 0x00, 0x00, 0x00]                   // declares more bytes than the file has
            );
            expect(parseLotusBuffer(buffer).worksheets[0].data).to.deep.equal([[1]]);
        });

        it('should fail cleanly when stray addresses would expand to a huge grid', () => {
            const records = [bof(0x1003)];
            for (let sheet = 0; sheet < 4; sheet++) {
                records.push(rec(0x27, [...addr(65535, sheet, 255), ...f64(1)]));
            }
            records.push(EOF);
            expect(() => parseLotusBuffer(file(...records))).to.throw('exceed the supported sheet size');
        });

        it('should ignore cell records too short for their layout', () => {
            const buffer = file(
                bof(0x1000),
                rec(0x17, [...addr(0, 0, 0), 1, 2]),
                rec(0x16, [0]),
                rec(0x19, addr(0, 0, 1)),
                rec(0x1A, [0]),
                rec(0x1B, [0x0B]),
                rec(0x18, [...addr(0, 0, 2), ...u16(smallInt(4))]),
                EOF
            );
            expect(parseLotusBuffer(buffer).worksheets[0].data).to.deep.equal([['', '', 4]]);
        });
    });
});
