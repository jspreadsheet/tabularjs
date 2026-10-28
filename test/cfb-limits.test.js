import { expect } from 'chai';
import fs from 'fs';
import CFB from '../src/utils/cfb.js';

/** Locate a UTF-16LE encoded name inside the raw container bytes */
function findUtf16(bytes, name) {
    const needle = [];
    for (const ch of name) {
        needle.push(ch.charCodeAt(0), 0);
    }
    outer: for (let i = 0; i <= bytes.length - needle.length; i++) {
        for (let j = 0; j < needle.length; j++) {
            if (bytes[i + j] !== needle[j]) continue outer;
        }
        return i;
    }
    return -1;
}

describe('CFB Hardening', () => {
    // samples/ is not tracked in git; skip fixture-based tests when absent
    beforeEach(function () {
        if (!fs.existsSync('samples/test.xls')) {
            this.skip();
        }
    });

    it('should parse a legitimate compound file', () => {
        const bytes = new Uint8Array(fs.readFileSync('samples/test.xls'));
        const cfb = CFB.read(bytes);
        const workbook = CFB.find(cfb, 'Workbook') || CFB.find(cfb, 'Book');
        expect(workbook).to.not.equal(null);
        expect(workbook.content.length).to.be.greaterThan(0);
        expect(workbook.content.length).to.be.lessThan(bytes.length + 1);
    });

    it('should clamp allocations when a directory entry declares a huge size (regression)', () => {
        // Corrupt a real file: point the mini-stream cutoff at 0xFFFFFFFF and
        // declare a ~4 GB stream size in the Workbook directory entry. The
        // old code did `new Uint8Array(entry.size)` straight from the file.
        const bytes = new Uint8Array(fs.readFileSync('samples/test.xls'));

        // Header offset 56: Mini Stream Cutoff Size
        bytes[56] = 0xFF;
        bytes[57] = 0xFF;
        bytes[58] = 0xFF;
        bytes[59] = 0xFF;

        // Directory entry: 64-byte name field, size at entry offset +120
        const namePos = findUtf16(bytes, 'Workbook');
        expect(namePos, 'Workbook directory entry should exist in sample').to.be.greaterThan(-1);
        bytes[namePos + 120] = 0xFE;
        bytes[namePos + 121] = 0xFF;
        bytes[namePos + 122] = 0xFF;
        bytes[namePos + 123] = 0xFF;

        // Must not attempt a ~4 GB allocation; content is clamped to what
        // the container actually holds
        const cfb = CFB.read(bytes);
        const workbook = CFB.find(cfb, 'Workbook');
        expect(workbook).to.not.equal(null);
        expect(workbook.content.length).to.be.at.most(bytes.length);
    });
});

describe('CFB Truncated Files', () => {
    const SECTOR = 512;
    const put32 = (b, at, v) => new DataView(b.buffer).setUint32(at, v >>> 0, true);

    /** Minimal v3 compound file: header, FAT (sector 0), directory (sector 1), "Workbook" stream from sector 2 */
    function build(streamBytes) {
        const nStream = Math.ceil(streamBytes.length / SECTOR);
        const file = new Uint8Array(SECTOR * (3 + nStream));
        file.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        file.set([0x3E, 0x00, 0x03, 0x00, 0xFE, 0xFF, 0x09, 0x00, 0x06, 0x00], 0x18);
        put32(file, 0x2C, 1);          // FAT sectors
        put32(file, 0x30, 1);          // first directory sector
        put32(file, 0x38, 4096);       // mini stream cutoff
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
        entry(dir + 128, 'Workbook', 2, 0xFFFFFFFF, 2, streamBytes.length);
        file.set(streamBytes, SECTOR * 3);
        return file;
    }

    const stream = Uint8Array.from({ length: 5000 }, (_, i) => i % 251);

    it('should read an intact file', () => {
        const wb = CFB.find(CFB.read(build(stream)), 'Workbook');
        expect(Array.from(wb.content)).to.deep.equal(Array.from(stream));
    });

    it('should return the readable part when the last sector is partial', () => {
        const cut = build(stream).slice(0, SECTOR * 3 + 4700);
        const wb = CFB.find(CFB.read(cut), 'Workbook');
        expect(Array.from(wb.content.slice(0, 4700))).to.deep.equal(Array.from(stream.slice(0, 4700)));
    });

    it('should stop where the chain runs past the end of the file', () => {
        const cut = build(stream).slice(0, SECTOR * 3 + SECTOR * 4);
        let wb;
        expect(() => { wb = CFB.find(CFB.read(cut), 'Workbook'); }).to.not.throw();
        expect(wb.content.length).to.equal(SECTOR * 4);
        expect(Array.from(wb.content)).to.deep.equal(Array.from(stream.slice(0, SECTOR * 4)));
    });
});
