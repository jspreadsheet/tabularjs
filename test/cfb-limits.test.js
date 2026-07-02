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
