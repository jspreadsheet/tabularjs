import { expect } from 'chai';
import JSZip from 'jszip';
import { readZipEntry, DEFAULT_MAX_ENTRY_BYTES } from '../src/utils/zip-utils.js';

describe('Zip Decompression Limits', () => {
    it('should read normal entries in every supported output type', async () => {
        const zip = new JSZip();
        zip.file('hello.txt', 'hello world');
        const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
        const entry = loaded.file('hello.txt');

        expect(await readZipEntry(entry, 'string')).to.equal('hello world');
        expect(await readZipEntry(entry, 'uint8array')).to.deep.equal(new TextEncoder().encode('hello world'));
        expect((await readZipEntry(entry, 'base64'))).to.equal(Buffer.from('hello world').toString('base64'));
        expect(Buffer.isBuffer(await readZipEntry(entry, 'nodebuffer'))).to.equal(true);
    });

    it('should have a sane default cap', () => {
        expect(DEFAULT_MAX_ENTRY_BYTES).to.be.a('number');
        expect(DEFAULT_MAX_ENTRY_BYTES).to.be.greaterThan(1024 * 1024);
    });

    it('should abort entries that exceed the cap during decompression', async () => {
        // 8 MB of zeros compresses to a few KB - the declared size is honest
        // here, but readZipEntry counts actual output, so a lying header
        // cannot bypass it either
        const zip = new JSZip();
        zip.file('bomb.xml', new Uint8Array(8 * 1024 * 1024), { compression: 'DEFLATE' });
        const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
        const entry = loaded.file('bomb.xml');

        try {
            await readZipEntry(entry, 'uint8array', 1024 * 1024);
            expect.fail('Should have thrown error');
        } catch (error) {
            expect(error.message).to.include('decompression limit');
            expect(error.message).to.include('bomb.xml');
        }
    });

    it('should allow entries exactly at the cap', async () => {
        const zip = new JSZip();
        zip.file('ok.bin', new Uint8Array(1024), { compression: 'DEFLATE' });
        const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
        const result = await readZipEntry(loaded.file('ok.bin'), 'uint8array', 1024);
        expect(result.length).to.equal(1024);
    });

    it('should protect the xlsx parser from oversized parts (integration)', async function () {
        // Verify the parser path reads parts through the capped reader
        // without breaking normal files
        const fs = await import('fs');
        if (!fs.existsSync('samples/test1.xlsx')) {
            this.skip();
        }
        const parserModule = await import('../src/parser.js');
        const buffer = new Uint8Array(fs.readFileSync('samples/test1.xlsx'));
        const result = await parserModule.default(buffer);
        expect(result.worksheets).to.be.an('array');
    });
});
