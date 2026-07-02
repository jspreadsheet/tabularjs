import { expect } from 'chai';
import { loadAsString, detectEncoding } from '../src/utils/loader.js';

describe('Encoding Support', () => {
    describe('loadAsString', () => {
        it('should decode UTF-8 by default', async () => {
            const bytes = new TextEncoder().encode('héllo, wörld');
            expect(await loadAsString(bytes)).to.equal('héllo, wörld');
        });

        it('should decode windows-1251 (Cyrillic) correctly (regression)', async () => {
            // "Привет" in windows-1251 - the old code decoded this as latin1
            const bytes = new Uint8Array([0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2]);
            expect(await loadAsString(bytes, 'windows-1251')).to.equal('Привет');
        });

        it('should decode windows-1252 smart quotes correctly (regression)', async () => {
            // 0x93/0x94 are curly quotes in windows-1252 but C1 controls in latin1
            const bytes = new Uint8Array([0x93, 0x48, 0x69, 0x94]);
            expect(await loadAsString(bytes, 'windows-1252')).to.equal('“Hi”');
        });

        it('should decode Shift_JIS via TextDecoder (regression)', async () => {
            // "テスト" in Shift_JIS - the old code fell back to utf-8
            const bytes = new Uint8Array([0x83, 0x65, 0x83, 0x58, 0x83, 0x67]);
            expect(await loadAsString(bytes, 'shift_jis')).to.equal('テスト');
        });

        it('should decode GBK via TextDecoder (regression)', async () => {
            // "中文" in GBK
            const bytes = new Uint8Array([0xD6, 0xD0, 0xCE, 0xC4]);
            expect(await loadAsString(bytes, 'gbk')).to.equal('中文');
        });

        it('should decode DOS codepages via the iconv-lite fallback', async () => {
            // "Café" in cp437 (0x82 = é) - TextDecoder has no cp437 label
            const bytes = new Uint8Array([0x43, 0x61, 0x66, 0x82]);
            expect(await loadAsString(bytes, 'cp437')).to.equal('Café');
        });

        it('should throw a helpful error for unknown encodings', async () => {
            const bytes = new Uint8Array([0x41]);
            try {
                await loadAsString(bytes, 'not-a-real-encoding');
                expect.fail('Should have thrown error');
            } catch (error) {
                expect(error.message.toLowerCase()).to.include('encoding');
            }
        });
    });

    describe('detectEncoding', () => {
        it('should not report Cyrillic text as latin1 (regression)', async () => {
            // A windows-1251 Cyrillic sample; the old map aliased the
            // detection result to latin1, producing mojibake downstream
            const text = 'Привет мир. Это тестовый файл с русским текстом для определения кодировки.';
            const bytes = new Uint8Array([...text].map((ch) => {
                const code = ch.codePointAt(0);
                if (code >= 0x0410 && code <= 0x044F) return code - 0x0410 + 0xC0;
                if (code === 0x0451) return 0xB8;
                return code < 0x80 ? code : 0x3F;
            }));
            const detected = await detectEncoding(bytes);
            expect(detected).to.not.equal('latin1');
            expect(detected).to.not.equal('utf-8');
            // Whatever chardet reports must be decodable to the original text
            const roundTrip = await loadAsString(bytes, detected);
            expect(roundTrip).to.include('Привет');
        });

        it('should detect UTF-8 text', async () => {
            const bytes = new TextEncoder().encode('Hello wörld, こんにちは, привет — plenty of UTF-8 here.');
            expect(await detectEncoding(bytes)).to.equal('utf-8');
        });
    });
});
