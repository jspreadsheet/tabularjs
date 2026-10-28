/**
 * Decryption of BIFF8 workbooks protected with Office 97 RC4 encryption ([MS-XLS] 2.2.10, [MS-OFFCRYPTO] 2.3.6).
 *
 * Excel encrypts "write-protected" / read-only recommended workbooks with the fixed password "VelvetSweatshop"
 * and opens them without asking; those are decrypted here. Other passwords, RC4 CryptoAPI and XOR obfuscation
 * raise an error that says the workbook is encrypted.
 *
 * @license MIT
 */

import { readUInt16LE } from './helpers.js';

export const DEFAULT_PASSWORD = 'VelvetSweatshop';

const FILEPASS = 0x002F;
const BOUNDSHEET = 0x0085;
// [MS-XLS] 2.2.10: records that are never encrypted
const PLAIN_RECORDS = new Set([0x0809, FILEPASS, 0x0194, 0x0195, 0x00E1, 0x0196, 0x0138]);
const BLOCK = 1024;

// ---------------------------------------------------------------------------------------------------- MD5 (RFC 1321)

const MD5_S = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
const MD5_K = new Int32Array(64);
for (let i = 0; i < 64; i++) {
    MD5_K[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 0x100000000) | 0;
}

/**
 * MD5 digest of a byte array -> Uint8Array(16)
 */
export function md5(bytes) {
    const n = bytes.length;
    const padded = new Uint8Array(((n + 8) >> 6) * 64 + 64);
    padded.set(bytes);
    padded[n] = 0x80;
    const bits = n * 8;
    for (let i = 0; i < 8; i++) {
        padded[padded.length - 8 + i] = i < 4 ? (bits >>> (8 * i)) & 0xFF : Math.floor(bits / 2 ** (8 * i)) & 0xFF;
    }
    let a0 = 0x67452301, b0 = 0xEFCDAB89 | 0, c0 = 0x98BADCFE | 0, d0 = 0x10325476;
    const m = new Int32Array(16);
    for (let off = 0; off < padded.length; off += 64) {
        for (let i = 0; i < 16; i++) {
            const p = off + i * 4;
            m[i] = padded[p] | (padded[p + 1] << 8) | (padded[p + 2] << 16) | (padded[p + 3] << 24);
        }
        let a = a0, b = b0, c = c0, d = d0;
        for (let i = 0; i < 64; i++) {
            let f, g;
            if (i < 16) {
                f = (b & c) | (~b & d);
                g = i;
            } else if (i < 32) {
                f = (d & b) | (~d & c);
                g = (5 * i + 1) & 15;
            } else if (i < 48) {
                f = b ^ c ^ d;
                g = (3 * i + 5) & 15;
            } else {
                f = c ^ (b | ~d);
                g = (7 * i) & 15;
            }
            const t = d;
            d = c;
            c = b;
            const x = (a + f + MD5_K[i] + m[g]) | 0;
            const s = MD5_S[(i >> 4) * 4 + (i & 3)];
            b = (b + ((x << s) | (x >>> (32 - s)))) | 0;
            a = t;
        }
        a0 = (a0 + a) | 0;
        b0 = (b0 + b) | 0;
        c0 = (c0 + c) | 0;
        d0 = (d0 + d) | 0;
    }
    const out = new Uint8Array(16);
    [a0, b0, c0, d0].forEach((w, k) => {
        for (let i = 0; i < 4; i++) out[k * 4 + i] = (w >>> (8 * i)) & 0xFF;
    });
    return out;
}

// ---------------------------------------------------------------------------------------------------- RC4

function rc4Stream(key, length) {
    const s = new Uint8Array(256);
    for (let i = 0; i < 256; i++) s[i] = i;
    for (let i = 0, j = 0; i < 256; i++) {
        j = (j + s[i] + key[i % key.length]) & 0xFF;
        const t = s[i];
        s[i] = s[j];
        s[j] = t;
    }
    const out = new Uint8Array(length);
    for (let k = 0, i = 0, j = 0; k < length; k++) {
        i = (i + 1) & 0xFF;
        j = (j + s[i]) & 0xFF;
        const t = s[i];
        s[i] = s[j];
        s[j] = t;
        out[k] = s[(s[i] + s[j]) & 0xFF];
    }
    return out;
}

// ---------------------------------------------------------------------------------------------------- key derivation

/**
 * [MS-OFFCRYPTO] 2.3.6.2: the 5-byte truncated hash of (password, salt) that every block key derives from
 */
function baseKey(password, salt) {
    const pw = new Uint8Array(password.length * 2);
    for (let i = 0; i < password.length; i++) {
        pw[i * 2] = password.charCodeAt(i) & 0xFF;
        pw[i * 2 + 1] = password.charCodeAt(i) >> 8;
    }
    const h0 = md5(pw).slice(0, 5);
    const buf = new Uint8Array(21 * 16);
    for (let i = 0; i < 16; i++) {
        buf.set(h0, i * 21);
        buf.set(salt, i * 21 + 5);
    }
    return md5(buf).slice(0, 5);
}

function blockKey(base, block) {
    const b = new Uint8Array(9);
    b.set(base);
    b[5] = block & 0xFF;
    b[6] = (block >>> 8) & 0xFF;
    b[7] = (block >>> 16) & 0xFF;
    b[8] = (block >>> 24) & 0xFF;
    return md5(b);
}

/**
 * [MS-OFFCRYPTO] 2.3.6.4 password verification: RC4(block 0) over EncryptedVerifier + EncryptedVerifierHash
 */
function verify(base, verifier, verifierHash) {
    const ks = rc4Stream(blockKey(base, 0), 32);
    const v = verifier.map((x, i) => x ^ ks[i]);
    const vh = verifierHash.map((x, i) => x ^ ks[16 + i]);
    const h = md5(v);
    return h.every((x, i) => x === vh[i]);
}

// ---------------------------------------------------------------------------------------------------- workbook stream

/**
 * The FILEPASS record of a workbook stream (in the globals, before the first sheet), or null
 */
export function findFilePass(data) {
    for (let off = 0; off + 4 <= data.length;) {
        const type = readUInt16LE(data, off);
        const len = readUInt16LE(data, off + 2);
        if (type === FILEPASS) return data.subarray(off + 4, off + 4 + len);
        if (type === 0x000A) return null; // EOF of the globals
        off += 4 + len;
    }
    return null;
}

/**
 * Decrypt a BIFF8 workbook stream that holds a FILEPASS record -> new Uint8Array (FILEPASS turned into an
 * ignorable record). Throws when the encryption is not RC4 with the default password.
 */
export function decryptWorkbook(data, password = DEFAULT_PASSWORD) {
    const fp = findFilePass(data);
    if (!fp) return data;
    const type = fp.length >= 2 ? readUInt16LE(fp, 0) : -1;
    const major = fp.length >= 4 ? readUInt16LE(fp, 2) : -1;
    const minor = fp.length >= 6 ? readUInt16LE(fp, 4) : -1;
    if (type !== 1 || major !== 1 || minor !== 1 || fp.length < 54) {
        const kind = type === 0 ? 'XOR obfuscation' : (type === 1 ? 'RC4 CryptoAPI' : 'unknown');
        throw new Error(`Workbook is encrypted (${kind}); decryption is not supported`);
    }
    const salt = fp.slice(6, 22);
    const base = baseKey(password, salt);
    if (!verify(base, fp.slice(22, 38), fp.slice(38, 54))) {
        throw new Error('Workbook is encrypted with a password');
    }

    // keystream position = stream offset, rekeyed every 1024 bytes; record headers and the plain records keep
    // their bytes but still advance the keystream
    const out = new Uint8Array(data);
    let ksBlock = -1;
    let ks = null;
    const xorRange = (from, to) => {
        for (let p = from; p < to; p++) {
            const blk = Math.floor(p / BLOCK);
            if (blk !== ksBlock) {
                ksBlock = blk;
                ks = rc4Stream(blockKey(base, blk), BLOCK);
            }
            out[p] ^= ks[p - blk * BLOCK];
        }
    };
    for (let off = 0; off + 4 <= data.length;) {
        const type = readUInt16LE(data, off);
        const len = readUInt16LE(data, off + 2);
        const start = off + 4;
        const end = Math.min(start + len, data.length);
        if (type === FILEPASS) {
            out[off] = 0;           // record type 0: skipped by the parser
            out[off + 1] = 0;
        } else if (type === BOUNDSHEET) {
            xorRange(start + 4, end); // lbPlyPos stays plain
        } else if (!PLAIN_RECORDS.has(type)) {
            xorRange(start, end);
        }
        off = start + len;
    }
    return out;
}
