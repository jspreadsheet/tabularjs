/**
 * Universal File Loader - Minimal overhead
 * Works in both Browser and Node.js environments
 */

let fs = null; // Cache fs module for Node.js
let iconv = null; // Cache iconv-lite module for extended encoding support

/**
 * Load file as buffer (Uint8Array) - Direct, no intermediate steps
 */
export async function loadAsBuffer(input) {
    // Already Uint8Array - return directly
    if (input instanceof Uint8Array) return input;

    // Node.js Buffer - single conversion
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
        return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }

    // File path (Node.js)
    if (typeof input === 'string') {
        if (!fs) fs = await import('fs');
        return new Uint8Array(fs.readFileSync(input));
    }

    // File/Blob (Browser) - direct read
    if (typeof File !== 'undefined' && input instanceof File ||
        typeof Blob !== 'undefined' && input instanceof Blob) {
        return new Uint8Array(await input.arrayBuffer());
    }

    throw new Error('Invalid input: expected file path, File, Blob, Uint8Array, or Buffer');
}

/**
 * Pick the encoding announced by a byte order mark, if any
 * @param {Uint8Array} buffer
 * @returns {string|null} 'utf-8', 'utf-16le', 'utf-16be' or null
 */
export function detectBOM(buffer) {
    if (buffer.length >= 3 && buffer[0] === 0xEF && buffer[1] === 0xBB && buffer[2] === 0xBF) return 'utf-8';
    if (buffer.length >= 2 && buffer[0] === 0xFF && buffer[1] === 0xFE) return 'utf-16le';
    if (buffer.length >= 2 && buffer[0] === 0xFE && buffer[1] === 0xFF) return 'utf-16be';
    return null;
}

/**
 * Load file as string
 *
 * Every input kind (path, File/Blob, bytes) goes through the same decoder, so
 * the result never depends on how the file was handed over. A byte order mark
 * wins over the requested encoding (as in browsers) and is never part of the
 * returned text.
 */
export async function loadAsString(input, encoding = 'utf-8') {
    const buffer = await loadAsBuffer(input);
    const enc = detectBOM(buffer) || (encoding || 'utf-8').toLowerCase();

    // TextDecoder covers all WHATWG Encoding Standard labels (windows-125x,
    // iso-8859-x, koi8, shift_jis, gbk, big5, euc-jp/kr, ...) in both
    // Node.js and browsers. It strips a BOM matching the encoding.
    let decoder = null;
    try {
        decoder = new TextDecoder(enc);
    } catch (e) {
        // Unknown label - fall back to iconv-lite below (cp437, cp850, ...)
    }
    if (decoder) {
        return stripBOM(decoder.decode(buffer));
    }

    if (!iconv) {
        try {
            const iconvModule = await import('iconv-lite');
            iconv = iconvModule.default || iconvModule;
        } catch (e) {
            throw new Error(`Encoding '${encoding}' requires iconv-lite package. Install with: npm install iconv-lite`);
        }
    }
    return stripBOM(iconv.decode(Buffer.from(buffer), encoding));
}

function stripBOM(text) {
    return text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
}

/**
 * Detect file encoding - Node.js only (uses chardet)
 * @param {string|Uint8Array|Buffer} input - File path or buffer
 * @returns {Promise<string>} Detected encoding (e.g., 'UTF-8', 'windows-1252', 'ISO-8859-1')
 */
export async function detectEncoding(input) {
    try {
        // Import chardet dynamically (Node.js only)
        const chardet = await import('chardet');

        // Get buffer for detection
        let buffer;
        if (typeof input === 'string') {
            // File path - read file
            if (!fs) fs = await import('fs');
            buffer = fs.readFileSync(input);
        } else if (input instanceof Uint8Array) {
            buffer = Buffer.from(input);
        } else if (typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
            buffer = input;
        } else {
            throw new Error('detectEncoding: expected file path, Uint8Array, or Buffer');
        }

        // Detect encoding
        const detected = chardet.default.detect(buffer);

        // Map chardet names to TextDecoder labels. Every label here is a
        // real decoder for that encoding - do NOT alias unrelated encodings
        // (the old map decoded Cyrillic and CJK files as latin1/utf-8).
        const encodingMap = {
            'UTF-8': 'utf-8',
            'UTF-16LE': 'utf-16le',
            'UTF-16BE': 'utf-16be',
            'ISO-8859-1': 'latin1',
            'windows-1252': 'windows-1252',
            'windows-1251': 'windows-1251',
            'GB2312': 'gb2312',
            'GB18030': 'gb18030',
            'Big5': 'big5',
            'EUC-JP': 'euc-jp',
            'EUC-KR': 'euc-kr',
            'Shift_JIS': 'shift_jis'
        };

        return encodingMap[detected] || detected?.toLowerCase() || 'utf-8';
    } catch (error) {
        // If chardet fails or not available (browser), fallback to utf-8
        console.warn('Encoding detection failed, using utf-8:', error.message);
        return 'utf-8';
    }
}

/**
 * Simple wrapper for onload callback - minimal overhead
 */
export async function parse(parserFn, input, options = {}) {
    try {
        const result = await parserFn(input, options);
        if (options.onload) options.onload(result);
        return result;
    } catch (error) {
        if (options.onerror) options.onerror(error);
        throw error;
    }
}
