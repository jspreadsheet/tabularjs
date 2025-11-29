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
 * Load file as string - Direct, no intermediate steps
 */
export async function loadAsString(input, encoding = 'utf-8') {
    // Encodings that require iconv-lite (not supported by Node.js or TextDecoder)
    const needsIconv = ['cp850', 'cp437', 'cp866', 'windows-1250', 'windows-1251',
                        'windows-1252', 'windows-1253', 'windows-1254', 'windows-1255',
                        'windows-1256', 'windows-1257', 'windows-1258', 'iso-8859-2',
                        'iso-8859-3', 'iso-8859-4', 'iso-8859-5', 'iso-8859-6',
                        'iso-8859-7', 'iso-8859-8', 'iso-8859-9', 'iso-8859-10',
                        'iso-8859-13', 'iso-8859-14', 'iso-8859-15', 'iso-8859-16',
                        'koi8-r', 'koi8-u', 'macintosh'].includes(encoding.toLowerCase());

    // If encoding requires iconv-lite
    if (needsIconv) {
        if (!iconv) {
            try {
                const iconvModule = await import('iconv-lite');
                iconv = iconvModule.default || iconvModule;
            } catch (e) {
                throw new Error(`Encoding '${encoding}' requires iconv-lite package. Install with: npm install iconv-lite`);
            }
        }

        // Load as buffer first
        const buffer = await loadAsBuffer(input);
        return iconv.decode(Buffer.from(buffer), encoding);
    }

    // Standard encodings supported by Node.js and browsers
    // File path (Node.js)
    if (typeof input === 'string') {
        if (!fs) fs = await import('fs');
        return fs.readFileSync(input, encoding);
    }

    // File/Blob (Browser)
    if (typeof File !== 'undefined' && input instanceof File ||
        typeof Blob !== 'undefined' && input instanceof Blob) {
        if (encoding === 'utf-8') {
            return input.text();
        } else {
            // For non-UTF-8 encodings in browser, read as buffer and decode
            const buffer = await input.arrayBuffer();
            return new TextDecoder(encoding).decode(buffer);
        }
    }

    // Buffer/Uint8Array
    if (input instanceof Uint8Array || typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
        return new TextDecoder(encoding).decode(input);
    }

    throw new Error('Invalid input: expected file path, File, Blob, Uint8Array, or Buffer');
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

        // Map common chardet names to Node.js encoding names
        const encodingMap = {
            'UTF-8': 'utf-8',
            'UTF-16LE': 'utf-16le',
            'UTF-16BE': 'utf-16be',
            'ISO-8859-1': 'latin1',
            'windows-1252': 'latin1', // Node.js treats windows-1252 as latin1
            'windows-1251': 'latin1',
            'GB2312': 'utf-8', // Fallback to utf-8 for Chinese
            'Big5': 'utf-8',   // Fallback to utf-8 for Traditional Chinese
            'EUC-JP': 'utf-8', // Fallback to utf-8 for Japanese
            'EUC-KR': 'utf-8', // Fallback to utf-8 for Korean
            'Shift_JIS': 'utf-8'
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
