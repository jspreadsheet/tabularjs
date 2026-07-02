/**
 * Safe JSZip entry readers.
 *
 * Zip-based formats (xlsx, ods, numbers) come from untrusted sources, and a
 * zip bomb can expand a few hundred KB into many GB. Reading entries through
 * a streaming accumulator lets us abort as soon as the *actual* uncompressed
 * output passes a cap - unlike checking the entry's declared size, which an
 * attacker can understate.
 */

// Per-entry decompression cap. Generous: real spreadsheet parts are far
// smaller, while a bomb blows past this within a few chunks.
export const DEFAULT_MAX_ENTRY_BYTES = 256 * 1024 * 1024;

function concatChunks(chunks, total) {
    const result = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.length;
    }
    return result;
}

function toBase64(bytes) {
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(bytes).toString('base64');
    }
    let binary = '';
    for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

/**
 * Read a JSZip entry with a decompression limit.
 *
 * @param {Object} entry - JSZip file entry (from zip.file(...))
 * @param {string} type - 'string' | 'uint8array' | 'nodebuffer' | 'base64'
 * @param {number} [maxBytes=DEFAULT_MAX_ENTRY_BYTES] - Uncompressed byte cap
 * @returns {Promise<string|Uint8Array|Buffer>}
 */
export function readZipEntry(entry, type, maxBytes = DEFAULT_MAX_ENTRY_BYTES) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let total = 0;
        let settled = false;

        const stream = entry.internalStream('uint8array');

        stream.on('data', (chunk) => {
            if (settled) {
                return;
            }
            total += chunk.length;
            if (total > maxBytes) {
                settled = true;
                stream.pause();
                reject(new Error(
                    `Zip entry '${entry.name}' exceeds the decompression limit of ${maxBytes} bytes (possible zip bomb)`
                ));
                return;
            }
            chunks.push(chunk);
        });

        stream.on('error', (error) => {
            if (!settled) {
                settled = true;
                reject(error);
            }
        });

        stream.on('end', () => {
            if (settled) {
                return;
            }
            settled = true;
            const bytes = concatChunks(chunks, total);
            switch (type) {
                case 'string':
                    resolve(new TextDecoder('utf-8').decode(bytes));
                    break;
                case 'base64':
                    resolve(toBase64(bytes));
                    break;
                case 'nodebuffer':
                    resolve(typeof Buffer !== 'undefined' ? Buffer.from(bytes) : bytes);
                    break;
                default:
                    resolve(bytes);
            }
        });

        stream.resume();
    });
}
