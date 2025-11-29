/**
 * Custom Compound File Binary (CFB) / OLE2 Parser
 *
 * Lightweight MIT-licensed implementation for reading Microsoft Compound File Binary format.
 * Based on [MS-CFB] Microsoft Compound File Binary File Format specification.
 *
 * Only implements READ functionality - no write support needed.
 * Focused on extracting streams from XLS files.
 *
 * @license MIT
 */

import { readUInt16LE, readUInt32LE } from './helpers.js';

// Constants from [MS-CFB] specification
const HEADER_SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const HEADER_CLSID_NULL = new Array(16).fill(0);
const DIFSECT = 0xFFFFFFFC;
const FATSECT = 0xFFFFFFFD;
const ENDOFCHAIN = 0xFFFFFFFE;
const FREESECT = 0xFFFFFFFF;
const MAXREGSECT = 0xFFFFFFFA;

// Directory entry types
const ENTRY_TYPES = {
    UNKNOWN: 0,
    STORAGE: 1,
    STREAM: 2,
    ROOT: 5
};

/**
 * Read a 64-bit integer as two 32-bit parts (we only care about low part for offsets)
 */
function readUInt64LE(buffer, offset) {
    const low = readUInt32LE(buffer, offset);
    const high = readUInt32LE(buffer, offset + 4);
    // For file sizes, we typically only need the low part
    // JavaScript can't safely represent full 64-bit integers
    return high === 0 ? low : low + (high * 0x100000000);
}

/**
 * Parse CFB Header (512 bytes)
 * [MS-CFB] 2.2 Compound File Header
 */
function parseHeader(buffer) {
    // Verify signature
    for (let i = 0; i < 8; i++) {
        if (buffer[i] !== HEADER_SIGNATURE[i]) {
            throw new Error('Invalid CFB signature');
        }
    }

    // Skip CLSID (16 bytes at offset 8)
    const minorVersion = readUInt16LE(buffer, 0x18);
    const majorVersion = readUInt16LE(buffer, 0x1A);
    const byteOrder = readUInt16LE(buffer, 0x1C); // Should be 0xFFFE (little-endian)

    if (byteOrder !== 0xFFFE) {
        throw new Error('Invalid byte order marker');
    }

    const sectorShift = readUInt16LE(buffer, 0x1E);
    const miniSectorShift = readUInt16LE(buffer, 0x20);

    // Calculate sector sizes
    const sectorSize = 1 << sectorShift; // Usually 512 or 4096
    const miniSectorSize = 1 << miniSectorShift; // Usually 64

    // Skip reserved (6 bytes at 0x22)
    const totalSectors = readUInt32LE(buffer, 0x28); // Only used in v4
    const fatSectors = readUInt32LE(buffer, 0x2C);
    const firstDirSector = readUInt32LE(buffer, 0x30);
    const transactionSignature = readUInt32LE(buffer, 0x34);
    const miniStreamCutoff = readUInt32LE(buffer, 0x38); // Usually 4096
    const firstMiniFatSector = readUInt32LE(buffer, 0x3C);
    const miniFatSectors = readUInt32LE(buffer, 0x40);
    const firstDifatSector = readUInt32LE(buffer, 0x44);
    const difatSectors = readUInt32LE(buffer, 0x48);

    // Read first 109 DIFAT entries (436 bytes starting at 0x4C)
    const difat = [];
    for (let i = 0; i < 109; i++) {
        const sector = readUInt32LE(buffer, 0x4C + i * 4);
        if (sector < MAXREGSECT) {
            difat.push(sector);
        }
    }

    return {
        minorVersion,
        majorVersion,
        sectorShift,
        miniSectorShift,
        sectorSize,
        miniSectorSize,
        totalSectors,
        fatSectors,
        firstDirSector,
        transactionSignature,
        miniStreamCutoff,
        firstMiniFatSector,
        miniFatSectors,
        firstDifatSector,
        difatSectors,
        difat
    };
}

/**
 * Read a sector from the file
 */
function readSector(buffer, sectorId, sectorSize) {
    const offset = (sectorId + 1) * sectorSize;
    if (offset + sectorSize > buffer.length) {
        throw new Error(`Sector ${sectorId} out of bounds`);
    }
    return buffer.slice(offset, offset + sectorSize);
}

/**
 * Build complete FAT (File Allocation Table)
 * [MS-CFB] 2.3 Compound File FAT Sectors
 */
function buildFAT(buffer, header) {
    const { sectorSize, difat, firstDifatSector, difatSectors } = header;
    const entriesPerSector = sectorSize / 4;
    const fat = [];

    // Read FAT sectors listed in DIFAT
    let fatSectors = [...difat];

    // If there are additional DIFAT sectors, read them
    let difatSector = firstDifatSector;
    for (let i = 0; i < difatSectors && difatSector < MAXREGSECT; i++) {
        const sector = readSector(buffer, difatSector, sectorSize);

        // Each DIFAT sector contains FAT sector numbers
        for (let j = 0; j < entriesPerSector - 1; j++) {
            const fatSectorNum = readUInt32LE(sector, j * 4);
            if (fatSectorNum < MAXREGSECT) {
                fatSectors.push(fatSectorNum);
            }
        }

        // Last entry points to next DIFAT sector
        difatSector = readUInt32LE(sector, (entriesPerSector - 1) * 4);
    }

    // Read all FAT sectors
    for (const fatSectorNum of fatSectors) {
        const sector = readSector(buffer, fatSectorNum, sectorSize);

        for (let i = 0; i < entriesPerSector; i++) {
            fat.push(readUInt32LE(sector, i * 4));
        }
    }

    return fat;
}

/**
 * Build Mini FAT for small streams
 * [MS-CFB] 2.5 Compound File MiniFAT Sectors
 */
function buildMiniFAT(buffer, header, fat) {
    const { firstMiniFatSector, miniFatSectors, sectorSize } = header;

    if (firstMiniFatSector >= MAXREGSECT || miniFatSectors === 0) {
        return [];
    }

    const entriesPerSector = sectorSize / 4;
    const miniFat = [];

    // Follow the chain of MiniFAT sectors
    let sector = firstMiniFatSector;
    let count = 0;

    while (sector < MAXREGSECT && count < miniFatSectors) {
        const sectorData = readSector(buffer, sector, sectorSize);

        for (let i = 0; i < entriesPerSector; i++) {
            miniFat.push(readUInt32LE(sectorData, i * 4));
        }

        sector = fat[sector];
        count++;
    }

    return miniFat;
}

/**
 * Read a chain of sectors
 */
function readChain(buffer, startSector, fat, sectorSize, maxSize = Infinity) {
    if (startSector >= MAXREGSECT) {
        return new Uint8Array(0);
    }

    const chunks = [];
    let sector = startSector;
    let totalSize = 0;

    while (sector < MAXREGSECT && totalSize < maxSize) {
        const sectorData = readSector(buffer, sector, sectorSize);
        chunks.push(sectorData);
        totalSize += sectorSize;
        sector = fat[sector];

        // Prevent infinite loops
        if (chunks.length > 100000) {
            throw new Error('FAT chain too long - possible corruption');
        }
    }

    // Combine chunks
    const result = new Uint8Array(Math.min(totalSize, maxSize));
    let offset = 0;
    for (const chunk of chunks) {
        const copySize = Math.min(chunk.length, maxSize - offset);
        result.set(chunk.slice(0, copySize), offset);
        offset += copySize;
        if (offset >= maxSize) break;
    }

    return result;
}

/**
 * Read mini stream (for streams smaller than cutoff size)
 */
function readMiniChain(miniStream, startSector, miniFat, miniSectorSize, size) {
    if (startSector >= MAXREGSECT) {
        return new Uint8Array(0);
    }

    const result = new Uint8Array(size);
    let sector = startSector;
    let offset = 0;

    while (sector < MAXREGSECT && offset < size) {
        const sectorOffset = sector * miniSectorSize;
        const copySize = Math.min(miniSectorSize, size - offset, miniStream.length - sectorOffset);

        if (sectorOffset + copySize > miniStream.length) {
            break;
        }

        result.set(miniStream.slice(sectorOffset, sectorOffset + copySize), offset);
        offset += copySize;
        sector = miniFat[sector];

        // Prevent infinite loops
        if (offset > size * 2) {
            throw new Error('MiniFAT chain too long - possible corruption');
        }
    }

    return result.slice(0, offset);
}

/**
 * Parse a directory entry (128 bytes)
 * [MS-CFB] 2.6 Compound File Directory Entry
 */
function parseDirectoryEntry(buffer, offset) {
    // Name (64 bytes, UTF-16LE)
    let nameLength = readUInt16LE(buffer, offset + 64);
    if (nameLength > 64) nameLength = 64;

    let name = '';
    for (let i = 0; i < nameLength - 2; i += 2) {
        const charCode = readUInt16LE(buffer, offset + i);
        if (charCode === 0) break;
        name += String.fromCharCode(charCode);
    }

    const type = buffer[offset + 66];
    const color = buffer[offset + 67]; // Red-Black tree color
    const leftSibling = readUInt32LE(buffer, offset + 68);
    const rightSibling = readUInt32LE(buffer, offset + 72);
    const childId = readUInt32LE(buffer, offset + 76);

    // CLSID (16 bytes at offset 80) - skip
    const stateBits = readUInt32LE(buffer, offset + 96);

    // Creation/modification times (8 bytes each) - skip for now

    const startSector = readUInt32LE(buffer, offset + 116);
    const sizeLow = readUInt32LE(buffer, offset + 120);
    const sizeHigh = readUInt32LE(buffer, offset + 124);
    const size = sizeHigh === 0 ? sizeLow : sizeLow + (sizeHigh * 0x100000000);

    return {
        name,
        type,
        color,
        leftSibling,
        rightSibling,
        childId,
        stateBits,
        startSector,
        size
    };
}

/**
 * Parse all directory entries
 * [MS-CFB] 2.6 Compound File Directory Sectors
 */
function parseDirectory(buffer, header, fat) {
    const { firstDirSector, sectorSize } = header;
    const entriesPerSector = sectorSize / 128;
    const entries = [];

    // Read directory sectors following the FAT chain
    const dirData = readChain(buffer, firstDirSector, fat, sectorSize);

    // Parse each 128-byte entry
    for (let i = 0; i < dirData.length; i += 128) {
        if (i + 128 > dirData.length) break;
        const entry = parseDirectoryEntry(dirData, i);
        entries.push(entry);
    }

    return entries;
}

/**
 * Build full paths for all directory entries
 */
function buildPaths(entries) {
    const paths = new Array(entries.length).fill('');

    if (entries.length === 0) return paths;

    // Root entry
    paths[0] = '/';

    // Build paths recursively
    function buildPath(index, parentPath) {
        if (index >= entries.length || index < 0) return;

        const entry = entries[index];
        if (!entry.name) return;

        const currentPath = parentPath === '/'
            ? '/' + entry.name
            : parentPath + '/' + entry.name;

        paths[index] = currentPath;

        // Process children
        if (entry.childId < entries.length && entry.childId !== index) {
            buildPath(entry.childId, currentPath);
        }

        // Process siblings
        if (entry.leftSibling < entries.length && entry.leftSibling !== index) {
            buildPath(entry.leftSibling, parentPath);
        }
        if (entry.rightSibling < entries.length && entry.rightSibling !== index) {
            buildPath(entry.rightSibling, parentPath);
        }
    }

    // Start from root's children
    if (entries[0].childId < entries.length) {
        buildPath(entries[0].childId, '/');
    }

    return paths;
}

/**
 * Read CFB file and return container object
 */
export function read(buffer, options = {}) {
    // Convert to Uint8Array if needed
    if (!(buffer instanceof Uint8Array)) {
        if (Buffer.isBuffer(buffer)) {
            buffer = new Uint8Array(buffer);
        } else {
            throw new Error('Buffer must be Uint8Array or Node Buffer');
        }
    }

    if (buffer.length < 512) {
        throw new Error(`File too small: ${buffer.length} bytes`);
    }

    // Parse header
    const header = parseHeader(buffer);

    // Build FAT
    const fat = buildFAT(buffer, header);

    // Build Mini FAT
    const miniFat = buildMiniFAT(buffer, header, fat);

    // Parse directory
    const entries = parseDirectory(buffer, header, fat);

    // Build full paths
    const paths = buildPaths(entries);

    // Read mini stream (from root entry)
    let miniStream = new Uint8Array(0);
    if (entries.length > 0 && entries[0].startSector < MAXREGSECT) {
        miniStream = readChain(
            buffer,
            entries[0].startSector,
            fat,
            header.sectorSize,
            entries[0].size
        );
    }

    // Create CFB container object
    const cfb = {
        header,
        fat,
        miniFat,
        entries,
        paths,
        miniStream,
        buffer
    };

    return cfb;
}

/**
 * Find an entry by path
 * [MS-CFB] 2.6.4 (case-insensitive comparison)
 */
export function find(cfb, path) {
    if (!path) return null;

    // Normalize path
    let searchPath = path.trim();
    if (!searchPath.startsWith('/')) {
        searchPath = '/' + searchPath;
    }
    searchPath = searchPath.toUpperCase();

    // Search for matching path (case-insensitive)
    for (let i = 0; i < cfb.paths.length; i++) {
        const entryPath = cfb.paths[i].toUpperCase();

        if (entryPath === searchPath) {
            return createEntryObject(cfb, i);
        }
    }

    // Try alternative formats
    const altPath = searchPath.replace(/^\//, '');
    for (let i = 0; i < cfb.paths.length; i++) {
        const entryPath = cfb.paths[i].toUpperCase().replace(/^\//, '');

        if (entryPath === altPath) {
            return createEntryObject(cfb, i);
        }
    }

    return null;
}

/**
 * Create entry object with content
 */
function createEntryObject(cfb, index) {
    const entry = cfb.entries[index];
    const path = cfb.paths[index];

    if (entry.type !== ENTRY_TYPES.STREAM) {
        // Only return streams with content
        return {
            name: entry.name,
            type: entry.type,
            size: entry.size
        };
    }

    // Read stream content
    let content;
    if (entry.size < cfb.header.miniStreamCutoff) {
        // Small stream - use mini FAT
        content = readMiniChain(
            cfb.miniStream,
            entry.startSector,
            cfb.miniFat,
            cfb.header.miniSectorSize,
            entry.size
        );
    } else {
        // Large stream - use regular FAT
        content = readChain(
            cfb.buffer,
            entry.startSector,
            cfb.fat,
            cfb.header.sectorSize,
            entry.size
        );
    }

    return {
        name: entry.name,
        type: entry.type,
        size: entry.size,
        content: content
    };
}

/**
 * List all entries in the container
 */
export function listEntries(cfb) {
    return cfb.entries.map((entry, index) => ({
        name: entry.name,
        path: cfb.paths[index],
        type: entry.type,
        size: entry.size
    })).filter(e => e.name && e.type !== ENTRY_TYPES.UNKNOWN);
}

export default { read, find, listEntries };
