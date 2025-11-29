import fs from 'fs';
import path from 'path';
import parser from './parser.js'

// Get file path from command line arguments
const filePath = process.argv[2];

if (!filePath) {
    console.log('Usage: node src/script.js <path-to-file>');
    process.exit(1);
}

if (!fs.existsSync(filePath)) {
    console.error(`Error: File not found: ${filePath}`);
    process.exit(1);
}

try {
    // Just pass the file path directly - the loader will handle it
    const result = await parser(filePath);
    console.log(JSON.stringify(result, null, 2));
} catch (error) {
    console.error('Error:', error.message);
    console.error(error.stack);
    process.exit(1);
}
