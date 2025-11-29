import { loadAsString, parse } from '../utils/loader.js';
import { getColumnName } from '../utils/helpers.js';

function parseCSV(str, delimiter) {
    // user-supplied delimiter or default comma
    delimiter = (delimiter || ",");

    // Final data
    let col = 0;
    let row = 0;
    let data = [[]];
    let limit = 0;
    let flag = null;
    let inside = false;
    let closed = false;

    // Go over all chars
    for (let i = 0; i < str.length; i++) {
        let r = data[row];
        // Create a new row
        if (! r) {
            r = data[row] = [];
        }
        // Create a new column
        if (! r[col]) {
            r[col] = '';
        }

        // Ignore
        let char = str[i];
        if (char === '\r') {
            continue;
        }

        // New row
        if ((char === '\n' || char === delimiter) && (inside === false || closed === true || ! flag)) {
            // Restart flags
            flag = null;
            inside = false;
            closed = false;

            if (r[col][0] === '"') {
                let val = r[col].trim();
                if (val[val.length-1] === '"') {
                    r[col] = val.substr(1, val.length-2);
                }
            }

            // Go to the next cell
            if (char === '\n') {
                // New line
                col = 0;
                row++;
            } else {
                // New column
                col++;
                if (col > limit) {
                    // Keep the reference of max column
                    limit = col;
                }
            }
        } else {
            // Inside quotes
            if (char === '"') {
                inside = ! inside;
            }

            if (flag === null) {
                flag = inside;
                if (flag === true) {
                    continue;
                }
            } else if (flag === true && ! closed) {
                if (char === '"') {
                    if (str[i+1] === '"') {
                        inside = true;
                        r[col] += char;
                        i++;
                    } else {
                        closed = true;
                    }
                    continue;
                }
            }

            r[col] += char;
        }
    }

    // Make sure a square matrix is generated
    for (let j = 0; j < data.length; j++) {
        for (let i = 0; i <= limit; i++) {
            if (typeof(data[j][i]) === 'undefined') {
                data[j][i] = '';
            }
        }
    }

    return data;
}

/**
 * Parse CSV from string content
 * @param {string} content - CSV string content
 * @param {string} delimiter - Column delimiter (default: ',')
 * @returns {object} Jspreadsheet Pro format
 */
export function parseCSVString(content, delimiter = ',') {
    // Parse the CSV content
    const data = parseCSV(content, delimiter);

    // Determine the number of columns
    const numColumns = data.length > 0 ? data[0].length : 0;

    // Create columns array with titles
    const columns = [];
    for (let i = 0; i < numColumns; i++) {
        columns.push({
            title: getColumnName(i)
        });
    }

    // Return in the specified format
    return {
        worksheets: [
            {
                data: data,
                columns: columns
            }
        ]
    };
}

/**
 * Parse CSV file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {string} options.delimiter - Column delimiter (default: ',')
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseCSVFile(input, options = {}) {
    return parse(async (inp, opts) => {
        // Handle delimiter as second parameter for backwards compatibility
        const delimiter = typeof options === 'string' ? options : (opts.delimiter || ',');

        // Load and parse directly
        const content = await loadAsString(inp);
        return parseCSVString(content, delimiter);
    }, input, typeof options === 'string' ? {} : options);
}
