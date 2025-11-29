# TabularJS

[![npm version](https://badge.fury.io/js/tabularjs.svg)](https://www.npmjs.com/package/tabularjs)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tests](https://img.shields.io/badge/tests-312%20passing-brightgreen.svg)](https://github.com/jspreadsheet/tabularjs)

A professional JavaScript library for converting spreadsheet files to JSON format. Supports 16+ file formats for use in Node.js and browser applications.

## Features

- **16+ File Formats** - Excel, OpenDocument, CSV, Lotus, SYLK, DIF, DBF, HTML tables, and more
- **Universal JSON Output** - Structured data format for any JavaScript application
- **Formula Preservation** - Maintains formulas where supported by the format
- **Styling & Formatting** - Preserves merged cells, styles, and comments
- **Zero Dependencies** - No xlsx, no SheetJS - pure JavaScript implementation
- **Framework Agnostic** - Works with Vanilla JS, React, Vue, Angular, and Node.js

## Supported Formats

| Format               | Extension                              | Features                                        |
|----------------------|----------------------------------------|-------------------------------------------------|
| Excel 97-2003        | `.xls`                                 | ✅ Formulas, ✅ Styles, ✅ Merged cells            |
| Excel 2007+          | `.xlsx`                                | ✅ Formulas, ✅ Styles, ✅ Merged cells            |
| OpenDocument         | `.ods`                                 | ✅ Formulas, ✅ Styles, ✅ Merged cells            |
| Lotus 1-2-3          | `.wks`, `.wk1`, `.wk3`, `.wk4`, `.123` | ✅ Data                                          |
| CSV                  | `.csv`                                 | ✅ Data                                          |
| TSV                  | `.tsv`, `.tab`                         | ✅ Data                                          |
| Plain Text           | `.txt`                                 | ✅ Data (tab-delimited)                          |
| XML Spreadsheet 2003 | `.xml`                                 | ✅ Formulas, ✅ Merged cells, ✅ Comments          |
| DIF                  | `.dif`                                 | ✅ Data, ✅ Labels, ✅ Comments                    |
| SYLK                 | `.slk`, `.sylk`                        | ✅ Formulas, ✅ Data                              |
| dBase                | `.dbf`                                 | ✅ Data, ✅ Column types, ✅ Metadata              |
| HTML Tables          | `.html`, `.htm`                        | ✅ Formulas, ✅ Merged cells, ✅ Styles            |

## Installation

```bash
npm install tabularjs
```

## Quick Start

### Browser

```javascript
import tabularjs from 'tabularjs';

// Parse a file
const result = await tabularjs(fileObject);
console.log(result.worksheets[0].data);
```

### Node.js

```javascript
import tabularjs from 'tabularjs';

const result = await tabularjs('path/to/file.xlsx');
console.log(result.worksheets[0].data);
```

**Command Line:**

```bash
node src/script.js samples/test1.xml
```

**Production Script Example:**

```javascript
import fs from 'fs';
import tabularjs from 'tabularjs';

const filePath = process.argv[2];

if (!filePath) {
    console.log('Usage: node script.js <path-to-file>');
    process.exit(1);
}

if (!fs.existsSync(filePath)) {
    console.error(`Error: File not found: ${filePath}`);
    process.exit(1);
}

try {
    const result = await tabularjs(filePath);
    console.log(JSON.stringify(result, null, 2));
} catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
}
```

## Example: Integration with Jspreadsheet

TabularJS is an independent file conversion tool. While it can be used with any application, here are examples showing integration with [Jspreadsheet CE](https://jspreadsheet.com) and [Jspreadsheet Pro](https://jspreadsheet.com) for creating online spreadsheets:

### Vanilla JavaScript

```html
<html>
<head>
    <script src="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/index.min.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/jspreadsheet.min.css" />
    <script src="https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js"></script>
</head>
<body>
    <div id="spreadsheet"></div>
    <input type="file" id="fileInput" />

    <script>
        document.getElementById('fileInput').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            const result = await tabularjs(file);

            // TabularJS output is directly compatible with Jspreadsheet
            jspreadsheet(document.getElementById('spreadsheet'), result);
        });
    </script>
</body>
</html>
```

### React

```jsx
import { useRef } from 'react';
import jspreadsheet from 'jspreadsheet-pro';
import tabularjs from 'tabularjs';

export default function App() {
    const spreadsheet = useRef(null);
    const inputRef = useRef(null);

    const load = async (e) => {
        const file = e.target.files[0];
        const result = await tabularjs(file);

        // Create spreadsheet from imported data
        jspreadsheet(spreadsheet.current, result);
    };

    return (
        <>
            <div ref={spreadsheet}></div>
            <input
                ref={inputRef}
                type="file"
                onChange={load}
                style={{ display: 'none' }}
            />
            <input
                type="button"
                value="Load file"
                onClick={() => inputRef.current.click()}
            />
        </>
    );
}
```

### Vue

```vue
<template>
    <div>
        <div ref="spreadsheet"></div>
        <input
            ref="inputRef"
            type="file"
            @change="load"
            style="display: none"
        />
        <input
            type="button"
            value="Load file"
            @click="$refs.inputRef.click()"
        />
    </div>
</template>

<script>
import jspreadsheet from 'jspreadsheet-pro';
import tabularjs from 'tabularjs';

export default {
    methods: {
        async load(e) {
            const file = e.target.files[0];
            const result = await tabularjs(file);

            // Create spreadsheet from imported data
            jspreadsheet(this.$refs.spreadsheet, result);
        }
    }
}
</script>
```

## Formula Support

| Format | Formula Support | Notes |
|--------|----------------|-------|
| XLS | ✅ Full | Decoded from BIFF PTG tokens |
| XLSX | ✅ Full | Native formula strings |
| ODS | ✅ Full | Native formula strings |
| XML 2003 | ✅ Full | Formula attribute |
| SYLK | ✅ Full | E parameter |
| HTML | ✅ Partial | Via `data-formula` attribute |
| CSV/TSV/TXT | ❌ No | Data-only format |
| DIF | ❌ No | Data-only format (stores calculated values) |
| DBF | ❌ No | Database format (stores calculated values) |
| Lotus | ❌ Limited | Complex token format |

## Scope & Limitations

TabularJS is designed as a lightweight parser focused on data extraction and essential spreadsheet features. The following advanced features are intentionally not supported:

**Not Supported:**
- Charts and graphs
- Pivot tables
- Macros and VBA code
- Images and embedded objects
- Print settings and page breaks

**Limited Support:**
- Advanced conditional formatting
- Complex data validations

TabularJS prioritizes data, formulas, basic styling, and structural elements for web-based applications.

## Testing

TabularJS includes comprehensive test coverage with 312+ test cases covering all supported formats.

**Run Tests:**

```bash
npm test
```

**Test Coverage:**
- CSV/TSV parsing with edge cases
- Excel (.xls, .xlsx) format validation
- OpenDocument (.ods) structure
- Legacy formats (Lotus, SYLK, DIF, DBF)
- HTML table parsing
- Formula preservation
- Style and formatting extraction

## Related Tools

TabularJS is part of a suite of JavaScript tools for building modern web applications:

- **[Jspreadsheet](https://jspreadsheet.com)** - Lightweight JavaScript spreadsheet component
- **[LemonadeJS](https://lemonadejs.com)** - Micro reactive JavaScript library
- **[CalendarJS](https://calendarjs.com)** - JavaScript calendar and date picker

## Technical Support

For technical support and questions, please contact: [support@jspreadsheet.com](mailto:support@jspreadsheet.com)

## Links

- **Website:** [tabularjs.com](https://tabularjs.com)
- **GitHub:** [github.com/jspreadsheet/tabularjs](https://github.com/jspreadsheet/tabularjs)
- **NPM:** `npm install tabularjs`

## License

MIT
