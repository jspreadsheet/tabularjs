export interface Example {
  slug: string;
  title: string;
  category: "basics" | "framework" | "integration" | "advanced";
  tagline: string;
  description: string;
  seoTitle: string;
  seoDescription: string;
  code: string;
  language: string;
  followUp?: { text: string; href: string; linkLabel: string };
}

export const examples: Example[] = [
  {
    slug: "nodejs-basic",
    title: "Parse a file in Node.js",
    category: "basics",
    tagline: "One line to turn any spreadsheet into JSON.",
    description:
      "Pass a file path to tabularjs and get a structured JSON document with every worksheet, its name, and its rows.",
    seoTitle: "Node.js: Convert XLSX/CSV/ODS to JSON — TabularJS",
    seoDescription:
      "Minimal Node.js example: parse any supported spreadsheet file to JSON with TabularJS in one call.",
    language: "js",
    code: `import tabularjs from 'tabularjs';

const result = await tabularjs('./data/sales.xlsx');

for (const sheet of result.worksheets) {
  console.log(sheet.name, '->', sheet.data.length, 'rows');
}`,
  },
  {
    slug: "browser-file-input",
    title: "Browser: parse from a file input",
    category: "basics",
    tagline: "Accept any spreadsheet via <input type=file>.",
    description:
      "Wire up a file input to TabularJS — the File object is passed straight through, no FileReader required.",
    seoTitle: "Browser file input to JSON — TabularJS",
    seoDescription:
      "Convert a user-uploaded spreadsheet (xlsx, csv, ods…) to JSON directly in the browser with TabularJS.",
    language: "html",
    code: `<input type="file" id="file" />
<pre id="out"></pre>

<script type="module">
  import tabularjs from 'https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js';

  document.getElementById('file').addEventListener('change', async (e) => {
    const result = await tabularjs(e.target.files[0]);
    document.getElementById('out').textContent = JSON.stringify(result, null, 2);
  });
</script>`,
  },
  {
    slug: "drag-and-drop",
    title: "Drag & drop spreadsheet upload",
    category: "basics",
    tagline: "Drop any supported file onto a zone and parse it.",
    description:
      "A drop zone that accepts files dropped from the desktop and converts them to JSON instantly.",
    seoTitle: "Drag and drop spreadsheet parser in JavaScript — TabularJS",
    seoDescription:
      "Build a drag-and-drop spreadsheet to JSON uploader in the browser with TabularJS.",
    language: "html",
    code: `<div id="drop" style="padding:40px;border:2px dashed #22d3ee;border-radius:12px">
  Drop a spreadsheet here
</div>

<script type="module">
  import tabularjs from 'tabularjs';

  const drop = document.getElementById('drop');
  drop.addEventListener('dragover', e => e.preventDefault());
  drop.addEventListener('drop', async (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    const result = await tabularjs(file);
    console.log(result);
  });
</script>`,
  },
  {
    slug: "react",
    title: "React: upload and render",
    category: "framework",
    tagline: "A React hook-based uploader.",
    description:
      "Use TabularJS in React to handle file uploads and display parsed data with useState and a file input ref.",
    seoTitle: "React spreadsheet to JSON converter — TabularJS",
    seoDescription:
      "React example: parse uploaded XLSX/CSV/ODS files to JSON and render them as a table.",
    language: "jsx",
    code: `import { useRef, useState } from 'react';
import tabularjs from 'tabularjs';

export default function Uploader() {
  const inputRef = useRef(null);
  const [data, setData] = useState(null);

  const onChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const result = await tabularjs(file);
    setData(result.worksheets[0].data);
  };

  return (
    <div>
      <input type="file" ref={inputRef} onChange={onChange} />
      {data && (
        <table>
          <tbody>
            {data.map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => <td key={j}>{cell}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}`,
  },
  {
    slug: "vue",
    title: "Vue 3 component",
    category: "framework",
    tagline: "Composition API file uploader.",
    description:
      "Use TabularJS inside a Vue 3 single-file component with reactive state and a file input.",
    seoTitle: "Vue.js spreadsheet to JSON — TabularJS",
    seoDescription:
      "Vue 3 example of parsing uploaded spreadsheet files to JSON with TabularJS.",
    language: "vue",
    code: `<script setup>
import { ref } from 'vue';
import tabularjs from 'tabularjs';

const data = ref(null);

async function onChange(e) {
  const file = e.target.files[0];
  const result = await tabularjs(file);
  data.value = result.worksheets[0].data;
}
</script>

<template>
  <input type="file" @change="onChange" />
  <pre v-if="data">{{ data }}</pre>
</template>`,
  },
  {
    slug: "nextjs-api-route",
    title: "Next.js App Router API route",
    category: "framework",
    tagline: "Server-side conversion via an uploaded file.",
    description:
      "Create a POST route that accepts multipart form data and returns the parsed JSON. Pure server code — the client never needs to load the parser.",
    seoTitle: "Next.js API route for spreadsheet upload — TabularJS",
    seoDescription:
      "Next.js 15 App Router example: server-side file upload and conversion to JSON with TabularJS.",
    language: "ts",
    code: `// app/api/convert/route.ts
import tabularjs from 'tabularjs';

export async function POST(req: Request) {
  const form = await req.formData();
  const file = form.get('file') as File;

  const buffer = Buffer.from(await file.arrayBuffer());
  const result = await tabularjs(buffer, { name: file.name });

  return Response.json(result);
}`,
  },
  {
    slug: "jspreadsheet-ce",
    title: "Load into Jspreadsheet CE",
    category: "integration",
    tagline: "TabularJS output is shaped for Jspreadsheet.",
    description:
      "Upload a file and instantly render it in the open-source Jspreadsheet CE component — no data transformation needed.",
    seoTitle: "Jspreadsheet CE file import — TabularJS",
    seoDescription:
      "Load .xlsx, .csv, .ods files directly into a Jspreadsheet CE spreadsheet component.",
    language: "html",
    code: `<script src="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/index.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/jspreadsheet-ce/dist/jspreadsheet.min.css" />
<script src="https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js"></script>

<div id="sheet"></div>
<input type="file" id="f" />

<script>
  document.getElementById('f').onchange = async (e) => {
    const result = await tabularjs(e.target.files[0]);
    jspreadsheet(document.getElementById('sheet'), result);
  };
</script>`,
    followUp: {
      text: "See more at",
      href: "https://jspreadsheet.com",
      linkLabel: "jspreadsheet.com",
    },
  },
  {
    slug: "jspreadsheet-pro",
    title: "Load into Jspreadsheet Pro",
    category: "integration",
    tagline: "Enterprise-grade spreadsheet with TabularJS import.",
    description:
      "Jspreadsheet Pro adds advanced features like pivot tables, charts, and formula engine. Use TabularJS to import files into a Pro workbook.",
    seoTitle: "Jspreadsheet Pro import integration — TabularJS",
    seoDescription:
      "Import XLSX, CSV, ODS, and more into a Jspreadsheet Pro workbook using TabularJS.",
    language: "html",
    code: `<script src="https://cdn.jsdelivr.net/npm/jspreadsheet-pro/dist/index.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js"></script>

<div id="workbook"></div>

<script>
  async function loadFile(file) {
    const result = await tabularjs(file);
    jspreadsheet(document.getElementById('workbook'), {
      worksheets: result.worksheets,
      toolbar: true,
      tabs: true
    });
  }
</script>`,
    followUp: {
      text: "Learn more about Pro features at",
      href: "https://bossanova.uk/jspreadsheet",
      linkLabel: "bossanova.uk/jspreadsheet",
    },
  },
  {
    slug: "lemonadejs",
    title: "Use with LemonadeJS",
    category: "integration",
    tagline: "Reactive rendering with tiny footprint.",
    description:
      "TabularJS pairs perfectly with LemonadeJS — a ~5kB reactive library from the same team — to build a file preview UI.",
    seoTitle: "LemonadeJS spreadsheet previewer — TabularJS",
    seoDescription:
      "Build a reactive spreadsheet preview with LemonadeJS and TabularJS.",
    language: "html",
    code: `<script src="https://cdn.jsdelivr.net/npm/lemonadejs/dist/lemonade.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js"></script>

<div id="root"></div>

<script>
  function Preview() {
    const self = this;
    self.sheets = [];

    self.onupload = async function(e) {
      const result = await tabularjs(e.target.files[0]);
      self.sheets = result.worksheets;
    };

    return \`<div>
      <input type="file" onchange="self.onupload(event)" />
      <section :loop="self.sheets">
        <h3>{{self.name}}</h3>
        <p>{{self.data.length}} rows</p>
      </section>
    </div>\`;
  }

  lemonade.render(Preview, document.getElementById('root'));
</script>`,
  },
  {
    slug: "express-upload",
    title: "Express.js upload endpoint",
    category: "framework",
    tagline: "Parse files uploaded to a Node.js server.",
    description:
      "An Express route that uses Multer for uploads and TabularJS for conversion — ready for production.",
    seoTitle: "Express.js file upload to JSON — TabularJS",
    seoDescription:
      "Handle spreadsheet uploads in Express.js with Multer and TabularJS, responding with JSON.",
    language: "js",
    code: `import express from 'express';
import multer from 'multer';
import tabularjs from 'tabularjs';

const app = express();
const upload = multer();

app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    const result = await tabularjs(req.file.buffer, { name: req.file.originalname });
    res.json(result);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.listen(3000);`,
  },
  {
    slug: "batch-directory",
    title: "Batch convert a directory",
    category: "advanced",
    tagline: "Walk a folder and emit JSON for every spreadsheet.",
    description:
      "A Node.js script that scans a directory, converts each supported file to JSON, and writes the result alongside.",
    seoTitle: "Batch convert spreadsheets to JSON in Node.js — TabularJS",
    seoDescription:
      "Recursively convert a directory of XLSX, CSV, and ODS files to JSON with a single Node.js script.",
    language: "js",
    code: `import fs from 'node:fs/promises';
import path from 'node:path';
import tabularjs from 'tabularjs';

const SUPPORTED = /\\.(xlsx?|ods|csv|tsv|dbf|xml|slk|dif|wk[1-4s])$/i;

async function walk(dir) {
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) await walk(full);
    else if (SUPPORTED.test(entry.name)) {
      const result = await tabularjs(full);
      await fs.writeFile(full + '.json', JSON.stringify(result, null, 2));
      console.log('wrote', full + '.json');
    }
  }
}

walk(process.argv[2] || '.');`,
  },
  {
    slug: "csv-streaming",
    title: "Convert a URL response",
    category: "advanced",
    tagline: "Fetch a remote spreadsheet and parse it.",
    description:
      "Pipe a fetch response through TabularJS without saving to disk — ideal for serverless functions and edge workers.",
    seoTitle: "Fetch remote CSV/XLSX to JSON — TabularJS",
    seoDescription:
      "Download and parse a remote spreadsheet directly to JSON with TabularJS and fetch.",
    language: "js",
    code: `import tabularjs from 'tabularjs';

const response = await fetch('https://example.com/report.xlsx');
const buffer = await response.arrayBuffer();

const result = await tabularjs(buffer, { name: 'report.xlsx' });
console.log(result.worksheets[0].data.slice(0, 5));`,
  },
  {
    slug: "html-table-scrape",
    title: "Scrape a page's tables",
    category: "advanced",
    tagline: "Extract HTML tables from a page as JSON.",
    description:
      "Point TabularJS at an HTML string and get every <table> on the page as a worksheet — perfect for quick scraping tasks.",
    seoTitle: "Scrape HTML tables to JSON in JavaScript — TabularJS",
    seoDescription:
      "Extract every HTML table from a web page into clean JSON using TabularJS.",
    language: "js",
    code: `import tabularjs from 'tabularjs';

const html = await fetch('https://example.com/prices').then(r => r.text());
const result = await tabularjs(html, { format: 'html' });

// One worksheet per <table>
for (const sheet of result.worksheets) {
  console.log(sheet.name, sheet.data);
}`,
  },
];

export function getExample(slug: string): Example | undefined {
  return examples.find((e) => e.slug === slug);
}
