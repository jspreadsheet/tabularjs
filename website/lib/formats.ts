export type FormatFeature = "formulas" | "styles" | "merged" | "comments" | "data" | "metadata";

export interface Format {
  slug: string;
  name: string;
  extensions: string[];
  category: "modern" | "legacy" | "text" | "database" | "web";
  tagline: string;
  description: string;
  features: FormatFeature[];
  seoTitle: string;
  seoDescription: string;
  example: string;
  notes?: string;
}

export const formats: Format[] = [
  {
    slug: "xlsx",
    name: "Excel 2007+ (XLSX)",
    extensions: [".xlsx"],
    category: "modern",
    tagline: "The modern Excel workbook format.",
    description:
      "OOXML-based spreadsheet format introduced with Microsoft Office 2007. TabularJS parses XLSX files natively in the browser and Node.js — no SheetJS, no xlsx dependency.",
    features: ["formulas", "styles", "merged", "data"],
    seoTitle: "Convert XLSX to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse Excel .xlsx files to JSON in Node.js and the browser. Preserves formulas, styles, and merged cells. Zero external dependencies.",
    example: `import tabularjs from 'tabularjs';

// Browser
const file = document.querySelector('input[type=file]').files[0];
const result = await tabularjs(file);

console.log(result.worksheets[0].data);
// => [ [ 'Name', 'Age' ], [ 'Alice', 30 ], ... ]`,
  },
  {
    slug: "xls",
    name: "Excel 97-2003 (XLS)",
    extensions: [".xls"],
    category: "legacy",
    tagline: "Legacy binary Excel format.",
    description:
      "Binary BIFF-based format used by Excel 97 through 2003. TabularJS decodes BIFF records and PTG tokens to reconstruct formulas and data.",
    features: ["formulas", "styles", "merged", "data"],
    seoTitle: "Convert XLS to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse legacy Excel .xls binary files to JSON with full formula decoding from BIFF PTG tokens.",
    example: `import tabularjs from 'tabularjs';

const result = await tabularjs('legacy-report.xls');
console.log(result.worksheets[0].data);`,
  },
  {
    slug: "ods",
    name: "OpenDocument (ODS)",
    extensions: [".ods"],
    category: "modern",
    tagline: "LibreOffice & OpenOffice spreadsheets.",
    description:
      "ISO-standardized spreadsheet format used by LibreOffice, OpenOffice, and Google Sheets exports. TabularJS reads the XML content directly.",
    features: ["formulas", "styles", "merged", "data"],
    seoTitle: "Convert ODS (OpenDocument) to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse OpenDocument .ods spreadsheets from LibreOffice and OpenOffice to JSON in Node.js or the browser.",
    example: `import tabularjs from 'tabularjs';

const result = await tabularjs('report.ods');
console.log(result.worksheets.map(w => w.name));`,
  },
  {
    slug: "csv",
    name: "CSV",
    extensions: [".csv"],
    category: "text",
    tagline: "Comma-separated values.",
    description:
      "The universal data exchange format. TabularJS handles escaping, quoted fields, BOM, and mixed line endings correctly.",
    features: ["data"],
    seoTitle: "Convert CSV to JSON in JavaScript — TabularJS",
    seoDescription:
      "Robust CSV to JSON conversion: quoted fields, escape sequences, BOM, CRLF/LF line endings, all handled.",
    example: `import tabularjs from 'tabularjs';

const result = await tabularjs('data.csv');
const rows = result.worksheets[0].data;`,
  },
  {
    slug: "tsv",
    name: "TSV",
    extensions: [".tsv", ".tab"],
    category: "text",
    tagline: "Tab-separated values.",
    description:
      "Delimited format using tabs. Common for clipboard paste from spreadsheets, database exports, and bioinformatics data.",
    features: ["data"],
    seoTitle: "Convert TSV (tab-separated) to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse tab-separated .tsv files to JSON. Ideal for clipboard-pasted spreadsheet data and database exports.",
    example: `const result = await tabularjs('export.tsv');`,
  },
  {
    slug: "txt",
    name: "Plain Text",
    extensions: [".txt"],
    category: "text",
    tagline: "Tab-delimited text files.",
    description:
      "Treats .txt files as tab-delimited data by default — ideal for legacy text dumps and exports from older business software.",
    features: ["data"],
    seoTitle: "Convert TXT (tab-delimited) to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse tab-delimited .txt text dumps to structured JSON worksheets.",
    example: `const result = await tabularjs('dump.txt');`,
  },
  {
    slug: "xml",
    name: "XML Spreadsheet 2003",
    extensions: [".xml"],
    category: "legacy",
    tagline: "The Office 2003 XML spreadsheet format.",
    description:
      "An XML-based spreadsheet format introduced with Microsoft Office 2003. TabularJS preserves formulas, merged cells, and comments.",
    features: ["formulas", "merged", "comments", "data"],
    seoTitle: "Convert XML Spreadsheet 2003 to JSON — TabularJS",
    seoDescription:
      "Parse the SpreadsheetML XML format (Office 2003) to JSON — formulas and comments preserved.",
    example: `const result = await tabularjs('legacy.xml');`,
  },
  {
    slug: "dif",
    name: "DIF",
    extensions: [".dif"],
    category: "legacy",
    tagline: "Data Interchange Format.",
    description:
      "A vintage format created for VisiCalc. Still found in archival data exchange. TabularJS handles headers, labels, and comments.",
    features: ["data", "comments"],
    seoTitle: "Convert DIF (Data Interchange Format) to JSON — TabularJS",
    seoDescription:
      "Parse the historic DIF (.dif) VisiCalc interchange format to modern JSON.",
    example: `const result = await tabularjs('archive.dif');`,
  },
  {
    slug: "sylk",
    name: "SYLK",
    extensions: [".slk", ".sylk"],
    category: "legacy",
    tagline: "Symbolic Link format.",
    description:
      "Microsoft's line-based spreadsheet format from the Multiplan era. TabularJS decodes E-parameter formulas and row/column records.",
    features: ["formulas", "data"],
    seoTitle: "Convert SYLK (.slk) to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse legacy SYLK/Multiplan .slk files to JSON with formula extraction.",
    example: `const result = await tabularjs('data.slk');`,
  },
  {
    slug: "dbf",
    name: "dBase (DBF)",
    extensions: [".dbf"],
    category: "database",
    tagline: "Legacy database files.",
    description:
      "dBase, FoxPro, and Clipper database files. TabularJS extracts column metadata (types, lengths) alongside the data rows.",
    features: ["data", "metadata"],
    seoTitle: "Convert DBF (dBase) to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse .dbf dBase/FoxPro database files to JSON with column type metadata preserved.",
    example: `const result = await tabularjs('customers.dbf');
console.log(result.worksheets[0].columns);
// [{ name: 'ID', type: 'N', length: 10 }, ...]`,
  },
  {
    slug: "lotus",
    name: "Lotus 1-2-3",
    extensions: [".wks", ".wk1", ".wk3", ".wk4", ".123"],
    category: "legacy",
    tagline: "Lotus 1-2-3 workbooks.",
    description:
      "Files from Lotus 1-2-3, the dominant spreadsheet before Excel. TabularJS extracts cell values from WKS, WK1, WK3, WK4, and .123 files.",
    features: ["data"],
    seoTitle: "Convert Lotus 1-2-3 (.wks, .wk1, .123) to JSON — TabularJS",
    seoDescription:
      "Read Lotus 1-2-3 workbook data from .wks, .wk1, .wk3, .wk4, .123 files in JavaScript.",
    example: `const result = await tabularjs('vintage.wk1');`,
    notes: "Formula tokens are complex and only data values are extracted.",
  },
  {
    slug: "html",
    name: "HTML Tables",
    extensions: [".html", ".htm"],
    category: "web",
    tagline: "Scrape tables from HTML.",
    description:
      "Convert any HTML <table> element — from a file, URL response, or string — into structured worksheet JSON. Supports data-formula attributes, merged cells (colspan/rowspan), and inline styles.",
    features: ["formulas", "styles", "merged", "data"],
    seoTitle: "Convert HTML Table to JSON in JavaScript — TabularJS",
    seoDescription:
      "Parse HTML table elements to JSON. Handles colspan/rowspan merged cells, inline styles, and data-formula attributes.",
    example: `import tabularjs from 'tabularjs';

const html = \`<table>
  <tr><th>Name</th><th>Score</th></tr>
  <tr><td>Alice</td><td data-formula="=SUM(A1:A10)">90</td></tr>
</table>\`;

const result = await tabularjs(html, { format: 'html' });`,
  },
];

export function getFormat(slug: string): Format | undefined {
  return formats.find((f) => f.slug === slug);
}
