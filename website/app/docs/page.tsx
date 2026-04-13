import type { Metadata } from "next";
import CodeBlock from "@/components/CodeBlock";
import Link from "next/link";
import { ArrowRight } from "lucide-react";

export const metadata: Metadata = {
  title: "Getting started",
  description:
    "Install TabularJS and parse your first spreadsheet file to JSON in minutes.",
};

export default function DocsPage() {
  return (
    <>
      <div className="mb-2 text-xs uppercase tracking-wider text-brand-600 dark:text-brand-300">Getting started</div>
      <h1 className="text-4xl font-bold mb-4 text-fg">Introduction</h1>
      <p className="text-lg text-fg-muted">
        TabularJS is a zero-dependency JavaScript library that converts spreadsheet files —
        Excel, OpenDocument, CSV, HTML tables, and many legacy formats — into clean structured
        JSON. It runs in both Node.js and the browser.
      </p>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Installation</h2>
      <p className="text-fg-muted">Install from NPM:</p>
      <div className="my-4">
        <CodeBlock code="npm install tabularjs" language="bash" />
      </div>

      <p className="text-fg-muted">Or use the CDN for browser-only projects:</p>
      <div className="my-4">
        <CodeBlock
          language="html"
          code={`<script src="https://cdn.jsdelivr.net/npm/tabularjs/dist/index.js"></script>`}
        />
      </div>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Your first parse</h2>
      <p className="text-fg-muted">
        The default export is a single async function. Pass it a file path, a File object, a
        Buffer, an ArrayBuffer, or a string — it detects the format automatically.
      </p>

      <div className="my-4">
        <CodeBlock
          filename="server.js"
          language="js"
          code={`import tabularjs from 'tabularjs';

// Node.js — file path
const result = await tabularjs('./sales.xlsx');

// Browser — File object
const result2 = await tabularjs(fileFromInput);

console.log(result.worksheets[0].data);`}
        />
      </div>

      <h2 className="mt-10 text-2xl font-semibold text-fg">What you get back</h2>
      <p className="text-fg-muted">
        TabularJS returns a document shape with a <code>worksheets</code> array — one entry per
        sheet. Each worksheet has a <code>name</code>, a two-dimensional <code>data</code> array,
        and optional metadata (styles, merged cells, formulas) depending on the source format.
      </p>

      <div className="my-4">
        <CodeBlock
          language="json"
          code={`{
  "worksheets": [
    {
      "name": "Sheet1",
      "data": [
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25]
      ],
      "mergeCells": {},
      "styles": {}
    }
  ]
}`}
        />
      </div>

      <div className="mt-10 p-5 rounded-xl border border-brand-500/20 bg-brand-500/5">
        <h3 className="font-semibold text-fg mb-1">Works great with Jspreadsheet</h3>
        <p className="text-sm text-fg-muted">
          TabularJS output is shape-compatible with{" "}
          <a href="https://jspreadsheet.com" target="_blank" rel="noreferrer" className="text-brand-600 dark:text-brand-300 hover:underline">
            Jspreadsheet
          </a>{" "}
          — pass the result straight in.
        </p>
      </div>

      <div className="mt-8 flex flex-wrap gap-3">
        <Link
          href="/docs/api"
          className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-fg text-surface text-sm font-medium hover:bg-brand-500 hover:text-white transition"
        >
          Read the API reference <ArrowRight size={14} />
        </Link>
        <Link
          href="/examples"
          className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-line text-fg text-sm font-medium hover:border-brand-500/40 transition"
        >
          Browse examples
        </Link>
      </div>
    </>
  );
}
