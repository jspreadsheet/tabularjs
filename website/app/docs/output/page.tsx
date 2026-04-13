import type { Metadata } from "next";
import CodeBlock from "@/components/CodeBlock";

export const metadata: Metadata = {
  title: "Output shape",
  description: "Understanding the JSON document that TabularJS returns.",
};

export default function OutputPage() {
  return (
    <>
      <div className="mb-2 text-xs uppercase tracking-wider text-brand-600 dark:text-brand-300">Reference</div>
      <h1 className="text-4xl font-bold mb-4 text-fg">Output shape</h1>
      <p className="text-lg text-fg-muted">
        TabularJS emits a predictable document structure that is directly compatible with{" "}
        <a href="https://jspreadsheet.com" target="_blank" rel="noreferrer" className="text-brand-600 dark:text-brand-300 hover:underline">
          Jspreadsheet
        </a>
        , so no intermediate shaping layer is needed.
      </p>

      <h2 className="mt-10 text-2xl font-semibold text-fg">TabularDocument</h2>
      <div className="my-4">
        <CodeBlock
          language="ts"
          code={`interface TabularDocument {
  worksheets: Worksheet[];
  meta?: {
    format: string;       // 'xlsx', 'csv', 'html', ...
    source?: string;      // filename when known
  };
}

interface Worksheet {
  name: string;
  data: Cell[][];
  mergeCells?: Record<string, [number, number]>;
  styles?: Record<string, CellStyle>;
  columns?: ColumnMeta[];
  comments?: Record<string, string>;
}

type Cell = string | number | boolean | null | {
  value: string | number | boolean | null;
  formula?: string;
};`}
        />
      </div>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Example output</h2>
      <div className="my-4">
        <CodeBlock
          language="json"
          code={`{
  "worksheets": [
    {
      "name": "Q4 Sales",
      "data": [
        ["Region", "Revenue", "Total"],
        ["North", 12000, { "value": 12000, "formula": "=SUM(B2)" }],
        ["South", 18500, { "value": 18500, "formula": "=SUM(B3)" }]
      ],
      "mergeCells": { "A1": [3, 1] },
      "styles": {
        "A1": { "font-weight": "bold", "background": "#f5f5f5" }
      }
    }
  ],
  "meta": { "format": "xlsx", "source": "q4.xlsx" }
}`}
        />
      </div>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Feature availability per format</h2>
      <p className="text-fg-muted">
        Not every format carries every feature. Text formats like CSV only emit <code>data</code>.
        Rich formats like XLSX and ODS populate <code>styles</code>, <code>mergeCells</code>, and
        formulas too. See the{" "}
        <a className="text-brand-600 dark:text-brand-300 hover:underline" href="/formats">format pages</a> for per-format details.
      </p>
    </>
  );
}
