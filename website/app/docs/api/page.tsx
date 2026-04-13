import type { Metadata } from "next";
import CodeBlock from "@/components/CodeBlock";

export const metadata: Metadata = {
  title: "API reference",
  description: "TabularJS API reference — function signature, options, and inputs.",
};

export default function ApiPage() {
  return (
    <>
      <div className="mb-2 text-xs uppercase tracking-wider text-brand-600 dark:text-brand-300">API reference</div>
      <h1 className="text-4xl font-bold mb-4 text-fg">API</h1>
      <p className="text-lg text-fg-muted">
        TabularJS exposes a single default export — an async function that takes an input and
        returns a worksheet document.
      </p>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Signature</h2>
      <div className="my-4">
        <CodeBlock
          language="ts"
          code={`declare function tabularjs(
  input: string | File | Blob | Buffer | ArrayBuffer | Uint8Array,
  options?: TabularOptions
): Promise<TabularDocument>;

interface TabularOptions {
  format?: string;      // force a format (e.g. 'html', 'csv')
  name?: string;        // hint for format detection
  encoding?: string;    // for text formats, default 'utf-8'
  delimiter?: string;   // for CSV/TSV, auto-detected by default
}`}
        />
      </div>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Accepted inputs</h2>
      <ul>
        <li><strong>File path (string)</strong> — Node.js only. Format is detected from the extension.</li>
        <li><strong>File / Blob</strong> — browser uploads. Format is detected from name and content.</li>
        <li><strong>Buffer / ArrayBuffer / Uint8Array</strong> — in-memory binary data.</li>
        <li><strong>String of markup</strong> — useful for HTML tables. Pass <code>{`{ format: 'html' }`}</code>.</li>
      </ul>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Options</h2>
      <h3 className="text-lg font-semibold text-fg mt-6">format</h3>
      <p className="text-fg-muted">
        Force a specific parser. Useful when passing raw content where extension is unavailable —
        for example, HTML fetched from a URL.
      </p>

      <h3 className="text-lg font-semibold text-fg mt-6">name</h3>
      <p className="text-fg-muted">
        A filename hint. When passing a Buffer, include a <code>name</code> so TabularJS can use
        the extension for format detection.
      </p>

      <h3 className="text-lg font-semibold text-fg mt-6">encoding</h3>
      <p className="text-fg-muted">
        Text encoding for CSV/TSV/TXT. Defaults to UTF-8. BOM is detected and stripped
        automatically.
      </p>

      <h3 className="text-lg font-semibold text-fg mt-6">delimiter</h3>
      <p className="text-fg-muted">
        Field delimiter for text formats. When omitted, TabularJS auto-detects comma, tab, or
        semicolon.
      </p>

      <h2 className="mt-10 text-2xl font-semibold text-fg">Errors</h2>
      <p className="text-fg-muted">
        Invalid inputs throw. Wrap calls in try/catch in user-facing code paths.
      </p>
      <div className="my-4">
        <CodeBlock
          language="js"
          code={`try {
  const result = await tabularjs(file);
} catch (err) {
  console.error('Could not parse:', err.message);
}`}
        />
      </div>
    </>
  );
}
