import type { Metadata } from "next";
import FormatBadge from "@/components/FormatBadge";
import { formats } from "@/lib/formats";

export const metadata: Metadata = {
  title: "Supported formats",
  description:
    "Every spreadsheet format TabularJS supports — XLSX, XLS, CSV, TSV, ODS, HTML, DBF, SYLK, DIF, XML SpreadsheetML, and Lotus 1-2-3.",
};

const categories = [
  { key: "modern", title: "Modern spreadsheets", desc: "The formats you use every day." },
  { key: "text", title: "Text formats", desc: "Delimited data — CSV, TSV, and plain text." },
  { key: "web", title: "Web formats", desc: "HTML tables from pages and documents." },
  { key: "database", title: "Database formats", desc: "dBase, FoxPro and Clipper files." },
  { key: "legacy", title: "Legacy formats", desc: "Vintage and archival spreadsheet formats." },
] as const;

export default function FormatsPage() {
  return (
    <div className="max-w-6xl mx-auto px-6 py-16">
      <div className="text-center mb-16">
        <h1 className="text-4xl md:text-5xl font-bold tracking-tight">Supported formats</h1>
        <p className="mt-4 text-lg text-fg-muted max-w-2xl mx-auto">
          TabularJS handles 16+ spreadsheet file formats out of the box — from the modern Office
          OOXML to vintage Lotus 1-2-3 dumps.
        </p>
      </div>

      {categories.map((cat) => {
        const items = formats.filter((f) => f.category === cat.key);
        if (items.length === 0) return null;
        return (
          <section key={cat.key} className="mb-16">
            <div className="mb-6">
              <h2 className="text-2xl font-semibold text-fg">{cat.title}</h2>
              <p className="text-fg-subtle">{cat.desc}</p>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
              {items.map((f, i) => (
                <FormatBadge key={f.slug} format={f} index={i} />
              ))}
            </div>
          </section>
        );
      })}
    </div>
  );
}
