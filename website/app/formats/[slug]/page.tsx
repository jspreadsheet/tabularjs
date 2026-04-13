import type { Metadata } from "next";
import { notFound } from "next/navigation";
import Link from "next/link";
import { ArrowLeft, ArrowRight, Check, X } from "lucide-react";
import CodeBlock from "@/components/CodeBlock";
import { formats, getFormat, type FormatFeature } from "@/lib/formats";

export async function generateStaticParams() {
  return formats.map((f) => ({ slug: f.slug }));
}

export async function generateMetadata({
  params,
}: {
  params: Promise<{ slug: string }>;
}): Promise<Metadata> {
  const { slug } = await params;
  const format = getFormat(slug);
  if (!format) return {};
  return {
    title: format.seoTitle,
    description: format.seoDescription,
  };
}

const FEATURE_LABELS: Record<FormatFeature, string> = {
  formulas: "Formulas",
  styles: "Styles",
  merged: "Merged cells",
  comments: "Comments",
  data: "Data values",
  metadata: "Column metadata",
};

const ALL_FEATURES: FormatFeature[] = ["data", "formulas", "styles", "merged", "comments", "metadata"];

export default async function FormatPage({ params }: { params: Promise<{ slug: string }> }) {
  const { slug } = await params;
  const format = getFormat(slug);
  if (!format) notFound();

  const idx = formats.findIndex((f) => f.slug === slug);
  const prev = idx > 0 ? formats[idx - 1] : null;
  const next = idx < formats.length - 1 ? formats[idx + 1] : null;

  return (
    <div className="max-w-4xl mx-auto px-6 py-16">
      <Link
        href="/formats"
        className="inline-flex items-center gap-2 text-sm text-fg-subtle hover:text-fg transition mb-8"
      >
        <ArrowLeft size={14} /> All formats
      </Link>

      <div className="flex items-center gap-4 mb-4 flex-wrap">
        {format.extensions.map((ext) => (
          <span
            key={ext}
            className="inline-flex items-center font-mono text-xs px-2.5 py-1 rounded-md border border-brand-500/30 bg-brand-500/10 text-brand-700 dark:text-brand-200"
          >
            {ext}
          </span>
        ))}
      </div>

      <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-fg">{format.name}</h1>
      <p className="mt-4 text-lg text-fg-muted">{format.tagline}</p>
      <p className="mt-4 text-fg-muted leading-relaxed">{format.description}</p>

      <h2 className="mt-12 text-2xl font-semibold text-fg">Feature support</h2>
      <div className="mt-4 grid grid-cols-2 md:grid-cols-3 gap-3">
        {ALL_FEATURES.map((feat) => {
          const supported = format.features.includes(feat);
          return (
            <div
              key={feat}
              className="flex items-center gap-2 rounded-lg border border-line bg-panel px-4 py-2.5"
            >
              {supported ? (
                <Check size={16} className="text-brand-600 dark:text-brand-300" />
              ) : (
                <X size={16} className="text-fg-faint" />
              )}
              <span className={supported ? "text-fg" : "text-fg-faint"}>
                {FEATURE_LABELS[feat]}
              </span>
            </div>
          );
        })}
      </div>

      {format.notes && (
        <p className="mt-6 text-sm text-fg-subtle italic border-l-2 border-line pl-4">
          {format.notes}
        </p>
      )}

      <h2 className="mt-12 text-2xl font-semibold text-fg">Example</h2>
      <div className="mt-4">
        <CodeBlock code={format.example} language="js" />
      </div>

      <h2 className="mt-12 text-2xl font-semibold text-fg">Why TabularJS for {format.name}?</h2>
      <ul className="mt-4 space-y-2 text-fg-muted">
        <li className="flex gap-3">
          <Check size={18} className="text-brand-600 dark:text-brand-300 shrink-0 mt-0.5" />
          <span>No external parser dependencies — faster installs, smaller bundles.</span>
        </li>
        <li className="flex gap-3">
          <Check size={18} className="text-brand-600 dark:text-brand-300 shrink-0 mt-0.5" />
          <span>Works identically in Node.js and modern browsers.</span>
        </li>
        <li className="flex gap-3">
          <Check size={18} className="text-brand-600 dark:text-brand-300 shrink-0 mt-0.5" />
          <span>
            Output is shape-compatible with{" "}
            <a href="https://jspreadsheet.com" target="_blank" rel="noreferrer" className="text-brand-600 dark:text-brand-300 hover:underline">
              Jspreadsheet
            </a>{" "}
            — render parsed data in a live workbook with one call.
          </span>
        </li>
      </ul>

      <div className="mt-16 pt-8 border-t border-line flex justify-between gap-4">
        {prev ? (
          <Link
            href={`/formats/${prev.slug}`}
            className="flex items-center gap-2 text-sm text-fg-subtle hover:text-fg"
          >
            <ArrowLeft size={14} />
            <span>
              <span className="block text-xs text-fg-faint">Previous</span>
              {prev.name}
            </span>
          </Link>
        ) : (
          <span />
        )}
        {next && (
          <Link
            href={`/formats/${next.slug}`}
            className="flex items-center gap-2 text-sm text-fg-subtle hover:text-fg text-right"
          >
            <span>
              <span className="block text-xs text-fg-faint">Next</span>
              {next.name}
            </span>
            <ArrowRight size={14} />
          </Link>
        )}
      </div>
    </div>
  );
}
