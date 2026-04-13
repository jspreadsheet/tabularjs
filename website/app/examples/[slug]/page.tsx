import type { Metadata } from "next";
import { notFound } from "next/navigation";
import Link from "next/link";
import { ArrowLeft, ArrowRight } from "lucide-react";
import CodeBlock from "@/components/CodeBlock";
import { examples, getExample } from "@/lib/examples";

export async function generateStaticParams() {
  return examples.map((e) => ({ slug: e.slug }));
}

export async function generateMetadata({
  params,
}: {
  params: Promise<{ slug: string }>;
}): Promise<Metadata> {
  const { slug } = await params;
  const ex = getExample(slug);
  if (!ex) return {};
  return {
    title: ex.seoTitle,
    description: ex.seoDescription,
  };
}

export default async function ExamplePage({ params }: { params: Promise<{ slug: string }> }) {
  const { slug } = await params;
  const ex = getExample(slug);
  if (!ex) notFound();

  const idx = examples.findIndex((e) => e.slug === slug);
  const prev = idx > 0 ? examples[idx - 1] : null;
  const next = idx < examples.length - 1 ? examples[idx + 1] : null;

  return (
    <div className="max-w-4xl mx-auto px-6 py-16">
      <Link
        href="/examples"
        className="inline-flex items-center gap-2 text-sm text-fg-subtle hover:text-fg transition mb-8"
      >
        <ArrowLeft size={14} /> All examples
      </Link>

      <div className="mb-2 inline-block text-xs uppercase tracking-wider text-brand-600 dark:text-brand-300 px-2 py-0.5 rounded bg-brand-500/10 border border-brand-500/20">
        {ex.category}
      </div>
      <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-fg">{ex.title}</h1>
      <p className="mt-4 text-lg text-fg-muted">{ex.tagline}</p>
      <p className="mt-4 text-fg-muted leading-relaxed">{ex.description}</p>

      <div className="mt-10">
        <CodeBlock code={ex.code} language={ex.language} />
      </div>

      {ex.followUp && (
        <div className="mt-8 p-5 rounded-xl border border-brand-500/20 bg-brand-500/5">
          <p className="text-sm text-fg-muted">
            {ex.followUp.text}{" "}
            <a
              href={ex.followUp.href}
              target="_blank"
              rel="noreferrer"
              className="text-brand-600 dark:text-brand-300 hover:text-brand-500 dark:hover:text-brand-200 underline underline-offset-2"
            >
              {ex.followUp.linkLabel}
            </a>
          </p>
        </div>
      )}

      <div className="mt-16 pt-8 border-t border-line flex justify-between gap-4">
        {prev ? (
          <Link
            href={`/examples/${prev.slug}`}
            className="flex items-center gap-2 text-sm text-fg-subtle hover:text-fg"
          >
            <ArrowLeft size={14} />
            <span>
              <span className="block text-xs text-fg-faint">Previous</span>
              {prev.title}
            </span>
          </Link>
        ) : (
          <span />
        )}
        {next && (
          <Link
            href={`/examples/${next.slug}`}
            className="flex items-center gap-2 text-sm text-fg-subtle hover:text-fg text-right"
          >
            <span>
              <span className="block text-xs text-fg-faint">Next</span>
              {next.title}
            </span>
            <ArrowRight size={14} />
          </Link>
        )}
      </div>
    </div>
  );
}
