import type { Metadata } from "next";
import Link from "next/link";
import { examples } from "@/lib/examples";
import { ArrowRight } from "lucide-react";

export const metadata: Metadata = {
  title: "Examples",
  description:
    "Real-world examples using TabularJS — Node.js, browser uploads, React, Vue, Next.js, Jspreadsheet integrations, and more.",
};

const categories = [
  { key: "basics", title: "Basics", desc: "Short, self-contained snippets to get started." },
  { key: "framework", title: "Frameworks", desc: "React, Vue, Next.js, Express and more." },
  { key: "integration", title: "Integrations", desc: "Wire TabularJS up to other tools." },
  { key: "advanced", title: "Advanced", desc: "Batch jobs, remote files, scraping." },
] as const;

export default function ExamplesPage() {
  return (
    <div className="max-w-6xl mx-auto px-6 py-16">
      <div className="text-center mb-16">
        <h1 className="text-4xl md:text-5xl font-bold tracking-tight">Examples</h1>
        <p className="mt-4 text-lg text-fg-muted max-w-2xl mx-auto">
          Copy-paste recipes for every common TabularJS use case — from a single-file Node.js
          script to a full Next.js upload endpoint.
        </p>
      </div>

      {categories.map((cat) => {
        const items = examples.filter((e) => e.category === cat.key);
        if (items.length === 0) return null;
        return (
          <section key={cat.key} className="mb-16">
            <div className="mb-6">
              <h2 className="text-2xl font-semibold text-fg">{cat.title}</h2>
              <p className="text-fg-subtle">{cat.desc}</p>
            </div>
            <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
              {items.map((ex) => (
                <Link
                  key={ex.slug}
                  href={`/examples/${ex.slug}`}
                  className="group relative rounded-xl border border-line bg-panel p-5 hover:border-brand-500/40 transition-all"
                >
                  <h3 className="font-semibold text-fg group-hover:text-brand-600 dark:group-hover:text-brand-300 transition-colors">
                    {ex.title}
                  </h3>
                  <p className="mt-1 text-sm text-fg-subtle line-clamp-2">{ex.tagline}</p>
                  <div className="mt-4 inline-flex items-center gap-1 text-xs text-brand-600 dark:text-brand-300 opacity-0 group-hover:opacity-100 transition-opacity">
                    View example <ArrowRight size={12} />
                  </div>
                </Link>
              ))}
            </div>
          </section>
        );
      })}
    </div>
  );
}
