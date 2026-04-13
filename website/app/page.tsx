"use client";

import Link from "next/link";
import { motion } from "motion/react";
import {
  ArrowRight,
  Github,
  Package,
  Zap,
  Shield,
  Layers,
  FileSpreadsheet,
  Code2,
  Sparkles,
} from "lucide-react";
import CodeBlock from "@/components/CodeBlock";
import FeatureCard from "@/components/FeatureCard";
import FormatBadge from "@/components/FormatBadge";
import Section from "@/components/Section";
import { formats } from "@/lib/formats";

const HERO_CODE = `import tabularjs from 'tabularjs';

// Parse any of 16+ formats
const result = await tabularjs(file);

// Clean JSON — ready for any app
console.log(result.worksheets[0].data);`;

export default function HomePage() {
  return (
    <>
      {/* HERO */}
      <section className="relative overflow-hidden">
        <div className="absolute inset-0 grid-pattern opacity-60" />
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[800px] h-[400px] bg-gradient-to-br from-brand-500/25 via-accent-500/15 to-transparent blur-3xl rounded-full pointer-events-none" />

        <div className="max-w-6xl mx-auto px-6 pt-20 md:pt-28 pb-16 relative">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6 }}
            className="flex justify-center mb-6"
          >
            <a
              href="https://github.com/jspreadsheet/tabularjs"
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full border border-line bg-panel/60 backdrop-blur text-xs text-fg-muted hover:border-brand-500/40 transition"
            >
              <Sparkles size={12} className="text-brand-500 dark:text-brand-300" />
              v1.0.1 is out — zero-dependency parser
              <ArrowRight size={12} />
            </a>
          </motion.div>

          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.7, delay: 0.1 }}
            className="text-5xl md:text-7xl font-bold text-center tracking-tight leading-[1.05]"
          >
            Free Spreadsheet to JSON.
            <br />
            <span className="gradient-text">Several Formats. One line.</span>
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.7, delay: 0.2 }}
            className="mt-6 text-lg md:text-xl text-fg-muted text-center max-w-2xl mx-auto"
          >
            A professional JavaScript library for converting XLSX, CSV, ODS, HTML, DBF, SYLK, DIF,
            Lotus 1-2-3 and more to clean JSON — in Node.js or the browser.
          </motion.p>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.7, delay: 0.3 }}
            className="mt-8 flex flex-wrap justify-center gap-3"
          >
            <Link
              href="/docs"
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-fg text-surface font-medium hover:bg-brand-500 hover:text-white transition"
            >
              Get started
              <ArrowRight size={16} />
            </Link>
            <a
              href="https://github.com/jspreadsheet/tabularjs"
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-line bg-panel text-fg font-medium hover:border-brand-500/40 transition"
            >
              <Github size={16} />
              Star on GitHub
            </a>
            <a
              href="https://www.npmjs.com/package/tabularjs"
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-line bg-panel text-fg font-medium hover:border-brand-500/40 transition"
            >
              <Package size={16} />
              npm install tabularjs
            </a>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 40 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8, delay: 0.4 }}
            className="mt-16 max-w-2xl mx-auto"
          >
            <div className="relative">
              <div className="absolute -inset-4 bg-gradient-to-r from-brand-500/25 via-accent-500/25 to-brand-500/25 blur-2xl rounded-2xl" />
              <div className="relative">
                <CodeBlock code={HERO_CODE} language="js" filename="example.js" />
              </div>
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.6 }}
            className="mt-12 flex flex-wrap justify-center gap-x-10 gap-y-2 text-sm text-fg-subtle"
          >
            <span>✓ Zero dependencies</span>
            <span>✓ 16+ formats</span>
            <span>✓ Node.js & browser</span>
            <span>✓ 312+ tests</span>
          </motion.div>
        </div>
      </section>

      {/* FEATURES */}
      <Section className="max-w-6xl mx-auto px-6 py-20">
        <div className="text-center mb-12">
          <h2 className="text-3xl md:text-4xl font-bold tracking-tight">
            Built for real spreadsheet workflows
          </h2>
          <p className="mt-3 text-fg-subtle max-w-2xl mx-auto">
            TabularJS handles the messy parts — legacy binary formats, inline styles, formulas, and
            edge cases — so you can focus on your data.
          </p>
        </div>

        <div className="grid md:grid-cols-3 gap-5">
          <FeatureCard icon={Zap} title="Zero dependencies" delay={0}>
            No SheetJS, no xlsx, no heavyweight parsers. Pure JavaScript with a tiny install size.
          </FeatureCard>
          <FeatureCard icon={Layers} title="16+ formats" delay={0.05}>
            Excel, OpenDocument, CSV, TSV, HTML tables, DBF, SYLK, DIF, and the full Lotus 1-2-3
            family.
          </FeatureCard>
          <FeatureCard icon={Shield} title="Framework agnostic" delay={0.1}>
            Works with Vanilla JS, React, Vue, Angular, Svelte, and Node.js — with one API.
          </FeatureCard>
          <FeatureCard icon={FileSpreadsheet} title="Formula-aware" delay={0.15}>
            Extracts formulas from XLSX, XLS (BIFF PTG), ODS, SYLK, HTML, and XML SpreadsheetML.
          </FeatureCard>
          <FeatureCard icon={Code2} title="Structured JSON" delay={0.2}>
            Output shape is stable and predictable — directly compatible with Jspreadsheet.
          </FeatureCard>
          <FeatureCard icon={Sparkles} title="Browser & server" delay={0.25}>
            Accept file inputs in the browser, or process files and buffers on the server.
          </FeatureCard>
        </div>
      </Section>

      {/* FORMATS GRID */}
      <Section className="max-w-6xl mx-auto px-6 py-20">
        <div className="flex items-end justify-between mb-8 flex-wrap gap-4">
          <div>
            <h2 className="text-3xl md:text-4xl font-bold tracking-tight">Every format you need</h2>
            <p className="mt-3 text-fg-subtle max-w-xl">
              From modern Excel workbooks to archival Lotus 1-2-3 data — one API, every format.
            </p>
          </div>
          <Link
            href="/formats"
            className="text-brand-600 dark:text-brand-300 hover:text-brand-500 dark:hover:text-brand-200 text-sm inline-flex items-center gap-1"
          >
            Browse all formats <ArrowRight size={14} />
          </Link>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
          {formats.map((f, i) => (
            <FormatBadge key={f.slug} format={f} index={i} />
          ))}
        </div>
      </Section>

      {/* CODE DEMO */}
      <Section className="max-w-6xl mx-auto px-6 py-20">
        <div className="grid lg:grid-cols-2 gap-12 items-center">
          <div>
            <div className="inline-block px-3 py-1 rounded-full text-xs bg-brand-500/10 text-brand-600 dark:text-brand-300 border border-brand-500/20 mb-4">
              Integration
            </div>
            <h2 className="text-3xl md:text-4xl font-bold tracking-tight">
              Drop into{" "}
              <a
                href="https://jspreadsheet.com"
                target="_blank"
                rel="noreferrer"
                className="gradient-text"
              >
                Jspreadsheet
              </a>{" "}
              with one call
            </h2>
            <p className="mt-4 text-fg-muted leading-relaxed">
              TabularJS output is shape-compatible with{" "}
              <a href="https://jspreadsheet.com" target="_blank" rel="noreferrer" className="text-brand-600 dark:text-brand-300 hover:underline">
                Jspreadsheet CE
              </a>{" "}
              and{" "}
              <a
                href="https://bossanova.uk/jspreadsheet"
                target="_blank"
                rel="noreferrer"
                className="text-brand-600 dark:text-brand-300 hover:underline"
              >
                Jspreadsheet Pro
              </a>
              . Parse a user&apos;s upload, pass the result straight in, and you have a live editable
              workbook — no glue code.
            </p>
            <div className="mt-6 flex gap-3">
              <Link
                href="/examples/jspreadsheet-ce"
                className="inline-flex items-center gap-2 text-sm px-4 py-2 rounded-lg bg-subtle hover:bg-subtle-strong text-fg transition"
              >
                See the example <ArrowRight size={14} />
              </Link>
              <Link
                href="/examples"
                className="inline-flex items-center gap-2 text-sm px-4 py-2 rounded-lg border border-line hover:border-brand-500/40 text-fg transition"
              >
                All examples
              </Link>
            </div>
          </div>

          <div>
            <CodeBlock
              filename="upload.html"
              language="html"
              code={`<script src="https://cdn.jsdelivr.net/npm/jspreadsheet-ce"></script>
<script src="https://cdn.jsdelivr.net/npm/tabularjs"></script>

<div id="sheet"></div>
<input type="file" id="f" />

<script>
  document.getElementById('f').onchange = async (e) => {
    const result = await tabularjs(e.target.files[0]);
    jspreadsheet(
      document.getElementById('sheet'),
      result
    );
  };
</script>`}
            />
          </div>
        </div>
      </Section>

      {/* CTA */}
      <Section className="max-w-6xl mx-auto px-6 py-20">
        <div className="relative rounded-2xl border border-line bg-gradient-to-br from-panel via-surface-soft to-panel p-10 md:p-16 overflow-hidden">
          <div className="absolute inset-0 grid-pattern opacity-30" />
          <div className="absolute top-0 right-0 w-96 h-96 bg-brand-500/15 blur-3xl rounded-full" />
          <div className="absolute bottom-0 left-0 w-80 h-80 bg-accent-500/10 blur-3xl rounded-full" />

          <div className="relative text-center">
            <h2 className="text-3xl md:text-5xl font-bold tracking-tight">
              Start parsing in{" "}
              <span className="gradient-text">under a minute</span>
            </h2>
            <p className="mt-4 text-fg-muted max-w-xl mx-auto">
              MIT licensed. Install from NPM, or drop the CDN script into any HTML page.
            </p>
            <div className="mt-8 flex flex-wrap justify-center gap-3">
              <Link
                href="/docs"
                className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-fg text-surface font-medium hover:bg-brand-500 hover:text-white transition"
              >
                Read the docs
                <ArrowRight size={16} />
              </Link>
              <Link
                href="/examples"
                className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg border border-line bg-panel text-fg font-medium hover:border-brand-500/40 transition"
              >
                Browse examples
              </Link>
            </div>
          </div>
        </div>
      </Section>
    </>
  );
}
