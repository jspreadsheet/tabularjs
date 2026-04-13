"use client";

import { motion } from "motion/react";
import Link from "next/link";
import { Format } from "@/lib/formats";

export default function FormatBadge({ format, index = 0 }: { format: Format; index?: number }) {
  const ext = format.extensions[0].replace(".", "").toUpperCase();

  return (
    <motion.div
      initial={{ opacity: 0, scale: 0.9 }}
      whileInView={{ opacity: 1, scale: 1 }}
      viewport={{ once: true }}
      transition={{ duration: 0.4, delay: index * 0.03 }}
      whileHover={{ y: -4, transition: { duration: 0.15 } }}
    >
      <Link
        href={`/formats/${format.slug}`}
        className="block group relative rounded-xl border border-line bg-panel p-5 hover:border-brand-500/50 transition-all"
      >
        <div className="absolute inset-0 rounded-xl bg-gradient-to-br from-brand-500/0 to-accent-500/0 group-hover:from-brand-500/10 group-hover:to-accent-500/5 transition-all" />
        <div className="relative">
          <div className="inline-flex items-center justify-center w-10 h-10 rounded-md bg-subtle text-brand-600 dark:text-brand-300 font-mono text-xs font-semibold mb-3">
            {ext}
          </div>
          <h3 className="font-semibold text-fg group-hover:text-brand-600 dark:group-hover:text-brand-300 transition-colors">
            {format.name}
          </h3>
          <p className="text-sm text-fg-subtle mt-1 line-clamp-2">{format.tagline}</p>
        </div>
      </Link>
    </motion.div>
  );
}
