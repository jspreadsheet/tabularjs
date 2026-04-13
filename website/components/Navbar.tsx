"use client";

import Link from "next/link";
import { motion } from "motion/react";
import { Github, Package, Menu, X } from "lucide-react";
import { useState } from "react";
import ThemeToggle from "./ThemeToggle";

const navLinks = [
  { href: "/docs", label: "Docs" },
  { href: "/formats", label: "Formats" },
  { href: "/examples", label: "Examples" },
];

export default function Navbar() {
  const [open, setOpen] = useState(false);

  return (
    <motion.header
      initial={{ y: -30, opacity: 0 }}
      animate={{ y: 0, opacity: 1 }}
      transition={{ duration: 0.5, ease: "easeOut" }}
      className="sticky top-0 z-50 backdrop-blur-xl bg-surface/70 border-b border-line"
    >
      <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
        <Link href="/" className="flex items-center gap-2 group">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-brand-400 to-accent-500 flex items-center justify-center font-bold text-black shadow-lg shadow-brand-500/20">
            T
          </div>
          <span className="font-semibold text-lg tracking-tight">
            Tabular<span className="text-brand-500 dark:text-brand-400">JS</span>
          </span>
        </Link>

        <nav className="hidden md:flex items-center gap-8">
          {navLinks.map((l) => (
            <Link
              key={l.href}
              href={l.href}
              className="text-sm text-fg-muted hover:text-fg transition-colors"
            >
              {l.label}
            </Link>
          ))}
        </nav>

        <div className="hidden md:flex items-center gap-2">
          <ThemeToggle />
          <a
            href="https://www.npmjs.com/package/tabularjs"
            target="_blank"
            rel="noreferrer"
            className="p-2 rounded-md text-fg-subtle hover:text-fg hover:bg-subtle transition"
            aria-label="NPM"
          >
            <Package size={18} />
          </a>
          <a
            href="https://github.com/jspreadsheet/tabularjs"
            target="_blank"
            rel="noreferrer"
            className="p-2 rounded-md text-fg-subtle hover:text-fg hover:bg-subtle transition"
            aria-label="GitHub"
          >
            <Github size={18} />
          </a>
          <Link
            href="/docs"
            className="ml-1 text-sm font-medium px-4 py-2 rounded-md bg-fg text-surface hover:bg-brand-500 hover:text-white transition"
          >
            Get started
          </Link>
        </div>

        <div className="md:hidden flex items-center gap-1">
          <ThemeToggle />
          <button
            className="text-fg p-2"
            onClick={() => setOpen(!open)}
            aria-label="Menu"
          >
            {open ? <X size={22} /> : <Menu size={22} />}
          </button>
        </div>
      </div>

      {open && (
        <motion.div
          initial={{ opacity: 0, height: 0 }}
          animate={{ opacity: 1, height: "auto" }}
          className="md:hidden border-t border-line bg-surface"
        >
          <nav className="flex flex-col p-4 gap-3">
            {navLinks.map((l) => (
              <Link
                key={l.href}
                href={l.href}
                className="text-fg-muted py-2"
                onClick={() => setOpen(false)}
              >
                {l.label}
              </Link>
            ))}
            <a
              href="https://github.com/jspreadsheet/tabularjs"
              target="_blank"
              rel="noreferrer"
              className="text-fg-muted py-2"
            >
              GitHub
            </a>
          </nav>
        </motion.div>
      )}
    </motion.header>
  );
}
