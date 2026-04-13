"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import clsx from "clsx";

const sections = [
  {
    title: "Getting started",
    links: [
      { href: "/docs", label: "Introduction" },
      { href: "/docs/api", label: "API reference" },
      { href: "/docs/output", label: "Output shape" },
    ],
  },
  {
    title: "Examples",
    links: [
      { href: "/examples/nodejs-basic", label: "Node.js basic" },
      { href: "/examples/browser-file-input", label: "Browser upload" },
      { href: "/examples/react", label: "React" },
      { href: "/examples/vue", label: "Vue" },
      { href: "/examples/nextjs-api-route", label: "Next.js API" },
      { href: "/examples/jspreadsheet-ce", label: "Jspreadsheet CE" },
    ],
  },
];

export default function DocsSidebar() {
  const pathname = usePathname();

  return (
    <aside className="hidden lg:block w-56 shrink-0">
      <div className="sticky top-20 space-y-8">
        {sections.map((s) => (
          <div key={s.title}>
            <h4 className="text-xs uppercase tracking-wider text-fg-subtle font-semibold mb-3">
              {s.title}
            </h4>
            <ul className="space-y-1">
              {s.links.map((l) => {
                const active = pathname === l.href;
                return (
                  <li key={l.href}>
                    <Link
                      href={l.href}
                      className={clsx(
                        "block px-3 py-1.5 text-sm rounded-md transition",
                        active
                          ? "bg-brand-500/10 text-brand-600 dark:text-brand-300 border-l-2 border-brand-400"
                          : "text-fg-muted hover:text-fg hover:bg-panel"
                      )}
                    >
                      {l.label}
                    </Link>
                  </li>
                );
              })}
            </ul>
          </div>
        ))}
      </div>
    </aside>
  );
}
