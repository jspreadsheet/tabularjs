import Link from "next/link";
import { Github, Package } from "lucide-react";

export default function Footer() {
  return (
    <footer className="border-t border-line mt-32">
      <div className="max-w-6xl mx-auto px-6 py-16 grid md:grid-cols-4 gap-10">
        <div>
          <div className="flex items-center gap-2 mb-3">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-brand-400 to-accent-500 flex items-center justify-center font-bold text-black">
              T
            </div>
            <span className="font-semibold text-lg">
              Tabular<span className="text-brand-500 dark:text-brand-400">JS</span>
            </span>
          </div>
          <p className="text-sm text-fg-subtle max-w-xs">
            A zero-dependency JavaScript library for converting spreadsheet files to JSON.
          </p>
        </div>

        <div>
          <h4 className="text-sm font-semibold mb-4 text-fg">Documentation</h4>
          <ul className="space-y-2 text-sm text-fg-subtle">
            <li><Link href="/docs" className="hover:text-fg">Getting started</Link></li>
            <li><Link href="/docs/api" className="hover:text-fg">API reference</Link></li>
            <li><Link href="/formats" className="hover:text-fg">Supported formats</Link></li>
            <li><Link href="/examples" className="hover:text-fg">Examples</Link></li>
          </ul>
        </div>

        <div>
          <h4 className="text-sm font-semibold mb-4 text-fg">Ecosystem</h4>
          <ul className="space-y-2 text-sm text-fg-subtle">
            <li><a href="https://jspreadsheet.com" target="_blank" rel="noreferrer" className="hover:text-fg">Jspreadsheet CE</a></li>
            <li><a href="https://bossanova.uk/jspreadsheet" target="_blank" rel="noreferrer" className="hover:text-fg">Jspreadsheet Pro</a></li>
            <li><a href="https://lemonadejs.com" target="_blank" rel="noreferrer" className="hover:text-fg">LemonadeJS</a></li>
            <li><a href="https://calendarjs.com" target="_blank" rel="noreferrer" className="hover:text-fg">CalendarJS</a></li>
          </ul>
        </div>

        <div>
          <h4 className="text-sm font-semibold mb-4 text-fg">Resources</h4>
          <ul className="space-y-2 text-sm text-fg-subtle">
            <li>
              <a href="https://github.com/jspreadsheet/tabularjs" target="_blank" rel="noreferrer" className="hover:text-fg flex items-center gap-2">
                <Github size={14} /> GitHub
              </a>
            </li>
            <li>
              <a href="https://www.npmjs.com/package/tabularjs" target="_blank" rel="noreferrer" className="hover:text-fg flex items-center gap-2">
                <Package size={14} /> NPM
              </a>
            </li>
            <li>
              <a href="https://github.com/jspreadsheet/tabularjs/issues" target="_blank" rel="noreferrer" className="hover:text-fg">
                Report an issue
              </a>
            </li>
            <li>
              <a href="mailto:support@jspreadsheet.com" className="hover:text-fg">Support</a>
            </li>
          </ul>
        </div>
      </div>

      <div className="border-t border-line py-6 text-center text-xs text-fg-subtle">
        MIT licensed — &copy; {new Date().getFullYear()} Jspreadsheet Team
      </div>
    </footer>
  );
}
