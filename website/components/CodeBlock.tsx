"use client";

import { Check, Copy } from "lucide-react";
import { useState } from "react";

interface Props {
  code: string;
  language?: string;
  filename?: string;
}

export default function CodeBlock({ code, language = "js", filename }: Props) {
  const [copied, setCopied] = useState(false);

  const copy = async () => {
    await navigator.clipboard.writeText(code);
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };

  return (
    <div className="relative rounded-xl border border-[#1a222d] bg-[#0b0f14] overflow-hidden">
      <div className="flex items-center justify-between px-4 py-2.5 border-b border-[#1a222d] bg-[#0f141b]">
        <div className="flex items-center gap-2">
          <div className="flex gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-[#ff5f57]" />
            <span className="w-2.5 h-2.5 rounded-full bg-[#febc2e]" />
            <span className="w-2.5 h-2.5 rounded-full bg-[#28c840]" />
          </div>
          {filename && (
            <span className="text-xs text-[#8b98a5] ml-2 font-mono">{filename}</span>
          )}
          <span className="text-[10px] uppercase tracking-wider text-[#8b98a5] ml-2 px-2 py-0.5 rounded bg-[#1a222d]">
            {language}
          </span>
        </div>
        <button
          onClick={copy}
          className="text-[#8b98a5] hover:text-white text-xs flex items-center gap-1.5 px-2 py-1 rounded hover:bg-[#1a222d] transition"
          aria-label="Copy"
        >
          {copied ? <Check size={14} /> : <Copy size={14} />}
          {copied ? "Copied" : "Copy"}
        </button>
      </div>
      <pre className="p-5 text-sm font-mono leading-relaxed overflow-x-auto text-[#e6edf3]">
        <code>{code}</code>
      </pre>
    </div>
  );
}
