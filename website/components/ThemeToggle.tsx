"use client";

import { Moon, Sun } from "lucide-react";
import { useEffect, useState } from "react";

export default function ThemeToggle() {
  const [theme, setTheme] = useState<"light" | "dark">("dark");
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    const saved = (localStorage.getItem("theme") as "light" | "dark" | null) ?? "dark";
    setTheme(saved);
    setMounted(true);
  }, []);

  const toggle = () => {
    const next = theme === "dark" ? "light" : "dark";
    setTheme(next);
    document.documentElement.classList.remove("light", "dark");
    document.documentElement.classList.add(next);
    localStorage.setItem("theme", next);
  };

  if (!mounted) {
    return <div className="w-8 h-8" aria-hidden />;
  }

  return (
    <button
      onClick={toggle}
      className="p-2 rounded-md text-fg-subtle hover:text-fg hover:bg-subtle transition"
      aria-label="Toggle theme"
      title={theme === "dark" ? "Switch to light" : "Switch to dark"}
    >
      {theme === "dark" ? <Sun size={18} /> : <Moon size={18} />}
    </button>
  );
}
