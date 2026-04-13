import type { Config } from "tailwindcss";

const config: Config = {
  darkMode: "class",
  content: [
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        surface: "rgb(var(--surface) / <alpha-value>)",
        "surface-soft": "rgb(var(--surface-soft) / <alpha-value>)",
        panel: "rgb(var(--panel) / <alpha-value>)",
        subtle: "rgb(var(--subtle) / <alpha-value>)",
        "subtle-strong": "rgb(var(--subtle-strong) / <alpha-value>)",
        line: "rgb(var(--border) / <alpha-value>)",
        "line-strong": "rgb(var(--border-strong) / <alpha-value>)",
        fg: "rgb(var(--fg) / <alpha-value>)",
        "fg-muted": "rgb(var(--fg-muted) / <alpha-value>)",
        "fg-subtle": "rgb(var(--fg-subtle) / <alpha-value>)",
        "fg-faint": "rgb(var(--fg-faint) / <alpha-value>)",
        brand: {
          50: "#ecfdf5",
          100: "#d1fae5",
          200: "#a7f3d0",
          300: "#6ee7b7",
          400: "#34d399",
          500: "#10b981",
          600: "#059669",
          700: "#047857",
          800: "#065f46",
          900: "#064e3b",
        },
        accent: {
          300: "#d9f99d",
          400: "#bef264",
          500: "#a3e635",
        },
      },
      fontFamily: {
        sans: ["Inter", "ui-sans-serif", "system-ui", "sans-serif"],
        mono: ["JetBrains Mono", "ui-monospace", "monospace"],
      },
      animation: {
        "gradient-x": "gradient-x 8s ease infinite",
        float: "float 6s ease-in-out infinite",
      },
      keyframes: {
        "gradient-x": {
          "0%, 100%": { "background-position": "0% 50%" },
          "50%": { "background-position": "100% 50%" },
        },
        float: {
          "0%, 100%": { transform: "translateY(0px)" },
          "50%": { transform: "translateY(-10px)" },
        },
      },
    },
  },
  plugins: [require("@tailwindcss/typography")],
};

export default config;
