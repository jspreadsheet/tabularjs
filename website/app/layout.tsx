import type { Metadata } from "next";
import "./globals.css";
import Navbar from "@/components/Navbar";
import Footer from "@/components/Footer";
import ThemeScript from "@/components/ThemeScript";

export const metadata: Metadata = {
  metadataBase: new URL("https://tabularjs.com"),
  title: {
    default: "TabularJS — Convert spreadsheets to JSON in JavaScript",
    template: "%s | TabularJS",
  },
  description:
    "A zero-dependency JavaScript library for converting spreadsheet files to JSON. Supports XLSX, XLS, ODS, CSV, HTML, DBF, SYLK, DIF, Lotus 1-2-3, and more. Works in Node.js and the browser.",
  keywords: [
    "xlsx to json",
    "excel to json javascript",
    "csv parser",
    "ods to json",
    "spreadsheet parser",
    "jspreadsheet",
    "tabularjs",
  ],
  openGraph: {
    type: "website",
    title: "TabularJS — Convert spreadsheets to JSON in JavaScript",
    description:
      "16+ file formats. Zero dependencies. Works in Node.js and the browser.",
    url: "https://tabularjs.com",
    siteName: "TabularJS",
  },
  twitter: {
    card: "summary_large_image",
    title: "TabularJS",
    description: "Convert spreadsheets to JSON in JavaScript.",
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <head>
        <ThemeScript />
      </head>
      <body>
        <Navbar />
        <main>{children}</main>
        <Footer />
      </body>
    </html>
  );
}
