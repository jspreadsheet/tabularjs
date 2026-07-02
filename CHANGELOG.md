# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Fixed
- **Packaging**: the npm package could not be imported in Node.js at all - `main` pointed at a UMD bundle inside a `"type": "module"` package, so the UMD wrapper ran as ESM (where top-level `this` is undefined) and threw. Added an `exports` map: Node resolves the ESM source, bundlers/browsers get a real ES module bundle (`dist/index.mjs`), and `<script>` tags keep the UMD build (now using `globalThis`)
- **XLS formulas**: BIFF8 row references were masked to 14 bits (rows above 16384 decoded wrong) and "relative" rows >= 8193 in normal formulas were corrupted by bogus sign-extension; `tRef`/`tArea` now decode as absolute storage per spec
- **XLS shared formulas**: `tAreaN` tokens (ranges in shared formulas) were not handled, derailing token parsing; `tRefN`/`tAreaN` offsets now sign-extend correctly (16-bit rows, 8-bit columns)
- **XLSX assets**: relationship targets with multiple `../` segments (images, charts) failed to resolve; replaced string hacking with proper OPC path resolution
- **Encoding**: `detectEncoding` aliased windows-1251 (Cyrillic) to latin1 and CJK encodings to utf-8, producing mojibake; decoding now goes through `TextDecoder` (all WHATWG labels) with an iconv-lite fallback for DOS codepages

### Added
- Buffer input: `tabularjs(uint8Array | arrayBuffer | buffer)` now works as documented, plus an `options.extension` override
- Content-based format detection (magic bytes) when the extension is missing or unrecognized - zip containers (xlsx/ods/numbers), OLE2, Lotus, dBase, SYLK, DIF, XML, HTML, CSV/TSV
- Apple Numbers (.numbers) parser wired into format dispatch (experimental)
- Zip decompression caps on all xlsx/ods/numbers entry reads (zip-bomb protection, counts actual inflated bytes)
- CFB (OLE2) stream allocations clamped to the container size; FAT chain length bounded (crafted .xls DoS protection)
- 62 new regression tests (374 total)

### Changed
- Browser bundles no longer inline iconv-lite/chardet code tables: 512 KiB -> 196 KiB; both are now declared `optionalDependencies` for Node users

## [1.0.1] - 2025-11-29

### Added
- Excel 2007+ (.xlsx) parser with formula support
- Excel 97-2003 (.xls) parser with formula support
- OpenDocument Spreadsheet (.ods) parser
- XML Spreadsheet 2003 (.xml) parser with R1C1 to A1 conversion
- CSV/TSV parser with custom delimiter support
- HTML table parser with merged cells support
- Lotus 1-2-3 (.wks, .wk1, .wk3, .wk4, .123) parser
- SYLK (.slk, .sylk) parser
- DIF (.dif) parser
- dBase (.dbf) parser
- Automatic file format detection by extension
- Formula preservation where supported
- Merged cells support (colspan, rowspan)
- Cell styling and formatting
- Comments and annotations
- Named ranges support
- Multiple worksheets support
- Browser and Node.js compatibility
- Zero external spreadsheet dependencies
- Direct Jspreadsheet CE/Pro compatibility
- Comprehensive test suite with 312 tests covering all parsers
- Test coverage reporting with c8
- GitHub Actions CI/CD pipeline (multi-OS, multi-Node)
- Contributing guidelines (CONTRIBUTING.md)
- Test watch mode for development


### Features
- Supports 16+ file formats
- Automatic format detection
- Formula preservation (XLS, XLSX, ODS, XML, SYLK, HTML)
- Merged cells handling
- Styling support (fonts, colors, borders, alignment)
- Comments and annotations
- Hidden rows/columns
- Column widths and row heights
- Worksheet visibility states
- Framework agnostic (Vanilla JS, React, Vue, Angular)

