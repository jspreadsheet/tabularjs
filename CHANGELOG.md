# Changelog

All notable changes to this project will be documented in this file.

## [2.1.1] - 2026-10-28

### Fixed
- **Lotus**: formula cells exposed the raw, undecoded token bytes as `cells[address].formula` (an array of numbers); only the calculated value is returned now, in `data`
- **Lotus**: string results of formulas included the record header bytes (`"ÿ\u0002ok"` instead of `"ok"`)
- **DBF**: the record count from the header was trusted: a crafted 41 KB file allocated 118 million cells; records are now limited to what the file holds
- **DBF**: text fields ignored the file's codepage (Cyrillic, Central European, CJK... came out garbled); they are decoded with the codepage of the language driver byte, or `options.encoding` when given
- **SYLK**: files with CR-only line endings (classic Mac Excel) parsed as empty
- **DIF / Numbers**: removed debug `console.log` output during parsing
- Found by running the LibreOffice import test corpus (`sc/qa/unit/data`: 123, wks, dif, dbf, slk); 6 new regression tests (433 total)

### Changed
- README formula table: CSV/TSV/TXT keep `=...` cells as text, Lotus returns calculated values only

## [2.1.0] - 2026-10-28

### Added
- **XLS formulas - references to other sheets**: `Sheet2!A1`, `'My Sheet'!A1:B2`, multi-sheet ranges (`Sheet1:Sheet3!A1`) and external workbooks (`'[Book2.xls]Prices'!A1`), resolved through SUPBOOK / EXTERNSHEET (`tRef3d`, `tArea3d`, `tRefErr3d`, `tAreaErr3d`); references to deleted sheets become `#REF!`
- **XLS formulas - names**: defined names (`=SUM(Total)`), built-in names (`Print_Area`...), external names (`[Book2.xls]!Rate`) and add-in / user-defined functions (`MYUDF(A1)`) via NAME / EXTERNNAME (`tName`, `tNameX`)
- **XLS formulas - arrays**: array formulas (ARRAY record) on every cell of their range, and array constants (`{1,2;"a",TRUE}`) read from RgbExtra (`tArray`)
- **XLS formulas**: memory tokens (`tMemArea`, `tMemFunc`, `tMemErr`, ...) and the full error set (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`)
- Encrypted .xls: workbooks protected with Office 97 RC4 and Excel's default password (write-protected / "read-only recommended" files) are decrypted; other passwords, RC4 CryptoAPI and XOR obfuscation raise a clear "Workbook is encrypted" error instead of parsing ciphertext (which ran out of memory)
- `parseWorkbookStream(bytes)` export: parse a raw BIFF8 Workbook stream
- TypeScript definitions for the byte input (`Uint8Array | ArrayBuffer`) and `options.extension` added in 2.0.0, for `onload` / `onerror`, and for `worksheetState`, `gridline`, `media` and `definedNames`, which the parsers already returned
- **HTML input**: `tabularjs('<table>...</table>')` and DOM elements now work as documented (a string starting with `<` is markup, not a file path); before, HTML strings failed with `Unsupported file type`
- 51 new regression tests (425 total): HTML strings and multiple tables, CSV encoding, TXT delimiter, XLS records (SST rich/phonetic/CONTINUE, embedded charts, chart sheets, RC4, sheet/name/external references, array formulas, cached results), PTG 3-D refs/names/arrays/tMem, truncated CFB files

### Fixed
- **XLS strings**: the Shared String Table ignored rich-text runs, phonetic (Far East) data and CONTINUE records; after the first formatted string every later cell read the wrong string (or garbage with control characters), and strings past the first 8 KB were lost. Measured on 400 real-world workbooks against xlrd: 0 value mismatches after the fix (was ~40% of workbooks with corrupted text)
- **XLS sheets**: cells were assigned to sheets by counting worksheet BOF records, so a chart embedded in a worksheet ended the sheet early (cells after it dropped) and a chart sheet shifted every following sheet's data; records are now mapped through the BOUNDSHEET stream offsets
- **XLS formula results**: boolean, error and empty-string cached results were read as numbers (wrong offset of the `fExprO` marker); BOOLERR errors now read as `#DIV/0!` etc. instead of `#ERROR7`
- **XLS**: LABEL records (inline strings written by older tools) were ignored
- **XLS formulas**: zero-argument functions (`NOW()`, `TODAY()`, `PI()`, `TRUE()`...) swallowed the preceding operand
- **XLS formulas**: a formula that cannot be decoded faithfully now shows the value Excel cached for the cell instead of `=ERROR` or a wrong formula
- **CSV**: `options.encoding` was ignored and files were always decoded as UTF-8; it now accepts any TextDecoder label (and iconv-lite codepages in Node)
- **TXT**: a `delimiter` passed for `.txt` files was overridden by tab
- Removed the undocumented `options.worksheetIndex` from the JSDoc: it was never implemented
- **CFB**: truncated files (last sector partial, or chains running past the end) threw `Sector out of bounds`; the readable part is now returned

### Changed
- **HTML**: every `<table>` on the page becomes a worksheet (`Table1`, `Table2`...); before, only the first table was returned. Pass `tableIndex` to get a single table as before
- Internal `decodePTG(tokens, debug, cellContext, ctx)` is now strict: it throws `PTGError` on unknown tokens, truncated input, stack underflow or leftover operands instead of guessing, and takes a 4th `ctx` argument (workbook context for sheet and name references). Not part of the public API

## [2.0.0] - 2026-10-09

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

