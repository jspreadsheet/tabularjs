# Changelog

All notable changes to this project will be documented in this file.

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

