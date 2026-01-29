# Changelog

All notable changes to Cuneiform will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2025-01-28

### Added

#### Core Features
- Complete Office Open XML SpreadsheetML (.xlsx) read and write support
- High-level `Workbook` API for reading .xlsx files with lazy sheet loading
- High-level `WorkbookWriter` API for creating and writing .xlsx files
- Full cell value resolution with `CellValue` enum (text, number, date, boolean, error, empty)
- Support for reading and writing formulas with cached values

#### Formula Engine
- Comprehensive formula evaluation engine with 467 Excel functions
- 97% full implementation rate (453/467 functions with complete logic)
- Support for 12 function categories:
  - Mathematical (67 functions): SUM, AVERAGE, ROUND, ABS, etc.
  - Statistical (134 functions): STDEV, VAR, NORMDIST, TDIST, etc.
  - Text (36 functions): CONCATENATE, LEFT, RIGHT, FIND, SUBSTITUTE, etc.
  - Date & Time (24 functions): DATE, NOW, YEAR, MONTH, NETWORKDAYS, etc.
  - Financial (52 functions): PV, FV, PMT, IRR, NPV, etc.
  - Logical (16 functions): IF, AND, OR, NOT, IFERROR, etc.
  - Lookup & Reference (21 functions): VLOOKUP, HLOOKUP, INDEX, MATCH, etc.
  - Engineering (56 functions): Complex number functions, base conversions, etc.
  - Database (10 functions): DSUM, DAVERAGE, DCOUNT, etc.
  - Information (23 functions): ISBLANK, ISERROR, TYPE, CELL, etc.
  - Compatibility (14 functions): FORECAST.LINEAR, PERCENTILE.INC, etc.
  - Web & Service (14 functions): WEBSERVICE, ENCODEURL (stubs)
- Advanced formula features: array formulas, range references, cross-sheet references
- Full Excel error propagation (#REF!, #VALUE!, #DIV/0!, #N/A, etc.)

#### Advanced Query Features
- Range iteration: `range("A1:C10")`
- Column and row access: `column("A")`, `row(1)`
- Filtering: `rows(where:)`, `find(where:)`, `findAll(where:)`
- Data validations API
- Named ranges support
- Hyperlinks (external and internal)

#### Advanced Write Features
- Sheet protection with password support and granular options
- Workbook protection (structure and windows)
- Charts (read-only access to embedded charts)
- Pivot tables (read-only access)
- Tables (read-only access)
- Multiple sheet support
- Hyperlink creation

#### Performance Optimizations
- Lazy sheet loading (sheets parsed on-demand)
- Streaming iteration for memory-efficient processing of large files
- Efficient shared strings handling
- Style-based date detection

#### Architecture
- Modern Swift 6 with Sendable types and value semantics
- Typed throws for comprehensive error handling
- OPC (Open Packaging Conventions) layer for ZIP/package handling
- Modular parser architecture (Workbook, Worksheet, SharedStrings, Styles)
- Modular builder architecture (XML generation for .xlsx creation)

#### Documentation
- **Complete DocC documentation catalog** with Apple-style formatting
- **Getting Started guide** with installation and quick start examples
- **7 comprehensive articles**:
  - Architecture: System design and layer overview
  - Formula Engine: 467 functions organized in 12 categories
  - Performance Tuning: Memory optimization and benchmarking
  - Writing Workbooks: Complete guide to creating Excel files
  - Advanced Queries: Filtering, searching, and range operations
  - Error Handling: Error patterns and recovery strategies
  - Migration Guide: Transitioning from CoreXLSX to Cuneiform
- **2 interactive tutorials**:
  - Data Analysis: Reading and analyzing spreadsheet data
  - Report Generation: Creating multi-sheet reports with formulas
- **13 formula reference files** documenting all 467 functions:
  - Complete syntax, parameters, return values for every function
  - Practical Swift code examples
  - Excel compatibility notes
  - Implementation status indicators
- **Comprehensive API documentation** for all public types:
  - Workbook, Sheet, WorkbookWriter with 100+ code examples
  - CellValue, CellReference, CuneiformError with usage patterns
  - FormulaEvaluator, FormulaParser with detailed explanations
  - Topics organization for DocC navigation
- **5 open source governance files** (LICENSE, CONTRIBUTING, CODE_OF_CONDUCT, SECURITY, CHANGELOG)
- **GitHub Actions workflows** for CI and documentation deployment
- **Issue templates** for bug reports, feature requests, and documentation issues
- **Example projects** with complete source code:
  - DataAnalysis: Extract and compute statistics
  - ReportGeneration: Create structured reports

#### Testing
- 834 passing tests across 45 test files
- Comprehensive coverage of read, write, formula evaluation, and edge cases
- Swift Testing framework integration

### Known Limitations
- 14 functions are stubs (LAMBDA family, Cube/OLAP functions): return #N/A
- Charts, pivot tables, and tables are read-only
- No VBA/macro support (by design)
- No external data connections

[Unreleased]: https://github.com/jramos57/cuneiform/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/jramos57/cuneiform/releases/tag/v0.1.0
