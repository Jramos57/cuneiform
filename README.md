# Cuneiform

Pure Swift library for reading Office Open XML SpreadsheetML (.xlsx) files.

## Quick Start

```swift
import Cuneiform

// Open an .xlsx file
let workbook = try Workbook.open(url: URL(fileURLWithPath: "data.xlsx"))

// List all sheets
for sheet in workbook.sheets {
    print("Sheet: \(sheet.name)")
}

// Access a sheet by name
if let sheet = try workbook.sheet(named: "Sheet1") {
    // Get a cell value
    if let value = sheet.cell(at: "A1") {
        print("A1: \(value)")  // text, number, date, boolean, error, or empty
    }
    
    // Get a row
    let row = sheet.row(1)
    print("Row 1: \(row)")
}

// Or access by index
if let sheet = try workbook.sheet(at: 0) {
    // Resolved cell values with types
    let cellA1 = sheet.cell(at: "A1")  // CellValue (text, number, date, etc.)
}
```

## Writing Workbooks

```swift
import Cuneiform

// Create a new workbook
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

// Write cells
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Name", to: "A1")
    sheet.writeNumber(42, to: "B1")
    sheet.writeBoolean(true, to: "C1")
    sheet.writeFormula("B1*2", cachedValue: 84, to: "D1")
}

// Save to file
try writer.save(to: URL(fileURLWithPath: "output.xlsx"))
```

## Advanced Queries

```swift
// Range access
let range = sheet.range("A1:C10")
for (ref, value) in range {
    print("\(ref): \(value)")
}

// Column access
let columnA = sheet.column("A")
let columnB = sheet.column(at: 1) // 0-based index

// Row filtering
let nycRows = sheet.rows { cells in
    cells.contains { $0.value == .text("NYC") }
}

// Find cells
if let cell = sheet.find(where: { _, value in value == .number(100) }) {
    print("Found at \(cell.reference)")
}

let allMatches = sheet.findAll { _, value in
    if case .number(let n) = value { return n > 50 }
    return false
}
```

## Performance

### Streaming Large Files

For large spreadsheets, use lazy iteration to minimize memory usage:

```swift
// Streaming iteration (memory-efficient)
for row in sheet.rows() {
    for (ref, value) in row {
        process(ref, value)
    }
}
```

### Lazy Sheet Loading

Sheets are loaded on-demand—only when accessed via `sheet(named:)` or `sheet(at:)`. This defers parsing until needed, improving startup time for workbooks with many sheets.

### Benchmark Results

From test suite on typical hardware:
- Read 1,000 rows: ~34ms
- Write 1,000 rows (10 columns): ~50ms
- Round-trip 500 rows: ~80ms
- Find operations in 1,000 rows: <1ms
- Range query 2,600 cells: ~660ms

### Performance Tips

1. **Use streaming**: `sheet.rows()` for large files (lower memory)
2. **Batch operations**: Write many cells at once rather than saving repeatedly
3. **Limit queries**: Use `find()` instead of `findAll()` when you only need the first match
4. **Defer loading**: Access sheets only when needed

## Resolving Cell Values

The `CellValue` enum represents fully-resolved cell content:
- `.text(String)` – text and inline strings
- `.number(Double)` – numbers
- `.date(String)` – dates (ISO 8601 format; numeric conversion is caller's responsibility)
- `.boolean(Bool)` – booleans
- `.error(String)` – spreadsheet errors
- `.empty` – empty cells

Cell values are resolved using:
1. **SharedStrings** for shared string cell references
2. **Styles** for date detection (numeric values with date formats become `.date()`)
3. **RawCellValue** for direct cell types and values

## Build & Test

```bash
cd /Users/jonathan/Desktop/garden/cuneiform
swift build
swift test
```

**Status:** All 146 tests pass.

## Components

### Read API
- **Workbook** – High-level API to open .xlsx files, access sheets, and resolve cells
- **Sheet** – Worksheet wrapper with cell value resolution, formulas, and advanced queries
- **Parsers:**
  - `SharedStringsParser` – Parse `/xl/sharedStrings.xml`
  - `WorkbookParser` – Parse `/xl/workbook.xml` for sheet metadata
  - `WorksheetParser` – Parse `/xl/worksheets/sheet*.xml` for cell data and formulas
  - `StylesParser` – Parse `/xl/styles.xml` for number formats and date detection

### Write API
- **WorkbookWriter** – Create new .xlsx files with multiple sheets
- **Builders:**
  - `ContentTypesBuilder` – Generate `[Content_Types].xml`
  - `RelationshipsBuilder` – Generate `.rels` files
  - `SharedStringsBuilder` – Manage string deduplication
  - `WorkbookBuilder` – Build workbook XML
  - `WorksheetBuilder` – Build worksheet XML with cells and formulas
- **ZipWriter** – Create ZIP archives for .xlsx output

### Core
- **OPC** – Package layer: opening ZIP archives, reading parts, relationships, content types

## Migration Notes: Swift 6 Testing

Swift 6 includes built-in Swift Testing, replacing the external `swift-testing` package. This project currently keeps the external dependency to ensure the test suite runs across toolchains that don't yet expose the built-in `Testing` module.

- Current state: All tests pass. You may see deprecation warnings from `swift-testing`.
- Rationale: Removing the dependency caused `_TestingInternals` errors on this toolchain.

### Migrate when your toolchain supports built-in Testing

1. Edit `Package.swift`:
   - Remove the `.package(url: "https://github.com/swiftlang/swift-testing.git", ...)` entry.
   - Remove `.product(name: "Testing", package: "swift-testing")` from the `CuneiformTests` target dependencies.
   - If needed for early toolchains, add `swiftSettings: [.enableExperimentalFeature("Testing")]` to the test target.
2. Run:
   ```bash
   swift build
   swift test
   ```
3. If you encounter errors like `missing required module '_TestingInternals'`, revert the changes and retain the external dependency until the toolchain exposes the built-in module.

## Ergonomic Helpers

### Named Ranges

```swift
// Resolve a named range like "MyRange" → (sheet, range)
if let (sheet, range) = workbook.definedNameRange("MyRange") {
    for (ref, value) in range { print("\(ref): \(value)") }
}

// Or fetch the raw defined name
if let name = workbook.definedName("MyRange") {
    print(name.name, name.refersTo)
}
```

### Data Validations

```swift
// All validations that intersect a range
let v = sheet.validations(for: "A2:B2")

// Validations that apply to a single cell
let atB2 = sheet.validations(at: "B2")

// Example: check kinds present (e.g., .list, .whole)
let kinds = Set(v.map(\.kind))
print(kinds)
```

### Hyperlinks

```swift
// External hyperlink
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.addHyperlinkExternal(at: "B2",
                               url: "https://example.com",
                               display: "Example",
                               tooltip: "Open example.com")
}

### Sheet Protection

```swift
// Protect a sheet with optional password
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.protectSheet(password: "secret123")
}

// Use preset options: default, strict (locks all), readonly
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.protectSheet(password: "pwd", options: .strict)
}

// Custom protection options
writer.modifySheet(at: sheetIndex) { sheet in
    var options = SheetProtectionOptions()
    options.formatCells = false   // Prevent cell formatting
    options.insertRows = false    // Prevent row insertion
    options.deleteColumns = true  // Allow column deletion
    sheet.protectSheet(password: "pwd", options: options)
}

// Read protection state
let workbook = try Workbook.open(url: url)
let sheet = try workbook.sheet(at: 0)
if let protection = sheet.protection {
    print("Sheet is protected")
    if protection.passwordHash != nil { print("Password-protected") }
}
```

### Charts

```swift
// Access charts embedded in a worksheet
let sheet = try workbook.sheet(named: "Dashboard")

// Get all charts in the sheet
for chart in sheet.charts {
    print("Chart: \(chart.title ?? "Untitled")")
    print("  Type: \(chart.type)")
    print("  Series: \(chart.seriesCount)")
}

// Charts are read-side only (parsing from `/xl/charts/chartN.xml`).
// Chart data includes:
// - type: The chart classification (column, bar, line, pie, area, etc.)
// - title: Optional chart title
// - seriesCount: Number of data series
// - dataRange: Optional reference to data source

// Internal hyperlink (to a location within the workbook)
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.addHyperlinkInternal(at: "C3",
                               location: "Sheet2!A1",
                               display: "Go to Sheet2 A1",
                               tooltip: "Jump to cell")
}
```

## Documentation

- 📖 [Performance Tuning Guide](Documentation/PERFORMANCE.md) - Optimize for your use case
- 🔄 [Migration Guide](Documentation/MIGRATION.md) - Migrate from other libraries
- 💡 [Data Analysis Example](Examples/DataAnalysis/) - Extract and analyze data
- 📊 [Report Generation Example](Examples/ReportGeneration/) - Create structured reports

## Notes

- Requires Swift 6 toolchain.
- macOS 13+ targets; iOS/tvOS/watchOS/visionOS supported via Swift Package.
- Built with modern idiomatic Swift: Sendable types, value semantics, typed throws, comprehensive error handling.
