# Cuneiform

[![Swift](https://img.shields.io/badge/Swift-6.0-orange.svg)](https://swift.org)
[![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20iOS%20%7C%20tvOS%20%7C%20watchOS%20%7C%20visionOS%20%7C%20Linux-lightgrey.svg)](https://github.com/jramos57/cuneiform)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![CI](https://github.com/jramos57/cuneiform/workflows/CI/badge.svg)](https://github.com/jramos57/cuneiform/actions)
[![Documentation](https://img.shields.io/badge/docs-DocC-blue.svg)](https://jramos57.github.io/cuneiform/documentation/cuneiform/)

Pure Swift library for reading and writing Office Open XML SpreadsheetML (.xlsx) files with comprehensive formula support.

---

## Features

✅ **Read & Write** - Open existing workbooks and create new ones  
✅ **467 Excel Functions** - Comprehensive formula engine (97% full implementations)  
✅ **Advanced Queries** - Built-in filtering, searching, and range operations  
✅ **High Performance** - Lazy loading and memory-efficient streaming  
✅ **Type Safety** - Swift 6 with Sendable types and typed throws  
✅ **Cross-Platform** - macOS, iOS, tvOS, watchOS, visionOS, and Linux  
✅ **Zero Dependencies** - Pure Swift with no external libraries  
✅ **Well-Tested** - 834 passing tests ensuring reliability  

---

## Quick Start

### Reading Workbooks

```swift
import Cuneiform

// Open an .xlsx file
let workbook = try Workbook.open(url: fileURL)

// Access a sheet and read cells
if let sheet = try workbook.sheet(named: "Sheet1") {
    // Get a cell value (type-safe enum)
    if let value = sheet.cell(at: "A1") {
        switch value {
        case .text(let text): print("Text: \(text)")
        case .number(let num): print("Number: \(num)")
        case .date(let date): print("Date: \(date)")
        case .boolean(let bool): print("Boolean: \(bool)")
        case .error(let err): print("Error: \(err)")
        case .empty: print("Empty cell")
        }
    }
    
    // Iterate rows efficiently
    for row in sheet.rows() {
        for (ref, value) in row {
            print("\(ref): \(value)")
        }
    }
}
```

### Writing Workbooks

```swift
import Cuneiform

// Create a new workbook
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

// Write cells with various types
writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Name", to: "A1")
    sheet.writeNumber(42, to: "B1")
    sheet.writeBoolean(true, to: "C1")
    sheet.writeFormula("SUM(B1:B10)", to: "B11")
    sheet.writeDate(Date(), to: "D1")
}

// Save to disk
try writer.save(to: outputURL)
```

### Advanced Queries

```swift
// Filter rows by condition
let activeRows = sheet.rows { cells in
    cells.contains { $0.value == .text("Active") }
}

// Find cells matching criteria
let highValues = sheet.findAll { ref, value in
    if case .number(let num) = value {
        return num > 100
    }
    return false
}

// Access entire columns
let columnB = sheet.column("B")
let total = columnB.reduce(0.0) { sum, (_, value) in
    if case .number(let num) = value {
        return sum + num
    }
    return sum
}

// Work with ranges
let range = sheet.range("A1:C10")
for (ref, value) in range {
    // Process each cell
}
```

### Formula Evaluation

```swift
// Evaluate formulas programmatically
let evaluator = FormulaEvaluator(workbook: workbook)

// Simple calculations
let result = try evaluator.evaluate("=SUM(1, 2, 3)")  // 6.0

// Reference cells
let total = try evaluator.evaluate("=SUM(A1:A10)")

// Complex formulas with nested functions
let analysis = try evaluator.evaluate(
    "=IF(AVERAGE(A1:A10) > 50, \"High\", \"Low\")"
)
```

---

## Documentation

Comprehensive documentation is available:

- 📚 **[API Documentation](https://jramos57.github.io/cuneiform/documentation/cuneiform/)** - Complete API reference
- 🚀 **[Getting Started Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/gettingstarted)** - Installation and quick start
- 📖 **[Architecture Overview](https://jramos57.github.io/cuneiform/documentation/cuneiform/architecture)** - System design and layers
- ⚡️ **[Performance Tuning](https://jramos57.github.io/cuneiform/documentation/cuneiform/performancetuning)** - Optimization strategies
- ✍️ **[Writing Workbooks](https://jramos57.github.io/cuneiform/documentation/cuneiform/writingworkbooks)** - Creating Excel files
- 🔍 **[Advanced Queries](https://jramos57.github.io/cuneiform/documentation/cuneiform/advancedqueries)** - Filtering and searching
- 🧮 **[Formula Engine](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulaengine)** - 467 supported functions
- 📊 **[Formula Reference](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulareference)** - Complete function catalog
- ⚠️ **[Error Handling](https://jramos57.github.io/cuneiform/documentation/cuneiform/errorhandling)** - Error patterns and recovery
- 🔄 **[Migration Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/migrationguide)** - Migrate from CoreXLSX

### Tutorials

Step-by-step tutorials for common tasks:

- 📊 **[Data Analysis Tutorial](https://jramos57.github.io/cuneiform/tutorials/cuneiform/dataanalysis)** - Reading and analyzing data
- 📝 **[Report Generation Tutorial](https://jramos57.github.io/cuneiform/tutorials/cuneiform/reportgeneration)** - Creating multi-sheet reports

### Examples

Complete example projects:

- 💡 **[Data Analysis Example](Examples/DataAnalysis/)** - Extract and compute statistics
- 📈 **[Report Generation Example](Examples/ReportGeneration/)** - Create structured reports

---

## Installation

### Swift Package Manager

Add Cuneiform to your `Package.swift`:

```swift
dependencies: [
    .package(url: "https://github.com/jramos57/cuneiform.git", from: "0.1.0")
]
```

Then add it to your target:

```swift
targets: [
    .target(
        name: "YourTarget",
        dependencies: ["Cuneiform"]
    )
]
```

### Xcode

1. File → Add Package Dependencies
2. Enter: `https://github.com/jramos57/cuneiform.git`
3. Select version: `0.1.0` or higher

---

## Requirements

- Swift 6.0 or later
- macOS 13.0+ / iOS 16.0+ / tvOS 16.0+ / watchOS 9.0+ / visionOS 1.0+
- Linux (Ubuntu 20.04+ or similar)

---

## Formula Support

Cuneiform includes a comprehensive formula engine with **467 Excel-compatible functions** across 12 categories:

| Category | Functions | Examples |
|----------|-----------|----------|
| **Mathematical** | 75+ | SUM, AVERAGE, ROUND, SIN, COS, SQRT, POWER |
| **Statistical** | 100+ | MIN, MAX, MEDIAN, STDEV, PERCENTILE, CORREL |
| **Text** | 40 | LEFT, RIGHT, MID, CONCAT, FIND, SUBSTITUTE |
| **Date & Time** | 27 | TODAY, DATE, YEAR, MONTH, WEEKDAY, EOMONTH |
| **Financial** | 55 | PMT, PV, FV, IRR, NPV, PRICE, YIELD |
| **Logical** | 11 | IF, AND, OR, NOT, IFS, SWITCH, IFERROR |
| **Lookup** | 35 | VLOOKUP, XLOOKUP, INDEX, MATCH, FILTER, SORT |
| **Engineering** | 60+ | CONVERT, HEX2DEC, COMPLEX, BESSELI, ERF |
| **Database** | 10 | DSUM, DAVERAGE, DCOUNT, DMAX, DMIN |
| **Information** | 28 | ISBLANK, ISERROR, TYPE, CELL, INFO |
| **Compatibility** | 23 | Legacy functions + LAMBDA, LET, MAP, REDUCE |
| **Web Service** | 8 | HYPERLINK, WEBSERVICE, ENCODEURL |

**Implementation Status**: 97% full implementations, 2% partial, 1% stubs

See the [Formula Reference](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulareference) for complete documentation.

---

## Performance

### Benchmark Results

Typical performance on modern hardware:

| Operation | Performance |
|-----------|-------------|
| Read 1,000 rows | ~34ms |
| Write 1,000 rows (10 columns) | ~50ms |
| Round-trip 500 rows | ~80ms |
| Find operations (1,000 rows) | <1ms |
| Range query (2,600 cells) | ~660ms |

### Optimization Tips

1. **Use lazy iteration** - `sheet.rows()` for memory-efficient processing
2. **Batch writes** - Modify multiple cells before saving
3. **Limit queries** - Use `find()` instead of `findAll()` when appropriate
4. **Defer loading** - Access sheets only when needed (automatic lazy loading)

See the [Performance Tuning Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/performancetuning) for detailed optimization strategies.

---

## Advanced Features

### Sheet Protection

```swift
// Protect a sheet with custom permissions
writer.modifySheet(at: sheetIndex) { sheet in
    var options = SheetProtectionOptions()
    options.formatCells = false
    options.insertRows = false
    sheet.protectSheet(password: "secret", options: options)
}
```

### Hyperlinks

```swift
// External hyperlink
sheet.addHyperlinkExternal(
    at: "A1",
    url: "https://example.com",
    display: "Visit Site"
)

// Internal hyperlink
sheet.addHyperlinkInternal(
    at: "B1",
    location: "Sheet2!A1",
    display: "Go to Sheet2"
)
```

### Data Validations

```swift
// Read validations from a sheet
let validations = sheet.validations(at: "A2")
for validation in validations {
    print("Type: \(validation.kind)")
    print("Formula: \(validation.formula1 ?? "")")
}
```

### Named Ranges

```swift
// Access defined names
if let (sheet, range) = workbook.definedNameRange("SalesData") {
    for (ref, value) in range {
        print("\(ref): \(value)")
    }
}
```

### Merge Cells

```swift
// Merge a range of cells
sheet.mergeCells(range: "A1:C1")
```

### Comments

```swift
// Read comments from cells
let comments = sheet.comments(at: "A1")
for comment in comments {
    print("Author: \(comment.author)")
    print("Text: \(comment.text)")
}
```

### Charts

```swift
// Access chart data (read-only)
for chart in sheet.charts {
    print("Chart: \(chart.title ?? "Untitled")")
    print("Type: \(chart.type)")
}
```

---

## Architecture

Cuneiform is built with a clean, layered architecture:

### Layer 1: OPC Package
- ZIP archive management
- Part reading and relationships
- Content type resolution

### Layer 2: Parsers
- XML parsing for all SpreadsheetML parts
- Shared strings, styles, workbooks, worksheets
- Charts, pivot tables, comments

### Layer 3: Domain
- High-level `Workbook` and `Sheet` APIs
- Cell value resolution
- Query and filtering operations

### Layer 4: Builders
- XML generation for all parts
- Workbook and worksheet construction
- ZIP packaging for output

### Layer 5: Formula Engine
- Formula parsing (tokenization + AST)
- Expression evaluation (467 functions)
- Cell reference resolution

See the [Architecture Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/architecture) for detailed information.

---

## Contributing

Bug reports are welcome! Please use the [issue tracker](https://github.com/jramos57/cuneiform/issues) to report bugs.

**Note**: This project does not accept pull requests. If you have a feature suggestion, please open a [feature request](https://github.com/jramos57/cuneiform/issues/new/choose) instead.

See [CONTRIBUTING.md](CONTRIBUTING.md) for more information.

---

## License

Cuneiform is released under the MIT License. See [LICENSE](LICENSE) for details.

---

## Credits

Created by [Jonathan Ramos](https://github.com/jramos57)

---

## Resources

- **Documentation**: https://jramos57.github.io/cuneiform/documentation/cuneiform/
- **GitHub**: https://github.com/jramos57/cuneiform
- **Issues**: https://github.com/jramos57/cuneiform/issues
- **Discussions**: https://github.com/jramos57/cuneiform/discussions
- **OOXML Standard**: [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)

---

Made with ❤️ using Swift
