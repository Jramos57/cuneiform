# Release Notes: Cuneiform v0.1.0

**Release Date**: January 28, 2025

Cuneiform is a pure Swift library for reading and writing Excel (.xlsx) files with comprehensive formula support. This initial release provides production-ready functionality for working with Office Open XML SpreadsheetML documents.

---

## 🎉 Highlights

### Complete Read/Write Support
- **Read any .xlsx file** with full cell value resolution
- **Create new workbooks** with multiple sheets and formulas
- **Lazy loading** for efficient memory usage with large files
- **Streaming iteration** for processing millions of rows

### 467 Excel Functions
- **97% full implementations** (453 functions with complete Excel compatibility)
- **12 categories**: Math, Stats, Text, DateTime, Financial, Logical, Lookup, Engineering, Database, Info, Compat, Web
- **Advanced features**: Array formulas, range references, cross-sheet references
- **Error propagation**: Full Excel error handling (#REF!, #VALUE!, etc.)

### Advanced Features
- **Sheet protection** with password and granular permissions
- **Workbook protection** for structure and windows
- **Hyperlinks** (external and internal)
- **Data validations**, **Named ranges**, **Merge cells**
- **Charts, Pivot tables, Tables** (read-only)
- **Comments and Rich text**

### Modern Swift
- **Swift 6** with Sendable types and value semantics
- **Typed throws** for comprehensive error handling
- **Cross-platform**: macOS, iOS, tvOS, watchOS, visionOS, Linux
- **Zero dependencies**: Pure Swift implementation

### Comprehensive Documentation
- **Complete DocC catalog** with 30+ documentation files
- **7 in-depth articles** covering architecture, performance, migration
- **2 interactive tutorials** for common workflows
- **467 functions documented** with syntax, examples, Excel links
- **100+ API examples** throughout the documentation

---

## 📦 Installation

### Swift Package Manager

```swift
dependencies: [
    .package(url: "https://github.com/jramos57/cuneiform.git", from: "0.1.0")
]
```

### Xcode

1. File → Add Package Dependencies
2. Enter: `https://github.com/jramos57/cuneiform.git`
3. Select version `0.1.0` or higher

---

## 🚀 Getting Started

### Read a Workbook

```swift
import Cuneiform

let workbook = try Workbook.open(url: fileURL)

// Access sheets
if let sheet = try workbook.sheet(named: "Sales") {
    // Read cells with type-safe values
    if let value = sheet.cell(at: "A1") {
        switch value {
        case .number(let num): print("Number: \(num)")
        case .text(let text): print("Text: \(text)")
        default: break
        }
    }
    
    // Iterate efficiently
    for row in sheet.rows() {
        // Process each row
    }
}
```

### Create a Workbook

```swift
var writer = WorkbookWriter()
let sheet = writer.addSheet(named: "Report")

writer.modifySheet(at: sheet) { s in
    s.writeText("Product", to: "A1")
    s.writeText("Sales", to: "B1")
    s.writeNumber(1500.0, to: "B2")
    s.writeFormula("SUM(B2:B10)", to: "B11")
}

try writer.save(to: outputURL)
```

---

## 📚 Documentation

Complete documentation is available at:  
**https://jramos57.github.io/cuneiform/documentation/cuneiform/**

### Key Resources

- **[Getting Started](https://jramos57.github.io/cuneiform/documentation/cuneiform/gettingstarted)** - Installation and basics
- **[Architecture](https://jramos57.github.io/cuneiform/documentation/cuneiform/architecture)** - System design
- **[Performance Tuning](https://jramos57.github.io/cuneiform/documentation/cuneiform/performancetuning)** - Optimization guide
- **[Formula Engine](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulaengine)** - 467 functions overview
- **[Formula Reference](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulareference)** - Complete function catalog
- **[Migration Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/migrationguide)** - From CoreXLSX

### Tutorials

- **[Data Analysis](https://jramos57.github.io/cuneiform/tutorials/cuneiform/dataanalysis)** - Reading and analyzing data
- **[Report Generation](https://jramos57.github.io/cuneiform/tutorials/cuneiform/reportgeneration)** - Creating reports

---

## ✨ What's New

### Core Reading API
- `Workbook.open(url:)` - Open .xlsx files with automatic lazy loading
- `workbook.sheet(named:)` / `sheet(at:)` - Access sheets by name or index
- `sheet.cell(at:)` - Get resolved cell values with type-safe enum
- `sheet.rows()` - Memory-efficient lazy iteration
- `sheet.range(_:)` - Access rectangular cell ranges
- `sheet.column(_:)` - Get entire columns efficiently

### Advanced Reading
- `sheet.find(where:)` / `findAll(where:)` - Search with predicates
- `sheet.rows(where:)` - Filter rows during iteration
- `sheet.formula(at:)` - Access cell formulas
- `sheet.richText(at:)` - Get formatted text with runs
- `sheet.validations(at:)` - Read data validations
- `sheet.hyperlinks(at:)` - Access hyperlinks
- `sheet.comments(at:)` - Read cell comments
- `workbook.definedNameRange(_:)` - Resolve named ranges

### Core Writing API
- `WorkbookWriter()` - Create new workbooks
- `addSheet(named:)` - Add sheets with custom names
- `writeText(_:to:)` - Write text values
- `writeNumber(_:to:)` - Write numeric values
- `writeBoolean(_:to:)` - Write boolean values
- `writeDate(_:to:)` - Write dates
- `writeFormula(_:to:)` - Write formulas with cached values

### Advanced Writing
- `protectSheet(password:options:)` - Password-protect sheets
- `protectWorkbook(password:options:)` - Protect workbook structure
- `addHyperlinkExternal(at:url:)` - Add web links
- `addHyperlinkInternal(at:location:)` - Add cross-sheet links
- `mergeCells(range:)` - Merge cell ranges
- `addDataValidation(range:)` - Add validation rules
- `addConditionalFormat(range:)` - Add conditional formatting
- `addComment(at:text:author:)` - Add cell comments

### Formula Engine (467 Functions)

**Mathematical (75+)**  
SUM, AVERAGE, ROUND, ABS, SQRT, POWER, SIN, COS, TAN, LOG, EXP, PI, RAND, GCD, LCM, FACT, COMBIN, PERMUT, MOD, QUOTIENT, CEILING, FLOOR, TRUNC, INT, SIGN, EVEN, ODD, MROUND, SUMPRODUCT, SUMSQ, PRODUCT, and more

**Statistical (100+)**  
MIN, MAX, COUNT, COUNTA, COUNTIF, MEDIAN, MODE, STDEV, VAR, PERCENTILE, QUARTILE, RANK, LARGE, SMALL, CORREL, PEARSON, SLOPE, INTERCEPT, FORECAST, NORM.DIST, NORM.INV, BINOM.DIST, POISSON.DIST, T.DIST, F.DIST, CHISQ.DIST, and more

**Text (40)**  
CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, FIND, SEARCH, SUBSTITUTE, REPLACE, TEXT, VALUE, CHAR, CODE, EXACT, REPT, TEXTJOIN, TEXTSPLIT, and more

**Date & Time (27)**  
TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, WEEKNUM, HOUR, MINUTE, SECOND, TIME, DAYS, EDATE, EOMONTH, NETWORKDAYS, WORKDAY, DATEDIF, YEARFRAC, and more

**Financial (55)**  
PMT, PV, FV, RATE, NPER, IRR, XIRR, NPV, XNPV, MIRR, DB, DDB, SLN, SYD, PRICE, YIELD, ACCRINT, DURATION, COUPDAYBS, COUPDAYS, and more

**Logical (11)**  
IF, IFS, AND, OR, NOT, XOR, IFERROR, IFNA, SWITCH, TRUE, FALSE

**Lookup & Reference (35)**  
VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, XMATCH, INDIRECT, OFFSET, CHOOSE, FILTER, SORT, UNIQUE, SEQUENCE, TRANSPOSE, and more

**Engineering (60+)**  
CONVERT, HEX2DEC, DEC2HEX, BIN2DEC, BITAND, BITOR, BITXOR, COMPLEX, IMREAL, IMAGINARY, IMABS, IMSUM, IMPRODUCT, BESSELI, BESSELJ, ERF, ERFC, and more

**Database (10)**  
DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DGET, DPRODUCT, DSTDEV, DVAR

**Information (28)**  
ISBLANK, ISERROR, ISERR, ISNA, ISTEXT, ISNUMBER, ISLOGICAL, TYPE, N, NA, CELL, ROW, COLUMN, ROWS, COLUMNS, ERROR.TYPE, and more

**Compatibility (23)**  
Legacy distribution functions, LAMBDA, LET, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY

**Web Service (8)**  
HYPERLINK, WEBSERVICE, ENCODEURL (most are stubs)

See the **[Formula Reference](https://jramos57.github.io/cuneiform/documentation/cuneiform/formulareference)** for complete documentation.

---

## 🎯 Performance

Typical benchmarks on modern hardware:

| Operation | Time |
|-----------|------|
| Read 1,000 rows | 34ms |
| Write 1,000 rows (10 cols) | 50ms |
| Round-trip 500 rows | 80ms |
| Find in 1,000 rows | <1ms |
| Range query 2,600 cells | 660ms |

See the **[Performance Tuning Guide](https://jramos57.github.io/cuneiform/documentation/cuneiform/performancetuning)** for optimization strategies.

---

## ⚠️ Known Limitations

### Formula Engine
- **14 functions are stubs** (return #N/A):
  - LAMBDA family: LAMBDA, LET, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY
  - OLAP functions: CUBEVALUE, CUBEMEMBER, CUBEMEMBERPROPERTY
  - Web functions: WEBSERVICE, FILTERXML, RTD
- Future releases will add implementations for these functions

### Advanced Excel Features (Read-Only)
- **Charts**: Can read chart metadata but cannot create/modify charts
- **Pivot Tables**: Can read pivot table structure but cannot create/modify
- **Excel Tables**: Can read table definitions but cannot create/modify tables
- Future releases may add write support for these features

### Not Supported (By Design)
- **VBA/Macros**: No support for Visual Basic or macro execution
- **External Data Connections**: No support for database connections, web queries, etc.
- **Embedded Objects**: No support for images, shapes, or OLE objects
- These features are outside the scope of Cuneiform

---

## 🐛 Bug Reports

Found an issue? Please report it on GitHub:  
**https://github.com/jramos57/cuneiform/issues**

Use the structured issue templates for:
- Bug reports
- Feature requests  
- Documentation issues

---

## 🙏 Acknowledgments

Cuneiform is built following the Office Open XML standard (ECMA-376 / ISO/IEC 29500).

Special thanks to the Swift community and the creators of the OOXML specification.

---

## 📄 License

Cuneiform is released under the MIT License.

Copyright (c) 2025 Jonathan Ramos

---

## 🔗 Links

- **Homepage**: https://github.com/jramos57/cuneiform
- **Documentation**: https://jramos57.github.io/cuneiform/documentation/cuneiform/
- **Issues**: https://github.com/jramos57/cuneiform/issues
- **Discussions**: https://github.com/jramos57/cuneiform/discussions
- **OOXML Standard**: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

---

Made with ❤️ using Swift
