# Architecture

Understanding Cuneiform's internal design and how it processes Office Open XML files.

## Overview

Cuneiform is built on a layered architecture that cleanly separates concerns:

1. **OPC Layer** - Open Packaging Conventions (ZIP and relationships)
2. **Parser Layer** - XML parsing and data extraction
3. **Domain Layer** - High-level APIs (Workbook, Sheet, FormulaEvaluator)
4. **Builder Layer** - XML generation for writing files

This design provides flexibility, testability, and clear separation between file format handling and business logic.

## Architectural Layers

```
┌─────────────────────────────────────────────────┐
│         High-Level API (Domain Layer)           │
│  Workbook, Sheet, WorkbookWriter, CellValue     │
└─────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────┐
│            Parser & Builder Layer                │
│  WorkbookParser, WorksheetParser, StylesParser   │
│  WorkbookBuilder, WorksheetBuilder               │
└─────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────┐
│              OPC Package Layer                   │
│  OPCPackage, Relationships, ContentTypes         │
│  ZIP archive handling, part reading              │
└─────────────────────────────────────────────────┘
                      ↓
┌─────────────────────────────────────────────────┐
│            Foundation Layer                      │
│  ZipArchive, XMLDocument, Data, URL              │
└─────────────────────────────────────────────────┘
```

## OPC Layer

### What is OPC?

Office Open XML documents are based on the **Open Packaging Conventions** (OPC) standard. An .xlsx file is actually a ZIP archive containing:

- **Parts**: XML files and other resources (e.g., `/xl/workbook.xml`, `/xl/worksheets/sheet1.xml`)
- **Relationships**: Connections between parts (stored in `_rels/` directories)
- **Content Types**: MIME type declarations (`[Content_Types].xml`)

### OPCPackage

The ``OPCPackage`` struct is the foundation of file reading:

```swift
public struct OPCPackage: Sendable {
    private let archiveData: Data
    private let archive: ZipArchive
    public let contentTypes: ContentTypes
    public let rootRelationships: Relationships
}
```

Key responsibilities:
- Opens ZIP archives from URLs or Data
- Provides access to parts (files within the archive)
- Manages relationships between parts
- Handles content type mapping

### Opening a Package

```swift
// From file
let package = try OPCPackage.open(url: fileURL)

// Check if parts exist
let hasStyles = package.partExists(.styles)

// Read raw part data
let workbookXML = try package.readPart(.workbook)

// Get relationships
let relationships = try package.relationships(for: .workbook)
```

### Part Paths

The ``PartPath`` enum defines standard Office Open XML part locations:

```swift
public enum PartPath {
    case workbook           // /xl/workbook.xml
    case worksheet(Int)     // /xl/worksheets/sheet1.xml
    case sharedStrings      // /xl/sharedStrings.xml
    case styles             // /xl/styles.xml
    // ... and more
}
```

This abstraction hides ZIP path details from higher layers.

### Relationships

Relationships connect parts together. They're defined in `_rels/` files:

```xml
<Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
```

The ``Relationships`` struct handles relationship lookups:

```swift
// Get all worksheet relationships
let worksheetRels = relationships[.worksheet]

// Get specific relationship by ID
if let rel = relationships["rId1"] {
    let targetPath = rel.resolveTarget(relativeTo: .workbook)
}
```

### Content Types

The ``ContentTypes`` struct maps part paths to MIME types, defined in `[Content_Types].xml`:

```xml
<Override
    PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
```

## Parser Layer

### Parser Responsibilities

Each parser is responsible for:
1. Parsing XML data into structured Swift types
2. Handling missing or optional elements
3. Providing sensible defaults
4. Throwing typed errors for malformed data

### WorkbookParser

Parses `/xl/workbook.xml` to extract:
- Sheet metadata (names, IDs, relationship IDs)
- Defined names (named ranges)
- Workbook-level protection settings

```swift
let workbookData = try package.readPart(.workbook)
let workbookInfo = try WorkbookParser.parse(data: workbookData)

for sheet in workbookInfo.sheets {
    print("\(sheet.name) - ID: \(sheet.sheetId)")
}
```

Located at: `Sources/Cuneiform/SpreadsheetML/WorkbookParser.swift:1`

### WorksheetParser

Parses `/xl/worksheets/sheetN.xml` to extract:
- Cell data (values, formulas, types)
- Merge cells
- Data validations
- Sheet protection
- Hyperlinks

```swift
let worksheetData = try package.readPart(.worksheet(1))
let worksheet = try WorksheetParser.parse(data: worksheetData)

for row in worksheet.rows {
    for cell in row.cells {
        print("\(cell.reference): \(cell.value)")
    }
}
```

Located at: `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift:1`

### SharedStringsParser

Parses `/xl/sharedStrings.xml` - a deduplication table for text:

```swift
let stringsData = try package.readPart(.sharedStrings)
let sharedStrings = try SharedStringsParser.parse(data: stringsData)

// Access by index
let text = sharedStrings[42]  // "Hello World"
```

Located at: `Sources/Cuneiform/SpreadsheetML/SharedStringsParser.swift:1`

### StylesParser

Parses `/xl/styles.xml` to extract:
- Number formats (for date detection)
- Cell formats (which number format to use)

This enables automatic date detection:

```swift
let stylesData = try package.readPart(.styles)
let styles = try StylesParser.parse(data: stylesData)

// Check if a cell format indicates a date
if styles.isDateFormat(cellFormatIndex: 5) {
    // Format cell value as date
}
```

Located at: `Sources/Cuneiform/SpreadsheetML/StylesParser.swift:1`

## Domain Layer

### Workbook

The ``Workbook`` struct is the primary entry point for reading .xlsx files:

```swift
public struct Workbook: Sendable {
    private let package: OPCPackage
    private let workbookInfo: WorkbookInfo
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    
    public var sheets: [SheetInfo]
    public var protection: WorkbookProtection?
}
```

Key features:
- **Lazy Loading**: Sheets are only parsed when accessed via `sheet(named:)` or `sheet(at:)`
- **Value Resolution**: Automatically resolves shared strings, styles, and dates
- **Error Handling**: Typed throws with ``CuneiformError``

Located at: `Sources/Cuneiform/SpreadsheetML/Workbook.swift:1`

### Sheet

The ``Sheet`` struct provides the querying API:

```swift
public struct Sheet: Sendable {
    public let name: String
    private let worksheet: Worksheet
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
}
```

Key features:
- **Cell Access**: `cell(at:)`, `row(_:)`, `column(_:)`, `range(_:)`
- **Queries**: `find(where:)`, `findAll(where:)`, `rows(where:)`
- **Formula Support**: `formula(at:)` returns raw formula strings
- **Validations**: `validations(at:)`, `validations(for:)`

Located at: `Sources/Cuneiform/SpreadsheetML/Sheet.swift:1`

### CellValue

The ``CellValue`` enum represents fully-resolved cell content:

```swift
public enum CellValue: Sendable, Equatable {
    case text(String)
    case number(Double)
    case date(String)     // ISO 8601 format
    case boolean(Bool)
    case error(String)
    case empty
}
```

This enum is the result of combining:
1. Raw cell value from worksheet XML
2. Shared string lookup (if applicable)
3. Style-based date detection (if applicable)

### Value Resolution Pipeline

```
RawCellValue (from XML)
        ↓
Is it a shared string reference?
    ↓ Yes: Look up in SharedStrings table
    ↓ No: Use raw value
        ↓
Is it a number with a date format?
    ↓ Yes: Convert to .date(String)
    ↓ No: Keep as .number(Double)
        ↓
Return CellValue
```

## Formula Engine

### FormulaParser

Parses formula strings into an Abstract Syntax Tree (AST):

```swift
let parser = FormulaParser()
let ast = try parser.parse("=SUM(A1:A10) + 5")

// AST structure:
// BinaryOp(+)
//   ├─ FunctionCall("SUM", args: [Range("A1:A10")])
//   └─ Literal(5.0)
```

Located at: `Sources/Cuneiform/SpreadsheetML/FormulaParser.swift:1`

### FormulaEvaluator

Evaluates formulas with 467 built-in functions:

```swift
var evaluator = FormulaEvaluator()
evaluator.setCell("A1", value: .number(10))
evaluator.setCell("A2", value: .number(20))

let result = evaluator.evaluate("=SUM(A1:A2)")  // 30.0
```

Architecture:
- **Function Registry**: Maps function names to implementations
- **Recursive Evaluation**: Evaluates AST nodes depth-first
- **Error Propagation**: Handles #REF!, #VALUE!, #DIV/0!, etc.
- **Type Coercion**: Automatic conversion between types where appropriate

Located at: `Sources/Cuneiform/SpreadsheetML/FormulaEvaluator.swift:1` (12,357 lines)

See <doc:FormulaEngine> for comprehensive details.

## Builder Layer

The builder layer creates XML for writing .xlsx files.

### WorkbookWriter

High-level API for creating workbooks:

```swift
var writer = WorkbookWriter()
let sheetIndex = writer.addSheet(named: "Data")

writer.modifySheet(at: sheetIndex) { sheet in
    sheet.writeText("Hello", to: "A1")
    sheet.writeNumber(42, to: "B1")
}

try writer.save(to: outputURL)
```

Located at: `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift:1`

### XML Builders

Each builder generates specific XML files:

- **WorkbookBuilder**: Creates `/xl/workbook.xml`
- **WorksheetBuilder**: Creates `/xl/worksheets/sheetN.xml`
- **SharedStringsBuilder**: Creates `/xl/sharedStrings.xml` with deduplication
- **ContentTypesBuilder**: Creates `[Content_Types].xml`
- **RelationshipsBuilder**: Creates `_rels/.rels` files

### Write Pipeline

```
1. User API calls (writeText, writeNumber, etc.)
        ↓
2. Builders accumulate data in memory
        ↓
3. save() triggers XML generation
        ↓
4. XML builders create part data
        ↓
5. ZipWriter creates ZIP archive
        ↓
6. Write to disk
```

## Memory Management

### Lazy Loading Strategy

Cuneiform defers parsing until data is actually accessed:

```swift
// Opening a workbook parses only:
// - [Content_Types].xml
// - _rels/.rels
// - /xl/workbook.xml
let workbook = try Workbook.open(url: fileURL)  // Fast!

// Sheet parsing happens here:
let sheet = try workbook.sheet(at: 0)  // Parses /xl/worksheets/sheet1.xml
```

This makes opening large workbooks with many sheets very fast.

### Streaming Iteration

The `rows()` iterator provides memory-efficient traversal:

```swift
// Only one row in memory at a time
for row in sheet.rows() {
    for (ref, value) in row {
        process(ref, value)
    }
    // Previous row data can be released
}
```

### Value Semantics

All public APIs use `struct` types with value semantics:

- **No retain cycles**: Automatic memory management
- **Thread-safe**: Sendable types work safely across concurrency boundaries
- **Predictable**: No hidden side effects from mutations

## Error Handling

### CuneiformError

All errors are represented by the ``CuneiformError`` enum:

```swift
public enum CuneiformError: Error, Sendable {
    case fileNotFound(path: String)
    case accessDenied(path: String)
    case invalidPackageStructure(reason: String)
    case missingPart(path: String)
    case malformedXML(part: String, detail: String)
    case invalidCellReference(String)
    case invalidFormulaExpression(String)
    case zipError(String)
}
```

All public APIs use typed throws:

```swift
public static func open(url: URL) throws(CuneiformError) -> Workbook
```

See <doc:ErrorHandling> for comprehensive error handling patterns.

## Design Principles

### 1. Separation of Concerns

Each layer has a single responsibility:
- **OPC**: ZIP and package structure
- **Parsers**: XML to domain types
- **Domain**: User-facing APIs
- **Builders**: Domain types to XML

### 2. Lazy Evaluation

Parse only what's needed:
- Sheets loaded on-demand
- Relationships cached per-part
- Streaming iteration available

### 3. Type Safety

Strong typing throughout:
- `CellValue` enum instead of raw strings
- `PartPath` enum instead of string paths
- Typed errors instead of generic `Error`

### 4. Value Semantics

Immutable, copyable types:
- All public APIs use `struct`
- `Sendable` conformance for concurrency
- No hidden mutable state

### 5. Performance by Default

Efficient out of the box:
- Minimal memory allocations
- Streaming-friendly APIs
- No unnecessary copying

## File Format Compliance

Cuneiform implements ECMA-376 / ISO/IEC 29500 (Office Open XML):

- **~90-92% compliance**: Covers most common spreadsheet features
- **Read support**: Full parsing of standard parts
- **Write support**: Generates valid .xlsx files
- **Formula engine**: 467 Excel functions (97% full implementations)

### Not Currently Supported

- Rich text formatting (partial)
- Charts (read-only)
- Pivot tables (read-only)
- VBA macros (by design)
- External data connections (by design)

## Testing Architecture

The test suite is organized by layer:

- **OPC Tests**: Package opening, part reading, relationships
- **Parser Tests**: XML parsing, error cases, edge cases
- **Domain Tests**: Workbook operations, queries, value resolution
- **Formula Tests**: 834 tests covering all 467 functions
- **Integration Tests**: Round-trip read/write validation

Located at: `Tests/CuneiformTests/`

## Performance Characteristics

### Time Complexity

| Operation | Complexity | Notes |
|-----------|-----------|-------|
| Open workbook | O(1) | Lazy loading |
| Access sheet | O(1) | Direct lookup |
| Cell lookup | O(log n) | Binary search on rows |
| Column iteration | O(n) | Linear scan |
| Range access | O(n × m) | Rectangle iteration |
| Find first | O(n) | Stops at first match |
| Find all | O(n) | Full traversal |

### Space Complexity

| Structure | Size | Notes |
|-----------|------|-------|
| Package | ~100KB | Content types + relationships |
| Sheet (parsed) | ~n cells | Row/cell data structures |
| Shared strings | ~m strings | Deduplicated text |
| Styles | ~50KB | Number formats + cell formats |

## Extension Points

Cuneiform is designed for extension:

### Custom Parsers

Add support for new part types:

```swift
extension OPCPackage {
    func parseCustomPart() throws -> CustomData {
        let data = try readPart(.custom("mypart.xml"))
        return try CustomParser.parse(data: data)
    }
}
```

### Custom Builders

Generate custom XML parts:

```swift
struct CustomBuilder {
    func build() -> Data {
        // Generate custom XML
    }
}
```

### Custom Functions

Extend the formula evaluator:

```swift
extension FormulaEvaluator {
    mutating func registerCustomFunction(_ name: String, implementation: @escaping ([CellValue]) -> CellValue) {
        // Register custom formula function
    }
}
```

## See Also

- <doc:FormulaEngine> - Deep dive into formula parsing and evaluation
- <doc:PerformanceTuning> - Optimization strategies
- <doc:ErrorHandling> - Error handling patterns
- ``OPCPackage`` - Package layer API reference
- ``Workbook`` - Domain layer API reference
