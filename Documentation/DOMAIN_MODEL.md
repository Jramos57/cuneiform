# Cuneiform Domain Model

This document defines the core entities, protocols, and relationships for the Cuneiform SpreadsheetML library.

## Core Entity Hierarchy

```
Package (ZIP Container)
└── Workbook
    ├── Worksheets[]
    │   └── Cells (sparse grid)
    │       ├── Value (number, string, boolean, error, formula)
    │       └── Style reference
    ├── SharedStringTable
    ├── StyleSheet
    │   ├── Fonts[]
    │   ├── Fills[]
    │   ├── Borders[]
    │   ├── NumberFormats[]
    │   └── CellFormats[]
    ├── Theme
    ├── DefinedNames[]
    └── Tables[]
```

---

## Protocol Definitions

### Package Layer

```swift
/// Represents an Open Packaging Conventions container
protocol OPCPackage {
    /// All parts in the package
    var parts: [PartPath: PackagePart] { get }

    /// Content types declaration
    var contentTypes: ContentTypes { get }

    /// Root relationships
    var relationships: Relationships { get }

    /// Extract a part by path
    func part(at path: PartPath) -> PackagePart?

    /// Get relationships for a specific part
    func relationships(for part: PartPath) -> Relationships?
}

/// A single part within an OPC package
protocol PackagePart {
    var path: PartPath { get }
    var contentType: ContentType { get }
    var data: Data { get }
}

/// Type-safe part path
struct PartPath: Hashable, ExpressibleByStringLiteral {
    let value: String

    static let contentTypes = PartPath("[Content_Types].xml")
    static let rootRelationships = PartPath("_rels/.rels")
    static let workbook = PartPath("xl/workbook.xml")
}
```

### Workbook Layer

```swift
/// The root spreadsheet document
protocol Workbook {
    /// All worksheets in order
    var worksheets: [Worksheet] { get }

    /// Shared string table (may be nil if all strings are inline)
    var sharedStrings: SharedStringTable? { get }

    /// Style definitions
    var styles: StyleSheet? { get }

    /// Named ranges and definitions
    var definedNames: [DefinedName] { get }

    /// Workbook-level properties
    var properties: WorkbookProperties { get }

    /// Get worksheet by name
    func worksheet(named: String) -> Worksheet?

    /// Get worksheet by index (0-based)
    func worksheet(at index: Int) -> Worksheet?
}

struct WorkbookProperties {
    var activeSheet: Int?
    var date1904: Bool  // Mac Excel date system
    var calcMode: CalculationMode
}

enum CalculationMode {
    case manual
    case automatic
    case automaticExceptTables
}
```

### Worksheet Layer

```swift
/// A single worksheet within a workbook
protocol Worksheet {
    /// Sheet name (displayed on tab)
    var name: String { get }

    /// Sheet index within workbook
    var index: Int { get }

    /// Used range (smallest rectangle containing all data)
    var dimension: CellRange? { get }

    /// Access cell at reference
    subscript(ref: CellReference) -> Cell? { get }

    /// Access cell at column/row (0-based)
    subscript(column: Int, row: Int) -> Cell? { get }

    /// Iterate over all non-empty rows
    var rows: [Row] { get }

    /// Column definitions (widths, hidden, etc.)
    var columns: [ColumnDefinition] { get }

    /// Merged cell regions
    var mergedCells: [CellRange] { get }

    /// Sheet visibility
    var visibility: SheetVisibility { get }
}

struct Row {
    let index: Int          // 0-based row index
    let cells: [Cell]       // Non-empty cells in this row
    let height: Double?     // Custom height if set
    let hidden: Bool
    let outlineLevel: Int
}

struct ColumnDefinition {
    let range: ClosedRange<Int>  // Column indices this applies to
    let width: Double?
    let hidden: Bool
    let outlineLevel: Int
    let bestFit: Bool
}

enum SheetVisibility {
    case visible
    case hidden
    case veryHidden  // Only unhideable via VBA
}
```

### Cell Layer

```swift
/// A single cell in a worksheet
protocol Cell {
    /// Cell reference (e.g., "A1")
    var reference: CellReference { get }

    /// The cell's value
    var value: CellValue { get }

    /// Style index (into StyleSheet.cellFormats)
    var styleIndex: Int? { get }

    /// Formula if present
    var formula: Formula? { get }
}

/// All possible cell value types
enum CellValue: Equatable {
    case empty
    case number(Double)
    case string(String)         // Resolved from shared strings or inline
    case boolean(Bool)
    case error(CellError)
    case date(Date)             // Stored as number, interpreted via format

    /// The raw value as stored in XML
    var rawValue: String? { get }
}

enum CellError: String {
    case null = "#NULL!"
    case div0 = "#DIV/0!"
    case value = "#VALUE!"
    case ref = "#REF!"
    case name = "#NAME?"
    case num = "#NUM!"
    case na = "#N/A"
    case gettingData = "#GETTING_DATA"
}

/// Cell reference handling
struct CellReference: Hashable, CustomStringConvertible {
    let column: Int     // 0-based (A=0, B=1, ..., Z=25, AA=26)
    let row: Int        // 0-based

    /// Create from A1 notation
    init?(_ string: String)

    /// Create from indices
    init(column: Int, row: Int)

    /// A1 notation string
    var description: String { get }

    /// Column letter(s)
    var columnLetter: String { get }

    /// Row number (1-based for display)
    var rowNumber: Int { get }
}

/// A rectangular range of cells
struct CellRange: Hashable, CustomStringConvertible {
    let start: CellReference
    let end: CellReference

    /// Create from A1:B2 notation
    init?(_ string: String)

    /// All cell references in the range
    var cells: [CellReference] { get }

    /// Number of columns
    var columnCount: Int { get }

    /// Number of rows
    var rowCount: Int { get }
}
```

### Formula Layer (Phase 3)

```swift
/// Represents a cell formula
struct Formula {
    /// The formula text (without leading =)
    let text: String

    /// Formula type
    let type: FormulaType

    /// For array formulas: the range it applies to
    let arrayRange: CellRange?

    /// For shared formulas: the reference to master cell
    let sharedIndex: Int?
}

enum FormulaType {
    case normal
    case array       // CSE formula (Ctrl+Shift+Enter)
    case shared      // Shared formula (optimization)
    case dataTable   // What-if data table
}
```

### Shared Strings Layer

```swift
/// The shared string table
protocol SharedStringTable {
    /// Total count of string references in workbook
    var count: Int { get }

    /// Number of unique strings
    var uniqueCount: Int { get }

    /// Get string by index
    subscript(index: Int) -> SharedString? { get }

    /// Find index of string (for writing)
    func index(of string: String) -> Int?

    /// Add string and return its index
    mutating func add(_ string: String) -> Int
}

/// A string item in the shared string table
struct SharedString {
    /// Plain text content
    var text: String

    /// Rich text runs (if formatted)
    var richText: [RichTextRun]?

    /// Whether this has formatting
    var isRichText: Bool { richText != nil }
}

struct RichTextRun {
    let text: String
    let properties: RunProperties?
}

struct RunProperties {
    var bold: Bool?
    var italic: Bool?
    var underline: UnderlineStyle?
    var strikethrough: Bool?
    var fontSize: Double?
    var fontName: String?
    var color: Color?
}
```

### Style Layer

```swift
/// Complete style information for a workbook
protocol StyleSheet {
    var fonts: [Font] { get }
    var fills: [Fill] { get }
    var borders: [Border] { get }
    var numberFormats: [NumberFormat] { get }
    var cellFormats: [CellFormat] { get }
    var cellStyles: [CellStyle] { get }
}

/// A cell format combines references to other style elements
struct CellFormat {
    var fontIndex: Int
    var fillIndex: Int
    var borderIndex: Int
    var numberFormatId: Int
    var alignment: Alignment?
    var protection: Protection?
}

struct Font {
    var name: String
    var size: Double
    var bold: Bool
    var italic: Bool
    var underline: UnderlineStyle?
    var strikethrough: Bool
    var color: Color?
    var family: FontFamily?
}

struct Fill {
    var pattern: PatternFill?
    var gradient: GradientFill?
}

struct PatternFill {
    var type: PatternType
    var foregroundColor: Color?
    var backgroundColor: Color?
}

enum PatternType: String {
    case none
    case solid
    case gray125
    case gray0625
    // ... many more patterns
}

struct Border {
    var left: BorderEdge?
    var right: BorderEdge?
    var top: BorderEdge?
    var bottom: BorderEdge?
    var diagonal: BorderEdge?
    var diagonalUp: Bool
    var diagonalDown: Bool
}

struct BorderEdge {
    var style: BorderStyle
    var color: Color?
}

enum BorderStyle: String {
    case none
    case thin
    case medium
    case thick
    case dashed
    case dotted
    case double
    // ... more styles
}

struct NumberFormat {
    var id: Int
    var formatCode: String
}

/// Color representation
enum Color {
    case indexed(Int)           // Palette index
    case rgb(r: UInt8, g: UInt8, b: UInt8, a: UInt8)
    case theme(index: Int, tint: Double)
    case auto
}

struct Alignment {
    var horizontal: HorizontalAlignment?
    var vertical: VerticalAlignment?
    var wrapText: Bool
    var textRotation: Int?      // 0-180 or 255 for vertical
    var indent: Int?
    var shrinkToFit: Bool
}

enum HorizontalAlignment: String {
    case general, left, center, right, fill, justify, centerContinuous, distributed
}

enum VerticalAlignment: String {
    case top, center, bottom, justify, distributed
}
```

### Table Layer (Phase 4)

```swift
/// An Excel Table (ListObject)
protocol Table {
    var name: String { get }
    var displayName: String { get }
    var range: CellRange { get }
    var columns: [TableColumn] { get }
    var hasHeaderRow: Bool { get }
    var hasTotalRow: Bool { get }
    var style: TableStyle? { get }
}

struct TableColumn {
    let id: Int
    let name: String
    let totalsRowFunction: TotalsRowFunction?
    let totalsRowFormula: String?
}

enum TotalsRowFunction: String {
    case none, sum, min, max, average, count, countNums, stdDev, `var`, custom
}
```

---

## Reading vs Writing Protocols

### Read-Only Access (Phase 1)

```swift
/// Read-only workbook access
protocol WorkbookReader {
    /// Open workbook from file path
    static func open(at path: URL) throws -> Workbook

    /// Open workbook from data
    static func open(data: Data) throws -> Workbook
}
```

### Read-Write Access (Phase 2)

```swift
/// Mutable workbook for writing
protocol MutableWorkbook: Workbook {
    /// Add a new worksheet
    mutating func addWorksheet(named: String) -> MutableWorksheet

    /// Remove worksheet by index
    mutating func removeWorksheet(at index: Int)

    /// Save to file
    func save(to path: URL) throws

    /// Export to data
    func export() throws -> Data
}

protocol MutableWorksheet: Worksheet {
    /// Set cell value
    mutating func setValue(_ value: CellValue, at ref: CellReference)

    /// Set cell with style
    mutating func setCell(_ cell: Cell, at ref: CellReference)

    /// Set column width
    mutating func setColumnWidth(_ width: Double, for column: Int)

    /// Set row height
    mutating func setRowHeight(_ height: Double, for row: Int)

    /// Merge cells
    mutating func mergeCells(_ range: CellRange)
}
```

---

## Type Mappings: XML to Swift

| XML Type | Swift Type | Notes |
|----------|------------|-------|
| `xs:string` | `String` | |
| `xs:boolean` | `Bool` | XML: "true"/"1" or "false"/"0" |
| `xs:double` | `Double` | |
| `xs:int` | `Int` | |
| `xs:unsignedInt` | `UInt32` | |
| `ST_CellRef` | `CellReference` | A1 notation |
| `ST_Ref` | `CellRange` | A1:B2 notation |
| `CT_Color` | `Color` | enum with cases |
| `ST_Xstring` | `String` | Escaped string |

---

## Error Handling

```swift
/// Errors that can occur during parsing
enum CuneiformError: Error {
    // Package errors
    case invalidPackage(reason: String)
    case missingPart(path: PartPath)
    case invalidContentType(expected: String, found: String)

    // XML parsing errors
    case malformedXML(part: PartPath, detail: String)
    case missingRequiredElement(element: String, in: String)
    case invalidAttributeValue(attribute: String, value: String)

    // Data errors
    case invalidCellReference(String)
    case sharedStringIndexOutOfRange(index: Int)
    case styleIndexOutOfRange(index: Int)

    // File errors
    case fileNotFound(URL)
    case accessDenied(URL)
    case notAnXlsxFile
}
```

---

## Implementation Strategy

### Phase 1: Reading
1. Implement `OPCPackage` with ZIP extraction
2. Parse `[Content_Types].xml` and relationships
3. Parse `workbook.xml` to get structure
4. Implement lazy worksheet loading
5. Parse shared strings on demand
6. Parse styles on demand
7. Implement cell value resolution

### Phase 2: Writing
1. Implement mutable protocols
2. Generate XML from domain models
3. Build relationship graph
4. Generate content types
5. Create valid ZIP package

### Phase 3: Formulas
1. Formula parser (tokenizer + grammar)
2. Reference resolution
3. Dependency graph
4. Evaluation engine (subset of functions)

### Phase 4: Tables
1. Parse table definitions
2. Support structured references in formulas
3. Generate valid table XML

---

## Open Questions

1. **Lazy vs Eager Loading:** Should worksheets be parsed on access or upfront?
   - Recommendation: Lazy loading for large files

2. **String Handling:** Keep shared string indices or resolve immediately?
   - Recommendation: Resolve to strings, track original indices for round-trip

3. **Style Handling:** Full style resolution or just indices?
   - Recommendation: Both - indices for round-trip, resolved for convenience

4. **Date Handling:** Dates are stored as numbers - when to interpret?
   - Recommendation: Provide both raw number and interpreted date based on format

5. **Memory Model:** Value types (struct) vs reference types (class)?
   - Recommendation: Structs for data, protocols for abstraction
