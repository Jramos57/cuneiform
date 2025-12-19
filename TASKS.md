# Cuneiform Parser Implementation Tasks

## Status: Phase 1-3 Complete ✓ / Phase 4 Planned (OOXML Toolkit)

**Phase 1 Implemented:** December 17, 2025
**Phase 2 Implemented:** December 18, 2025
**Phase 3 (3.1-3.4) Implemented:** December 18, 2025
**Current Verification:** All 168 tests pass (Dec 19, 2025)
**Current Compliance:** ~60% ISO/IEC 29500
**Target Compliance:** ~85% (after Phase 4)

Tasks 1-4 (SpreadsheetML parsers) are complete:
- SharedStringsParser - Parse shared string table with rich text support
- WorkbookParser - Extract sheet metadata and visibility states
- WorksheetParser - Parse cells with all value types, formulas, merged cells
- StylesParser - Number formats and date format detection

See [VERIFICATION_CHECKLIST.md](VERIFICATION_CHECKLIST.md) for detailed status.

## Phase 2 Ergonomics

Developer-facing helpers to make parsed data easier to use:
- [x] `Sheet.validations(for:)` and `Sheet.validations(at:)` to query data validations by A1 range or single cell.
- [x] `Workbook.definedName(_:)` and `Workbook.definedNameRange(_:)` to fetch named ranges and resolve `Sheet!$A$1:$B$10` into `(sheet, range)`.
 - [x] README updated with examples for these helpers (see Ergonomic Helpers section).
 - [x] Hyperlinks (read-side minimal): parse `<hyperlink>` entries (`ref`, `r:id`, `location`, `display`, `tooltip`) and expose via `Sheet.hyperlinks`.
 - [x] Hyperlinks (write-side minimal): emit `<hyperlinks>` in worksheet XML and add worksheet relationships `_rels/sheetN.xml.rels` with `Type=hyperlink` + `TargetMode="External"` for external URLs; APIs: `SheetWriter.addHyperlinkExternal(at:url:display:tooltip:)`, `SheetWriter.addHyperlinkInternal(at:location:display:tooltip:)`.
 - [x] Comments (read-side): resolve worksheet comments via relationships and expose as `Sheet.comments` / `comments(at:)`.
 - [x] Comments (write-side minimal): emit `/xl/commentsN.xml` with authors and text runs; add worksheet `Type=comments` relationship pointing to `../commentsN.xml`; API: `SheetWriter.addComment(at:text:author:)`.
 - [x] Comments (display via VML): emit `/xl/drawings/vmlDrawingN.vml` with legacy VML shapes anchored to each comment; add worksheet `Type=vmlDrawing` relationship and `<legacyDrawing r:id>` element in worksheet XML so Excel renders comment indicators and bubbles.
 - [x] Sheet protection (read-side): parse `<sheetProtection>` element exposing all protection flags and password hash via `Sheet.protection`.
 - [x] Sheet protection (write-side): emit `<sheetProtection>` element in worksheet XML with customizable flags; API: `SheetWriter.protectSheet(password:options:)` with `SheetProtectionOptions` struct supporting `.default`, `.strict`, and `.readonly` presets.

## Phase 3 Advanced Features

### Phase 3.1: Sheet Protection ✓
Completed (Dec 18, 2025):
- [x] Sheet protection (read-side): Parse `<sheetProtection>` with 16 flags + password hash
- [x] Sheet protection (write-side): Emit `<sheetProtection>` with `SheetWriter.protectSheet(password:options:)` API
- [x] 13 new tests; total 143 tests passing

### Phase 3.2: Charts ✓
- [x] Chart parser: Parse `/xl/charts/chart*.xml` to extract type, title, series count, data ranges
- [x] ChartData struct: type (enum: column, bar, line, pie, area, etc.), title, seriesCount, dataRange
- [x] Content types: Added `.chart` and `.drawing` to ContentType
- [x] Relationship types: Added `.chart` and `.drawing` to RelationshipType
- [x] 6 new tests (chart type parsing); 3 new integration tests; total 146 tests passing
- [x] Integrate charts into WorksheetParser to expose via `Sheet.charts` property
- [x] Add chart discovery via `/xl/drawings/drawing*.xml` relationships
- [x] **Status**: Phase 3.2 Complete (9 tests)

### Phase 3.3: Workbook Protection ✓
Completed (Dec 18, 2025):
- [x] Workbook protection (read-side): Parse `<workbookProtection sheet="1" windows="1" password="hash"/>` element
- [x] WorkbookProtection struct: structureProtected, windowsProtected, passwordHash fields
- [x] Workbook.protection property: Expose parsed protection data
- [x] 6 new parser tests; total 152 tests passing
- [x] Workbook protection (write-side): Emit `<workbookProtection>` element via WorkbookBuilder
- [x] WorkbookProtectionOptions struct: `.default`, `.structureOnly`, `.strict` presets
- [x] WorkbookWriter.protectWorkbook(password:options:) API
- [x] 8 new write/round-trip tests; total 160 tests passing
- [x] **Status**: Phase 3.3 Complete (14 tests: 6 read + 8 write)

### Phase 3.4: Pivot Tables ✓
Completed (Dec 18, 2025):
- [x] Pivot table parser: Parse `/xl/pivotTables/pivotTableN.xml` to extract name, cacheId, location, field counts
- [x] PivotTableData struct: Sendable, Equatable with all metadata fields
- [x] Relationship type: Added `.pivotTable` to RelationshipType
- [x] Workbook integration: Automatic discovery from worksheet relationships via `Workbook.pivotTables` property
- [x] 7 parser tests + 1 integration test validating discovery from real XLSX with 22 pivot tables; total 168 tests passing
- [x] **Status**: Phase 3.4 Complete (8 tests)

---

## Phase 4: OOXML Toolkit Compliance

**Goal:** Increase ISO/IEC 29500 compliance from ~60% to ~85%+ for production-ready OOXML toolkit.

### Phase 4.1: Full Styles Support (§18.8)
Expand `StylesParser` and `StylesBuilder` to expose complete cell formatting.

**Read-side (StylesParser.swift):**
- [ ] Parse `<fonts>` section: family, size, color (theme/rgb/indexed), bold, italic, underline, strike
- [ ] Parse `<fills>` section: pattern type, foreground/background colors
- [ ] Parse `<borders>` section: left/right/top/bottom/diagonal styles and colors
- [ ] Parse `<cellStyleXfs>` for named style definitions
- [ ] Parse alignment: horizontal, vertical, wrapText, textRotation, indent

**Write-side (StylesBuilder.swift):**
- [ ] Emit full `<font>` elements with all attributes
- [ ] Emit `<fill>` with patternFill and gradient support
- [ ] Emit `<border>` with style (thin, medium, thick, dashed, etc.) and color
- [ ] Emit `<alignment>` element in cellXfs
- [ ] Support theme color references (scheme colors)

**High-level API:**
- [ ] `CellStyle` struct: font, fill, border, alignment, numberFormat
- [ ] `Sheet.cellStyle(at:)` returns full formatting
- [ ] `SheetWriter.setCellStyle(_:at:)` for applying styles

**Tests:**
- [ ] Round-trip: write styled cells → read back → verify all properties
- [ ] Parse real Excel files with complex formatting
- [ ] Theme color resolution

---

### Phase 4.2: Tables/ListObjects (§18.5)
Excel Tables with headers, totals, and structured references.

**Read-side:**
- [ ] Parse `/xl/tables/tableN.xml`
- [ ] Extract: name, displayName, ref (range), headerRowCount, totalsRowCount
- [ ] Parse `<tableColumn>` elements: id, name, totalsRowFunction
- [ ] Parse `<autoFilter>` within table
- [ ] Parse table styles: tableStyleInfo element

**Write-side:**
- [ ] Emit `table.xml` with columns and range
- [ ] Add table relationship to worksheet
- [ ] Emit `<tableParts>` in worksheet XML
- [ ] Register table content type override

**High-level API:**
- [ ] `TableData` struct: name, range, columns, hasHeaders, hasTotals
- [ ] `Sheet.tables` property
- [ ] `SheetWriter.addTable(name:range:columns:)`

**Tests:**
- [ ] Create table, read back, verify structure
- [ ] Parse Excel-created tables
- [ ] Table with totals row formulas

---

### Phase 4.3: Conditional Formatting (§18.3.1)
Data bars, color scales, icon sets, and formula-based rules.

**Read-side:**
- [ ] Parse `<conditionalFormatting>` elements
- [ ] Extract sqref (affected ranges)
- [ ] Parse rule types: cellIs, colorScale, dataBar, iconSet, expression
- [ ] Parse operator and formula for cellIs rules
- [ ] Parse cfvo (conditional format value objects) for gradients

**Write-side:**
- [ ] Emit `<conditionalFormatting>` with rules
- [ ] Support highlight cells rules (greater than, less than, between, etc.)
- [ ] Support data bar generation
- [ ] Support color scale (2-color, 3-color)
- [ ] Support icon sets

**High-level API:**
- [ ] `ConditionalRule` enum with associated values per rule type
- [ ] `Sheet.conditionalFormats` property
- [ ] `SheetWriter.addConditionalFormat(range:rule:)`

**Tests:**
- [ ] Highlight cells greater than value
- [ ] Data bar with min/max
- [ ] 3-color scale
- [ ] Icon set (arrows, traffic lights)

---

### Phase 4.4: AutoFilter (§18.3.2.1)
Column filtering without full table.

**Read-side:**
- [ ] Parse `<autoFilter ref="A1:D100">` element
- [ ] Parse `<filterColumn>` with colId
- [ ] Parse filter types: filters (discrete values), customFilters, top10, colorFilter

**Write-side:**
- [ ] Emit `<autoFilter>` element in worksheet
- [ ] Emit `<filterColumn>` with filter criteria

**High-level API:**
- [ ] `AutoFilter` struct: range, columnFilters
- [ ] `Sheet.autoFilter` property
- [ ] `SheetWriter.setAutoFilter(range:)`

**Tests:**
- [ ] Set autofilter range
- [ ] Parse filtered worksheet
- [ ] Column with discrete value filter

---

### Phase 4.5: Rich Text (§18.4)
Full rich text support for cells and comments.

**Read-side:**
- [ ] Preserve `<r>` (run) elements in SharedStrings
- [ ] Parse `<rPr>` (run properties): font, size, color, bold, italic, underline
- [ ] Expose as `[TextRun]` array instead of plain String

**Write-side:**
- [ ] Emit rich text runs in sharedStrings.xml
- [ ] Emit rich text in comments

**High-level API:**
- [ ] `TextRun` struct: text, font, size, color, bold, italic, underline
- [ ] `RichText` = `[TextRun]`
- [ ] `CellValue.richText([TextRun])` case
- [ ] `Sheet.richText(at:)` returns formatted runs

**Tests:**
- [ ] Parse cell with multiple formatted runs
- [ ] Write rich text, read back, verify formatting
- [ ] Rich text in comments

---

### Phase 4.6: Shared Strings Optimization
Use shared strings table for write efficiency.

**Write-side:**
- [ ] Track unique strings during workbook write
- [ ] Generate optimized sharedStrings.xml
- [ ] Reference shared string indices in cells (`t="s"`)
- [ ] Option to force inline strings for small files

**Tests:**
- [ ] Large file with repeated strings uses shared table
- [ ] Shared string indices resolve correctly on read-back

---

### Phase 4.7: Page Setup & Print Areas (§18.3.1.63)
Print configuration for worksheets.

**Read-side:**
- [ ] Parse `<pageSetup>` element: orientation, paperSize, scale, fitToPage
- [ ] Parse `<pageMargins>`: left, right, top, bottom, header, footer
- [ ] Parse print area from defined names (`_xlnm.Print_Area`)
- [ ] Parse print titles (`_xlnm.Print_Titles`)

**Write-side:**
- [ ] Emit `<pageSetup>` element
- [ ] Emit `<pageMargins>` element
- [ ] Add print area/titles as defined names

**High-level API:**
- [ ] `PageSetup` struct: orientation, paperSize, margins, scale
- [ ] `Sheet.pageSetup` property
- [ ] `SheetWriter.setPageSetup(_:)`

---

### Phase 4 Priority Order

1. **4.1 Full Styles** - Most requested, enables formatting round-trip
2. **4.2 Tables** - Common Excel feature, structured data
3. **4.3 Conditional Formatting** - Visual data analysis
4. **4.4 AutoFilter** - Essential for data worksheets
5. **4.5 Rich Text** - Text formatting preservation
6. **4.6 Shared Strings** - Performance optimization
7. **4.7 Page Setup** - Print support

---

### Future Phases (Post-4.x)

- [ ] Sparklines (Excel 2010+ extension)
- [ ] Slicers (Excel 2010+ extension)
- [ ] DrawingML shapes and images
- [ ] Full chart generation (not just parsing)
- [ ] Pivot table generation
- [ ] External data connections
- [ ] Strict format support (ISO namespace)


## Swift Style Requirements

**Write modern, idiomatic Swift 6.** Use every trick in the book:

### Concurrency & Safety
- All types must be `Sendable`
- Use `sending` parameter modifier where appropriate
- Prefer value types (`struct`, `enum`) over `class`
- Only use `class` when reference semantics are required (like `XMLParserDelegate`)
- Mark classes `final` unless inheritance is intended

### Type System
- Use `some Protocol` (opaque types) for return types when hiding implementation
- Leverage generic constraints: `where Element: Sendable`
- Use `@frozen` on public enums if ABI stability matters
- Prefer `Int` over `Int64`/`Int32` unless specific size needed
- Use `String` not `NSString`

### Syntax & Idioms
```swift
// Use if-let shorthand (Swift 5.7+)
if let value { ... }  // NOT: if let value = value

// Use guard for early exits
guard let data = optionalData else { return nil }

// Use trailing closure syntax
array.map { $0.uppercased() }

// Use implicit returns in single-expression closures/properties
var isEmpty: Bool { count == 0 }

// Use keypath expressions
sheets.map(\.name)  // NOT: sheets.map { $0.name }

// Use first(where:) not filter().first
array.first { $0.isValid }
```

### Protocol Conformances
```swift
// Synthesize when possible
struct Foo: Hashable, Sendable { ... }  // Compiler generates ==, hash(into:)

// Use protocol extensions for defaults
extension Collection where Element == RawCell {
    func cell(at ref: String) -> RawCell? { ... }
}

// Conform to ExpressibleBy*Literal for ergonomics
extension CellReference: ExpressibleByStringLiteral {
    public init(stringLiteral value: String) {
        self.init(value)!  // Only if safe
    }
}

// Conform to CustomStringConvertible for debugging
extension CellReference: CustomStringConvertible {
    public var description: String { "\(column)\(row)" }
}
```

### Enums
```swift
// Use associated values
enum RawCellValue: Sendable {
    case number(Double)
    case sharedString(index: Int)
}

// Use raw values for parsing
enum SheetState: String, Sendable, CaseIterable {
    case visible
    case hidden
    case veryHidden
}

// Use static properties for well-known values
extension ContentType {
    static let worksheet = ContentType("application/...")
}
```

### Computed Properties & Subscripts
```swift
// Prefer computed properties over methods for simple accessors
var count: Int { strings.count }
var isEmpty: Bool { strings.isEmpty }

// Use subscripts for indexed/keyed access
subscript(index: Int) -> String? {
    guard index >= 0, index < strings.count else { return nil }
    return strings[index]
}

// Use labeled subscripts for clarity
subscript(column col: String, row row: Int) -> RawCell? { ... }
```

### Error Handling
```swift
// Extend existing error enum, don't create new types
// Add cases to CuneiformError if needed

// Use typed throws (Swift 6)
func parse(data: Data) throws(CuneiformError) -> Result { ... }

// Provide context in errors
throw CuneiformError.malformedXML(
    part: partPath,
    detail: "Expected <si> element, found <\(elementName)>"
)
```

### Documentation
```swift
/// A reference to a cell like "A1" or "BC123"
///
/// Cell references consist of a column (letters) and row (1-based number).
///
/// ```swift
/// let ref = CellReference("A1")
/// print(ref?.columnIndex)  // 0
/// print(ref?.row)          // 1
/// ```
public struct CellReference: Hashable, Sendable { ... }
```

### Avoid
- `NSArray`, `NSDictionary`, `NSString` (use Swift types)
- Force unwrapping `!` except in tests or truly invariant cases
- `var` when `let` suffices
- Stringly-typed APIs (use enums, typed wrappers)
- Mutable state where immutable works
- `Any` or `AnyObject` (prefer generics or protocols)
- Inheritance hierarchies (prefer composition)

### Performance Considerations
- Use `lazy` collections for deferred computation
- Use `reserveCapacity` for known-size collections
- Prefer `ContiguousArray` for homogeneous value types in hot paths
- Use `@inlinable` sparingly and only for genuine hot paths

---

## Context

You are implementing SpreadsheetML parsers for the Cuneiform library. The OPC package layer is complete - you can open .xlsx files and read raw XML parts. Your job is to parse the SpreadsheetML XML into Swift domain types.

**Existing Infrastructure:**
- `OPCPackage` - Opens .xlsx, reads parts as `Data` or `String`
- `PartPath` - Well-known paths like `.workbook`, `.sharedStrings`, `.styles`
- `ContentType` - MIME types for SpreadsheetML parts
- `Relationship` / `Relationships` - Part relationships with type-based lookup
- `CuneiformError` - Error enum (add new cases as needed)

**Code Location:** `Sources/Cuneiform/`
**Tests Location:** `Tests/CuneiformTests/`

**Build/Test Commands:**
```bash
cd /Users/jonathan/Desktop/garden/cuneiform
swift build
swift test
```

---

## Task 1: SharedStrings Parser

**Priority:** HIGH (other parsers depend on this)

**Purpose:** Parse `/xl/sharedStrings.xml` to resolve string cell values. Excel stores strings in a shared table and cells reference them by index.

### Input XML Structure
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="5" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si>
    <r><t>Rich </t></r>
    <r><rPr><b/></rPr><t>Text</t></r>
  </si>
</sst>
```

### Deliverables

**File:** `Sources/Cuneiform/SpreadsheetML/SharedStringsParser.swift`

```swift
/// Parsed shared strings table
public struct SharedStrings: Sendable {
    /// All strings in order (index = position)
    public let strings: [String]

    /// Get string by index
    public subscript(index: Int) -> String? { get }

    /// Number of strings
    public var count: Int { get }

    /// Empty table
    public static let empty: SharedStrings
}

/// Parser for sharedStrings.xml
public enum SharedStringsParser {
    /// Parse shared strings from XML data
    public static func parse(data: Data) throws -> SharedStrings
}
```

### Requirements

1. Parse `<si>` elements in order - index is position in file
2. Handle simple strings: `<si><t>text</t></si>`
3. Handle rich text: `<si><r><t>part1</t></r><r><t>part2</t></r></si>` → concatenate all `<t>` elements
4. Handle whitespace preservation: `<t xml:space="preserve"> text </t>`
5. Handle empty strings: `<si><t/></si>` or `<si></si>`
6. Return `SharedStrings.empty` if file doesn't exist (optional part)

### Verification Tests

Create `Tests/CuneiformTests/SharedStringsTests.swift`:

```swift
@Suite struct SharedStringsTests {
    @Test func parseSimpleStrings()
    // Input: <sst><si><t>A</t></si><si><t>B</t></si></sst>
    // Expect: strings[0] == "A", strings[1] == "B", count == 2

    @Test func parseRichText()
    // Input: <sst><si><r><t>Hello </t></r><r><t>World</t></r></si></sst>
    // Expect: strings[0] == "Hello World"

    @Test func parsePreservedWhitespace()
    // Input with xml:space="preserve"
    // Expect: whitespace preserved

    @Test func parseEmptyString()
    // Input: <sst><si><t/></si></sst>
    // Expect: strings[0] == ""

    @Test func subscriptOutOfBounds()
    // Expect: returns nil, doesn't crash
}
```

---

## Task 2: Workbook Parser

**Priority:** HIGH (entry point for sheet discovery)

**Purpose:** Parse `/xl/workbook.xml` to discover sheets and their relationship IDs.

### Input XML Structure
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
    <sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
</workbook>
```

### Deliverables

**File:** `Sources/Cuneiform/SpreadsheetML/WorkbookParser.swift`

```swift
/// Sheet metadata from workbook.xml
public struct SheetInfo: Sendable {
    /// Display name
    public let name: String

    /// Internal sheet ID
    public let sheetId: Int

    /// Relationship ID (used to find actual sheet part)
    public let relationshipId: String

    /// Visibility state
    public let state: SheetState
}

/// Sheet visibility
public enum SheetState: String, Sendable {
    case visible
    case hidden
    case veryHidden
}

/// Parsed workbook metadata
public struct WorkbookInfo: Sendable {
    /// All sheets in workbook order
    public let sheets: [SheetInfo]

    /// Get sheet by name
    public func sheet(named: String) -> SheetInfo?
}

/// Parser for workbook.xml
public enum WorkbookParser {
    public static func parse(data: Data) throws -> WorkbookInfo
}
```

### Requirements

1. Parse all `<sheet>` elements in order
2. Extract: `name`, `sheetId`, `r:id` (relationship ID)
3. Handle `state` attribute: missing = visible, "hidden", "veryHidden"
4. Throw `CuneiformError.missingRequiredElement` if name or r:id missing
5. Handle namespace prefix `r:id` correctly

### Verification Tests

Create `Tests/CuneiformTests/WorkbookParserTests.swift`:

```swift
@Suite struct WorkbookParserTests {
    @Test func parseSheets()
    // Expect: correct count, names, sheetIds, relationshipIds

    @Test func parseSheetStates()
    // Expect: visible (default), hidden, veryHidden parsed correctly

    @Test func sheetByName()
    // Expect: lookup works, returns nil for missing

    @Test func missingNameThrows()
    // Expect: CuneiformError.missingRequiredElement

    @Test func emptyWorkbook()
    // Input: <workbook><sheets/></workbook>
    // Expect: empty sheets array, no crash
}
```

---

## Task 3: Worksheet Parser

**Priority:** HIGH (core data extraction)

**Purpose:** Parse `/xl/worksheets/sheet*.xml` to extract cell data.

### Input XML Structure
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C3"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="n"><v>42</v></c>
      <c r="C1"><v>3.14</v></c>
    </row>
    <row r="2">
      <c r="A2" t="b"><v>1</v></c>
      <c r="B2" t="str"><v>inline</v></c>
      <c r="C2" t="e"><v>#DIV/0!</v></c>
    </row>
    <row r="3">
      <c r="A3" s="1"><v>44562</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A5:C5"/>
  </mergeCells>
</worksheet>
```

### Cell Types (`t` attribute)
- `s` = shared string (value is index into SharedStrings)
- `n` or missing = number
- `b` = boolean (0 or 1)
- `str` = inline string
- `e` = error
- `d` = date (ISO 8601 format)

### Deliverables

**File:** `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift`

```swift
/// A cell reference like "A1" or "BC123"
public struct CellReference: Hashable, Sendable {
    public let column: String  // "A", "B", ... "AA", "AB", etc.
    public let row: Int        // 1-based

    public init?(_ reference: String)  // Parse "A1" format
    public init(column: String, row: Int)

    /// Column as 0-based index (A=0, B=1, ..., Z=25, AA=26)
    public var columnIndex: Int { get }
}

/// Raw cell value before type resolution
public enum RawCellValue: Sendable {
    case sharedString(index: Int)
    case number(Double)
    case boolean(Bool)
    case inlineString(String)
    case error(String)
    case date(String)  // ISO format, conversion happens later
    case empty
}

/// A parsed cell
public struct RawCell: Sendable {
    public let reference: CellReference
    public let value: RawCellValue
    public let styleIndex: Int?  // References styles.xml
}

/// A parsed row
public struct RawRow: Sendable {
    public let index: Int  // 1-based row number
    public let cells: [RawCell]
}

/// Parsed worksheet data
public struct WorksheetData: Sendable {
    /// Declared dimension (may be inaccurate)
    public let dimension: String?

    /// All rows with data
    public let rows: [RawRow]

    /// Merged cell ranges
    public let mergedCells: [String]  // "A1:C3" format

    /// Get cell by reference
    public func cell(at ref: CellReference) -> RawCell?
    public func cell(at ref: String) -> RawCell?
}

/// Parser for worksheet XML
public enum WorksheetParser {
    public static func parse(data: Data) throws -> WorksheetData
}
```

### Requirements

1. Parse `<dimension ref="..."/>` if present
2. Parse all `<row>` elements, extract `r` attribute for row number
3. Parse all `<c>` (cell) elements within rows
4. Determine cell type from `t` attribute (default = number)
5. Extract `<v>` (value) element content
6. Extract `s` attribute for style index
7. Parse `<mergeCells>` section
8. Handle missing rows/cells gracefully (sparse data)
9. Cell reference parsing must handle columns AA, AB, ... ZZ, AAA, etc.

### Verification Tests

Create `Tests/CuneiformTests/WorksheetParserTests.swift`:

```swift
@Suite struct WorksheetParserTests {
    @Test func parseCellReference()
    // "A1" -> column: "A", row: 1, columnIndex: 0
    // "Z1" -> columnIndex: 25
    // "AA1" -> columnIndex: 26
    // "AZ1" -> columnIndex: 51

    @Test func parseSharedStringCell()
    // <c r="A1" t="s"><v>0</v></c>
    // Expect: RawCellValue.sharedString(index: 0)

    @Test func parseNumberCell()
    // <c r="A1"><v>42.5</v></c>
    // Expect: RawCellValue.number(42.5)

    @Test func parseBooleanCell()
    // <c r="A1" t="b"><v>1</v></c>
    // Expect: RawCellValue.boolean(true)

    @Test func parseInlineStringCell()
    // <c r="A1" t="str"><v>hello</v></c>
    // Expect: RawCellValue.inlineString("hello")

    @Test func parseErrorCell()
    // <c r="A1" t="e"><v>#DIV/0!</v></c>
    // Expect: RawCellValue.error("#DIV/0!")

    @Test func parseEmptyCell()
    // <c r="A1"/>
    // Expect: RawCellValue.empty

    @Test func parseMergedCells()
    // Expect: mergedCells array populated

    @Test func cellLookup()
    // Expect: cell(at: "A1") works, returns nil for missing

    @Test func sparseData()
    // Rows 1 and 5 only, cells A1 and C1 only
    // Expect: no crash, correct lookup
}
```

---

## Task 4: Styles Parser (Minimal)

**Priority:** MEDIUM (needed for date detection)

**Purpose:** Parse `/xl/styles.xml` to extract number formats. Full style parsing is complex - start minimal.

### Why This Matters

Excel stores dates as numbers (days since 1900). The only way to know a number is a date is by checking its number format. Format codes like `"mm/dd/yyyy"` or built-in format IDs 14-22 indicate dates.

### Input XML Structure (simplified)
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <cellXfs count="2">
    <xf numFmtId="0"/>
    <xf numFmtId="164"/>
  </cellXfs>
</styleSheet>
```

### Built-in Date Format IDs
IDs 14-22 are built-in date/time formats:
- 14: `mm-dd-yy`
- 15: `d-mmm-yy`
- 16: `d-mmm`
- 17: `mmm-yy`
- 18: `h:mm AM/PM`
- 19: `h:mm:ss AM/PM`
- 20: `h:mm`
- 21: `h:mm:ss`
- 22: `m/d/yy h:mm`

### Deliverables

**File:** `Sources/Cuneiform/SpreadsheetML/StylesParser.swift`

```swift
/// Number format information
public struct NumberFormat: Sendable {
    public let id: Int
    public let formatCode: String?  // nil for built-in formats

    /// Is this a date/time format?
    public var isDateFormat: Bool { get }
}

/// Minimal styles information
public struct StylesInfo: Sendable {
    /// Custom number formats (id -> format code)
    public let numberFormats: [Int: String]

    /// Cell format records (index = style index from cell `s` attribute)
    /// Each entry is the numFmtId used by that style
    public let cellFormats: [Int]  // numFmtId for each xf entry

    /// Get number format for a style index
    public func numberFormat(forStyleIndex index: Int) -> NumberFormat?

    /// Is the style index a date format?
    public func isDateFormat(styleIndex: Int) -> Bool

    /// Empty/default styles
    public static let empty: StylesInfo
}

/// Parser for styles.xml
public enum StylesParser {
    public static func parse(data: Data) throws -> StylesInfo
}
```

### Requirements

1. Parse `<numFmts>` section - map numFmtId to formatCode
2. Parse `<cellXfs>` section - extract numFmtId for each `<xf>` element
3. Implement `isDateFormat` logic:
   - Built-in IDs 14-22 are dates
   - Custom formats containing `y`, `m`, `d`, `h`, `s` patterns (but not `[Red]`, `#,##0`, etc.)
4. Return `StylesInfo.empty` if styles.xml doesn't exist
5. Handle missing numFmtId (defaults to 0 = General)

### Date Format Detection Heuristic

A format code is likely a date if it contains date/time tokens but isn't a pure number format:
- Contains: `y`, `m`, `d`, `h`, `s` (case insensitive)
- Does NOT start with: `#`, `0`, `?` (number format indicators)
- Does NOT contain only: `@` (text format)

This is a heuristic - exact detection is complex.

### Verification Tests

Create `Tests/CuneiformTests/StylesParserTests.swift`:

```swift
@Suite struct StylesParserTests {
    @Test func parseCustomNumberFormats()
    // Expect: numberFormats dictionary populated

    @Test func parseCellFormats()
    // Expect: cellFormats array matches xf order

    @Test func builtInDateFormats()
    // Style with numFmtId 14
    // Expect: isDateFormat(styleIndex:) returns true

    @Test func customDateFormat()
    // numFmtId 164 with formatCode "yyyy-mm-dd"
    // Expect: isDateFormat returns true

    @Test func numberFormatNotDate()
    // formatCode "#,##0.00"
    // Expect: isDateFormat returns false

    @Test func missingStylesFile()
    // Expect: StylesInfo.empty, no crash
}
```

---

## Integration Notes

### File Organization
```
Sources/Cuneiform/
├── Core/
│   ├── Errors.swift         (existing - add cases if needed)
│   ├── PartPath.swift       (existing)
│   ├── ContentType.swift    (existing)
│   └── Relationship.swift   (existing)
├── Package/
│   ├── OPCPackage.swift     (existing)
│   ├── ZipArchive.swift     (existing)
│   ├── ContentTypesParser.swift  (existing)
│   └── RelationshipsParser.swift (existing)
└── SpreadsheetML/           (NEW - create this directory)
    ├── SharedStringsParser.swift
    ├── WorkbookParser.swift
    ├── WorksheetParser.swift
    └── StylesParser.swift
```

### Using Existing Infrastructure

```swift
// Example: How to get XML data for parsing
var package = try OPCPackage.open(url: xlsxURL)

// Read shared strings
if package.partExists(.sharedStrings) {
    let data = try package.readPart(.sharedStrings)
    let sharedStrings = try SharedStringsParser.parse(data: data)
}

// Read workbook
let workbookData = try package.readPart(.workbook)
let workbookInfo = try WorkbookParser.parse(data: workbookData)

// Get workbook relationships to find sheet paths
let workbookRels = try package.relationships(for: .workbook)

// For each sheet, resolve its path and parse
for sheet in workbookInfo.sheets {
    if let rel = workbookRels[sheet.relationshipId] {
        let sheetPath = rel.resolveTarget(relativeTo: .workbook)
        let sheetData = try package.readPart(sheetPath)
        let worksheet = try WorksheetParser.parse(data: sheetData)
    }
}
```

### XML Parsing Pattern

Use Foundation's `XMLParser` with delegate (see existing `ContentTypesParser.swift` and `RelationshipsParser.swift` for examples):

```swift
final class MyParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    // ... state variables ...

    static func parse(data: Data) throws -> MyResult {
        let parser = MyParser()
        let xmlParser = XMLParser(data: data)
        xmlParser.delegate = parser

        guard xmlParser.parse() else {
            throw CuneiformError.malformedXML(...)
        }

        return MyResult(...)
    }

    func parser(_ parser: XMLParser,
                didStartElement elementName: String,
                namespaceURI: String?,
                qualifiedName qName: String?,
                attributes attributeDict: [String: String]) {
        // Handle element start
    }

    func parser(_ parser: XMLParser,
                foundCharacters string: String) {
        // Accumulate text content
    }

    func parser(_ parser: XMLParser,
                didEndElement elementName: String,
                namespaceURI: String?,
                qualifiedName qName: String?) {
        // Handle element end
    }
}
```

---

## Verification Checklist

After implementation, verify:

- [x] `swift build` succeeds with no warnings
- [x] `swift test` passes all tests (168/168)
- [x] All types are `Sendable`
- [x] All public APIs have doc comments
- [x] Error cases throw appropriate `CuneiformError` variants
- [x] Empty/missing optional parts handled gracefully (no crashes)
- [x] Code follows existing patterns in the codebase

---

## Test Data

For integration testing, use any .xlsx file created by Excel. The library should handle files from:
- Microsoft Excel (Windows/Mac)
- Google Sheets (exported as .xlsx)
- LibreOffice Calc
- Numbers (exported as .xlsx)

Create a simple test file with:
- Multiple sheets
- Text, numbers, dates, booleans
- Shared strings
- Some formatting
- A merged cell range

---

## Questions?

If requirements are unclear, check:
1. Existing parser implementations in `Package/` directory
2. ISO/IEC 29500-1 specification in `Documentation/iEC 29500/`
3. [MS-OI29500] for Excel-specific behaviors

Do not deviate from the specified APIs without documenting why.

---

## Phase 2: High-Level API, Writing, Queries, Performance, Styling

The following Phase 2 items are implemented and verified:

- [x] High-level read API: `Workbook.open(url:)`, `workbook.sheet(named:)`, `Sheet.cell(at:)`
- [x] `CellValue` resolution (shared strings, numbers, booleans, inline strings, errors, dates via styles)
- [x] Advanced queries: `Sheet.range(_:)`, `Sheet.column(_:)`, `Sheet.rows(where:)`, `Sheet.find`, `Sheet.findAll`
- [x] Performance: lazy sheet loading, streaming `RowIterator` for rows; benchmarks added
- [x] Write API: `WorkbookWriter`, worksheet XML emitting (numbers, strings, booleans, formulas)
- [x] Write-side styling: `StylesBuilder` for `styles.xml`; style indices threaded into cells (`s` attribute)
- [x] Round-trip tests for write + styling (bold, fills, borders, dates, large datasets)
- [x] Merged cells on write: `<mergeCells>` emission via `WorksheetBuilder` and `SheetWriter.mergeCells(_:)`, round-trip validated
- [x] Data validations on write: `<dataValidations>` emission with list and numeric constraints; operator support (e.g., `between` with `op` and `formula1`/`formula2`)
- [x] Workbook defined names: `<definedNames>` emission in `workbook.xml` for named ranges
 - [x] Expanded validation variants: decimal `greaterThanOrEqual`, date `between` with two formulas, whole `lessThanOrEqual`, list using range references
 - [x] Read-side parsing: workbook `definedNames` and worksheet `dataValidations` exposed via high-level APIs

### Upcoming Tasks

**Completed:**
- [x] Add "How to Use" doc snippets for named ranges and data validations
- [x] Hyperlinks and cell comments (read/write minimal)
- [x] Charts/drawings metadata parsing (Phase 3.2 complete)

**Next - Phase 4 (OOXML Toolkit Compliance):**
- [ ] **Phase 4.1: Full Styles** - fonts, fills, borders, alignment (read + write)
- [ ] **Phase 4.2: Tables** - Excel Tables/ListObjects
- [ ] **Phase 4.3: Conditional Formatting** - data bars, color scales, icon sets
- [ ] **Phase 4.4: AutoFilter** - column filtering
- [ ] **Phase 4.5: Rich Text** - formatted text runs
- [ ] **Phase 4.6: Shared Strings** - write optimization
- [ ] **Phase 4.7: Page Setup** - print configuration

**Release Prep (after Phase 4):**
- [ ] Exporters: CSV/JSON/HTML with streaming
- [ ] Version bump, CI (macOS/Linux), DocC preview, README polish, CHANGELOG

### References

- Read API and queries: `Sources/Cuneiform/SpreadsheetML/Workbook.swift`, `Sheet.swift`
- Writing: `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift`, `SpreadsheetMLBuilders.swift`
- Styling: `StylesBuilder` in `SpreadsheetMLBuilders.swift`; `StylesParser.swift` for read-side date detection
- Tests: `Tests/CuneiformTests/StylingTests.swift`, `WorkbookWriterTests.swift`, `AdvancedQueryTests.swift`, `PerformanceBenchmarks.swift`, `WorksheetWriteExtrasTests.swift`, `DataValidationVariantTests.swift`, `NamedRangesWriteTests.swift`

---
