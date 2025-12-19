# Cuneiform Implementation Verification Checklist

Use this checklist to verify work is complete and correct.

## Quick Verification Commands

```bash
cd /Users/jonathan/Desktop/garden/cuneiform
swift build   # Must succeed with no errors
swift test    # All tests must pass
```

---

## Status Summary (Dec 19, 2025)

**STATUS: GREEN** - All 191 tests pass.

Note: Swift 6 includes built-in Swift Testing. This toolchain on macOS currently requires the external `swift-testing` package for the `Testing` module; removing it led to missing `_TestingInternals`. We have retained the dependency to keep the suite green and accept the deprecation warnings. See [README.md](README.md#migration-notes-swift-6-testing) for migration steps when your toolchain supports the built-in module.

- [x] Parsers implemented: SharedStrings, Workbook, Worksheet, Styles (full), Charts, Workbook Protection, Pivot Tables, Tables/ListObjects
- [x] Full styles support: fonts, fills, borders, alignment read/write (Phase 4.1)
- [x] Tables discovery: parse /xl/tables/tableN.xml with columns, autoFilter, style info (Phase 4.2 read-side)
- [x] All parser tests pass
- [x] Build succeeds
- [x] Entire test suite passes (191/191)

---

## New Additions (Dec 18, 2025)

Write-side styling and formatting are implemented and verified.

- [x] StylesBuilder for `styles.xml` (fonts, fills, borders, number formats, xf cell formats)
- [x] `WorksheetBuilder` supports style indices per cell (`s` attribute)
- [x] `WorkbookWriter` now adds `xl/styles.xml` and a workbook styles relationship
- [x] Round-trip styling tests added: bold text, fills/colors, borders, date formats, large mixed-format datasets
- [x] Alignment fixes in tests to match `CellValue` API

Additional write-side features:
- [x] Merged cells: `<mergeCells>` emission and round-trip verification
- [x] Data validations: `<dataValidations>` section with list and numeric constraints, including operator and dual-formula support
- [x] Named ranges: `<definedNames>` emitted in `workbook.xml`
- [x] Hyperlinks (read-side): `<hyperlinks>` parsed with `ref`, `r:id`, `display`, `tooltip`, `location`
 - [x] Hyperlinks (write-side): `<hyperlinks>` emitted in worksheet XML; worksheet relationships `_rels/sheetN.xml.rels` include `Type=hyperlink` entries with `TargetMode="External"` for external URLs; internal hyperlinks use `location` only (no relationship)
 - [x] Comments (read-side): worksheet comments parsed via relationships and exposed through `Sheet.comments` helpers
 - [x] Comments (write-side minimal): emits `/xl/commentsN.xml` with authors and comment text; worksheet relationships include `Type=comments` targets pointing to `../commentsN.xml`
 - [x] Comments (display via VML): emits `/xl/drawings/vmlDrawingN.vml` with VML shapes anchored to each comment; worksheet relationships include `Type=vmlDrawing` targets; worksheet XML includes `<legacyDrawing r:id>` element; enabling Excel comment indicators and bubbles

New validation variants:
- [x] Decimal validation with `greaterThanOrEqual`
- [x] Date validation with `between` and dual formulas
- [x] Whole-number validation with `lessThanOrEqual`
- [x] List validation using range references (e.g., `Sheet!$A$1:$A$10`)

### Files Updated/Added
- [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (added `StylesBuilder`, style-aware cells; merged cells and data validations emission)
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (integrated `styles.xml` and relationships; merge/data validation APIs; defined names)
- [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (date detection via styles retained)
- [x] `Tests/CuneiformTests/StylingTests.swift` (new write-side styling tests)
- [x] `Tests/CuneiformTests/WorksheetWriteExtrasTests.swift` (merged cells, list validations presence)
- [x] `Tests/CuneiformTests/DataValidationVariantTests.swift` (numeric between validation with operator and two formulas)
- [x] `Tests/CuneiformTests/NamedRangesWriteTests.swift` (defined names emission in workbook)
- [x] `Tests/CuneiformTests/NamedRangesReadTests.swift` (read defined names from workbook)
- [x] `Tests/CuneiformTests/DataValidationReadTests.swift` (read data validations from worksheet)
 - [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (ergonomic helpers `validations(for:)`, `validations(at:)`)
 - [x] `Sources/Cuneiform/SpreadsheetML/Workbook.swift` (helpers `definedName(_:)`, `definedNameRange(_:)`)
 - [x] `Tests/CuneiformTests/DataValidationHelpersTests.swift` (tests for `Sheet.validations(for:)` and `Sheet.validations(at:)`)
 - [x] `README.md` (added ergonomic examples for named ranges and data validations)
 - [x] `Tests/CuneiformTests/HyperlinksReadTests.swift` (parse worksheet `<hyperlink>` entries)
 - [x] `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift` (hyperlinks parsing)
 - [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (hyperlinks exposure via `Sheet.hyperlinks` and `hyperlinks(at:)`)
 - [x] `Tests/CuneiformTests/HyperlinksWriteTests.swift` (write-side `<hyperlinks>` emission and worksheet `.rels` external hyperlink relationships)
 - [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (worksheet-level `.rels` emission for external hyperlinks; `SheetWriter.addHyperlinkExternal` / `addHyperlinkInternal` APIs)
 - [x] `Sources/Cuneiform/SpreadsheetML/CommentsParser.swift`, `CommentsBuilder.swift` (comments parse/build)
 - [x] `Sources/Cuneiform/Core/ContentType.swift` (comments content type)
 - [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (comments part overrides and relationships)
 - [x] `Tests/CuneiformTests/CommentsParserTests.swift`, `Tests/CuneiformTests/CommentsWriteTests.swift`
 - [x] `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift` (sheet protection parsing from `<sheetProtection>` element)
 - [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (Sheet.protection property)
 - [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (WorksheetBuilder protection support, XML emission)
 - [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (SheetProtectionOptions struct and SheetWriter.protectSheet API)
 - [x] `Tests/CuneiformTests/SheetProtectionParserTests.swift` (6 tests for read-side protection parsing)
 - [x] `Tests/CuneiformTests/SheetProtectionWriteTests.swift` (7 tests for write-side protection emission and round-trip)

### Phase 3.1: Sheet Protection (Dec 18, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift` (sheet protection parsing from `<sheetProtection>` element)
- [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (Sheet.protection property)
- [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (WorksheetBuilder protection support, XML emission)
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (SheetProtectionOptions struct and SheetWriter.protectSheet API)
- [x] `Tests/CuneiformTests/SheetProtectionParserTests.swift` (6 tests for read-side protection parsing)
- [x] `Tests/CuneiformTests/SheetProtectionWriteTests.swift` (7 tests for write-side protection emission and round-trip)

### Phase 3.2: Charts - Basic Parser (Dec 18, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/ChartParser.swift` (parse `/xl/charts/chart*.xml`)
- [x] `Sources/Cuneiform/Core/ContentType.swift` (added `.chart`, `.drawing` content types)
- [x] `Sources/Cuneiform/Core/Relationship.swift` (added `.chart`, `.drawing` relationship types)
- [x] `Tests/CuneiformTests/ChartParserTests.swift` (6 tests for chart type parsing: column, line, pie, area, bar, unknown)
- [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (added chartsList and charts property)
- [x] `Sources/Cuneiform/SpreadsheetML/Workbook.swift` (integrated chart discovery via relationships)
- [x] `Tests/CuneiformTests/ChartIntegrationTests.swift` (3 tests for chart discovery and integration)

### Phase 3.3: Workbook Protection (Dec 18, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookParser.swift` (parse `<workbookProtection>` element with sheet/windows/password attributes)
- [x] `Sources/Cuneiform/SpreadsheetML/Workbook.swift` (protection property exposure)
- [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (WorkbookBuilder.setProtection() and XML emission)
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (WorkbookProtectionOptions struct and protectWorkbook API)
- [x] `Tests/CuneiformTests/WorkbookProtectionParserTests.swift` (6 tests for read-side protection parsing)
- [x] `Tests/CuneiformTests/WorkbookProtectionWriteTests.swift` (8 tests for write-side protection and round-trip)

### Phase 3.4: Pivot Tables (Dec 18, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/PivotTableParser.swift` (parse `/xl/pivotTables/pivotTableN.xml`)
- [x] `Sources/Cuneiform/Core/Relationship.swift` (added `.pivotTable` relationship type)
- [x] `Sources/Cuneiform/SpreadsheetML/Workbook.swift` (automatic discovery via worksheet relationships; `Workbook.pivotTables` property)
- [x] `Tests/CuneiformTests/PivotTableParserTests.swift` (7 tests for pivot metadata)
- [x] `Tests/CuneiformTests/WorkbookIntegrationTests.swift` (1 integration test validating discovery from real XLSX with 22 pivot tables)

### Phase 4.1: Full Styles Support (Dec 19, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/StylesParser.swift` (expanded with CellFont, CellFill, CellBorder, CellBorderSide, CellAlignment, CellStyle, CellColor types; parse fonts, fills, borders, alignment)
- [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (added `cellStyle(at:)` methods for CellReference and String)
- [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (added `StylesBuilder.addCellStyle()` convenience method)
- [x] `Tests/CuneiformTests/Phase41StylesTests.swift` (8 tests: parseFonts, parseFontsWithThemeColor, parseFills, parseAlignment, getCellStyle, roundTripCellStyle, multipleStylesWithDedup, parseBorders)
- [x] All types marked Sendable and Equatable
- [x] Full round-trip styling verified

### Phase 4.2: Tables/ListObjects (Dec 19, 2025) ✓
- [x] `Sources/Cuneiform/SpreadsheetML/TableParser.swift` (parse `/xl/tables/tableN.xml` with XMLParser delegate)
- [x] Domain types: `TableData`, `TableColumn`, `TableAutoFilter`, `TableStyleInfo` (all Sendable, Equatable)
- [x] `Sources/Cuneiform/SpreadsheetML/Workbook.swift` (added `Workbook.tables` property with automatic discovery via worksheet relationships)
- [x] `Tests/CuneiformTests/TableParserTests.swift` (10 comprehensive parser tests: simple table, totals row, autoFilter, style, multiple functions, edge cases, equatable, sendable)
- [x] `Tests/CuneiformTests/TableIntegrationTests.swift` (5 integration tests: table discovery, multiple tables, column ordering, complex table with all features, malformed XML)
- [x] Read-side implementation complete; write-side deferred to next iteration

### Verification
- [x] `swift build` succeeds
- [x] `swift test` succeeds: 191 tests passing (176 baseline + 8 Phase 4.1 + 15 Phase 4.2 read-side)

---

## Phase 4: OOXML Toolkit Compliance

**Goal:** Increase ISO/IEC 29500 compliance from ~60% to ~85%+

**Current Compliance:** ~70% (after Phase 4.2 read-side)

| Area | Current | Target | Status |
|------|---------|--------|--------|
| Styles (§18.8) | 90% | 90% | ✓ Complete (Phase 4.1) |
| Tables (§18.5) | 50% | 80% | ⚡ Phase 4.2 Read-Side Complete |
| Conditional Formatting (§18.3.1) | 0% | 80% | Planned (Phase 4.3) |
| AutoFilter | 0% | 80% | Planned (Phase 4.4) |
| Rich Text (§18.4) | 20% | 90% | Planned (Phase 4.5) |

### Phase 4.1: Full Styles Support ✓ (Dec 19, 2025)
- [x] Read: fonts, fills, borders, alignment from `styles.xml`
- [x] Write: complete style elements with all attributes via `StylesBuilder.addCellStyle()`
- [x] API: `CellStyle` struct, `Sheet.cellStyle(at:)` methods (CellReference and String variants)
- [x] Types: `CellFont`, `CellFill`, `CellBorder`, `CellBorderSide`, `CellAlignment`, `CellColor`, `NumberFormat` (all Sendable, Equatable)
- [x] Tests: 8 comprehensive tests for round-trip formatting verification

### Phase 4.2: Tables/ListObjects ⚡ (Dec 19, 2025 - Read-Side Complete)
- [x] Read: parse `/xl/tables/tableN.xml` with XMLParser delegate
- [x] Extract: name, displayName, ref (range), headerRowCount, totalsRowCount
- [x] Parse: `<tableColumn>` elements with id, name, totalsRowFunction
- [x] Parse: `<autoFilter>` within table
- [x] Parse: `tableStyleInfo` element with all flags
- [x] API: `TableData` struct, `Sheet.tables` property discovery via relationships
- [x] Types: `TableColumn`, `TableAutoFilter`, `TableStyleInfo` (all Sendable, Equatable)
- [x] Tests: 15 comprehensive tests for parser and integration
- [ ] Write-side: deferred to Phase 4.2 continuation (emit `table.xml`, `<tableParts>`, relationships)

### Phase 4.2: Tables/ListObjects
- [ ] Read: parse `/xl/tables/tableN.xml`
- [ ] Write: emit table.xml with columns
- [ ] API: `TableData`, `Sheet.tables`, `SheetWriter.addTable()`
- [ ] Tests: create and parse Excel tables

### Phase 4.3: Conditional Formatting
- [ ] Read: parse `<conditionalFormatting>` rules
- [ ] Write: emit highlight, data bar, color scale, icon set rules
- [ ] API: `ConditionalRule` enum, `Sheet.conditionalFormats`
- [ ] Tests: various rule types

### Phase 4.4: AutoFilter
- [ ] Read: parse `<autoFilter>` element
- [ ] Write: emit filter ranges and criteria
- [ ] API: `AutoFilter`, `Sheet.autoFilter`
- [ ] Tests: filter range setup

### Phase 4.5: Rich Text
- [ ] Read: preserve `<r>` runs with formatting
- [ ] Write: emit rich text in cells and comments
- [ ] API: `TextRun`, `RichText`, `CellValue.richText`
- [ ] Tests: formatted text round-trip

### Phase 4.6: Shared Strings Optimization
- [ ] Write: use shared strings table for efficiency
- [ ] Tests: large file optimization

### Phase 4.7: Page Setup & Print
- [ ] Read: parse `<pageSetup>`, `<pageMargins>`
- [ ] Write: emit print configuration
- [ ] API: `PageSetup`, `Sheet.pageSetup`

---

## Ergonomics

- [x] Data validations helpers: `Sheet.validations(for:)` to filter by A1 range and `Sheet.validations(at:)` for a single cell.
- [x] Named range helpers: `Workbook.definedName(_:)` to fetch by name and `Workbook.definedNameRange(_:)` to resolve `Sheet!$A$1:$B$10` into `(sheet, range)`.

## Phase 2 Test Suites

### WorkbookIntegrationTests
- [x] High-level `Workbook.open(url:)` API
- [x] Sheet access by name and index
- [x] Cell value resolution (shared strings, numbers, dates)

### AdvancedQueryTests
- [x] `Sheet.range(_:)` extracts cell ranges
- [x] `Sheet.column(_:)` retrieves column data
- [x] `Sheet.rows(where:)` filters rows by predicate
- [x] `Sheet.find` / `Sheet.findAll` cell search

### WorkbookWriterTests
- [x] Create new workbook with sheets
- [x] Write numbers, strings, booleans, formulas
- [x] Round-trip: write then read back

### StylingTests
- [x] Bold text styling
- [x] Fill colors and patterns
- [x] Border styles
- [x] Date number formats
- [x] Large mixed-format datasets

### PerformanceBenchmarks
- [x] Lazy sheet loading performance
- [x] Streaming row iteration
- [x] Large file handling

### WorksheetWriteExtrasTests
- [x] Merged cells round-trip
- [x] List data validation presence

### DataValidationVariantTests
- [x] Numeric between validation with operator and two formulas

### NamedRangesWriteTests
- [x] Defined names emitted in `workbook.xml`

### NamedRangesReadTests
- [x] Read defined names via `Workbook.definedNamesList`

### DataValidationReadTests
- [x] Read data validations via `Sheet.dataValidations`

---

## Phase 1: Parser Tasks

## Task 1: SharedStrings Parser

### Files Created
- [x] `Sources/Cuneiform/SpreadsheetML/SharedStringsParser.swift`
- [x] `Tests/CuneiformTests/SharedStringsTests.swift`

### API Exists
- [x] `SharedStrings` struct with `strings: [String]`, `subscript`, `count`, `empty`
- [x] `SharedStringsParser.parse(data:)` static method

### Functionality
- [x] Simple strings: `<si><t>text</t></si>` → "text"
- [x] Rich text: `<si><r><t>a</t></r><r><t>b</t></r></si>` → "ab"
- [x] Empty strings handled
- [x] Out-of-bounds subscript returns nil (no crash)

### Tests Pass
- [x] `parseSimpleStrings`
- [x] `parseRichText`
- [x] `parsePreservedWhitespace`
- [x] `parseEmptyString`
- [x] `subscriptOutOfBounds`

---

## Task 2: Workbook Parser

### Files Created
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookParser.swift`
- [x] `Tests/CuneiformTests/WorkbookParserTests.swift`

### API Exists
- [x] `SheetInfo` struct with `name`, `sheetId`, `relationshipId`, `state`
- [x] `SheetState` enum: `visible`, `hidden`, `veryHidden` + `CaseIterable`
- [x] `WorkbookInfo` struct with `sheets: [SheetInfo]`, `sheet(named:)`
- [x] `WorkbookParser.parse(data:)` static method

### Functionality
- [x] Parses all `<sheet>` elements in order
- [x] Extracts `name`, `sheetId`, `r:id` attributes
- [x] Handles `state` attribute (default = visible)
- [x] `sheet(named:)` lookup works
- [x] Empty workbook doesn't crash

### Tests Pass
- [x] `parseSheets`
- [x] `parseSheetStates`
- [x] `sheetByName`
- [x] `missingNameThrows`
- [x] `emptyWorkbook`

---

## Task 3: Worksheet Parser

### Files Created
- [x] `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift`
- [x] `Tests/CuneiformTests/WorksheetParserTests.swift`

### API Exists
- [x] `CellReference` struct with `column`, `row`, `columnIndex`, `init?(_ reference:)`
- [x] `CellReference: CustomStringConvertible, ExpressibleByStringLiteral`
- [x] `RawCellValue` enum: `sharedString`, `number`, `boolean`, `inlineString`, `error`, `date`, `empty`
- [x] `RawCell` struct with `reference`, `value`, `styleIndex`
- [x] `RawRow` struct with `index`, `cells`
- [x] `WorksheetData` struct with `dimension`, `rows`, `mergedCells`, `cell(at:)`
- [x] `WorksheetParser.parse(data:)` static method

### Functionality
- [x] Cell reference parsing: A1, Z1, AA1, AZ1, etc.
- [x] Column index calculation correct (A=0, Z=25, AA=26)
- [x] Shared string cells (`t="s"`) parsed correctly
- [x] Number cells (no `t` attribute) parsed correctly
- [x] Boolean cells (`t="b"`) parsed correctly
- [x] Inline string cells (`t="str"`) parsed correctly
- [x] Error cells (`t="e"`) parsed correctly
- [x] Empty cells handled
- [x] Merged cells extracted
- [x] Sparse data doesn't crash

### Tests Pass
- [x] `parseCellReference`
- [x] `parseSharedStringCell`
- [x] `parseNumberCell`
- [x] `parseBooleanCell`
- [x] `parseInlineStringCell`
- [x] `parseErrorCell`
- [x] `parseEmptyCell`
- [x] `parseMergedCells`
- [x] `cellLookup`
- [x] `sparseData`

---

## Task 4: Styles Parser

### Files Created
- [x] `Sources/Cuneiform/SpreadsheetML/StylesParser.swift`
- [x] `Tests/CuneiformTests/StylesParserTests.swift`

### API Exists
- [x] `NumberFormat` struct with `id`, `formatCode`, `isDateFormat`
- [x] `StylesInfo` struct with `numberFormats`, `cellFormats`, `numberFormat(forStyleIndex:)`, `isDateFormat(styleIndex:)`, `empty`
- [x] `StylesParser.parse(data:)` static method

### Functionality
- [x] Custom number formats parsed from `<numFmts>`
- [x] Cell formats parsed from `<cellXfs>`
- [x] Built-in date formats (14-22) detected
- [x] Custom date formats detected (contains y/m/d/h/s)
- [x] Number formats not mistaken for dates
- [x] Missing styles.xml returns `StylesInfo.empty`

### Tests Pass
- [x] `parseCustomNumberFormats`
- [x] `parseCellFormats`
- [x] `builtInDateFormats`
- [x] `customDateFormat`
- [x] `numberFormatNotDate`
- [x] `missingStylesFile`

---

## Overall Quality

- [x] All types marked `Sendable`
- [x] All public APIs have doc comments
- [x] Code style matches existing codebase
- [x] Errors use `CuneiformError` enum

Note: Deprecation warnings from swift-testing package (cosmetic, Swift 6 issue).

## Swift Style (Modern Idioms)

### Must Use
- [x] `if let value { }` shorthand - N/A (no redundant bindings)
- [x] Keypath expressions: `array.map(\.name)` - Used in `flatMap(\.cells)`
- [x] Implicit returns in single-expression computed properties
- [x] `guard` for early exits
- [x] Trailing closure syntax
- [x] `first(where:)` not `filter().first`

### Protocol Conformances
- [x] `Hashable` where equality comparison needed
- [x] `CustomStringConvertible` for debugging (`CellReference`)
- [x] `ExpressibleByStringLiteral` where ergonomic (`CellReference`)
- [x] `CaseIterable` on enums where useful (`SheetState`)
- [x] Synthesized conformances (not manual `==` or `hash(into:)`)

### Types
- [x] Value types (`struct`) preferred over `class`
- [x] Classes marked `final`
- [x] Enums with associated values for variants (`RawCellValue`)
- [x] Enums with raw values for parsing known strings (`SheetState`)
- [x] Computed properties over getter methods
- [x] Subscripts for indexed/keyed access

### Avoid (Reject if Present)
- [x] No `NSString`, `NSArray`, `NSDictionary`
- [x] No unnecessary force unwraps `!`
- [x] No `var` where `let` works
- [x] No `Any` or `AnyObject` (use generics)
- [x] No stringly-typed APIs (use enums/wrappers)
- [x] No mutable state where immutable works

---

## Integration Smoke Test

After all parsers are implemented, this should work:

```swift
import Cuneiform

// Open any .xlsx file
var package = try OPCPackage.open(url: testFileURL)

// Parse shared strings
let ssData = try package.readPart(.sharedStrings)
let sharedStrings = try SharedStringsParser.parse(data: ssData)
print("Shared strings: \(sharedStrings.count)")

// Parse workbook
let wbData = try package.readPart(.workbook)
let workbook = try WorkbookParser.parse(data: wbData)
print("Sheets: \(workbook.sheets.map(\.name))")

// Parse first worksheet
let wbRels = try package.relationships(for: .workbook)
if let firstSheet = workbook.sheets.first,
   let rel = wbRels[firstSheet.relationshipId] {
    let sheetPath = rel.resolveTarget(relativeTo: .workbook)
    let sheetData = try package.readPart(sheetPath)
    let worksheet = try WorksheetParser.parse(data: sheetData)
    print("Rows: \(worksheet.rows.count)")
}

// Parse styles
if package.partExists(.styles) {
    let stylesData = try package.readPart(.styles)
    let styles = try StylesParser.parse(data: stylesData)
    print("Cell formats: \(styles.cellFormats.count)")
}
```

---

## Sign-Off

- [x] **Build passes**: `swift build` exits 0
- [x] **Tests pass**: `swift test` shows all green (168/168)
- [x] **Checklist complete**: All boxes above checked
- [x] **Code reviewed**: Matches existing style

**Phase 1 Verified by:** Vek (December 17, 2025)
**Phase 2 Verified by:** Vek (December 18, 2025)
