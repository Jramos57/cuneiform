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

**STATUS: GREEN** - All 111 tests pass.

Note: Swift 6 includes built-in Swift Testing. This toolchain on macOS currently requires the external `swift-testing` package for the `Testing` module; removing it led to missing `_TestingInternals`. We have retained the dependency to keep the suite green and accept the deprecation warnings. See [README.md](README.md#migration-notes-swift-6-testing) for migration steps when your toolchain supports the built-in module.

- [x] Parsers implemented: SharedStrings, Workbook, Worksheet, Styles
- [x] All parser tests pass
- [x] Build succeeds
- [x] Entire test suite passes (111/111)

---

## New Additions (Dec 19, 2025)

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

### Files Updated/Added
- [x] `Sources/Cuneiform/SpreadsheetML/SpreadsheetMLBuilders.swift` (added `StylesBuilder`, style-aware cells; merged cells and data validations emission)
- [x] `Sources/Cuneiform/SpreadsheetML/WorkbookWriter.swift` (integrated `styles.xml` and relationships; merge/data validation APIs; defined names)
- [x] `Sources/Cuneiform/SpreadsheetML/Sheet.swift` (date detection via styles retained)
- [x] `Tests/CuneiformTests/StylingTests.swift` (new write-side styling tests)
- [x] `Tests/CuneiformTests/WorksheetWriteExtrasTests.swift` (merged cells, list validations presence)
- [x] `Tests/CuneiformTests/DataValidationVariantTests.swift` (numeric between validation with operator and two formulas)
- [x] `Tests/CuneiformTests/NamedRangesWriteTests.swift` (defined names emission in workbook)

### Verification
- [x] `swift build` succeeds
- [x] `swift test` succeeds: 111 tests passing including new styling suite, write extras, named ranges, and performance benchmarks

---

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
- [x] **Tests pass**: `swift test` shows all green (111/111)
- [x] **Checklist complete**: All boxes above checked
- [x] **Code reviewed**: Matches existing style

**Phase 1 Verified by:** Vek (December 17, 2025)
**Phase 2 Verified by:** Vek (December 18-19, 2025)
