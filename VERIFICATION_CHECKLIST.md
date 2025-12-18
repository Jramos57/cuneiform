# Parser Implementation Verification Checklist

Use this checklist to verify the delegated work is complete and correct.

## Quick Verification Commands

```bash
cd /Users/jonathan/Desktop/garden/cuneiform
swift build   # Must succeed with no errors
swift test    # All tests must pass
```

---

## Task 1: SharedStrings Parser

### Files Created
- [ ] `Sources/Cuneiform/SpreadsheetML/SharedStringsParser.swift`
- [ ] `Tests/CuneiformTests/SharedStringsTests.swift`

### API Exists
- [ ] `SharedStrings` struct with `strings: [String]`, `subscript`, `count`, `empty`
- [ ] `SharedStringsParser.parse(data:)` static method

### Functionality
- [ ] Simple strings: `<si><t>text</t></si>` → "text"
- [ ] Rich text: `<si><r><t>a</t></r><r><t>b</t></r></si>` → "ab"
- [ ] Empty strings handled
- [ ] Out-of-bounds subscript returns nil (no crash)

### Tests Pass
- [ ] `parseSimpleStrings`
- [ ] `parseRichText`
- [ ] `parsePreservedWhitespace`
- [ ] `parseEmptyString`
- [ ] `subscriptOutOfBounds`

---

## Task 2: Workbook Parser

### Files Created
- [ ] `Sources/Cuneiform/SpreadsheetML/WorkbookParser.swift`
- [ ] `Tests/CuneiformTests/WorkbookParserTests.swift`

### API Exists
- [ ] `SheetInfo` struct with `name`, `sheetId`, `relationshipId`, `state`
- [ ] `SheetState` enum: `visible`, `hidden`, `veryHidden`
- [ ] `WorkbookInfo` struct with `sheets: [SheetInfo]`, `sheet(named:)`
- [ ] `WorkbookParser.parse(data:)` static method

### Functionality
- [ ] Parses all `<sheet>` elements in order
- [ ] Extracts `name`, `sheetId`, `r:id` attributes
- [ ] Handles `state` attribute (default = visible)
- [ ] `sheet(named:)` lookup works
- [ ] Empty workbook doesn't crash

### Tests Pass
- [ ] `parseSheets`
- [ ] `parseSheetStates`
- [ ] `sheetByName`
- [ ] `missingNameThrows`
- [ ] `emptyWorkbook`

---

## Task 3: Worksheet Parser

### Files Created
- [ ] `Sources/Cuneiform/SpreadsheetML/WorksheetParser.swift`
- [ ] `Tests/CuneiformTests/WorksheetParserTests.swift`

### API Exists
- [ ] `CellReference` struct with `column`, `row`, `columnIndex`, `init?(_ reference:)`
- [ ] `RawCellValue` enum: `sharedString`, `number`, `boolean`, `inlineString`, `error`, `date`, `empty`
- [ ] `RawCell` struct with `reference`, `value`, `styleIndex`
- [ ] `RawRow` struct with `index`, `cells`
- [ ] `WorksheetData` struct with `dimension`, `rows`, `mergedCells`, `cell(at:)`
- [ ] `WorksheetParser.parse(data:)` static method

### Functionality
- [ ] Cell reference parsing: A1, Z1, AA1, AZ1, etc.
- [ ] Column index calculation correct (A=0, Z=25, AA=26)
- [ ] Shared string cells (`t="s"`) parsed correctly
- [ ] Number cells (no `t` attribute) parsed correctly
- [ ] Boolean cells (`t="b"`) parsed correctly
- [ ] Inline string cells (`t="str"`) parsed correctly
- [ ] Error cells (`t="e"`) parsed correctly
- [ ] Empty cells handled
- [ ] Merged cells extracted
- [ ] Sparse data doesn't crash

### Tests Pass
- [ ] `parseCellReference`
- [ ] `parseSharedStringCell`
- [ ] `parseNumberCell`
- [ ] `parseBooleanCell`
- [ ] `parseInlineStringCell`
- [ ] `parseErrorCell`
- [ ] `parseEmptyCell`
- [ ] `parseMergedCells`
- [ ] `cellLookup`
- [ ] `sparseData`

---

## Task 4: Styles Parser

### Files Created
- [ ] `Sources/Cuneiform/SpreadsheetML/StylesParser.swift`
- [ ] `Tests/CuneiformTests/StylesParserTests.swift`

### API Exists
- [ ] `NumberFormat` struct with `id`, `formatCode`, `isDateFormat`
- [ ] `StylesInfo` struct with `numberFormats`, `cellFormats`, `numberFormat(forStyleIndex:)`, `isDateFormat(styleIndex:)`, `empty`
- [ ] `StylesParser.parse(data:)` static method

### Functionality
- [ ] Custom number formats parsed from `<numFmts>`
- [ ] Cell formats parsed from `<cellXfs>`
- [ ] Built-in date formats (14-22) detected
- [ ] Custom date formats detected (contains y/m/d/h/s)
- [ ] Number formats not mistaken for dates
- [ ] Missing styles.xml returns `StylesInfo.empty`

### Tests Pass
- [ ] `parseCustomNumberFormats`
- [ ] `parseCellFormats`
- [ ] `builtInDateFormats`
- [ ] `customDateFormat`
- [ ] `numberFormatNotDate`
- [ ] `missingStylesFile`

---

## Overall Quality

- [ ] All types marked `Sendable`
- [ ] All public APIs have doc comments
- [ ] No compiler warnings
- [ ] Code style matches existing codebase
- [ ] Errors use `CuneiformError` enum

## Swift Style (Modern Idioms)

### Must Use
- [ ] `if let value { }` shorthand (not `if let value = value`)
- [ ] Keypath expressions: `array.map(\.name)` not `array.map { $0.name }`
- [ ] Implicit returns in single-expression computed properties
- [ ] `guard` for early exits
- [ ] Trailing closure syntax
- [ ] `first(where:)` not `filter().first`

### Protocol Conformances
- [ ] `Hashable` where equality comparison needed
- [ ] `CustomStringConvertible` for debugging
- [ ] `ExpressibleByStringLiteral` where ergonomic (e.g., `CellReference`)
- [ ] `CaseIterable` on enums where useful
- [ ] Synthesized conformances (not manual `==` or `hash(into:)`)

### Types
- [ ] Value types (`struct`) preferred over `class`
- [ ] Classes marked `final`
- [ ] Enums with associated values for variants
- [ ] Enums with raw values for parsing known strings
- [ ] Computed properties over getter methods
- [ ] Subscripts for indexed/keyed access

### Avoid (Reject if Present)
- [ ] No `NSString`, `NSArray`, `NSDictionary`
- [ ] No unnecessary force unwraps `!`
- [ ] No `var` where `let` works
- [ ] No `Any` or `AnyObject` (use generics)
- [ ] No stringly-typed APIs (use enums/wrappers)
- [ ] No mutable state where immutable works

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

- [ ] **Build passes**: `swift build` exits 0
- [ ] **Tests pass**: `swift test` shows all green
- [ ] **Checklist complete**: All boxes above checked
- [ ] **Code reviewed**: Matches existing style

**Verified by:** _______________
**Date:** _______________
