# Excel Quirks and Compatibility Notes

This document captures Microsoft Excel-specific behaviors, deviations from the ISO standard, and practical compatibility requirements for the Cuneiform library.

## Critical Compatibility Rules

### 1. Always Use Transitional Format

**Rule:** Default to Transitional namespace, not Strict.

Excel versions before 2013 cannot open Strict format files. Even modern Excel opens Transitional by default and saves Transitional unless specifically configured otherwise.

```xml
<!-- Transitional (use this) -->
xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"

<!-- Strict (avoid unless explicitly requested) -->
xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
```

### 2. Relationship ID Generation

**Rule:** Always use format `rId{N}` where N is a positive integer.

Excel expects relationship IDs to follow this pattern. While the spec allows any NCName, using non-standard IDs may cause issues.

```xml
<!-- Good -->
<Relationship Id="rId1" .../>
<Relationship Id="rId2" .../>

<!-- Bad (may work but risky) -->
<Relationship Id="sheet1" .../>
<Relationship Id="rel-abc" .../>
```

### 3. Part Ordering in Content Types

**Rule:** `Default` elements must come before `Override` elements.

While not strictly required by spec, some older Excel versions expect this ordering:

```xml
<Types xmlns="...">
    <Default Extension="rels" ContentType="..."/>  <!-- Defaults first -->
    <Default Extension="xml" ContentType="..."/>
    <Override PartName="/xl/workbook.xml" ContentType="..."/>  <!-- Overrides second -->
</Types>
```

---

## Date and Time Handling

### The 1900 vs 1904 Date System

**Issue:** Excel has two incompatible date systems.

- **1900 System (Windows default):** Day 1 = January 1, 1900
- **1904 System (Mac legacy):** Day 1 = January 1, 1904

Check `workbook.xml` for `<workbookPr date1904="1"/>` to determine which system is in use.

**The 1900 Leap Year Bug:**

Excel incorrectly treats 1900 as a leap year (it wasn't). This means:
- Serial number 60 = February 29, 1900 (non-existent date)
- Serial numbers after 60 are off by one day from actual dates

**Implementation:**
```swift
func serialToDate(_ serial: Double, date1904: Bool) -> Date {
    let baseDate: Date
    if date1904 {
        // January 1, 1904
        baseDate = DateComponents(year: 1904, month: 1, day: 1).date!
    } else {
        // January 1, 1900 (but Excel thinks it's day 1, not day 0)
        // Also account for the fake Feb 29, 1900
        baseDate = DateComponents(year: 1899, month: 12, day: 30).date!
    }
    return baseDate.addingTimeInterval(serial * 86400)
}
```

### Time Zone Handling

**Rule:** Excel stores dates/times with no timezone information.

All dates are local time with no UTC offset. When reading/writing, do not apply timezone conversions.

---

## Number Formats

### Built-in Number Format IDs

Excel reserves format IDs 0-163 for built-in formats. Custom formats must use ID 164+.

| ID | Format | Notes |
|----|--------|-------|
| 0 | General | |
| 1 | 0 | Integer |
| 2 | 0.00 | Two decimals |
| 3 | #,##0 | Thousands separator |
| 4 | #,##0.00 | |
| 9 | 0% | Percentage |
| 10 | 0.00% | |
| 11 | 0.00E+00 | Scientific |
| 12 | # ?/? | Fraction |
| 13 | # ??/?? | Fraction (two digits) |
| 14 | mm-dd-yy | Date |
| 15 | d-mmm-yy | |
| 16 | d-mmm | |
| 17 | mmm-yy | |
| 18 | h:mm AM/PM | Time |
| 19 | h:mm:ss AM/PM | |
| 20 | h:mm | 24-hour |
| 21 | h:mm:ss | |
| 22 | m/d/yy h:mm | Date+time |
| 37-40 | Accounting formats | |
| 41-44 | Currency formats | |
| 45-48 | Time formats | mm:ss, [h]:mm:ss, etc. |
| 49 | @ | Text |

**Locale Dependency:** Many built-in formats are locale-dependent. Format ID 14 may display differently based on system locale.

### Detecting Date Cells

**Problem:** Dates are stored as numbers. How to know if a cell is a date?

Check the number format. If it contains date/time tokens (`y`, `m`, `d`, `h`, `s`, `AM/PM`) and NOT in brackets, it's likely a date format.

```swift
func isDateFormat(_ formatCode: String) -> Bool {
    // Remove bracketed sections (locale codes, colors, conditions)
    let withoutBrackets = formatCode.replacing(/\[.*?\]/, with: "")

    // Check for date/time tokens
    let dateTokens = ["y", "m", "d", "h", "s"]
    return dateTokens.contains { withoutBrackets.lowercased().contains($0) }
}
```

**Caveat:** Format `mm` could be months or minutes depending on context.

---

## Shared Strings

### Empty Strings

**Quirk:** Excel may or may not include empty strings in the shared string table.

An empty cell might be:
- Missing from XML entirely
- Present with `t="s"` pointing to an empty string
- Present with `t="inlineStr"` and empty `<is><t/></is>`

**Handle all cases:**
```swift
// Empty cell (no <c> element)
// Cell with shared string index pointing to ""
// Cell with inline empty string
```

### Whitespace Preservation

**Rule:** Use `xml:space="preserve"` for strings with leading/trailing whitespace.

```xml
<si><t xml:space="preserve">  has spaces  </t></si>
```

Without this attribute, XML parsers may strip whitespace.

### Rich Text Complexity

**Issue:** Rich text can have formatting at any level of granularity.

A single cell might have:
```xml
<si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><t> and </t></r>
    <r><rPr><i/></rPr><t>italic</t></r>
</si>
```

For round-trip fidelity, preserve the exact run structure even if you display it as plain text.

---

## Cell References

### Maximum Dimensions

- **Columns:** A to XFD (16,384 columns, index 0-16383)
- **Rows:** 1 to 1,048,576

### Column Letter Conversion

```swift
func columnIndex(from letters: String) -> Int {
    var result = 0
    for char in letters.uppercased() {
        result = result * 26 + Int(char.asciiValue! - Character("A").asciiValue! + 1)
    }
    return result - 1  // Convert to 0-based
}

func columnLetter(from index: Int) -> String {
    var result = ""
    var n = index + 1  // Convert to 1-based
    while n > 0 {
        n -= 1
        result = String(Character(UnicodeScalar(65 + n % 26)!)) + result
        n /= 26
    }
    return result
}
```

### R1C1 vs A1 Notation

Excel supports both but `.xlsx` files always use A1 notation in the XML. R1C1 is only used in:
- Some formula contexts
- VBA code
- User preferences

---

## Styles

### Default Styles

**Rule:** Excel expects certain built-in styles to exist.

The styles.xml must include:
- At least 2 fills: `none` and `gray125`
- At least 1 border: empty (no borders)
- At least 1 font: default font
- At least 1 cellXf: default cell format

```xml
<fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
</fills>
```

### Style Index 0

**Rule:** Style index 0 must be the default "Normal" style.

All cells without explicit `s` attribute use style 0.

### Font Family Values

| Value | Meaning |
|-------|---------|
| 0 | Not applicable |
| 1 | Roman |
| 2 | Swiss |
| 3 | Modern |
| 4 | Script |
| 5 | Decorative |

---

## Formulas

### Formula Prefix

**Quirk:** Formulas in XML do not include the leading `=`.

In Excel UI: `=SUM(A1:A10)`
In XML: `<f>SUM(A1:A10)</f>`

### Array Formulas

**Format:** Array formulas need special attributes:

```xml
<!-- Master cell of array formula -->
<c r="A1">
    <f t="array" ref="A1:B2">A1:A2*B1:B2</f>
    <v>10</v>
</c>
<!-- Other cells in array just reference the master -->
```

### Shared Formulas

**Optimization:** Excel may use shared formulas to reduce file size:

```xml
<!-- Master cell defines the formula -->
<c r="B2">
    <f t="shared" ref="B2:B100" si="0">A2*2</f>
    <v>4</v>
</c>
<!-- Other cells reference the shared formula -->
<c r="B3">
    <f t="shared" si="0"/>
    <v>6</v>
</c>
```

The `si` (shared index) links cells to their master formula.

### Volatile Functions

These functions recalculate every time, even if inputs haven't changed:
- `NOW()`, `TODAY()`
- `RAND()`, `RANDBETWEEN()`
- `OFFSET()`, `INDIRECT()`
- `INFO()`, `CELL()`

---

## Worksheets

### Sheet Names

**Restrictions:**
- Maximum 31 characters
- Cannot contain: `\ / ? * [ ] :`
- Cannot be empty
- Cannot start or end with apostrophe
- Cannot be "History" (reserved)

### Hidden Sheets

Two levels of hiding:
- `state="hidden"`: Can be unhidden via UI
- `state="veryHidden"`: Only unhideable via VBA

```xml
<sheet name="Data" sheetId="1" state="veryHidden" r:id="rId1"/>
```

### Tab Colors

Sheet tab colors are stored in the sheet's properties:

```xml
<sheetPr>
    <tabColor rgb="FF0000FF"/>
</sheetPr>
```

---

## Dimensions and Sparse Storage

### Dimension Element

**Quirk:** The `<dimension>` element may be inaccurate.

Excel sets it to the used range, but:
- Deleted content may leave stale dimensions
- Some tools don't update it properly

**Recommendation:** Don't trust `<dimension>` for actual data bounds. Scan the sheet data instead.

### Sparse Row/Cell Storage

Rows and cells are sparse - only non-empty cells are stored:

```xml
<sheetData>
    <row r="1" spans="1:3">  <!-- Row 1 -->
        <c r="A1"><v>1</v></c>
        <c r="C1"><v>3</v></c>  <!-- B1 is empty, not stored -->
    </row>
    <!-- Row 2 is empty, not stored -->
    <row r="3" spans="1:1">
        <c r="A3"><v>5</v></c>
    </row>
</sheetData>
```

---

## Merged Cells

### Storage Format

```xml
<mergeCells count="2">
    <mergeCell ref="A1:B2"/>
    <mergeCell ref="D1:F1"/>
</mergeCells>
```

### Value Location

**Rule:** Only the top-left cell of a merged region stores the value.

Other cells in the merged region should be empty in the XML.

---

## Performance Considerations

### Calculation Chain

**`calcChain.xml`** stores the order cells should be calculated.

**Recommendation:** When writing, omit `calcChain.xml`. Excel will rebuild it on first open. This avoids complex dependency tracking during write.

### Large Shared String Tables

For files with many unique strings, the shared string table can be huge.

**Optimization:** Consider:
- Streaming XML parsing (don't load entire file)
- Lazy loading of shared strings
- Memory-mapping for very large files

---

## XML Quirks

### Self-Closing Tags

Excel generates inconsistent XML:
```xml
<c r="A1"><v>1</v></c>  <!-- Sometimes -->
<c r="A1"><v>1</v></c>  <!-- Or this -->
<c r="A1"/>            <!-- Empty cell as self-closing -->
```

Accept all valid XML forms.

### Namespace Prefixes

Excel uses unprefixed default namespace for SpreadsheetML:
```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
```

But relationship references use `r:` prefix:
```xml
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
```

### Extension Lists

Excel 2010+ may include extension lists with newer features:

```xml
<extLst>
    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
         uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
        <!-- Sparklines, slicers, etc. -->
    </ext>
</extLst>
```

**Strategy:** Preserve unknown extensions for round-trip fidelity.

---

## Testing Recommendations

### Validation Tools

1. **Open in Excel** - The ultimate test
2. **Open XML SDK Validator** - Checks spec compliance
3. **LibreOffice Calc** - Tests interoperability

### Test File Corpus

Create test files covering:
- Empty workbook
- Single cell with each value type
- Dates in both 1900 and 1904 systems
- All built-in number formats
- Rich text in shared strings
- Merged cells
- Hidden rows/columns
- Multiple sheets with various visibilities
- Large files (10,000+ rows)
- Files with formulas
- Files with tables

### Round-Trip Testing

Read → Write → Read should produce identical data:
```swift
let original = try Workbook.open(path)
try original.save(to: tempPath)
let roundTripped = try Workbook.open(tempPath)
XCTAssertEqual(original, roundTripped)
```

---

## References

- [MS-OI29500] Microsoft Office Implementation Notes
- [MS-XLSX] Excel Binary File Format
- ECMA-376 Office Open XML File Formats
- OpenXML SDK Documentation
