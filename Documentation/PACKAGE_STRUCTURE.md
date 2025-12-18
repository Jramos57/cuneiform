# Office Open XML Package Structure

This document describes the ZIP-based package structure of .xlsx files as defined by ISO/IEC 29500 (Office Open XML) and implemented by Microsoft Excel.

## Overview

An .xlsx file is a ZIP archive containing XML documents and their relationships. The package follows the **Open Packaging Conventions (OPC)** which define:

1. How content is organized into "parts" (files within the ZIP)
2. How parts relate to each other via "relationships"
3. How content types are declared

## Minimal Valid .xlsx Structure

```
example.xlsx (ZIP archive)
├── [Content_Types].xml          # REQUIRED: Declares content types for all parts
├── _rels/
│   └── .rels                    # REQUIRED: Package-level relationships
├── docProps/                    # Optional but typical
│   ├── app.xml                  # Application properties
│   └── core.xml                 # Core properties (author, title, etc.)
└── xl/
    ├── _rels/
    │   └── workbook.xml.rels    # REQUIRED: Workbook relationships
    ├── workbook.xml             # REQUIRED: Main workbook definition
    ├── worksheets/
    │   └── sheet1.xml           # REQUIRED: At least one worksheet
    ├── sharedStrings.xml        # Optional: Shared string table
    ├── styles.xml               # Optional: Cell styles and formatting
    └── theme/
        └── theme1.xml           # Optional: Theme definitions
```

## Core Package Files

### 1. [Content_Types].xml

**Purpose:** Declares the content type (MIME type) for every part in the package.

**Location:** Root of ZIP (required name with brackets)

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <!-- Default content types by extension -->
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>

    <!-- Override content types for specific parts -->
    <Override PartName="/xl/workbook.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/theme/theme1.xml"
              ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/docProps/core.xml"
              ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml"
              ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

**Key Content Types for SpreadsheetML:**

| Part Type | Content Type |
|-----------|--------------|
| Workbook (main) | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| Worksheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| Shared Strings | `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml` |
| Styles | `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Table | `application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml` |
| Pivot Table | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml` |
| Relationships | `application/vnd.openxmlformats-package.relationships+xml` |

---

### 2. _rels/.rels (Package Relationships)

**Purpose:** Defines top-level relationships from the package root.

**Location:** `_rels/.rels`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                  Target="xl/workbook.xml"/>
    <Relationship Id="rId2"
                  Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                  Target="docProps/core.xml"/>
    <Relationship Id="rId3"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                  Target="docProps/app.xml"/>
</Relationships>
```

**Relationship Types:**

| Type URI | Purpose |
|----------|---------|
| `.../officeDocument` | Points to main document (workbook.xml) |
| `.../metadata/core-properties` | Points to core.xml |
| `.../extended-properties` | Points to app.xml |

---

### 3. xl/workbook.xml (Main Workbook)

**Purpose:** Defines the workbook structure, sheet listing, and workbook-level properties.

**Location:** `xl/workbook.xml`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
    <workbookPr defaultThemeVersion="124226"/>
    <bookViews>
        <workbookView xWindow="120" yWindow="120" windowWidth="19095" windowHeight="11475"/>
    </bookViews>
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    </sheets>
    <calcPr calcId="125725"/>
</workbook>
```

**Key Elements:**
- `<sheets>`: Lists all worksheets with name, ID, and relationship reference
- `<bookViews>`: Window position and active tab
- `<definedNames>`: Named ranges and print areas
- `<calcPr>`: Calculation properties

---

### 4. xl/_rels/workbook.xml.rels (Workbook Relationships)

**Purpose:** Maps relationship IDs (rId) to actual part paths.

**Location:** `xl/_rels/workbook.xml.rels`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                  Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                  Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                  Target="theme/theme1.xml"/>
    <Relationship Id="rId4"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                  Target="styles.xml"/>
    <Relationship Id="rId5"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                  Target="sharedStrings.xml"/>
</Relationships>
```

---

### 5. xl/worksheets/sheet1.xml (Worksheet)

**Purpose:** Contains all data for a single worksheet.

**Location:** `xl/worksheets/sheet{N}.xml`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <dimension ref="A1:L10"/>
    <sheetViews>
        <sheetView workbookViewId="0">
            <selection activeCell="A1" sqref="A1"/>
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15"/>
    <cols>
        <col min="1" max="1" width="17" customWidth="1"/>
    </cols>
    <sheetData>
        <row r="1" spans="1:12">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1"><v>42</v></c>
            <c r="C1" t="b"><v>1</v></c>
        </row>
    </sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
```

**Cell Value Types (`t` attribute):**

| Type | Meaning | Value Format |
|------|---------|--------------|
| (none) | Number | Numeric literal in `<v>` |
| `s` | Shared String | Index into sharedStrings.xml |
| `str` | Inline String | String literal in `<v>` |
| `b` | Boolean | `0` or `1` |
| `e` | Error | Error code (#REF!, #VALUE!, etc.) |
| `inlineStr` | Inline Rich Text | `<is><t>text</t></is>` |

**Cell Reference Format:**
- Column: A-Z, AA-ZZ, AAA-XFD (up to 16,384 columns)
- Row: 1 to 1,048,576
- Example: `A1`, `B2`, `AA100`, `XFD1048576`

---

### 6. xl/sharedStrings.xml (Shared String Table)

**Purpose:** Deduplicates string values across all cells. Cells reference strings by index.

**Location:** `xl/sharedStrings.xml`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="580" uniqueCount="77">
    <si><t>First Name</t></si>
    <si><t>Last Name</t></si>
    <si><t>Email</t></si>
    <si><t xml:space="preserve">  Preserved Whitespace  </t></si>
    <si>
        <r><rPr><b/></rPr><t>Bold</t></r>
        <r><t> and normal</t></r>
    </si>
</sst>
```

**Key Points:**
- `count`: Total string references in workbook
- `uniqueCount`: Number of unique strings in table
- `<si>`: String item (0-indexed)
- `<t>`: Text content
- `<r>`: Rich text run (for formatted text)
- `xml:space="preserve"`: Preserves leading/trailing whitespace

---

### 7. xl/styles.xml (Styles)

**Purpose:** Defines fonts, fills, borders, number formats, and cell styles.

**Location:** `xl/styles.xml`

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="1">
        <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    </numFmts>
    <fonts count="2">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
        <font>
            <b/>
            <sz val="11"/>
            <name val="Calibri"/>
        </font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border>
            <left/><right/><top/><bottom/><diagonal/>
        </border>
    </borders>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellXfs>
</styleSheet>
```

**Style Hierarchy:**
1. `<numFmts>`: Custom number formats
2. `<fonts>`: Font definitions
3. `<fills>`: Fill patterns and colors
4. `<borders>`: Border definitions
5. `<cellStyleXfs>`: Named cell styles
6. `<cellXfs>`: Cell format combinations (referenced by cell `s` attribute)

**Built-in Number Formats (numFmtId):**

| ID | Format |
|----|--------|
| 0 | General |
| 1 | 0 |
| 2 | 0.00 |
| 9 | 0% |
| 10 | 0.00% |
| 14 | mm-dd-yy |
| 22 | m/d/yy h:mm |

---

## Relationship System

### How Relationships Work

1. **Part A** needs to reference **Part B**
2. A `.rels` file in `_rels/` directory stores the mapping
3. Relationship file name: `{partname}.rels`
4. Each relationship has:
   - `Id`: Unique identifier (e.g., `rId1`)
   - `Type`: URI describing relationship type
   - `Target`: Relative or absolute path to target part

### Relationship File Locations

| Source Part | Relationship File |
|-------------|-------------------|
| Package root | `_rels/.rels` |
| `xl/workbook.xml` | `xl/_rels/workbook.xml.rels` |
| `xl/worksheets/sheet1.xml` | `xl/worksheets/_rels/sheet1.xml.rels` |

### Common Relationship Types

| Short Name | Full URI |
|------------|----------|
| officeDocument | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| worksheet | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet` |
| sharedStrings | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings` |
| styles | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| theme | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| table | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/table` |

---

## Namespaces

### Primary Namespaces

| Prefix | URI | Usage |
|--------|-----|-------|
| (default) | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | SpreadsheetML elements |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship references |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility |
| `x14` | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` | Excel 2010 extensions |

### Transitional vs Strict

**Transitional (Default):**
- Uses `http://schemas.openxmlformats.org/...` namespaces
- Supports legacy features
- Maximum Excel compatibility

**Strict:**
- Uses `http://purl.oclc.org/ooxml/...` namespaces
- Removed deprecated features
- Better for long-term archival

---

## Implementation Notes for Cuneiform

### Reading Strategy

1. **Open ZIP archive**
2. **Parse `[Content_Types].xml`** to understand what parts exist
3. **Parse `_rels/.rels`** to find workbook location
4. **Parse `xl/workbook.xml`** to get sheet listing
5. **Parse `xl/_rels/workbook.xml.rels`** to map rIds to sheet paths
6. **Parse shared strings** (if present) before worksheets
7. **Parse each worksheet** on demand (lazy loading)

### Writing Strategy

1. **Build all parts in memory**
2. **Generate relationship IDs** sequentially (rId1, rId2, ...)
3. **Write parts to ZIP** in any order
4. **Generate `[Content_Types].xml`** from known parts
5. **Generate relationship files** from relationship graph

### Required vs Optional Parts

**Required for valid .xlsx:**
- `[Content_Types].xml`
- `_rels/.rels`
- `xl/workbook.xml`
- `xl/_rels/workbook.xml.rels`
- At least one `xl/worksheets/sheet{N}.xml`

**Optional but common:**
- `xl/sharedStrings.xml`
- `xl/styles.xml`
- `xl/theme/theme1.xml`
- `docProps/core.xml`
- `docProps/app.xml`

---

## References

- ISO/IEC 29500-1:2016 - Part 1: Fundamentals and Markup Language Reference
- ISO/IEC 29500-2:2012 - Part 2: Open Packaging Conventions
- [MS-OI29500] - Microsoft Office Implementation Notes
