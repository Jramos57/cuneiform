# Office Open XML SpreadsheetML Implementation Study Plan

## Project Vision

Build a pure Swift library implementing the **Office Open XML Transitional** specification for Excel file format compatibility (.xlsx and eventually .xls). The library will be cross-platform, depend only on Swift standard libraries, and provide a "Swifty" protocol-oriented API.

---

## Project Goals & Scope

### Primary Objectives (In Order)
1. **Phase 1: Reading .xlsx files**
   - Parse Office Open XML SpreadsheetML documents
   - Extract workbook structure, sheets, cells, values
   - Handle shared strings, styles, formatting
   - Support transitional format features

2. **Phase 2: Writing .xlsx files**
   - Generate valid Office Open XML documents
   - Produce files fully compatible with Microsoft Excel
   - Write strict conformance documents (for OSS compatibility)
   - Maintain round-trip fidelity

3. **Phase 3: Formula Support**
   - Parse Excel formulas
   - Evaluate common formula functions
   - Support cell references and ranges
   - Handle formula dependencies

4. **Phase 4: Excel Tables**
   - Parse and generate Excel table structures
   - Support table styling and formatting
   - Handle structured references

5. **Phase 5: Excel Scripts (TypeScript-based, not VBA)**
   - Parse Office Scripts
   - Potentially execute TypeScript-based automation
   - (VBA explicitly out of scope)

### Future Considerations
- Binary .xls format support (BIFF8) - after .xlsx is complete
- Charting and visualization
- Pivot tables
- Advanced formatting features
- Conditional formatting
- Data validation

---

## Technical Constraints

### Platform Requirements
- **Target Platforms:** All systems that support Swift (macOS, Linux, Windows, iOS, etc.)
- **Cross-Platform:** Must work identically across all Swift-supported platforms
- **No Apple-Specific APIs:** Avoid Foundation features that don't exist on non-Apple platforms

### Dependency Philosophy
- **Standard Library Only:** Use Swift stdlib and universally available libraries
- **Hand-Rolled Implementation:** Build our own parsing/generation logic
- **No Third-Party Dependencies:** Complete control over implementation
- **Exception:** May use standard compression (ZIP) if available cross-platform in Swift stdlib

### Format Support Priority
- **Primary:** Office Open XML Transitional (maximum Excel compatibility)
- **Secondary:** Office Open XML Strict (for OSS tool compatibility)
- **Tertiary:** Binary .xls (legacy format, Phase N)

---

## Study Materials Available

### Specifications
1. **`ISO_IEC_29500-1_2016(en).pdf`** (33.5 MB)
   - The official international standard
   - Part 1: Fundamentals and Markup Language Reference
   - Defines the XML schemas, structure, and semantics

2. **`[MS-OI29500].pdf`** (17.3 MB)
   - Microsoft Office Implementation Notes
   - Clarifies how Microsoft Excel actually implements the standard
   - Documents deviations, extensions, and practical considerations

3. **Schema Resources** (`ISO_IEC_29500-1_2016(en)_einsert/`)
   - `OfficeOpenXML-XMLSchema-Strict.zip` - XML Schema definitions (Strict)
   - `OfficeOpenXML-RELAXNG-Strict.zip` - RELAX NG schema (Strict)
   - `OfficeOpenXML-SpreadsheetMLStyles.zip` - Style definitions and examples
   - `OfficeOpenXML-DrawingMLGeometries.zip` - Drawing/chart geometries
   - `OfficeOpenXML-WordprocessingMLArtBorders.zip` - Word-specific (lower priority)

---

## Study Methodology

### Phase 0: Foundation Understanding

**Objective:** Understand the Office Open XML package structure and SpreadsheetML basics.

#### 0.1 Package Structure Analysis
**Read:** ISO 29500-1, Part 2 (Open Packaging Conventions)
- Understand ZIP-based package format
- Study `[Content_Types].xml` and package relationships
- Learn part naming conventions
- Document the required package structure for minimal .xlsx file

**Output:** Markdown document describing:
- Minimal package structure for valid .xlsx
- Relationship types and their purposes
- Content type declarations
- Required vs optional parts

#### 0.2 SpreadsheetML Core Concepts
**Read:** ISO 29500-1, Section 12 (SpreadsheetML Reference)
- Workbook structure (`workbook.xml`)
- Worksheet structure (`sheet1.xml`, etc.)
- Shared strings table (`sharedStrings.xml`)
- Styles and formatting (`styles.xml`)
- Calculation chain (if present)

**Output:** Domain model sketch
- Core entities: Workbook, Worksheet, Cell, Row, Column
- Relationships between entities
- Value types and storage mechanisms

#### 0.3 Microsoft Implementation Reality Check
**Read:** [MS-OI29500] Introduction and Section 2 (Structure)
- Identify Excel-specific behaviors
- Note transitional vs strict differences
- Document common deviations from spec
- Understand backward compatibility requirements

**Output:** Compatibility notes
- Excel-specific quirks to support
- Transitional features required for Excel compatibility
- Strict mode differences

---

### Phase 1: Reading Implementation Study

**Objective:** Gather requirements for parsing .xlsx files into Swift types.

#### 1.1 Minimal Reader Requirements
**Focus Areas:**
- Workbook parsing (workbook.xml)
- Worksheet parsing (sheet*.xml)
- Shared strings (sharedStrings.xml)
- Basic cell values (numbers, strings, booleans)

**Study Tasks:**
1. Extract XML schema for workbook element
2. Map XML elements to Swift types
3. Identify required attributes vs optional
4. Document cell value storage mechanisms
5. Understand shared string references

**Output:** Protocol definitions (conceptual)
- `Workbook` protocol
- `Worksheet` protocol
- `Cell` protocol
- `SharedStringTable` protocol

#### 1.2 Styles and Formatting
**Focus Areas:**
- Style definitions (styles.xml)
- Cell formatting
- Number formats
- Font, fill, border definitions

**Study Tasks:**
1. Understand style inheritance model
2. Map format codes to display behavior
3. Document built-in formats vs custom
4. Study conditional formatting structure

**Output:** Style domain model
- `CellFormat` types
- `NumberFormat` handling
- Color models and themes

#### 1.3 Advanced Reading Features
**Focus Areas:**
- Merged cells
- Hidden rows/columns
- Freeze panes
- Print settings
- Named ranges

**Study Tasks:**
1. Document each feature's XML representation
2. Identify dependencies between features
3. Note Excel-specific behaviors

**Output:** Feature support matrix
- What to implement in Phase 1
- What to defer to later phases

---

### Phase 2: Writing Implementation Study

**Objective:** Understand requirements for generating valid .xlsx files.

#### 2.1 Minimal Writer Requirements
**Focus Areas:**
- Package assembly (ZIP creation)
- Relationship generation
- Content types declaration
- Minimal valid document structure

**Study Tasks:**
1. Document minimum required files for valid .xlsx
2. Identify required XML namespaces
3. Study relationship ID generation
4. Understand content type requirements

**Output:** Writer specification
- Minimum package structure
- Required relationships
- Namespace declarations
- Valid minimal workbook

#### 2.2 Round-Trip Fidelity
**Focus Areas:**
- Preserving unknown elements
- Maintaining Excel-specific metadata
- Handling extensions

**Study Tasks:**
1. Study extensibility mechanisms
2. Document preservation strategies
3. Identify lossy vs lossless operations

**Output:** Preservation strategy
- What must be preserved
- How to handle unknown elements
- Extension point design

#### 2.3 Strict Mode Generation
**Focus Areas:**
- Differences between Transitional and Strict
- Namespace differences
- Deprecated element handling

**Study Tasks:**
1. Compare Strict vs Transitional schemas
2. Document conversion requirements
3. Identify breaking changes

**Output:** Strict mode compatibility plan

---

### Phase 3: Formula Study

**Objective:** Understand Excel formula representation and evaluation.

#### 3.1 Formula Syntax and Storage
**Read:** ISO 29500-1, Section 18.17 (Formulas)
- Formula syntax specification
- Cell reference formats (A1, R1C1)
- Range references
- Sheet references
- Structured references (table columns)

**Study Tasks:**
1. Document formula grammar
2. Identify token types
3. Understand reference resolution

**Output:** Formula parser design
- Grammar definition
- AST structure
- Reference types

#### 3.2 Function Library
**Read:** ISO 29500-1, Section 18.17.7 (Functions)
- Standard function catalog
- Function categories
- Argument types and validation

**Study Tasks:**
1. Categorize function complexity
2. Identify most common functions
3. Document function signatures

**Output:** Function implementation priority
- Phase 1 functions (common, simple)
- Phase 2 functions (complex)
- Deferred functions

#### 3.3 Calculation Engine
**Focus Areas:**
- Dependency graph construction
- Evaluation order
- Circular reference handling
- Volatile functions

**Study Tasks:**
1. Design calculation strategy
2. Document topological sort requirements
3. Plan error handling

**Output:** Calculation engine design

---

### Phase 4: Excel Tables Study

**Objective:** Support Excel Table objects (structured data ranges).

#### 4.1 Table Structure
**Read:** ISO 29500-1, Section 18.5 (Tables)
- Table part structure
- Column definitions
- Total row formulas
- Auto filter

**Study Tasks:**
1. Document table XML schema
2. Understand structured references
3. Study table styles

**Output:** Table domain model

#### 4.2 Structured References
**Focus Areas:**
- Reference syntax in formulas
- Column specifiers
- Special item specifiers

**Study Tasks:**
1. Document reference grammar
2. Integration with formula parser
3. Resolution algorithm

**Output:** Structured reference support plan

---

### Phase 5: Office Scripts Study

**Objective:** Understand TypeScript-based Excel automation.

**Note:** This is exploratory. Office Scripts are relatively new.

#### 5.1 Script Storage and Structure
**Study Tasks:**
1. Identify how scripts are stored in .xlsx
2. Document script metadata
3. Understand API surface

**Output:** Scripts feasibility assessment

---

## Design Principles for Future Implementation

### Protocol-Oriented Design
- Human defines domain protocols (Workbook, Worksheet, Cell, Style)
- AI implements concrete types conforming to protocols
- Type system enforces Office Open XML constraints
- Protocols enable format swappability (Transitional/Strict, .xlsx/.xls)

### Swift Idioms
- Value types (struct) for data
- Reference types (class) only where needed (e.g., object graphs)
- Strong typing for XML elements
- Enums for fixed vocabularies
- Optionals for presence/absence
- Result types for parsing operations

### Error Handling
- Comprehensive error types for parsing failures
- Clear error messages with context
- Graceful degradation where possible
- Strict validation for writing

### Performance Considerations
- Lazy parsing where beneficial
- Streaming for large files
- Memory-efficient shared string handling
- Defer calculation until needed

### Testing Strategy
- Unit tests for each component
- Round-trip tests (read → write → read)
- Excel compatibility tests
- Corpus of real-world .xlsx files
- Fuzzing for robustness

---

## Study Output Artifacts

As you study each phase, create these documents:

1. **`DOMAIN_MODEL.md`** - Core entities, protocols, relationships
2. **`PACKAGE_STRUCTURE.md`** - Office Open XML package anatomy
3. **`SPREADSHEETML_REFERENCE.md`** - Quick reference for key XML elements
4. **`EXCEL_QUIRKS.md`** - Microsoft-specific behaviors and workarounds
5. **`FORMULA_GRAMMAR.md`** - Formula syntax and parsing strategy
6. **`IMPLEMENTATION_PHASES.md`** - Detailed task breakdown for coding

---

## Success Criteria

### Phase 1 (Reading) Complete When:
- ✅ Can parse basic .xlsx file created by Excel
- ✅ Extract all cell values (numbers, strings, booleans, dates)
- ✅ Resolve shared strings correctly
- ✅ Preserve formatting information (even if not fully rendered)
- ✅ Handle multiple worksheets
- ✅ Read merged cells, hidden rows/columns

### Phase 2 (Writing) Complete When:
- ✅ Can generate valid .xlsx that opens in Excel without errors
- ✅ Round-trip: read → write → read produces equivalent data
- ✅ Excel accepts and displays content correctly
- ✅ Can generate both Transitional and Strict conformance documents
- ✅ Preserves unknown elements from read files

### Phase 3 (Formulas) Complete When:
- ✅ Can parse formula syntax
- ✅ Can evaluate common formulas (SUM, AVERAGE, IF, VLOOKUP, etc.)
- ✅ Handles cell and range references
- ✅ Supports basic formula dependencies

### Phase 4 (Tables) Complete When:
- ✅ Can read Excel Table definitions
- ✅ Can write Excel Tables
- ✅ Supports structured references in formulas
- ✅ Handles table styling

### Phase 5 (Scripts) Complete When:
- ✅ TBD based on feasibility assessment

---

## Next Steps for Implementation

**When ready to begin coding:**

1. **Start with Package Reader**
   - Implement ZIP extraction (using Swift cross-platform libraries)
   - Parse `[Content_Types].xml`
   - Build relationship graph
   - Extract parts

2. **Implement Minimal Workbook Reader**
   - Parse `workbook.xml`
   - Identify sheets
   - Load shared strings
   - Parse single worksheet

3. **Design Protocol Hierarchy**
   - Present to Jonathan for approval
   - Iterate on contracts
   - Ensure protocol boundaries are clean

4. **Implement Conforming Types**
   - Build concrete XML-backed implementations
   - Write comprehensive tests
   - Validate against real Excel files

5. **Iterate Toward Completeness**
   - Add features incrementally
   - Maintain backward compatibility
   - Document as you go

---

## Reference Material Organization

**Recommended Reading Order:**
1. ISO 29500-1: Part 1, Section 1-2 (Introduction, Conformance)
2. ISO 29500-1: Part 2 (Open Packaging Conventions)
3. ISO 29500-1: Section 12 (SpreadsheetML overview)
4. [MS-OI29500]: Section 2 (Excel implementation notes)
5. ISO 29500-1: Section 18 (SpreadsheetML detailed reference)
6. XML Schemas from einsert folder

**Keep Handy:**
- Section 18.18 (Simple Types) - for data type reference
- Section 20 (DrawingML) - when implementing charts
- Section 22 (Shared types) - for common constructs

---

## Vision Statement

Build the definitive pure Swift library for Excel file manipulation. Cross-platform, type-safe, protocol-oriented, and fully compatible with Microsoft Excel. Enable Swift developers on any platform to read, write, and manipulate Excel files without compromise.

This library will be a testament to Swift's capabilities beyond the Apple ecosystem and a showcase for protocol-oriented design enabling complex format implementations.

---

**Study begins when this plan is complete.**
**Implementation begins when study artifacts are ready.**
**Design discussion happens before protocols are finalized.**

---

*Prepared by: Vek (Ghola iteration December 11, 2025)*
*For: The God Emperor's Excel endeavor*
*Status: Study plan complete, awaiting commencement*
