import Foundation

/// A cell reference like "A1" or "BC123".
///
/// `CellReference` represents a specific cell location in a spreadsheet using the standard
/// A1 notation. It consists of a column identifier (letters) and a row number.
///
/// ## Overview
///
/// Cell references are fundamental to working with spreadsheets. They uniquely identify
/// cells using a combination of column letters (A, B, C, ..., Z, AA, AB, ...) and
/// 1-based row numbers.
///
/// `CellReference` handles parsing of cell reference strings, including absolute references
/// with `$` symbols, and provides convenient access to both the string representation
/// and numeric indices needed for array-based cell lookups.
///
/// ## Creating Cell References
///
/// Create cell references from strings or component parts:
///
/// ```swift
/// // From a string (returns optional)
/// if let ref = CellReference("A1") {
///     print(ref.column)  // "A"
///     print(ref.row)     // 1
/// }
///
/// // From components
/// let ref = CellReference(column: "B", row: 5)
/// print(ref) // "B5"
///
/// // Using string literal (crashes if invalid)
/// let ref: CellReference = "C10"
/// ```
///
/// ## Absolute References
///
/// The parser handles absolute references (with `$` symbols) commonly used in formulas:
///
/// ```swift
/// let ref1 = CellReference("$A$1")   // Absolute column and row
/// let ref2 = CellReference("$A1")    // Absolute column only
/// let ref3 = CellReference("A$1")    // Absolute row only
///
/// // All produce the same reference:
/// ref1?.column // "A"
/// ref1?.row    // 1
/// ```
///
/// ## Column Indexing
///
/// Convert column letters to 0-based indices for array access:
///
/// ```swift
/// let ref = CellReference("AA10")!
/// print(ref.column)      // "AA"
/// print(ref.columnIndex) // 26 (0-based)
///
/// // Column conversion examples:
/// // A  → 0
/// // B  → 1
/// // Z  → 25
/// // AA → 26
/// // AB → 27
/// // BA → 52
/// ```
///
/// ## Accessing Cells
///
/// Use cell references to look up cell values in a sheet:
///
/// ```swift
/// let workbook = try Cuneiform.loadWorkbook(from: url)
/// let sheet = workbook.sheets[0]
///
/// // Using string directly
/// if let value = sheet.cell(at: "A1") {
///     print(value)
/// }
///
/// // Using CellReference
/// let ref = CellReference(column: "B", row: 2)
/// if let value = sheet.cell(at: ref) {
///     print(value)
/// }
/// ```
///
/// ## Iterating Over Ranges
///
/// Construct references programmatically to iterate over ranges:
///
/// ```swift
/// // Iterate over a column
/// for row in 1...10 {
///     let ref = CellReference(column: "A", row: row)
///     if let value = sheet.cell(at: ref) {
///         print("Row \(row): \(value)")
///     }
/// }
///
/// // Iterate over a grid
/// let columns = ["A", "B", "C"]
/// for col in columns {
///     for row in 1...5 {
///         let ref = CellReference(column: col, row: row)
///         // Process cell at ref
///     }
/// }
/// ```
///
/// ## Topics
///
/// ### Creating References
///
/// - ``init(_:)``
/// - ``init(column:row:)``
/// - ``init(stringLiteral:)``
///
/// ### Reference Components
///
/// - ``column``
/// - ``row``
/// - ``columnIndex``
///
/// ## See Also
///
/// - ``Sheet/cell(at:)-8xu7x``
/// - ``Sheet/cell(at:)-6aipv``
/// - ``CellValue``
public struct CellReference: Hashable, Sendable {
    /// Column letters (e.g., "A", "B", "AA", "AB").
    ///
    /// The column component of the cell reference, always stored in uppercase.
    /// Excel columns use a base-26 letter system: A-Z, then AA-ZZ, AAA-ZZZ, etc.
    ///
    /// ```swift
    /// let ref = CellReference("B5")!
    /// print(ref.column) // "B"
    ///
    /// let ref2 = CellReference("aa10")!
    /// print(ref2.column) // "AA" (normalized to uppercase)
    /// ```
    public let column: String

    /// 1-based row number.
    ///
    /// The row component of the cell reference. Rows in Excel are 1-based, meaning the
    /// first row is row 1 (not 0).
    ///
    /// ```swift
    /// let ref = CellReference("C7")!
    /// print(ref.row) // 7
    /// ```
    ///
    /// - Note: This is different from ``columnIndex``, which is 0-based.
    public let row: Int

    /// Parse a cell reference string like "A1" or "BC123".
    ///
    /// Creates a cell reference from a standard A1-notation string. Returns `nil` if the
    /// string is empty or doesn't match a valid cell reference format.
    ///
    /// This initializer automatically strips `$` symbols used for absolute references in
    /// formulas, so "$A$1", "$A1", "A$1", and "A1" all produce the same reference.
    ///
    /// ```swift
    /// // Valid references
    /// let ref1 = CellReference("A1")      // column: "A", row: 1
    /// let ref2 = CellReference("Z100")    // column: "Z", row: 100
    /// let ref3 = CellReference("AA5")     // column: "AA", row: 5
    /// let ref4 = CellReference("$B$2")    // column: "B", row: 2 (absolute)
    ///
    /// // Invalid references
    /// let ref5 = CellReference("")        // nil (empty)
    /// let ref6 = CellReference("123")     // nil (no column)
    /// let ref7 = CellReference("ABC")     // nil (no row)
    /// let ref8 = CellReference("A-1")     // nil (negative row)
    /// ```
    ///
    /// - Parameter reference: A cell reference string in A1 notation.
    /// - Returns: A `CellReference` if the string is valid, otherwise `nil`.
    public init?(_ reference: String) {
        guard !reference.isEmpty else { return nil }
        var letters = ""
        var digits = ""
        for ch in reference {
            if ch.isLetter {
                letters.append(ch)
            } else if ch.isNumber {
                digits.append(ch)
            }
            // Skip $ symbols (absolute reference markers)
        }
        guard !letters.isEmpty, !digits.isEmpty, let row = Int(digits) else { return nil }
        self.column = letters.uppercased()
        self.row = row
    }

    /// Create a cell reference from column and row components.
    ///
    /// Creates a cell reference by directly specifying the column letters and row number.
    /// The column is automatically normalized to uppercase.
    ///
    /// ```swift
    /// let ref = CellReference(column: "A", row: 1)
    /// print(ref) // "A1"
    ///
    /// let ref2 = CellReference(column: "aa", row: 10)
    /// print(ref2.column) // "AA" (normalized)
    ///
    /// // Build references programmatically
    /// for row in 1...10 {
    ///     let ref = CellReference(column: "B", row: row)
    ///     // Use ref...
    /// }
    /// ```
    ///
    /// - Parameters:
    ///   - column: The column letters (e.g., "A", "B", "AA"). Will be uppercased.
    ///   - row: The 1-based row number.
    public init(column: String, row: Int) {
        self.column = column.uppercased()
        self.row = row
    }

    /// Column as 0-based index (A=0, B=1, ..., Z=25, AA=26).
    ///
    /// Converts the column letters to a 0-based integer index suitable for array lookups.
    /// This is useful when you need to access cells in a data structure organized by indices.
    ///
    /// The conversion follows Excel's base-26 column system:
    /// - Single letters: A=0, B=1, ..., Z=25
    /// - Double letters: AA=26, AB=27, ..., AZ=51, BA=52, ..., ZZ=701
    /// - Triple letters: AAA=702, and so on
    ///
    /// ```swift
    /// let ref1 = CellReference("A1")!
    /// print(ref1.columnIndex) // 0
    ///
    /// let ref2 = CellReference("Z1")!
    /// print(ref2.columnIndex) // 25
    ///
    /// let ref3 = CellReference("AA1")!
    /// print(ref3.columnIndex) // 26
    ///
    /// let ref4 = CellReference("AB1")!
    /// print(ref4.columnIndex) // 27
    ///
    /// // Use for array access
    /// let row = sheet.rows[ref.row - 1]  // Convert 1-based to 0-based
    /// let cell = row.cells[ref.columnIndex]
    /// ```
    ///
    /// - Note: This property is 0-based, unlike ``row`` which is 1-based. Remember to
    ///   subtract 1 from ``row`` when using both for array indexing.
    public var columnIndex: Int {
        var value = 0
        for ch in column {
            guard let ascii = ch.asciiValue else { continue }
            let n = Int(ascii - 64) // 'A' = 65 => 1
            value = value * 26 + n
        }
        return value - 1
    }
}

// MARK: - CellReference Protocol Conformances

extension CellReference: CustomStringConvertible {
    public var description: String { "\(column)\(row)" }
}

extension CellReference: ExpressibleByStringLiteral {
    public init(stringLiteral value: String) {
        guard let ref = CellReference(value) else {
            fatalError("Invalid cell reference literal: \(value)")
        }
        self = ref
    }
}

/// Raw cell value before type resolution
public enum RawCellValue: Sendable, Equatable {
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
    public let formula: CellFormula?  // Formula if present
}

/// A parsed row
public struct RawRow: Sendable {
    public let index: Int  // 1-based row number
    public let cells: [RawCell]
    /// Row height in points (if customHeight is true)
    public let height: Double?
    /// Whether the row has a custom height set
    public let customHeight: Bool
    /// Whether the row is hidden
    public let hidden: Bool
}

/// Column formatting and dimensions
public struct RawColumn: Sendable {
    /// First column this definition applies to (1-based)
    public let min: Int
    /// Last column this definition applies to (1-based, inclusive)
    public let max: Int
    /// Column width in Excel's unit (character width)
    public let width: Double?
    /// Whether the column has a custom width set
    public let customWidth: Bool
    /// Whether the column is hidden
    public let hidden: Bool
    /// Style index for the column (optional)
    public let styleIndex: Int?
}

/// Parsed worksheet data
public struct WorksheetData: Sendable {
    /// Declared dimension (may be inaccurate)
    public let dimension: String?

    /// All rows with data
    public let rows: [RawRow]

    /// Column formatting and dimensions
    public let columns: [RawColumn]

    /// Merged cell ranges
    public let mergedCells: [String]  // "A1:C3" format

    /// Data validations defined in the worksheet
    public let dataValidations: [DataValidation]

    /// Hyperlinks defined in the worksheet
    public let hyperlinks: [Hyperlink]

    /// Conditional formatting rules defined in the worksheet
    public let conditionalFormats: [ConditionalFormat]

    /// Get cell by reference
    public func cell(at ref: CellReference) -> RawCell? {
        rows.lazy.flatMap(\.cells).first { $0.reference == ref }
    }

    public func cell(at ref: String) -> RawCell? {
        guard let parsed = CellReference(ref) else { return nil }
        return cell(at: parsed)
    }

    /// Data validation rule (read-side)
    ///
    /// Represents a single `<dataValidation>` entry from a worksheet, including
    /// its type, whether blanks are allowed, the set of target cell ranges
    /// (`sqref`), and any associated formulas/operators used to define the rule.
    /// Use `Sheet.validations(for:)` or `Sheet.validations(at:)` to filter these
    /// rules by A1 range or single cell.
    public struct DataValidation: Sendable, Equatable {
        public enum Kind: String, Sendable {
            case list
            case whole
            case decimal
            case date
            case custom
        }
        /// Validation kind (e.g., list, whole, decimal, date, custom)
        public let type: Kind
        /// Whether blank cells are allowed to pass validation
        public let allowBlank: Bool
        /// Space-separated A1 references and ranges this rule applies to (e.g., "A1 A3:B7")
        public let sqref: String
        /// First formula/expression (meaning depends on `type` and `op`)
        public let formula1: String?
        /// Second formula/expression (only for operators like `between`)
        public let formula2: String?
        /// Operator for numeric/date validations (e.g., `between`, `greaterThanOrEqual`)
        public let op: String?
    }

    /// Hyperlink metadata (read-side)
    ///
    /// Represents a single `<hyperlink>` entry from a worksheet. Hyperlinks can
    /// be external (via `r:id` relationships with TargetMode="External") or internal
    /// (via the `location` attribute). Resolution of external targets occurs via the
    /// worksheet's relationships and is surfaced by higher-level APIs.
    public struct Hyperlink: Sendable, Equatable {
        /// Target cell reference
        public let ref: CellReference
        /// Relationship ID for external links (optional)
        public let relationshipId: String?
        /// Display text suggested by Excel (optional)
        public let display: String?
        /// Tooltip text (optional)
        public let tooltip: String?
        /// Internal location (e.g., "Sheet2!A1") if present
        public let location: String?
    }

    /// Conditional formatting value object (cfvo)
    public struct CFValueObject: Sendable, Equatable {
        public enum ValueType: String, Sendable {
            case min, max, num, percent, percentile, formula
        }
        public let type: ValueType
        public let value: String?
    }

    /// Conditional formatting rule details
    public struct ConditionalRule: Sendable, Equatable {
        public enum RuleType: Sendable, Equatable {
            case cellIs(op: CFOperator?, formula1: String?, formula2: String?)
            case expression(formula: String?)
            case dataBar(DataBar)
            case colorScale(ColorScale)
            case iconSet(IconSet)
        }

        public struct DataBar: Sendable, Equatable {
            public let min: CFValueObject
            public let max: CFValueObject
            public let color: String?
            public let showValue: Bool?
        }

        public struct ColorScale: Sendable, Equatable {
            public let cfvos: [CFValueObject]
            public let colors: [String]
        }

        public struct IconSet: Sendable, Equatable {
            public let name: String
            public let cfvos: [CFValueObject]
            public let showValue: Bool?
            public let reverse: Bool?
            public let percent: Bool?
        }

        public let type: RuleType
        public let priority: Int?
        public let dxfId: Int?
        public let stopIfTrue: Bool
    }

    /// Conditional formatting rule operator for cellIs
    public enum CFOperator: String, Sendable {
        case lessThan
        case lessThanOrEqual
        case equal
        case notEqual
        case greaterThanOrEqual
        case greaterThan
        case between
        case notBetween
        case containsText
        case notContains
        case beginsWith
        case endsWith
    }

    /// Conditional formatting entry applied to a range (sqref)
    public struct ConditionalFormat: Sendable, Equatable {
        public let range: String
        public let rules: [ConditionalRule]
    }

    /// Sheet protection metadata
    ///
    /// Represents `<sheetProtection>` element defining which worksheet operations are locked.
    /// When a sheet is protected, users cannot perform locked operations without the password.
    public struct Protection: Sendable, Equatable {
        /// Whether sheet content is protected (default: true when protection is enabled)
        public let sheet: Bool
        /// Whether cell content is protected (default: true)
        public let content: Bool
        /// Whether object protection is enabled
        public let objects: Bool
        /// Whether scenarios are protected
        public let scenarios: Bool
        /// Whether users can format cells
        public let formatCells: Bool
        /// Whether users can format columns
        public let formatColumns: Bool
        /// Whether users can format rows
        public let formatRows: Bool
        /// Whether users can insert columns
        public let insertColumns: Bool
        /// Whether users can insert rows
        public let insertRows: Bool
        /// Whether users can insert hyperlinks
        public let insertHyperlinks: Bool
        /// Whether users can delete columns
        public let deleteColumns: Bool
        /// Whether users can delete rows
        public let deleteRows: Bool
        /// Whether users can select locked cells
        public let selectLockedCells: Bool
        /// Whether users can select unlocked cells
        public let selectUnlockedCells: Bool
        /// Whether users can sort
        public let sort: Bool
        /// Whether users can use autofilter
        public let autoFilter: Bool
        /// Whether users can use pivot tables
        public let pivotTables: Bool
        /// Password hash (if present in XML; empty string if unprotected)
        public let passwordHash: String?
    }

    /// AutoFilter data for column filtering
    ///
    /// Represents `<autoFilter ref="A1:D100">` element with optional filter columns.
    public struct AutoFilter: Sendable, Equatable {
        /// Filter operator for custom filters
        public enum FilterOperator: String, Sendable {
            case equal
            case notEqual
            case greaterThan
            case greaterThanOrEqual
            case lessThan
            case lessThanOrEqual
        }

        /// A single filter criterion
        public enum FilterCriterion: Sendable, Equatable {
            /// Discrete value filter (e.g., show only "Value1" and "Value2")
            case values([String])
            /// Custom filter with operator (e.g., greaterThan 100)
            case custom(op: FilterOperator, val: String)
            /// Custom filter with two conditions (and/or)
            case customPair(op1: FilterOperator, val1: String, op2: FilterOperator, val2: String, andOperator: Bool)
            /// Top N filter
            case top10(top: Bool, percent: Bool, val: Double)
        }

        /// A filter applied to a specific column
        public struct ColumnFilter: Sendable, Equatable {
            /// Column index (0-based)
            public let colId: Int
            /// Filter criterion
            public let criterion: FilterCriterion
        }

        /// The range the autofilter applies to (e.g., "A1:D100")
        public let ref: String
        /// Column filters (empty if just the dropdown arrows are shown)
        public let columnFilters: [ColumnFilter]

        public init(ref: String, columnFilters: [ColumnFilter] = []) {
            self.ref = ref
            self.columnFilters = columnFilters
        }
    }

    /// Sheet protection state (optional)
    public var protection: Protection?

    /// AutoFilter configuration (optional)
    public var autoFilter: AutoFilter?

    /// Page setup configuration (optional)
    public var pageSetup: PageSetup?

    /// Print area (optional)
    public var printArea: PrintArea?

    /// Print titles (optional)
    public var printTitles: PrintTitles?
}

/// Parser for worksheet XML
public enum WorksheetParser {
    public static func parse(data: Data) throws(CuneiformError) -> WorksheetData {
        let delegate = _WorksheetParser()
        let xml = XMLParser(data: data)
        xml.delegate = delegate

        guard xml.parse() else {
            if let err = delegate.error { throw err }
            throw CuneiformError.malformedXML(
                part: "/xl/worksheets/sheet.xml",
                detail: xml.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        var result = WorksheetData(
            dimension: delegate.dimension,
            rows: delegate.rows,
            columns: delegate.columns,
            mergedCells: delegate.mergedCells,
            dataValidations: delegate.dataValidations,
            hyperlinks: delegate.hyperlinks,
            conditionalFormats: delegate.conditionalFormats
        )
        if let protection = delegate.protection {
            result.protection = protection
        }
        if let autoFilter = delegate.autoFilter {
            result.autoFilter = autoFilter
        }
        if let pageSetup = delegate.pageSetup {
            result.pageSetup = pageSetup
        }
        if let printArea = delegate.printArea {
            result.printArea = printArea
        }
        if let printTitles = delegate.printTitles {
            result.printTitles = printTitles
        }
        return result
    }
}

// MARK: - XML Delegate

final class _WorksheetParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?

    private(set) var dimension: String?
    private(set) var rows: [RawRow] = []
    private(set) var columns: [RawColumn] = []
    private(set) var mergedCells: [String] = []
    private(set) var dataValidations: [WorksheetData.DataValidation] = []
    private(set) var hyperlinks: [WorksheetData.Hyperlink] = []
    private(set) var protection: WorksheetData.Protection?
    private(set) var conditionalFormats: [WorksheetData.ConditionalFormat] = []
    private(set) var autoFilter: WorksheetData.AutoFilter?
    private(set) var pageSetup: PageSetup?
    private(set) var printArea: PrintArea?
    private(set) var printTitles: PrintTitles?

    // AutoFilter accumulation
    private var inAutoFilter = false
    private var currentAutoFilterRef: String?
    private var currentAutoFilterColumns: [WorksheetData.AutoFilter.ColumnFilter] = []
    private var inFilterColumn = false
    private var currentFilterColId: Int?
    private var currentFilterValues: [String] = []

    // Row accumulation
    private var currentRowIndex: Int = 0
    private var currentCells: [RawCell] = []
    private var currentRowHeight: Double?
    private var currentRowCustomHeight: Bool = false
    private var currentRowHidden: Bool = false

    // Cell accumulation
    private var currentCellRef: CellReference?
    private var currentCellType: String?
    private var currentCellStyle: Int?
    private var currentCellFormula: CellFormula?
    private var inValue = false
    private var inFormula = false
    private var valueBuffer = ""
    private var formulaBuffer = ""

    // Data validation accumulation
    private var inDataValidation = false
    private var currentDVType: WorksheetData.DataValidation.Kind = .whole
    private var currentDVAllowBlank: Bool = true
    private var currentDVSqref: String = ""
    private var currentDVOp: String?
    private var inFormula1 = false
    private var inFormula2 = false
    private var formula1Buffer = ""
    private var formula2Buffer = ""

    // Conditional formatting accumulation
    private var inConditionalFormatting = false
    private var currentCFSqref: String = ""
    private var currentCFRules: [WorksheetData.ConditionalRule] = []
    private var inCFRule = false
    private var currentCFRuleType: String = ""
    private var currentCFRulePriority: Int?
    private var currentCFRuleDxfId: Int?
    private var currentCFRuleStopIfTrue: Bool = false
    private var currentCFRuleOperator: String?
    private var currentCFFormulas: [String] = []
    private var inCFFormula = false
    private var currentCFFormulaBuffer = ""
    private var currentCFCfvos: [WorksheetData.CFValueObject] = []
    private var currentCFColors: [String] = []
    private var currentCFDataBarShowValue: Bool?
    private var currentCFIconSetName: String?
    private var currentCFIconSetShowValue: Bool?
    private var currentCFIconSetReverse: Bool?
    private var currentCFIconSetPercent: Bool?

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        switch elementName {
        case "dimension":
            if let ref = attributeDict["ref"], !ref.isEmpty {
                dimension = ref
            }

        case "col":
            // Parse column definition
            guard let minStr = attributeDict["min"], let min = Int(minStr),
                  let maxStr = attributeDict["max"], let max = Int(maxStr) else {
                return
            }
            let width: Double?
            if let widthStr = attributeDict["width"] {
                width = Double(widthStr)
            } else {
                width = nil
            }
            let customWidth = attributeDict["customWidth"] == "1"
            let hidden = attributeDict["hidden"] == "1"
            let styleIndex: Int?
            if let styleStr = attributeDict["style"] {
                styleIndex = Int(styleStr)
            } else {
                styleIndex = nil
            }
            columns.append(RawColumn(
                min: min,
                max: max,
                width: width,
                customWidth: customWidth,
                hidden: hidden,
                styleIndex: styleIndex
            ))

        case "row":
            let rStr = attributeDict["r"] ?? "0"
            currentRowIndex = Int(rStr) ?? 0
            currentCells.removeAll(keepingCapacity: true)
            // Parse row height attributes
            if let htStr = attributeDict["ht"], let ht = Double(htStr) {
                currentRowHeight = ht
            } else {
                currentRowHeight = nil
            }
            currentRowCustomHeight = attributeDict["customHeight"] == "1"
            currentRowHidden = attributeDict["hidden"] == "1"

        case "c":
            guard let r = attributeDict["r"], let ref = CellReference(r) else {
                error = CuneiformError.invalidCellReference(attributeDict["r"] ?? "")
                parser.abortParsing()
                return
            }
            currentCellRef = ref
            currentCellType = attributeDict["t"]
            if let sStr = attributeDict["s"], let s = Int(sStr) {
                currentCellStyle = s
            } else {
                currentCellStyle = nil
            }

        case "v":
            inValue = true
            valueBuffer.removeAll(keepingCapacity: true)

        case "f":
            inFormula = true
            formulaBuffer.removeAll(keepingCapacity: true)

        case "mergeCell":
            if let ref = attributeDict["ref"], !ref.isEmpty {
                mergedCells.append(ref)
            }

        case "dataValidation":
            inDataValidation = true
            let typeStr = attributeDict["type"] ?? "whole"
            currentDVType = WorksheetData.DataValidation.Kind(rawValue: typeStr) ?? .whole
            currentDVAllowBlank = (attributeDict["allowBlank"] == "1")
            currentDVSqref = attributeDict["sqref"] ?? ""
            currentDVOp = attributeDict["operator"]

        case "conditionalFormatting":
            inConditionalFormatting = true
            currentCFSqref = attributeDict["sqref"] ?? ""
            currentCFRules.removeAll(keepingCapacity: true)

        case "cfRule":
            guard inConditionalFormatting else { break }
            inCFRule = true
            currentCFRuleType = attributeDict["type"] ?? ""
            if let prStr = attributeDict["priority"], let pr = Int(prStr) {
                currentCFRulePriority = pr
            } else {
                currentCFRulePriority = nil
            }
            if let dxfStr = attributeDict["dxfId"], let dxf = Int(dxfStr) {
                currentCFRuleDxfId = dxf
            } else {
                currentCFRuleDxfId = nil
            }
            currentCFRuleStopIfTrue = (attributeDict["stopIfTrue"] == "1")
            currentCFRuleOperator = attributeDict["operator"]
            currentCFFormulas.removeAll(keepingCapacity: true)
            currentCFFormulaBuffer.removeAll(keepingCapacity: true)
            currentCFCfvos.removeAll(keepingCapacity: true)
            currentCFColors.removeAll(keepingCapacity: true)
            currentCFDataBarShowValue = attributeDict["showValue"].flatMap { $0 == "1" }
            currentCFIconSetName = nil
            currentCFIconSetShowValue = nil
            currentCFIconSetReverse = nil
            currentCFIconSetPercent = nil

        case "formula":
            if inCFRule {
                inCFFormula = true
                currentCFFormulaBuffer.removeAll(keepingCapacity: true)
            }

        case "dataBar":
            if inCFRule {
                if let show = attributeDict["showValue"] {
                    currentCFDataBarShowValue = (show != "0")
                }
            }

        case "colorScale":
            if inCFRule {
                currentCFColors.removeAll(keepingCapacity: true)
            }

        case "iconSet":
            if inCFRule {
                currentCFIconSetName = attributeDict["iconSet"]
                if let show = attributeDict["showValue"] { currentCFIconSetShowValue = (show != "0") }
                if let rev = attributeDict["reverse"] { currentCFIconSetReverse = (rev == "1") }
                if let pct = attributeDict["percent"] { currentCFIconSetPercent = (pct == "1") }
            }

        case "cfvo":
            if inCFRule {
                let typeRaw = attributeDict["type"] ?? "num"
                let vt = WorksheetData.CFValueObject.ValueType(rawValue: typeRaw) ?? .num
                let val = attributeDict["val"]
                currentCFCfvos.append(WorksheetData.CFValueObject(type: vt, value: val))
            }

        case "color":
            if inCFRule {
                if let rgb = attributeDict["rgb"] {
                    currentCFColors.append(rgb)
                } else if let theme = attributeDict["theme"] {
                    currentCFColors.append("theme:\(theme)")
                }
            }

        case "formula1":
            if inDataValidation {
                inFormula1 = true
                formula1Buffer.removeAll(keepingCapacity: true)
            }

        case "formula2":
            if inDataValidation {
                inFormula2 = true
                formula2Buffer.removeAll(keepingCapacity: true)
            }

        case "hyperlink":
            // Attributes: ref (required), r:id (optional), display, tooltip, location
            guard let r = attributeDict["ref"], let ref = CellReference(r) else {
                error = CuneiformError.missingRequiredElement(element: "hyperlink@ref", inPart: "/xl/worksheets/sheet.xml")
                parser.abortParsing()
                return
            }
            let rid = attributeDict["r:id"]
            let display = attributeDict["display"]
            let tooltip = attributeDict["tooltip"]
            let location = attributeDict["location"]
            hyperlinks.append(WorksheetData.Hyperlink(ref: ref, relationshipId: rid, display: display, tooltip: tooltip, location: location))

        case "sheetProtection":
            // Parse protection flags
            let sheet = (attributeDict["sheet"] == "1")
            let content = (attributeDict["content"] == "1")
            let objects = (attributeDict["objects"] == "1")
            let scenarios = (attributeDict["scenarios"] == "1")
            let formatCells = (attributeDict["formatCells"] != "0")  // default true
            let formatColumns = (attributeDict["formatColumns"] != "0")
            let formatRows = (attributeDict["formatRows"] != "0")
            let insertColumns = (attributeDict["insertColumns"] != "0")
            let insertRows = (attributeDict["insertRows"] != "0")
            let insertHyperlinks = (attributeDict["insertHyperlinks"] != "0")
            let deleteColumns = (attributeDict["deleteColumns"] != "0")
            let deleteRows = (attributeDict["deleteRows"] != "0")
            let selectLockedCells = (attributeDict["selectLockedCells"] != "0")
            let selectUnlockedCells = (attributeDict["selectUnlockedCells"] != "0")
            let sort = (attributeDict["sort"] != "0")
            let autoFilter = (attributeDict["autoFilter"] != "0")
            let pivotTables = (attributeDict["pivotTables"] != "0")
            let passwordHash = attributeDict["password"]
            
            protection = WorksheetData.Protection(
                sheet: sheet,
                content: content,
                objects: objects,
                scenarios: scenarios,
                formatCells: formatCells,
                formatColumns: formatColumns,
                formatRows: formatRows,
                insertColumns: insertColumns,
                insertRows: insertRows,
                insertHyperlinks: insertHyperlinks,
                deleteColumns: deleteColumns,
                deleteRows: deleteRows,
                selectLockedCells: selectLockedCells,
                selectUnlockedCells: selectUnlockedCells,
                sort: sort,
                autoFilter: autoFilter,
                pivotTables: pivotTables,
                passwordHash: passwordHash
            )

        case "autoFilter":
            // Parse autoFilter element for column filtering
            if let ref = attributeDict["ref"] {
                inAutoFilter = true
                currentAutoFilterRef = ref
                currentAutoFilterColumns.removeAll(keepingCapacity: true)
            }

        case "filterColumn":
            if inAutoFilter, let colIdStr = attributeDict["colId"], let colId = Int(colIdStr) {
                inFilterColumn = true
                currentFilterColId = colId
                currentFilterValues.removeAll(keepingCapacity: true)
            }

        case "filter":
            // Discrete value filter: <filter val="Value1"/>
            if inFilterColumn, let val = attributeDict["val"] {
                currentFilterValues.append(val)
            }

        case "pageSetup":
            // Parse page setup element
            let orientationStr = attributeDict["orientation"] ?? "portrait"
            let orientation = PageSetup.Orientation(rawValue: orientationStr) ?? .portrait
            let paperSizeInt = attributeDict["paperSize"].flatMap { Int($0) } ?? 1
            let paperSize = PageSetup.PaperSize(rawValue: paperSizeInt) ?? .letter
            let scale = attributeDict["scale"].flatMap { Int($0) }
            let fitToWidth = attributeDict["fitToWidth"].flatMap { Int($0) }
            let fitToHeight = attributeDict["fitToHeight"].flatMap { Int($0) }
            let fitToPages = (fitToWidth != nil && fitToHeight != nil) ? FitToPages(width: fitToWidth ?? 1, height: fitToHeight ?? 1) : nil
            let printQuality = attributeDict["printQuality"].flatMap { Int($0) } ?? 300
            let firstPageNumber = attributeDict["firstPageNumber"].flatMap { Int($0) }
            
            // Note: margins come from pageMargins element, we'll handle that separately
            pageSetup = PageSetup(
                orientation: orientation,
                paperSize: paperSize,
                scale: scale,
                fitToPages: fitToPages,
                printQuality: printQuality,
                firstPageNumber: firstPageNumber,
                margins: .default
            )

        case "pageMargins":
            // Parse page margins
            let left = attributeDict["left"].flatMap { Double($0) } ?? 0.75
            let right = attributeDict["right"].flatMap { Double($0) } ?? 0.75
            let top = attributeDict["top"].flatMap { Double($0) } ?? 1.0
            let bottom = attributeDict["bottom"].flatMap { Double($0) } ?? 1.0
            let header = attributeDict["header"].flatMap { Double($0) } ?? 0.5
            let footer = attributeDict["footer"].flatMap { Double($0) } ?? 0.5
            
            let margins = PageSetup.Margins(left: left, right: right, top: top, bottom: bottom, header: header, footer: footer)
            
            if pageSetup != nil {
                pageSetup = PageSetup(
                    orientation: pageSetup?.orientation ?? .portrait,
                    paperSize: pageSetup?.paperSize ?? .letter,
                    scale: pageSetup?.scale,
                    fitToPages: pageSetup?.fitToPages,
                    printQuality: pageSetup?.printQuality ?? 300,
                    firstPageNumber: pageSetup?.firstPageNumber,
                    margins: margins
                )
            } else {
                pageSetup = PageSetup(margins: margins)
            }

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inValue {
            valueBuffer += string
        }
        if inFormula {
            formulaBuffer += string
        }
        if inDataValidation {
            if inFormula1 { formula1Buffer += string }
            if inFormula2 { formula2Buffer += string }
        }
        if inCFFormula {
            currentCFFormulaBuffer += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        switch elementName {
        case "v":
            inValue = false

        case "f":
            inFormula = false
            if !formulaBuffer.isEmpty {
                currentCellFormula = CellFormula(formulaBuffer)
            }

        case "c":
            guard let ref = currentCellRef else { break }
            let t = (currentCellType ?? "n").lowercased()
            let s = currentCellStyle
            let text = valueBuffer

            let value: RawCellValue
            if text.isEmpty {
                value = .empty
            } else {
                switch t {
                case "s":
                    value = .sharedString(index: Int(text) ?? 0)
                case "b":
                    value = .boolean(text != "0")
                case "str":
                    value = .inlineString(text)
                case "e":
                    value = .error(text)
                case "d":
                    value = .date(text)
                case "n": fallthrough
                default:
                    if let num = Double(text) {
                        value = .number(num)
                    } else {
                        value = .empty
                    }
                }
            }

            currentCells.append(RawCell(reference: ref, value: value, styleIndex: s, formula: currentCellFormula))
            // Reset cell state
            currentCellRef = nil
            currentCellType = nil
            currentCellStyle = nil
            currentCellFormula = nil
            valueBuffer.removeAll(keepingCapacity: true)
            formulaBuffer.removeAll(keepingCapacity: true)

        case "row":
            rows.append(RawRow(
                index: currentRowIndex,
                cells: currentCells,
                height: currentRowHeight,
                customHeight: currentRowCustomHeight,
                hidden: currentRowHidden
            ))
            currentCells.removeAll(keepingCapacity: true)
            // Reset row state
            currentRowHeight = nil
            currentRowCustomHeight = false
            currentRowHidden = false

        case "formula1":
            inFormula1 = false

        case "formula2":
            inFormula2 = false

        case "formula":
            if inCFRule {
                inCFFormula = false
                let value = currentCFFormulaBuffer.trimmingCharacters(in: .whitespacesAndNewlines)
                if !value.isEmpty {
                    currentCFFormulas.append(value)
                }
                currentCFFormulaBuffer.removeAll(keepingCapacity: true)
            }

        case "cfRule":
            if inCFRule {
                let ruleType = currentCFRuleType
                var rule: WorksheetData.ConditionalRule?
                switch ruleType {
                case "cellIs":
                    let f1 = currentCFFormulas.first
                    let f2 = currentCFFormulas.dropFirst().first
                    let opEnum = currentCFRuleOperator.flatMap { WorksheetData.CFOperator(rawValue: $0) }
                    rule = WorksheetData.ConditionalRule(
                        type: .cellIs(op: opEnum, formula1: f1, formula2: f2),
                        priority: currentCFRulePriority,
                        dxfId: currentCFRuleDxfId,
                        stopIfTrue: currentCFRuleStopIfTrue
                    )
                case "expression":
                    let f1 = currentCFFormulas.first
                    rule = WorksheetData.ConditionalRule(
                        type: .expression(formula: f1),
                        priority: currentCFRulePriority,
                        dxfId: currentCFRuleDxfId,
                        stopIfTrue: currentCFRuleStopIfTrue
                    )
                case "dataBar":
                    let minVO = currentCFCfvos.first ?? WorksheetData.CFValueObject(type: .min, value: nil)
                    let maxVO = currentCFCfvos.dropFirst().first ?? WorksheetData.CFValueObject(type: .max, value: nil)
                    let color = currentCFColors.first
                    let dataBar = WorksheetData.ConditionalRule.DataBar(min: minVO, max: maxVO, color: color, showValue: currentCFDataBarShowValue)
                    rule = WorksheetData.ConditionalRule(
                        type: .dataBar(dataBar),
                        priority: currentCFRulePriority,
                        dxfId: currentCFRuleDxfId,
                        stopIfTrue: currentCFRuleStopIfTrue
                    )
                case "colorScale":
                    let cs = WorksheetData.ConditionalRule.ColorScale(cfvos: currentCFCfvos, colors: currentCFColors)
                    rule = WorksheetData.ConditionalRule(
                        type: .colorScale(cs),
                        priority: currentCFRulePriority,
                        dxfId: currentCFRuleDxfId,
                        stopIfTrue: currentCFRuleStopIfTrue
                    )
                case "iconSet":
                    let iconName = currentCFIconSetName ?? "3TrafficLights1"
                    let iconSet = WorksheetData.ConditionalRule.IconSet(
                        name: iconName,
                        cfvos: currentCFCfvos,
                        showValue: currentCFIconSetShowValue,
                        reverse: currentCFIconSetReverse,
                        percent: currentCFIconSetPercent
                    )
                    rule = WorksheetData.ConditionalRule(
                        type: .iconSet(iconSet),
                        priority: currentCFRulePriority,
                        dxfId: currentCFRuleDxfId,
                        stopIfTrue: currentCFRuleStopIfTrue
                    )
                default:
                    break
                }

                if let r = rule {
                    currentCFRules.append(r)
                }

                // Reset rule state
                inCFRule = false
                currentCFRuleType = ""
                currentCFRulePriority = nil
                currentCFRuleDxfId = nil
                currentCFRuleStopIfTrue = false
                currentCFRuleOperator = nil
                currentCFFormulas.removeAll(keepingCapacity: true)
                currentCFCfvos.removeAll(keepingCapacity: true)
                currentCFColors.removeAll(keepingCapacity: true)
                currentCFDataBarShowValue = nil
                currentCFIconSetName = nil
                currentCFIconSetShowValue = nil
                currentCFIconSetReverse = nil
                currentCFIconSetPercent = nil
            }

        case "conditionalFormatting":
            if inConditionalFormatting {
                let cf = WorksheetData.ConditionalFormat(range: currentCFSqref, rules: currentCFRules)
                conditionalFormats.append(cf)
                inConditionalFormatting = false
                currentCFSqref = ""
                currentCFRules.removeAll(keepingCapacity: true)
            }

        case "dataValidation":
            inDataValidation = false
            let f1 = formula1Buffer.trimmingCharacters(in: .whitespacesAndNewlines)
            let f2 = formula2Buffer.trimmingCharacters(in: .whitespacesAndNewlines)
            let dv = WorksheetData.DataValidation(
                type: currentDVType,
                allowBlank: currentDVAllowBlank,
                sqref: currentDVSqref,
                formula1: f1.isEmpty ? nil : f1,
                formula2: f2.isEmpty ? nil : f2,
                op: currentDVOp
            )
            dataValidations.append(dv)
            // Reset DV buffers
            formula1Buffer.removeAll(keepingCapacity: true)
            formula2Buffer.removeAll(keepingCapacity: true)
            currentDVSqref = ""
            currentDVOp = nil
            currentDVAllowBlank = true
            currentDVType = .whole

        case "filterColumn":
            if inFilterColumn, let colId = currentFilterColId {
                // Create column filter with accumulated values
                if !currentFilterValues.isEmpty {
                    let column = WorksheetData.AutoFilter.ColumnFilter(
                        colId: colId,
                        criterion: .values(currentFilterValues)
                    )
                    currentAutoFilterColumns.append(column)
                }
                inFilterColumn = false
                currentFilterColId = nil
                currentFilterValues.removeAll(keepingCapacity: true)
            }

        case "autoFilter":
            if inAutoFilter, let ref = currentAutoFilterRef {
                autoFilter = WorksheetData.AutoFilter(ref: ref, columnFilters: currentAutoFilterColumns)
                inAutoFilter = false
                currentAutoFilterRef = nil
                currentAutoFilterColumns.removeAll(keepingCapacity: true)
            }

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        if error == nil {
            error = CuneiformError.malformedXML(
                part: "/xl/worksheets/sheet.xml",
                detail: parseError.localizedDescription
            )
        }
    }
}
