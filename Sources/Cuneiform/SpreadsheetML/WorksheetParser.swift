import Foundation

/// A cell reference like "A1" or "BC123"
///
/// Cell references consist of a column (letters) and row (1-based number).
///
/// ```swift
/// let ref = CellReference("A1")
/// print(ref?.columnIndex)  // 0
/// print(ref?.row)          // 1
/// ```
public struct CellReference: Hashable, Sendable {
    /// Column letters (e.g., "A", "B", "AA", "AB")
    public let column: String

    /// 1-based row number
    public let row: Int

    /// Parse a cell reference string like "A1" or "BC123"
    public init?(_ reference: String) {
        guard !reference.isEmpty else { return nil }
        var letters = ""
        var digits = ""
        for ch in reference {
            if ch.isLetter { letters.append(ch) } else { digits.append(ch) }
        }
        guard !letters.isEmpty, !digits.isEmpty, let row = Int(digits) else { return nil }
        self.column = letters.uppercased()
        self.row = row
    }

    /// Create a cell reference from column and row components
    public init(column: String, row: Int) {
        self.column = column.uppercased()
        self.row = row
    }

    /// Column as 0-based index (A=0, B=1, ..., Z=25, AA=26)
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
}

/// Parsed worksheet data
public struct WorksheetData: Sendable {
    /// Declared dimension (may be inaccurate)
    public let dimension: String?

    /// All rows with data
    public let rows: [RawRow]

    /// Merged cell ranges
    public let mergedCells: [String]  // "A1:C3" format

    /// Data validations defined in the worksheet
    public let dataValidations: [DataValidation]

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

        return WorksheetData(
            dimension: delegate.dimension,
            rows: delegate.rows,
            mergedCells: delegate.mergedCells,
            dataValidations: delegate.dataValidations
        )
    }
}

// MARK: - XML Delegate

final class _WorksheetParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?

    private(set) var dimension: String?
    private(set) var rows: [RawRow] = []
    private(set) var mergedCells: [String] = []
    private(set) var dataValidations: [WorksheetData.DataValidation] = []

    // Row accumulation
    private var currentRowIndex: Int = 0
    private var currentCells: [RawCell] = []

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

        case "row":
            let rStr = attributeDict["r"] ?? "0"
            currentRowIndex = Int(rStr) ?? 0
            currentCells.removeAll(keepingCapacity: true)

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
            rows.append(RawRow(index: currentRowIndex, cells: currentCells))
            currentCells.removeAll(keepingCapacity: true)

        case "formula1":
            inFormula1 = false

        case "formula2":
            inFormula2 = false

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
