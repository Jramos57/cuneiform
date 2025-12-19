/// A fully-resolved cell value with proper type conversion.
public enum CellValue: Sendable, Hashable {
    /// Text content
    case text(String)

    /// Numeric value
    case number(Double)

    /// Boolean value
    case boolean(Bool)

    /// Date (ISO 8601 string, conversion to Date is caller's responsibility)
    case date(String)

    /// Error value from spreadsheet
    case error(String)

    /// Empty cell
    case empty

    /// Description for debugging
    public var description: String {
        switch self {
        case .text(let s):
            return "text(\(s))"
        case .number(let n):
            return "number(\(n))"
        case .boolean(let b):
            return "boolean(\(b))"
        case .date(let d):
            return "date(\(d))"
        case .error(let e):
            return "error(\(e))"
        case .empty:
            return "empty"
        }
    }
}

extension CellValue: CustomStringConvertible {
    public var debugDescription: String { description }
}
