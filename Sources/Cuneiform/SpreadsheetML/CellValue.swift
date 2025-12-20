/// A fully-resolved cell value with proper type conversion.
public enum CellValue: Sendable {
    /// Text content
    case text(String)

    /// Rich text with formatting (multiple runs)
    case richText(RichText)

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
        case .richText(let runs):
            let preview = runs.plainText.prefix(50)
            return "richText(\(preview)...)"
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

extension CellValue: Equatable {
    public static func == (lhs: CellValue, rhs: CellValue) -> Bool {
        switch (lhs, rhs) {
        case (.text(let a), .text(let b)):
            return a == b
        case (.richText(let a), .richText(let b)):
            return a == b
        case (.number(let a), .number(let b)):
            return a == b
        case (.boolean(let a), .boolean(let b)):
            return a == b
        case (.date(let a), .date(let b)):
            return a == b
        case (.error(let a), .error(let b)):
            return a == b
        case (.empty, .empty):
            return true
        default:
            return false
        }
    }
}

extension CellValue: Hashable {
    public func hash(into hasher: inout Hasher) {
        switch self {
        case .text(let s):
            hasher.combine(0)
            hasher.combine(s)
        case .richText(let runs):
            hasher.combine(1)
            // Hash based on concatenated text since arrays aren't directly hashable
            hasher.combine(runs.plainText)
        case .number(let n):
            hasher.combine(2)
            hasher.combine(n)
        case .boolean(let b):
            hasher.combine(3)
            hasher.combine(b)
        case .date(let d):
            hasher.combine(4)
            hasher.combine(d)
        case .error(let e):
            hasher.combine(5)
            hasher.combine(e)
        case .empty:
            hasher.combine(6)
        }
    }
}
