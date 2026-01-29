/// A fully-resolved cell value with proper type conversion.
///
/// `CellValue` represents the data contained in a spreadsheet cell after type resolution and
/// shared string lookup. This is the primary type you'll interact with when reading cell data
/// from a workbook.
///
/// ## Overview
///
/// Each cell in a spreadsheet can contain different types of data. `CellValue` provides a
/// type-safe enumeration of all possible cell value types, from simple text and numbers to
/// rich formatted text and error values.
///
/// After parsing an XLSX file, Cuneiform resolves raw cell references and shared string
/// indices into concrete `CellValue` instances. This means you don't need to manually look
/// up shared strings or resolve cell types—everything is ready to use.
///
/// ## Working with Cell Values
///
/// Use pattern matching to extract and work with cell data:
///
/// ```swift
/// let workbook = try Cuneiform.loadWorkbook(from: url)
/// let sheet = workbook.sheets[0]
///
/// if let value = sheet.cell(at: "A1") {
///     switch value {
///     case .text(let text):
///         print("Text: \(text)")
///     case .number(let num):
///         print("Number: \(num)")
///     case .boolean(let bool):
///         print("Boolean: \(bool)")
///     case .empty:
///         print("Cell is empty")
///     default:
///         print("Other value type")
///     }
/// }
/// ```
///
/// ## Type Checking
///
/// Check the type of a cell value before processing:
///
/// ```swift
/// let value = sheet.cell(at: "B2")
///
/// if case .number(let num) = value {
///     // Process numeric value
///     let result = num * 2
/// }
///
/// if case .text = value {
///     // Process text value
///     print("Cell contains text")
/// }
/// ```
///
/// ## Converting Values
///
/// Extract values with optional binding for safe type conversion:
///
/// ```swift
/// // Extract text from various cell types
/// func extractText(from value: CellValue) -> String? {
///     switch value {
///     case .text(let s):
///         return s
///     case .richText(let rt):
///         return rt.plainText
///     case .number(let n):
///         return String(n)
///     case .boolean(let b):
///         return String(b)
///     case .date(let d):
///         return d
///     case .error(let e):
///         return "Error: \(e)"
///     case .empty:
///         return nil
///     }
/// }
/// ```
///
/// ## Handling Rich Text
///
/// Rich text cells contain multiple text runs with different formatting:
///
/// ```swift
/// if case .richText(let richText) = value {
///     // Get plain text without formatting
///     let plain = richText.plainText
///     
///     // Access individual formatted runs
///     for run in richText.runs {
///         print("Text: \(run.text)")
///         if let bold = run.bold {
///             print("Bold: \(bold)")
///         }
///     }
/// }
/// ```
///
/// ## Date Handling
///
/// Date values are provided as ISO 8601 strings. Convert them to `Date` objects as needed:
///
/// ```swift
/// if case .date(let dateString) = value {
///     let formatter = ISO8601DateFormatter()
///     if let date = formatter.date(from: dateString) {
///         print("Date: \(date)")
///     }
/// }
/// ```
///
/// ## Topics
///
/// ### Cell Value Cases
///
/// - ``text(_:)``
/// - ``richText(_:)``
/// - ``number(_:)``
/// - ``boolean(_:)``
/// - ``date(_:)``
/// - ``error(_:)``
/// - ``empty``
///
/// ### Inspecting Values
///
/// - ``description``
///
/// ## See Also
///
/// - ``Sheet/cell(at:)-8xu7x``
/// - ``RichText``
/// - ``CellReference``
public enum CellValue: Sendable {
    /// Text content stored in the cell.
    ///
    /// Represents plain text without any formatting. This is the most common cell value type
    /// for non-numeric data.
    ///
    /// ```swift
    /// let value = CellValue.text("Hello, World!")
    ///
    /// if case .text(let content) = value {
    ///     print(content) // "Hello, World!"
    /// }
    /// ```
    case text(String)

    /// Rich text with multiple formatting runs.
    ///
    /// Rich text cells contain text with varying formatting properties like bold, italic,
    /// font size, or color. Use the ``RichText/plainText`` property to extract text without
    /// formatting.
    ///
    /// ```swift
    /// if case .richText(let rt) = value {
    ///     // Get plain text
    ///     let plain = rt.plainText
    ///     
    ///     // Access formatted runs
    ///     for run in rt.runs {
    ///         print("\(run.text) - Bold: \(run.bold ?? false)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``RichText``
    case richText(RichText)

    /// Numeric value stored as a double-precision floating point.
    ///
    /// All numbers in Excel are stored as 64-bit floating point values, including integers.
    /// Convert to `Int` if you need integer values.
    ///
    /// ```swift
    /// if case .number(let num) = value {
    ///     let integer = Int(num)
    ///     let formatted = String(format: "%.2f", num)
    /// }
    /// ```
    case number(Double)

    /// Boolean value (true or false).
    ///
    /// Represents cells containing boolean values from formulas or direct entry.
    ///
    /// ```swift
    /// if case .boolean(let flag) = value {
    ///     print(flag ? "Yes" : "No")
    /// }
    /// ```
    case boolean(Bool)

    /// Date value as an ISO 8601 string.
    ///
    /// Excel stores dates as numbers, but Cuneiform detects date-formatted cells and provides
    /// them as ISO 8601 strings. Convert to `Date` using `ISO8601DateFormatter`.
    ///
    /// ```swift
    /// if case .date(let dateString) = value {
    ///     let formatter = ISO8601DateFormatter()
    ///     if let date = formatter.date(from: dateString) {
    ///         print(date)
    ///     }
    /// }
    /// ```
    ///
    /// - Note: The conversion from Excel's numeric date format to ISO 8601 is performed
    ///   during parsing. The string format is standardized and locale-independent.
    case date(String)

    /// Error value from a spreadsheet formula.
    ///
    /// Represents Excel error values like `#DIV/0!`, `#N/A`, `#REF!`, etc. that result
    /// from invalid formulas or data.
    ///
    /// ```swift
    /// if case .error(let errorCode) = value {
    ///     switch errorCode {
    ///     case "#DIV/0!":
    ///         print("Division by zero error")
    ///     case "#N/A":
    ///         print("Value not available")
    ///     default:
    ///         print("Error: \(errorCode)")
    ///     }
    /// }
    /// ```
    case error(String)

    /// Empty cell with no value.
    ///
    /// Represents a cell that contains no data. Note that empty cells may still have
    /// formatting or styles applied.
    ///
    /// ```swift
    /// if case .empty = value {
    ///     print("Cell is empty")
    /// }
    /// ```
    case empty

    /// A human-readable description of the cell value for debugging purposes.
    ///
    /// Provides a string representation of the cell value and its type, useful for logging
    /// and debugging. For rich text, only the first 50 characters are shown.
    ///
    /// ```swift
    /// let value = CellValue.number(42.5)
    /// print(value.description) // "number(42.5)"
    ///
    /// let text = CellValue.text("Hello")
    /// print(text.description) // "text(Hello)"
    /// ```
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
