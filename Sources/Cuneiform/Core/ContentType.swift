/// Represents a content type (MIME type) for a package part
public struct ContentType: Hashable, Sendable {
    /// The content type string
    public let value: String

    public init(_ value: String) {
        self.value = value
    }
}

// MARK: - Well-Known Content Types

extension ContentType {
    // Package-level
    public static let relationships = ContentType("application/vnd.openxmlformats-package.relationships+xml")
    public static let coreProperties = ContentType("application/vnd.openxmlformats-package.core-properties+xml")
    public static let xml = ContentType("application/xml")

    // SpreadsheetML
    public static let workbook = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    public static let worksheet = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    public static let sharedStrings = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
    public static let styles = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    public static let theme = ContentType("application/vnd.openxmlformats-officedocument.theme+xml")
    public static let table = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    public static let pivotTable = ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml")

    // Document properties
    public static let extendedProperties = ContentType("application/vnd.openxmlformats-officedocument.extended-properties+xml")
}

// MARK: - ExpressibleByStringLiteral

extension ContentType: ExpressibleByStringLiteral {
    public init(stringLiteral value: String) {
        self.init(value)
    }
}

// MARK: - CustomStringConvertible

extension ContentType: CustomStringConvertible {
    public var description: String {
        value
    }
}
