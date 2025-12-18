/// Represents a path to a part within an OPC package
///
/// Part paths are absolute paths starting with `/` that identify
/// parts within the ZIP archive.
public struct PartPath: Hashable, Sendable {
    /// The raw path string (always starts with `/`)
    public let value: String

    /// Create a part path from a string
    /// - Parameter value: The path string. A leading `/` is added if not present.
    public init(_ value: String) {
        if value.hasPrefix("/") {
            self.value = value
        } else {
            self.value = "/" + value
        }
    }

    /// The path without the leading slash (for ZIP entry lookup)
    public var zipEntryPath: String {
        String(value.dropFirst())
    }

    /// The file name component
    public var fileName: String {
        value.split(separator: "/").last.map(String.init) ?? ""
    }

    /// The directory containing this part
    public var directory: String {
        let components = value.split(separator: "/")
        if components.count <= 1 {
            return "/"
        }
        return "/" + components.dropLast().joined(separator: "/")
    }

    /// The path to this part's relationships file
    public var relationshipsPath: PartPath {
        let dir = directory
        let name = fileName
        if dir == "/" {
            return PartPath("/_rels/\(name).rels")
        }
        return PartPath("\(dir)/_rels/\(name).rels")
    }
}

// MARK: - Well-Known Paths

extension PartPath {
    /// Content types declaration
    public static let contentTypes = PartPath("/[Content_Types].xml")

    /// Package-level relationships
    public static let rootRelationships = PartPath("/_rels/.rels")

    /// Main workbook
    public static let workbook = PartPath("/xl/workbook.xml")

    /// Workbook relationships
    public static let workbookRelationships = PartPath("/xl/_rels/workbook.xml.rels")

    /// Shared strings table
    public static let sharedStrings = PartPath("/xl/sharedStrings.xml")

    /// Styles
    public static let styles = PartPath("/xl/styles.xml")
}

// MARK: - ExpressibleByStringLiteral

extension PartPath: ExpressibleByStringLiteral {
    public init(stringLiteral value: String) {
        self.init(value)
    }
}

// MARK: - CustomStringConvertible

extension PartPath: CustomStringConvertible {
    public var description: String {
        value
    }
}
