/// Represents a relationship between package parts
public struct Relationship: Sendable {
    /// The relationship ID (e.g., "rId1")
    public let id: String

    /// The relationship type URI
    public let type: RelationshipType

    /// The target part path (relative or absolute)
    public let target: String

    /// Whether the target is external
    public let isExternal: Bool

    public init(id: String, type: RelationshipType, target: String, isExternal: Bool = false) {
        self.id = id
        self.type = type
        self.target = target
        self.isExternal = isExternal
    }

    /// Resolve the target path relative to the source part's directory
    public func resolveTarget(relativeTo sourcePath: PartPath) -> PartPath {
        if target.hasPrefix("/") {
            return PartPath(target)
        }

        let sourceDir = sourcePath.directory
        if sourceDir == "/" {
            return PartPath(target)
        }

        // Handle relative paths
        var components = sourceDir.split(separator: "/").map(String.init)
        for part in target.split(separator: "/") {
            if part == ".." {
                components.removeLast()
            } else if part != "." {
                components.append(String(part))
            }
        }
        return PartPath("/" + components.joined(separator: "/"))
    }
}

// MARK: - RelationshipType

/// Well-known relationship types
public struct RelationshipType: Hashable, Sendable {
    public let uri: String

    public init(_ uri: String) {
        self.uri = uri
    }
}

extension RelationshipType {
    // Package-level relationships
    public static let officeDocument = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    )
    public static let coreProperties = RelationshipType(
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    )
    public static let extendedProperties = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    )

    // Workbook relationships
    public static let worksheet = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    )
    public static let sharedStrings = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    )
    public static let styles = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    )
    public static let theme = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    )
    public static let table = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
    )
    public static let pivotCacheDefinition = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    )
    public static let hyperlink = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    )
    public static let comments = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    )
    public static let vmlDrawing = RelationshipType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
    )
}

extension RelationshipType: ExpressibleByStringLiteral {
    public init(stringLiteral value: String) {
        self.init(value)
    }
}

extension RelationshipType: CustomStringConvertible {
    public var description: String {
        uri
    }
}

// MARK: - Relationships Collection

/// A collection of relationships from a single source
public struct Relationships: Sendable {
    /// All relationships keyed by ID
    public let byId: [String: Relationship]

    /// All relationships keyed by type
    public let byType: [RelationshipType: [Relationship]]

    /// All relationships in order
    public let all: [Relationship]

    public init(_ relationships: [Relationship]) {
        self.all = relationships
        self.byId = Dictionary(uniqueKeysWithValues: relationships.map { ($0.id, $0) })
        self.byType = Dictionary(grouping: relationships, by: \.type)
    }

    /// Get relationship by ID
    public subscript(id: String) -> Relationship? {
        byId[id]
    }

    /// Get relationships by type
    public subscript(type: RelationshipType) -> [Relationship] {
        byType[type] ?? []
    }

    /// Empty relationships
    public static let empty = Relationships([])
}
