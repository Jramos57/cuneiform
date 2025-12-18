import Foundation

/// An Open Packaging Conventions (OPC) package
///
/// This is the core container format for Office Open XML documents.
/// An OPC package is a ZIP archive containing parts (XML files and other resources)
/// connected by relationships.
public struct OPCPackage: Sendable {
    /// The raw archive data
    private let archiveData: Data

    /// The ZIP archive structure
    private let archive: ZipArchive

    /// Content types for all parts
    public let contentTypes: ContentTypes

    /// Root-level relationships
    public let rootRelationships: Relationships

    /// Cache of parsed relationships per part
    private var relationshipsCache: [PartPath: Relationships] = [:]

    /// Open an OPC package from a file URL
    public static func open(url: URL) throws -> OPCPackage {
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw CuneiformError.fileNotFound(path: url.path)
        }

        let data: Data
        do {
            data = try Data(contentsOf: url)
        } catch {
            throw CuneiformError.accessDenied(path: url.path)
        }

        return try open(data: data)
    }

    /// Open an OPC package from data
    public static func open(data: Data) throws -> OPCPackage {
        // Parse ZIP structure
        let archive = try ZipArchive.read(from: data)

        // Parse content types
        guard let contentTypesEntry = archive.entries["[Content_Types].xml"] else {
            throw CuneiformError.missingPart(path: "[Content_Types].xml")
        }
        let contentTypesData = try archive.extractData(for: contentTypesEntry, from: data)
        let contentTypes = try ContentTypesParser.parse(data: contentTypesData)

        // Parse root relationships
        let rootRelationships: Relationships
        if let relsEntry = archive.entries["_rels/.rels"] {
            let relsData = try archive.extractData(for: relsEntry, from: data)
            rootRelationships = try RelationshipsParser.parse(data: relsData, partPath: "_rels/.rels")
        } else {
            rootRelationships = .empty
        }

        return OPCPackage(
            archiveData: data,
            archive: archive,
            contentTypes: contentTypes,
            rootRelationships: rootRelationships
        )
    }

    /// Check if a part exists in the package
    public func partExists(_ path: PartPath) -> Bool {
        archive.entries[path.zipEntryPath] != nil
    }

    /// Get the content type for a part
    public func contentType(for path: PartPath) -> ContentType? {
        contentTypes.contentType(for: path)
    }

    /// Read raw data for a part
    public func readPart(_ path: PartPath) throws -> Data {
        guard let entry = archive.entries[path.zipEntryPath] else {
            throw CuneiformError.missingPart(path: path.value)
        }
        return try archive.extractData(for: entry, from: archiveData)
    }

    /// Read a part as a string
    public func readPartAsString(_ path: PartPath, encoding: String.Encoding = .utf8) throws -> String {
        let data = try readPart(path)
        guard let string = String(data: data, encoding: encoding) else {
            throw CuneiformError.malformedXML(part: path.value, detail: "Invalid string encoding")
        }
        return string
    }

    /// Get relationships for a specific part
    public mutating func relationships(for path: PartPath) throws -> Relationships {
        // Check cache
        if let cached = relationshipsCache[path] {
            return cached
        }

        // Find relationships file
        let relsPath = path.relationshipsPath

        guard let entry = archive.entries[relsPath.zipEntryPath] else {
            // No relationships file is valid - return empty
            relationshipsCache[path] = .empty
            return .empty
        }

        let data = try archive.extractData(for: entry, from: archiveData)
        let relationships = try RelationshipsParser.parse(data: data, partPath: relsPath.value)
        relationshipsCache[path] = relationships
        return relationships
    }

    /// List all part paths in the package
    public var partPaths: [PartPath] {
        archive.entries.keys
            .filter { !$0.hasSuffix(".rels") && $0 != "[Content_Types].xml" }
            .map { PartPath($0) }
    }

    /// Find the main document (workbook) relationship
    public func findMainDocument() -> Relationship? {
        rootRelationships[.officeDocument].first
    }
}
