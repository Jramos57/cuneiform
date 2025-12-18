import Foundation

/// Parsed content types from [Content_Types].xml
public struct ContentTypes: Sendable {
    /// Default content types by file extension
    public let defaults: [String: ContentType]

    /// Override content types by part path
    public let overrides: [PartPath: ContentType]

    /// Get content type for a part path
    public func contentType(for path: PartPath) -> ContentType? {
        // Check overrides first
        if let override = overrides[path] {
            return override
        }

        // Fall back to extension-based default
        let ext = (path.fileName as NSString).pathExtension.lowercased()
        return defaults[ext]
    }

    /// Empty content types
    public static let empty = ContentTypes(defaults: [:], overrides: [:])
}

// MARK: - Parser

/// Parses [Content_Types].xml
final class ContentTypesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private var defaults: [String: ContentType] = [:]
    private var overrides: [PartPath: ContentType] = [:]
    private var error: Error?

    static func parse(data: Data) throws -> ContentTypes {
        let parser = ContentTypesParser()
        let xmlParser = XMLParser(data: data)
        xmlParser.delegate = parser

        guard xmlParser.parse() else {
            if let error = parser.error {
                throw error
            }
            throw CuneiformError.malformedXML(
                part: "[Content_Types].xml",
                detail: xmlParser.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return ContentTypes(defaults: parser.defaults, overrides: parser.overrides)
    }

    // MARK: - XMLParserDelegate

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        switch elementName {
        case "Default":
            if let ext = attributeDict["Extension"],
               let contentType = attributeDict["ContentType"] {
                defaults[ext.lowercased()] = ContentType(contentType)
            }

        case "Override":
            if let partName = attributeDict["PartName"],
               let contentType = attributeDict["ContentType"] {
                overrides[PartPath(partName)] = ContentType(contentType)
            }

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        self.error = CuneiformError.malformedXML(
            part: "[Content_Types].xml",
            detail: parseError.localizedDescription
        )
    }
}
