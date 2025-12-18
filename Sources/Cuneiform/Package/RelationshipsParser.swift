import Foundation

/// Parses .rels files
final class RelationshipsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private var relationships: [Relationship] = []
    private var error: Error?
    private let partPath: String

    private init(partPath: String) {
        self.partPath = partPath
    }

    static func parse(data: Data, partPath: String) throws -> Relationships {
        let parser = RelationshipsParser(partPath: partPath)
        let xmlParser = XMLParser(data: data)
        xmlParser.delegate = parser

        guard xmlParser.parse() else {
            if let error = parser.error {
                throw error
            }
            throw CuneiformError.malformedXML(
                part: partPath,
                detail: xmlParser.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return Relationships(parser.relationships)
    }

    // MARK: - XMLParserDelegate

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        guard elementName == "Relationship" else { return }

        guard let id = attributeDict["Id"] else {
            error = CuneiformError.missingRequiredElement(element: "Id attribute", inPart: partPath)
            parser.abortParsing()
            return
        }

        guard let type = attributeDict["Type"] else {
            error = CuneiformError.missingRequiredElement(element: "Type attribute", inPart: partPath)
            parser.abortParsing()
            return
        }

        guard let target = attributeDict["Target"] else {
            error = CuneiformError.missingRequiredElement(element: "Target attribute", inPart: partPath)
            parser.abortParsing()
            return
        }

        let targetMode = attributeDict["TargetMode"]
        let isExternal = targetMode?.lowercased() == "external"

        relationships.append(Relationship(
            id: id,
            type: RelationshipType(type),
            target: target,
            isExternal: isExternal
        ))
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        if self.error == nil {
            self.error = CuneiformError.malformedXML(
                part: partPath,
                detail: parseError.localizedDescription
            )
        }
    }
}
