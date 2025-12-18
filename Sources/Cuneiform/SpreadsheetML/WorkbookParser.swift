import Foundation

/// Sheet visibility state
public enum SheetState: String, Sendable, CaseIterable {
    case visible
    case hidden
    case veryHidden
}

/// Sheet metadata from workbook.xml
public struct SheetInfo: Sendable {
    /// Display name
    public let name: String

    /// Internal sheet ID
    public let sheetId: Int

    /// Relationship ID (used to find actual sheet part)
    public let relationshipId: String

    /// Visibility state
    public let state: SheetState
}

/// Parsed workbook metadata
public struct WorkbookInfo: Sendable {
    /// All sheets in workbook order
    public let sheets: [SheetInfo]

    /// Get sheet by name
    public func sheet(named name: String) -> SheetInfo? {
        sheets.first { $0.name == name }
    }
}

/// Parser for workbook.xml
public enum WorkbookParser {
    public static func parse(data: Data) throws(CuneiformError) -> WorkbookInfo {
        let delegate = _WorkbookParser()
        let xml = XMLParser(data: data)
        xml.delegate = delegate

        guard xml.parse() else {
            if let err = delegate.error { throw err }
            throw CuneiformError.malformedXML(
                part: "/xl/workbook.xml",
                detail: xml.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return WorkbookInfo(sheets: delegate.sheets)
    }
}

// MARK: - XML Delegate

final class _WorkbookParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var sheets: [SheetInfo] = []
    fileprivate var error: CuneiformError?

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        guard elementName == "sheet" else { return }

        // Required: name and r:id
        guard let name = attributeDict["name"], !name.isEmpty else {
            error = CuneiformError.missingRequiredElement(element: "sheet@name", inPart: "/xl/workbook.xml")
            parser.abortParsing()
            return
        }

        // r:id lives in the relationships namespace
        guard let rid = attributeDict["r:id"], !rid.isEmpty else {
            error = CuneiformError.missingRequiredElement(element: "sheet@r:id", inPart: "/xl/workbook.xml")
            parser.abortParsing()
            return
        }

        let sheetIdStr = attributeDict["sheetId"] ?? "0"
        let sheetId = Int(sheetIdStr) ?? 0

        let stateStr = attributeDict["state"]?.lowercased()
        let state: SheetState
        switch stateStr {
        case nil: state = .visible
        case "visible": state = .visible
        case "hidden": state = .hidden
        case "veryhidden": state = .veryHidden
        default:
            // Unknown values treated as visible
            state = .visible
        }

        sheets.append(SheetInfo(name: name, sheetId: sheetId, relationshipId: rid, state: state))
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        if error == nil {
            error = CuneiformError.malformedXML(
                part: "/xl/workbook.xml",
                detail: parseError.localizedDescription
            )
        }
    }
}
