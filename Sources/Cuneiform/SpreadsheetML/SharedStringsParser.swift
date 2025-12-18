import Foundation

/// Parsed shared strings table
public struct SharedStrings: Sendable {
    /// All strings in order (index = position)
    public let strings: [String]

    /// Get string by index
    public subscript(index: Int) -> String? {
        guard index >= 0, index < strings.count else { return nil }
        return strings[index]
    }

    /// Number of strings
    public var count: Int { strings.count }

    /// Empty table
    public static let empty = SharedStrings(strings: [])
}

/// Parser for sharedStrings.xml
public enum SharedStringsParser {
    /// Parse shared strings from XML data
    public static func parse(data: Data) throws(CuneiformError) -> SharedStrings {
        let parser = _SharedStringsParser()
        let xml = XMLParser(data: data)
        xml.delegate = parser

        guard xml.parse() else {
            throw CuneiformError.malformedXML(
                part: "/xl/sharedStrings.xml",
                detail: xml.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return SharedStrings(strings: parser.strings)
    }
}

// MARK: - XML Delegate

/// Internal XML parser for shared strings
final class _SharedStringsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var strings: [String] = []

    private var inSI = false
    private var currentTextParts: [String] = []
    private var inText = false
    private var preserveWhitespace = false
    private var textBuffer = ""

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        switch elementName {
        case "si":
            inSI = true
            currentTextParts.removeAll(keepingCapacity: true)
        case "t":
            guard inSI else { return }
            inText = true
            preserveWhitespace = attributeDict["xml:space"]?.lowercased() == "preserve"
            textBuffer.removeAll(keepingCapacity: true)
        case "r":
            // Rich text run container; we only care about inner <t> elements
            break
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inText {
            textBuffer += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        switch elementName {
        case "t":
            guard inText else { return }
            // Preserve text exactly as provided; Excel shared strings expect
            // whitespace to be retained, and rich text concatenation should
            // not trim spaces between runs.
            currentTextParts.append(textBuffer)
            inText = false
            preserveWhitespace = false
            textBuffer.removeAll(keepingCapacity: true)
        case "si":
            if inSI {
                // Concatenate all collected <t> pieces; empty <si> becomes ""
                let combined = currentTextParts.joined()
                strings.append(combined)
                currentTextParts.removeAll(keepingCapacity: true)
                inSI = false
            }
        default:
            break
        }
    }
}
