import Foundation

/// Number format information
public struct NumberFormat: Sendable {
    public let id: Int
    public let formatCode: String?  // nil for built-in formats

    /// Is this a date/time format?
    public var isDateFormat: Bool {
        // Built-in date/time formats
        if (14...22).contains(id) { return true }
        guard let code = formatCode?.lowercased() else { return false }

        // Heuristic: contains any of y/m/d/h/s and doesn't look like pure number or text format
        if code.trimmingCharacters(in: .whitespaces).isEmpty { return false }
        let first = code.trimmingCharacters(in: .whitespaces).first
        if let f = first, ["#", "0", "?"].contains(String(f)) { return false }
        if code == "@" { return false }

        // Remove quoted literals and escaped characters to focus on tokens
        var filtered = ""
        var inQuote = false
        for ch in code {
            if ch == "\"" { inQuote.toggle(); continue }
            if inQuote { continue }
            filtered.append(ch)
        }

        let tokens = ["y", "m", "d", "h", "s"]
        return tokens.contains { filtered.contains($0) }
    }
}

/// Minimal styles information
public struct StylesInfo: Sendable {
    /// Custom number formats (id -> format code)
    public let numberFormats: [Int: String]

    /// Cell format records (index = style index from cell `s` attribute)
    /// Each entry is the numFmtId used by that style
    public let cellFormats: [Int]  // numFmtId for each xf entry

    /// Get number format for a style index
    public func numberFormat(forStyleIndex index: Int) -> NumberFormat? {
        guard index >= 0, index < cellFormats.count else { return nil }
        let id = cellFormats[index]
        let code = numberFormats[id]
        return NumberFormat(id: id, formatCode: code)
    }

    /// Is the style index a date format?
    public func isDateFormat(styleIndex index: Int) -> Bool {
        guard let nf = numberFormat(forStyleIndex: index) else { return false }
        return nf.isDateFormat
    }

    /// Empty/default styles
    public static let empty = StylesInfo(numberFormats: [:], cellFormats: [])
}

/// Parser for styles.xml
public enum StylesParser {
    public static func parse(data: Data) throws(CuneiformError) -> StylesInfo {
        let delegate = _StylesParser()
        let xml = XMLParser(data: data)
        xml.delegate = delegate

        guard xml.parse() else {
            if let err = delegate.error { throw err }
            throw CuneiformError.malformedXML(
                part: "/xl/styles.xml",
                detail: xml.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return StylesInfo(numberFormats: delegate.numberFormats, cellFormats: delegate.cellFormats)
    }
}

// MARK: - XML Delegate

final class _StylesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?
    private(set) var numberFormats: [Int: String] = [:]
    private(set) var cellFormats: [Int] = []

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        switch elementName {
        case "numFmt":
            if let idStr = attributeDict["numFmtId"], let id = Int(idStr) {
                let code = attributeDict["formatCode"] ?? ""
                numberFormats[id] = code
            }
        case "xf":
            // Only capture within <cellXfs> section; parser doesn't track parent, but it's okay to collect all xf numFmtIds
            let idStr = attributeDict["numFmtId"] ?? "0"
            let id = Int(idStr) ?? 0
            cellFormats.append(id)
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        if error == nil {
            error = CuneiformError.malformedXML(
                part: "/xl/styles.xml",
                detail: parseError.localizedDescription
            )
        }
    }
}
