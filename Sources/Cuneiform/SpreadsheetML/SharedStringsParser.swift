import Foundation

/// A shared string entry that can be plain text or rich text with formatting
public enum SharedStringEntry: Sendable, Equatable {
    /// Plain text string
    case plain(String)

    /// Rich text with formatting runs
    case rich(RichText)

    /// Get plain text regardless of type
    public var plainText: String {
        switch self {
        case .plain(let s):
            return s
        case .rich(let runs):
            return runs.plainText
        }
    }
}

/// Parsed shared strings table
public struct SharedStrings: Sendable {
    /// All strings in order (index = position), preserving formatting if present
    public let entries: [SharedStringEntry]

    /// Get string by index (returns plain text)
    public subscript(index: Int) -> String? {
        guard index >= 0, index < entries.count else { return nil }
        return entries[index].plainText
    }

    /// Get rich text entry by index
    public func richText(at index: Int) -> RichText? {
        guard index >= 0, index < entries.count else { return nil }
        if case .rich(let runs) = entries[index] {
            return runs
        }
        return nil
    }

    /// Number of strings
    public var count: Int { entries.count }

    /// Empty table
    public static let empty = SharedStrings(entries: [])
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

        return SharedStrings(entries: parser.entries)
    }
}

// MARK: - XML Delegate

/// Internal XML parser for shared strings
final class _SharedStringsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var entries: [SharedStringEntry] = []

    private var inSI = false
    private var currentRuns: [TextRun] = []
    private var currentTextParts: [String] = []

    private var inRun = false
    private var inText = false
    private var inRunProperties = false

    private var textBuffer = ""
    private var currentRunFont: String?
    private var currentRunFontSize: Double?
    private var currentRunColor: String?
    private var currentRunThemeColor: Int?
    private var currentRunBold = false
    private var currentRunItalic = false
    private var currentRunUnderline: String?
    private var currentRunStrikethrough = false
    private var currentRunVerticalAlign: String?

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
            currentRuns.removeAll(keepingCapacity: true)
            currentTextParts.removeAll(keepingCapacity: true)

        case "r":
            // Start of a rich text run
            guard inSI else { return }
            inRun = true
            resetRunProperties()
            textBuffer.removeAll(keepingCapacity: true)

        case "rPr":
            // Run properties (formatting)
            guard inRun else { return }
            inRunProperties = true

        case "rFont":
            // Font name
            guard inRunProperties, let val = attributeDict["val"] else { return }
            currentRunFont = val

        case "sz":
            // Font size
            guard inRunProperties, let val = attributeDict["val"], let size = Double(val) else { return }
            currentRunFontSize = size

        case "color":
            // Color
            if let rgb = attributeDict["rgb"] {
                currentRunColor = rgb
            } else if let theme = attributeDict["theme"], let themeIdx = Int(theme) {
                currentRunThemeColor = themeIdx
            }

        case "b":
            // Bold
            guard inRunProperties else { return }
            currentRunBold = attributeDict["val"]?.lowercased() != "false"

        case "i":
            // Italic
            guard inRunProperties else { return }
            currentRunItalic = attributeDict["val"]?.lowercased() != "false"

        case "u":
            // Underline
            guard inRunProperties else { return }
            let val = attributeDict["val"] ?? "single"
            currentRunUnderline = val

        case "strike":
            // Strikethrough
            guard inRunProperties else { return }
            currentRunStrikethrough = attributeDict["val"]?.lowercased() != "false"

        case "vertAlign":
            // Vertical alignment
            guard inRunProperties, let val = attributeDict["val"] else { return }
            currentRunVerticalAlign = val

        case "t":
            guard inSI else { return }
            inText = true
            textBuffer.removeAll(keepingCapacity: true)

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
            // Check if this is part of a rich text run or plain text
            if inRun {
                // This is part of a run; will be finalized in </r>
                currentTextParts.append(textBuffer)
            } else {
                // This is plain text (no run wrapper)
                currentTextParts.append(textBuffer)
            }
            inText = false
            textBuffer.removeAll(keepingCapacity: true)

        case "rPr":
            guard inRunProperties else { return }
            inRunProperties = false

        case "r":
            guard inRun else { return }
            // End of rich text run; create TextRun from collected parts and properties
            let runText = currentTextParts.joined()
            let run = TextRun(
                text: runText,
                fontName: currentRunFont,
                fontSize: currentRunFontSize,
                color: currentRunColor,
                themeColor: currentRunThemeColor,
                bold: currentRunBold,
                italic: currentRunItalic,
                underline: currentRunUnderline,
                strikethrough: currentRunStrikethrough,
                verticalAlign: currentRunVerticalAlign
            )
            currentRuns.append(run)
            inRun = false
            currentTextParts.removeAll(keepingCapacity: true)
            resetRunProperties()

        case "si":
            guard inSI else { return }
            // Determine if this is rich text or plain text
            if !currentRuns.isEmpty {
                // Rich text: use collected runs
                entries.append(.rich(currentRuns))
            } else if !currentTextParts.isEmpty {
                // Plain text: join parts (for simple <si><t>...</t></si>)
                entries.append(.plain(currentTextParts.joined()))
            } else {
                // Empty string
                entries.append(.plain(""))
            }
            inSI = false
            currentRuns.removeAll(keepingCapacity: true)
            currentTextParts.removeAll(keepingCapacity: true)

        default:
            break
        }
    }

    private func resetRunProperties() {
        currentRunFont = nil
        currentRunFontSize = nil
        currentRunColor = nil
        currentRunThemeColor = nil
        currentRunBold = false
        currentRunItalic = false
        currentRunUnderline = nil
        currentRunStrikethrough = false
        currentRunVerticalAlign = nil
    }
}
