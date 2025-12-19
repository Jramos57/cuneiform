import Foundation

/// A color specification (RGB hex, theme, or indexed)
public enum CellColor: Sendable, Equatable {
    case rgb(String)           // "FF000000" format
    case theme(Int)            // Theme color index (0-12)
    case indexed(Int)          // Indexed palette (0-63)

    /// Display as RGB hex string if possible
    var rgbHex: String? {
        if case .rgb(let hex) = self { return hex }
        return nil
    }
}

/// Font properties
public struct CellFont: Sendable, Equatable {
    public let name: String?       // Font family name
    public let size: Double?       // Font size in points
    public let bold: Bool          // Bold flag
    public let italic: Bool        // Italic flag
    public let underline: Bool     // Underline flag
    public let strike: Bool        // Strike-through flag
    public let color: CellColor?   // Font color

    public init(
        name: String? = nil,
        size: Double? = nil,
        bold: Bool = false,
        italic: Bool = false,
        underline: Bool = false,
        strike: Bool = false,
        color: CellColor? = nil
    ) {
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strike = strike
        self.color = color
    }
}

/// Fill (background) properties
public struct CellFill: Sendable, Equatable {
    public enum Pattern: String, Sendable {
        case none, solid, medGray, darkGray, lightGray, darkHorizontal, darkVertical
        case darkDown, darkUp, darkGrid, darkTrellis, lightHorizontal, lightVertical
        case lightDown, lightUp, lightGrid, lightTrellis, gray0625, gray125
    }

    public let pattern: Pattern
    public let foregroundColor: CellColor?
    public let backgroundColor: CellColor?

    public init(
        pattern: Pattern = .none,
        foregroundColor: CellColor? = nil,
        backgroundColor: CellColor? = nil
    ) {
        self.pattern = pattern
        self.foregroundColor = foregroundColor
        self.backgroundColor = backgroundColor
    }
}

/// Border properties (individual side)
public struct CellBorderSide: Sendable, Equatable {
    public enum Style: String, Sendable {
        case none, thin, medium, thick, dashed, dotted, double
        case hair, mediumDashed, dashDot, mediumDashDot, dashDotDot, mediumDashDotDot
        case slantDashDot
    }

    public let style: Style
    public let color: CellColor?

    public init(style: Style = .none, color: CellColor? = nil) {
        self.style = style
        self.color = color
    }
}

/// Border properties (all sides)
public struct CellBorder: Sendable, Equatable {
    public let left: CellBorderSide?
    public let right: CellBorderSide?
    public let top: CellBorderSide?
    public let bottom: CellBorderSide?
    public let diagonal: CellBorderSide?

    public init(
        left: CellBorderSide? = nil,
        right: CellBorderSide? = nil,
        top: CellBorderSide? = nil,
        bottom: CellBorderSide? = nil,
        diagonal: CellBorderSide? = nil
    ) {
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
    }
}

/// Alignment properties
public struct CellAlignment: Sendable, Equatable {
    public enum Horizontal: String, Sendable {
        case general, left, center, right, fill, distributed, centerContinuous, justified
    }

    public enum Vertical: String, Sendable {
        case top, center, bottom, distributed, justified
    }

    public let horizontal: Horizontal?
    public let vertical: Vertical?
    public let wrapText: Bool
    public let textRotation: Int?    // 0-180 degrees
    public let indent: Int?          // Indentation level

    public init(
        horizontal: Horizontal? = nil,
        vertical: Vertical? = nil,
        wrapText: Bool = false,
        textRotation: Int? = nil,
        indent: Int? = nil
    ) {
        self.horizontal = horizontal
        self.vertical = vertical
        self.wrapText = wrapText
        self.textRotation = textRotation
        self.indent = indent
    }
}

/// Complete cell style (format) information
public struct CellStyle: Sendable, Equatable {
    public let numberFormat: NumberFormat?
    public let font: CellFont?
    public let fill: CellFill?
    public let border: CellBorder?
    public let alignment: CellAlignment?

    public init(
        numberFormat: NumberFormat? = nil,
        font: CellFont? = nil,
        fill: CellFill? = nil,
        border: CellBorder? = nil,
        alignment: CellAlignment? = nil
    ) {
        self.numberFormat = numberFormat
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
    }
}

/// Number format information
public struct NumberFormat: Sendable, Equatable {
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

/// Complete styles information
public struct StylesInfo: Sendable {
    /// Custom number formats (id -> format code)
    public let numberFormats: [Int: String]

    /// Font definitions (index = font ID)
    public let fonts: [CellFont]

    /// Fill definitions (index = fill ID)
    public let fills: [CellFill]

    /// Border definitions (index = border ID)
    public let borders: [CellBorder]

    /// Cell format records (index = style index from cell `s` attribute)
    public let cellFormats: [CellFormatRecord]

    /// Cell format record
    public struct CellFormatRecord: Sendable, Equatable {
        public let numFmtId: Int
        public let fontId: Int
        public let fillId: Int
        public let borderId: Int
        public let alignment: CellAlignment?
    }

    /// Get complete cell style for a style index
    public func cellStyle(forStyleIndex index: Int) -> CellStyle? {
        guard index >= 0, index < cellFormats.count else { return nil }
        let record = cellFormats[index]

        let numberFmt = NumberFormat(id: record.numFmtId, formatCode: numberFormats[record.numFmtId])
        let font = fonts.indices.contains(record.fontId) ? fonts[record.fontId] : nil
        let fill = fills.indices.contains(record.fillId) ? fills[record.fillId] : nil
        let border = borders.indices.contains(record.borderId) ? borders[record.borderId] : nil

        return CellStyle(
            numberFormat: numberFmt,
            font: font,
            fill: fill,
            border: border,
            alignment: record.alignment
        )
    }

    /// Get number format for a style index
    public func numberFormat(forStyleIndex index: Int) -> NumberFormat? {
        guard index >= 0, index < cellFormats.count else { return nil }
        let id = cellFormats[index].numFmtId
        let code = numberFormats[id]
        return NumberFormat(id: id, formatCode: code)
    }

    /// Is the style index a date format?
    public func isDateFormat(styleIndex index: Int) -> Bool {
        guard let nf = numberFormat(forStyleIndex: index) else { return false }
        return nf.isDateFormat
    }

    /// Empty/default styles
    public static let empty = StylesInfo(
        numberFormats: [:],
        fonts: [CellFont()],  // Default font
        fills: [CellFill(pattern: .none), CellFill(pattern: .gray125)],  // Built-in fills
        borders: [CellBorder()],  // Default border
        cellFormats: [CellFormatRecord(numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: nil)]
    )
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

        return StylesInfo(
            numberFormats: delegate.numberFormats,
            fonts: delegate.fonts,
            fills: delegate.fills,
            borders: delegate.borders,
            cellFormats: delegate.cellFormats
        )
    }
}

// MARK: - XML Delegate

final class _StylesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?
    private(set) var numberFormats: [Int: String] = [:]
    private(set) var fonts: [CellFont] = []
    private(set) var fills: [CellFill] = []
    private(set) var borders: [CellBorder] = []
    private(set) var cellFormats: [StylesInfo.CellFormatRecord] = []

    // Parsing state
    private var inFonts = false
    private var inFills = false
    private var inBorders = false
    private var inCellXfs = false
    private var currentFont: CellFont?
    private var currentFill: CellFill?
    private var currentBorder: CellBorder?
    private var currentFontColor: CellColor?
    private var currentFillFgColor: CellColor?
    private var currentFillBgColor: CellColor?
    private var currentBorderLeft: CellBorderSide?
    private var currentBorderRight: CellBorderSide?
    private var currentBorderTop: CellBorderSide?
    private var currentBorderBottom: CellBorderSide?
    private var currentBorderDiagonal: CellBorderSide?

    override init() {
        super.init()
        // Don't add defaults here - they are implicit if missing
    }

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

        case "fonts":
            inFonts = true

        case "font":
            if inFonts {
                currentFont = CellFont()
            }

        case "sz":
            if inFonts, let valStr = attributeDict["val"], let size = Double(valStr) {
                if currentFont != nil {
                    currentFont = CellFont(
                        name: currentFont?.name,
                        size: size,
                        bold: currentFont?.bold ?? false,
                        italic: currentFont?.italic ?? false,
                        underline: currentFont?.underline ?? false,
                        strike: currentFont?.strike ?? false,
                        color: currentFont?.color
                    )
                }
            }

        case "b":
            if inFonts, currentFont != nil {
                currentFont = CellFont(
                    name: currentFont?.name,
                    size: currentFont?.size,
                    bold: true,
                    italic: currentFont?.italic ?? false,
                    underline: currentFont?.underline ?? false,
                    strike: currentFont?.strike ?? false,
                    color: currentFont?.color
                )
            }

        case "i":
            if inFonts, currentFont != nil {
                currentFont = CellFont(
                    name: currentFont?.name,
                    size: currentFont?.size,
                    bold: currentFont?.bold ?? false,
                    italic: true,
                    underline: currentFont?.underline ?? false,
                    strike: currentFont?.strike ?? false,
                    color: currentFont?.color
                )
            }

        case "u":
            if inFonts, currentFont != nil {
                currentFont = CellFont(
                    name: currentFont?.name,
                    size: currentFont?.size,
                    bold: currentFont?.bold ?? false,
                    italic: currentFont?.italic ?? false,
                    underline: true,
                    strike: currentFont?.strike ?? false,
                    color: currentFont?.color
                )
            }

        case "strike":
            if inFonts, currentFont != nil {
                currentFont = CellFont(
                    name: currentFont?.name,
                    size: currentFont?.size,
                    bold: currentFont?.bold ?? false,
                    italic: currentFont?.italic ?? false,
                    underline: currentFont?.underline ?? false,
                    strike: true,
                    color: currentFont?.color
                )
            }

        case "rFont":
            if inFonts, let name = attributeDict["val"] {
                if currentFont != nil {
                    currentFont = CellFont(
                        name: name,
                        size: currentFont?.size,
                        bold: currentFont?.bold ?? false,
                        italic: currentFont?.italic ?? false,
                        underline: currentFont?.underline ?? false,
                        strike: currentFont?.strike ?? false,
                        color: currentFont?.color
                    )
                }
            }

        case "color":
            if inFonts {
                currentFontColor = parseColor(attributeDict)
            }

        case "fills":
            inFills = true

        case "fill":
            if inFills {
                currentFill = CellFill()
                currentFillFgColor = nil
                currentFillBgColor = nil
            }

        case "patternFill":
            if inFills {
                let patternStr = attributeDict["patternType"] ?? "none"
                let pattern = CellFill.Pattern(rawValue: patternStr) ?? .none
                currentFill = CellFill(
                    pattern: pattern,
                    foregroundColor: currentFill?.foregroundColor,
                    backgroundColor: currentFill?.backgroundColor
                )
            }

        case "fgColor":
            if inFills {
                currentFillFgColor = parseColor(attributeDict)
            }

        case "bgColor":
            if inFills {
                currentFillBgColor = parseColor(attributeDict)
            }

        case "borders":
            inBorders = true

        case "border":
            if inBorders {
                currentBorder = CellBorder()
                currentBorderLeft = nil
                currentBorderRight = nil
                currentBorderTop = nil
                currentBorderBottom = nil
                currentBorderDiagonal = nil
            }

        case "left":
            if inBorders {
                currentBorderLeft = parseBorderSide(attributeDict)
            }

        case "right":
            if inBorders {
                currentBorderRight = parseBorderSide(attributeDict)
            }

        case "top":
            if inBorders {
                currentBorderTop = parseBorderSide(attributeDict)
            }

        case "bottom":
            if inBorders {
                currentBorderBottom = parseBorderSide(attributeDict)
            }

        case "diagonal":
            if inBorders {
                currentBorderDiagonal = parseBorderSide(attributeDict)
            }

        case "cellXfs":
            inCellXfs = true

        case "xf":
            if inCellXfs {
                let numFmtId = Int(attributeDict["numFmtId"] ?? "0") ?? 0
                let fontId = Int(attributeDict["fontId"] ?? "0") ?? 0
                let fillId = Int(attributeDict["fillId"] ?? "0") ?? 0
                let borderId = Int(attributeDict["borderId"] ?? "0") ?? 0

                // Alignment details come in nested <alignment> element
                cellFormats.append(StylesInfo.CellFormatRecord(
                    numFmtId: numFmtId,
                    fontId: fontId,
                    fillId: fillId,
                    borderId: borderId,
                    alignment: nil
                ))
            }

        case "alignment":
            if inCellXfs && !cellFormats.isEmpty {
                let horizontal = attributeDict["horizontal"].flatMap { CellAlignment.Horizontal(rawValue: $0) }
                let vertical = attributeDict["vertical"].flatMap { CellAlignment.Vertical(rawValue: $0) }
                let wrapText = attributeDict["wrapText"] == "1"
                let textRotation = attributeDict["textRotation"].flatMap { Int($0) }
                let indent = attributeDict["indent"].flatMap { Int($0) }

                let alignment = CellAlignment(
                    horizontal: horizontal,
                    vertical: vertical,
                    wrapText: wrapText,
                    textRotation: textRotation,
                    indent: indent
                )

                // Update the last cellFormat with alignment
                var lastRecord = cellFormats.removeLast()
                lastRecord = StylesInfo.CellFormatRecord(
                    numFmtId: lastRecord.numFmtId,
                    fontId: lastRecord.fontId,
                    fillId: lastRecord.fillId,
                    borderId: lastRecord.borderId,
                    alignment: alignment
                )
                cellFormats.append(lastRecord)
            }

        default:
            break
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        switch elementName {
        case "font":
            if inFonts, var font = currentFont {
                if currentFontColor != nil {
                    font = CellFont(
                        name: font.name,
                        size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        underline: font.underline,
                        strike: font.strike,
                        color: currentFontColor
                    )
                }
                fonts.append(font)
                currentFont = nil
                currentFontColor = nil
            }

        case "fonts":
            inFonts = false

        case "fill":
            if inFills, var fill = currentFill {
                if currentFillFgColor != nil || currentFillBgColor != nil {
                    fill = CellFill(
                        pattern: fill.pattern,
                        foregroundColor: currentFillFgColor,
                        backgroundColor: currentFillBgColor
                    )
                }
                fills.append(fill)
                currentFill = nil
                currentFillFgColor = nil
                currentFillBgColor = nil
            }

        case "fills":
            inFills = false

        case "border":
            if inBorders, var border = currentBorder {
                border = CellBorder(
                    left: currentBorderLeft,
                    right: currentBorderRight,
                    top: currentBorderTop,
                    bottom: currentBorderBottom,
                    diagonal: currentBorderDiagonal
                )
                borders.append(border)
                currentBorder = nil
                currentBorderLeft = nil
                currentBorderRight = nil
                currentBorderTop = nil
                currentBorderBottom = nil
                currentBorderDiagonal = nil
            }

        case "borders":
            inBorders = false

        case "cellXfs":
            inCellXfs = false

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

    // MARK: - Helpers

    private func parseColor(_ attributes: [String: String]) -> CellColor? {
        if let rgb = attributes["rgb"] {
            return .rgb(rgb)
        }
        if let themeStr = attributes["theme"], let theme = Int(themeStr) {
            return .theme(theme)
        }
        if let indexedStr = attributes["indexed"], let indexed = Int(indexedStr) {
            return .indexed(indexed)
        }
        return nil
    }

    private func parseBorderSide(_ attributes: [String: String]) -> CellBorderSide? {
        guard let styleStr = attributes["style"] else { return nil }
        let style = CellBorderSide.Style(rawValue: styleStr) ?? .none
        return CellBorderSide(style: style)
    }
}
