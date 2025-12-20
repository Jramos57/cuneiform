import Foundation

/// A run of text with formatting properties
public struct TextRun: Sendable, Equatable {
    /// The text content
    public let text: String

    /// Font name (e.g., "Calibri")
    public let fontName: String?

    /// Font size in points
    public let fontSize: Double?

    /// RGB color hex (e.g., "FF0000" for red)
    public let color: String?

    /// Theme color index (1-12)
    public let themeColor: Int?

    /// Bold flag
    public let bold: Bool

    /// Italic flag
    public let italic: Bool

    /// Underline style: "single", "double", "singleAccounting", "doubleAccounting"
    public let underline: String?

    /// Strikethrough flag
    public let strikethrough: Bool

    /// Vertical alignment: "superscript", "subscript"
    public let verticalAlign: String?

    /// Initializer with all properties
    public init(
        text: String,
        fontName: String? = nil,
        fontSize: Double? = nil,
        color: String? = nil,
        themeColor: Int? = nil,
        bold: Bool = false,
        italic: Bool = false,
        underline: String? = nil,
        strikethrough: Bool = false,
        verticalAlign: String? = nil
    ) {
        self.text = text
        self.fontName = fontName
        self.fontSize = fontSize
        self.color = color
        self.themeColor = themeColor
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.verticalAlign = verticalAlign
    }

    /// Plain text run (no formatting)
    public static func plain(_ text: String) -> TextRun {
        TextRun(text: text)
    }

    /// Concatenated text from all runs
    public var plainText: String { text }
}

/// Rich text content as an array of formatted runs
public typealias RichText = [TextRun]

extension RichText {
    /// Concatenate all runs into plain text
    public var plainText: String {
        map(\.text).joined()
    }

    /// Check if any run has formatting
    public var hasFormatting: Bool {
        contains { run in
            run.fontName != nil || run.fontSize != nil || run.color != nil ||
            run.themeColor != nil || run.bold || run.italic ||
            run.underline != nil || run.strikethrough || run.verticalAlign != nil
        }
    }
}
