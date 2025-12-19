import Foundation

/// Builds the `/xl/commentsN.xml` part for a worksheet.
public struct CommentsBuilder {
    private var authors: [String] = []
    private var authorIndex: [String: Int] = [:]
    private var entries: [Entry] = []

    /// A single comment entry to emit.
    public struct Entry {
        public let reference: CellReference
        public let authorId: Int
        public let text: String
    }

    public init() {}

    /// Adds a comment at a cell reference.
    /// - Parameters:
    ///   - reference: Target cell for the comment.
    ///   - text: Plain text body of the comment.
    ///   - author: Optional author name; defaults to an empty author entry.
    public mutating func addComment(at reference: CellReference, text: String, author: String? = nil) {
        let trimmedAuthor = author?.trimmingCharacters(in: .whitespacesAndNewlines)
        let authorId = indexForAuthor(trimmedAuthor ?? "")
        entries.append(Entry(reference: reference, authorId: authorId, text: text))
    }

    /// Returns true when comments have been added.
    public var hasComments: Bool { !entries.isEmpty }

    /// All comment entries accumulated (for VML generation)
    var allEntries: [Entry] { entries }

    /// All authors accumulated (for VML generation)
    var allAuthors: [String] { authors }

    /// Builds the comments XML data.
    public func build() -> Data {
        precondition(!entries.isEmpty, "No comments to build")

        var xml = """
        <?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
        <comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
          <authors>
        """

        for author in authors {
            xml += "\n    <author>\(xmlEscape(author))</author>"
        }

        xml += "\n  </authors>\n  <commentList>"

        for entry in entries {
            xml += """
\n    <comment ref=\"\(entry.reference)\" authorId=\"\(entry.authorId)\">\n      <text>\n        <r><t>\(xmlEscape(entry.text))</t></r>\n      </text>\n    </comment>
"""
        }

        xml += "\n  </commentList>\n</comments>\n"

        return xml.data(using: .utf8)!
    }

    // MARK: - Helpers

    private mutating func indexForAuthor(_ author: String) -> Int {
        if let existing = authorIndex[author] { return existing }
        let next = authors.count
        authors.append(author)
        authorIndex[author] = next
        return next
    }

    private func xmlEscape(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}
