import Foundation

/// Parsed comments data for a worksheet
public struct CommentsData: Sendable {
    /// Authors list referenced by comments via `authorId`
    public let authors: [String]
    /// All comments in the worksheet
    public let comments: [Comment]
}

/// A single cell comment (a.k.a. note)
public struct Comment: Sendable, Equatable {
    /// Target cell reference for the comment
    public let ref: CellReference
    /// Optional author name (resolved via `authorId`)
    public let author: String?
    /// Plain text content of the comment
    public let text: String
}

/// Parser for `/xl/comments*.xml` parts
public enum CommentsParser {
    public static func parse(data: Data) throws(CuneiformError) -> CommentsData {
        let delegate = _CommentsParser()
        let xml = XMLParser(data: data)
        xml.delegate = delegate

        guard xml.parse() else {
            if let err = delegate.error { throw err }
            throw CuneiformError.malformedXML(part: "/xl/comments.xml", detail: xml.parserError?.localizedDescription ?? "Unknown error")
        }

        return CommentsData(authors: delegate.authors, comments: delegate.comments)
    }
}

// MARK: - XML Delegate

final class _CommentsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?

    private(set) var authors: [String] = []
    private(set) var comments: [Comment] = []

    private var inAuthor = false
    private var authorBuffer = ""

    private var inText = false
    private var textBuffer = ""

    private var currentRef: CellReference?
    private var currentAuthorId: Int?

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String]) {
        switch elementName {
        case "author":
            inAuthor = true
            authorBuffer.removeAll(keepingCapacity: true)
        case "comment":
            if let r = attributeDict["ref"], let ref = CellReference(r) {
                currentRef = ref
            } else {
                error = CuneiformError.missingRequiredElement(element: "comment@ref", inPart: "/xl/comments.xml")
                parser.abortParsing()
                return
            }
            if let aStr = attributeDict["authorId"], let aid = Int(aStr) { currentAuthorId = aid } else { currentAuthorId = nil }
        case "text":
            inText = true
            textBuffer.removeAll(keepingCapacity: true)
        case "t":
            if inText { /* capture text in foundCharacters */ }
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inAuthor { authorBuffer += string }
        if inText { textBuffer += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        switch elementName {
        case "author":
            inAuthor = false
            let author = authorBuffer.trimmingCharacters(in: .whitespacesAndNewlines)
            authors.append(author)
            authorBuffer.removeAll(keepingCapacity: true)
        case "text":
            inText = false
        case "comment":
            guard let ref = currentRef else { break }
            let author = (currentAuthorId != nil && currentAuthorId! >= 0 && currentAuthorId! < authors.count) ? authors[currentAuthorId!] : nil
            let text = textBuffer.trimmingCharacters(in: .whitespacesAndNewlines)
            comments.append(Comment(ref: ref, author: author, text: text))
            // Reset
            currentRef = nil
            currentAuthorId = nil
            textBuffer.removeAll(keepingCapacity: true)
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        if error == nil {
            error = CuneiformError.malformedXML(part: "/xl/comments.xml", detail: parseError.localizedDescription)
        }
    }
}
