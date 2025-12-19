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

    /// Defined names (named ranges) declared in the workbook
    public let definedNames: [DefinedName]

    /// Workbook protection (if any)
    public let protection: WorkbookProtection?

    /// Get sheet by name
    public func sheet(named name: String) -> SheetInfo? {
        sheets.first { $0.name == name }
    }
}

/// A workbook-level defined name (named range)
///
/// Represents a `<definedName>` entry from `workbook.xml`. The `refersTo` value
/// typically has the form `Sheet!$A$1:$B$10` (quotes around sheet name if needed).
/// Use `Workbook.definedName(_:)` to fetch by name and `Workbook.definedNameRange(_:)`
/// to split into `(sheet, range)` components.
public struct DefinedName: Sendable, Equatable {
    /// The user-facing name of the range (e.g., "SalesData")
    public let name: String
    /// The formula-like reference describing the target cells (e.g., "Sheet1!$A$1:$B$10")
    public let refersTo: String
}

/// Workbook-level protection settings
///
/// Represents a `<workbookProtection>` element from `workbook.xml`.
/// These settings control sheet visibility, workbook window behavior, and structure changes.
public struct WorkbookProtection: Sendable, Equatable {
    /// Whether users can modify sheet structure (insert/delete/rename sheets)
    public let structureProtected: Bool
    /// Whether users can modify window settings
    public let windowsProtected: Bool
    /// Password hash (if present)
    public let passwordHash: String?

    public init(structureProtected: Bool = false, windowsProtected: Bool = false, passwordHash: String? = nil) {
        self.structureProtected = structureProtected
        self.windowsProtected = windowsProtected
        self.passwordHash = passwordHash
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

        return WorkbookInfo(sheets: delegate.sheets, definedNames: delegate.definedNames, protection: delegate.protection)
    }
}

// MARK: - XML Delegate

final class _WorkbookParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var sheets: [SheetInfo] = []
    private(set) var definedNames: [DefinedName] = []
    private(set) var protection: WorkbookProtection?
    fileprivate var error: CuneiformError?

    private var inDefinedName = false
    private var currentDefinedNameName: String = ""
    private var currentDefinedNameBuffer: String = ""

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        switch elementName {
        case "sheet":
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

        case "workbookProtection":
            // Excel protection attributes: sheet (lock structure), windows (lock window), password
            let structureProtected = (attributeDict["sheet"] != "0")
            let windowsProtected = (attributeDict["windows"] != "0")
            let passwordHash = attributeDict["password"]
            protection = WorkbookProtection(structureProtected: structureProtected, windowsProtected: windowsProtected, passwordHash: passwordHash)

        case "definedName":
            inDefinedName = true
            currentDefinedNameName = attributeDict["name"] ?? ""
            currentDefinedNameBuffer.removeAll(keepingCapacity: true)

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inDefinedName {
            currentDefinedNameBuffer += string
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        if elementName == "definedName" {
            inDefinedName = false
            let refers = currentDefinedNameBuffer.trimmingCharacters(in: .whitespacesAndNewlines)
            definedNames.append(DefinedName(name: currentDefinedNameName, refersTo: refers))
            currentDefinedNameName = ""
            currentDefinedNameBuffer.removeAll(keepingCapacity: true)
        }
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
