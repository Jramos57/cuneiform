import Foundation

// MARK: - Domain Types

/// A column within a table
public struct TableColumn: Sendable, Equatable {
    /// Unique ID within the table
    public let id: Int
    /// Column name (displayed in header)
    public let name: String
    /// Optional totals row function for this column (e.g., "sum", "average", "count")
    public let totalsRowFunction: String?
    
    public init(id: Int, name: String, totalsRowFunction: String? = nil) {
        self.id = id
        self.name = name
        self.totalsRowFunction = totalsRowFunction
    }
}

/// Auto-filter configuration for a table
public struct TableAutoFilter: Sendable, Equatable {
    /// Range that the autofilter covers
    public let ref: String?
    
    public init(ref: String? = nil) {
        self.ref = ref
    }
}

/// Table style information
public struct TableStyleInfo: Sendable, Equatable {
    /// Name of the table style (e.g., "TableStyleMedium2")
    public let name: String?
    /// Whether the first column is formatted differently (for names)
    public let showFirstColumn: Bool
    /// Whether the last column is formatted differently (for totals)
    public let showLastColumn: Bool
    /// Whether row stripes are shown
    public let showRowStripes: Bool
    /// Whether column stripes are shown
    public let showColumnStripes: Bool
    
    public init(
        name: String? = nil,
        showFirstColumn: Bool = false,
        showLastColumn: Bool = false,
        showRowStripes: Bool = true,
        showColumnStripes: Bool = false
    ) {
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes
    }
}

/// Parsed table (Excel Table or ListObject) from table.xml
public struct TableData: Sendable, Equatable {
    /// Unique internal table ID (e.g., 1 for table1)
    public let id: Int
    /// Display name of the table
    public let displayName: String
    /// Table name (used in formulas)
    public let name: String
    /// Range covered by the table (e.g., "A1:D100")
    public let ref: String
    /// Number of header rows (typically 1)
    public let headerRowCount: Int
    /// Number of totals rows (typically 0 or 1)
    public let totalsRowCount: Int
    /// Columns in the table
    public let columns: [TableColumn]
    /// Auto-filter configuration if present
    public let autoFilter: TableAutoFilter?
    /// Table style information
    public let tableStyleInfo: TableStyleInfo?
    
    public init(
        id: Int,
        displayName: String,
        name: String,
        ref: String,
        headerRowCount: Int = 1,
        totalsRowCount: Int = 0,
        columns: [TableColumn] = [],
        autoFilter: TableAutoFilter? = nil,
        tableStyleInfo: TableStyleInfo? = nil
    ) {
        self.id = id
        self.displayName = displayName
        self.name = name
        self.ref = ref
        self.headerRowCount = headerRowCount
        self.totalsRowCount = totalsRowCount
        self.columns = columns
        self.autoFilter = autoFilter
        self.tableStyleInfo = tableStyleInfo
    }
}

// MARK: - Parser

/// Parser for table.xml files
public enum TableParser {
    /// Parse a table from table.xml data
    /// - Parameter data: XML data from /xl/tables/tableN.xml
    /// - Parameter id: The table ID (extracted from filename, e.g., 1 for table1.xml)
    /// - Returns: Parsed TableData
    public static func parse(data: Data, id: Int) throws -> TableData {
        let parser = _TableParser(data: data, id: id)
        try parser.parse()
        return parser.result
    }
}

// MARK: - Private Parser Implementation

private final class _TableParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    let id: Int
    var result: TableData!
    
    private var parser: XMLParser
    private var currentElement = ""
    
    private var tableAttributes: [String: String] = [:]
    private var columns: [TableColumn] = []
    private var currentColumn: (id: Int, name: String, totalsRowFunction: String?)?
    private var autoFilter: TableAutoFilter?
    private var tableStyleInfo: TableStyleInfo?
    
    init(data: Data, id: Int) {
        self.id = id
        self.parser = XMLParser(data: data)
        super.init()
        self.parser.delegate = self
    }
    
    func parse() throws {
        if !parser.parse() {
            throw CuneiformError.malformedXML(part: "tables/table\(id).xml", detail: parser.parserError?.localizedDescription ?? "Unknown XML error")
        }
    }
    
    // MARK: - XMLParserDelegate
    
    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String] = [:]
    ) {
        currentElement = elementName
        
        switch elementName {
        case "table":
            tableAttributes = attributeDict
            
        case "tableColumns":
            // Container, no action needed
            break
            
        case "tableColumn":
            let columnId = Int(attributeDict["id"] ?? "0") ?? 0
            let name = attributeDict["name"] ?? ""
            let totalsRowFunction = attributeDict["totalsRowFunction"]
            currentColumn = (id: columnId, name: name, totalsRowFunction: totalsRowFunction)
            
        case "autoFilter":
            let ref = attributeDict["ref"]
            autoFilter = TableAutoFilter(ref: ref)
            
        case "tableStyleInfo":
            let name = attributeDict["name"]
            let showFirstColumn = attributeDict["showFirstColumn"] == "1"
            let showLastColumn = attributeDict["showLastColumn"] == "1"
            let showRowStripes = attributeDict["showRowStripes"] != "0" // Default true
            let showColumnStripes = attributeDict["showColumnStripes"] == "1"
            tableStyleInfo = TableStyleInfo(
                name: name,
                showFirstColumn: showFirstColumn,
                showLastColumn: showLastColumn,
                showRowStripes: showRowStripes,
                showColumnStripes: showColumnStripes
            )
            
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
        case "tableColumn":
            if let column = currentColumn {
                columns.append(TableColumn(
                    id: column.id,
                    name: column.name,
                    totalsRowFunction: column.totalsRowFunction
                ))
                currentColumn = nil
            }
            
        case "table":
            // Build final table data
            let displayName = tableAttributes["displayName"] ?? ""
            let name = tableAttributes["name"] ?? ""
            let ref = tableAttributes["ref"] ?? ""
            let headerRowCount = Int(tableAttributes["headerRowCount"] ?? "1") ?? 1
            let totalsRowCount = Int(tableAttributes["totalsRowCount"] ?? "0") ?? 0
            
            result = TableData(
                id: id,
                displayName: displayName,
                name: name,
                ref: ref,
                headerRowCount: headerRowCount,
                totalsRowCount: totalsRowCount,
                columns: columns,
                autoFilter: autoFilter,
                tableStyleInfo: tableStyleInfo
            )
            
        default:
            break
        }
    }
}
