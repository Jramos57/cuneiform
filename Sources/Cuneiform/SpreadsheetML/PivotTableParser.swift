import Foundation

/// Metadata about a pivot table
public struct PivotTableData: Sendable, Equatable {
    /// Name of the pivot table
    public let name: String
    
    /// Cache ID identifying the data source
    public let cacheId: Int
    
    /// Location of the pivot table on the worksheet (e.g., "B4:E20")
    public let location: String
    
    /// Field names in the pivot table
    public let fieldNames: [String]
    
    /// Number of row fields
    public let rowFieldCount: Int
    
    /// Number of column fields
    public let colFieldCount: Int
    
    /// Number of data fields
    public let dataFieldCount: Int
    
    /// Indicates if the pivot table uses autoformatting
    public let useAutoFormatting: Bool
    
    public init(
        name: String,
        cacheId: Int,
        location: String,
        fieldNames: [String],
        rowFieldCount: Int,
        colFieldCount: Int,
        dataFieldCount: Int,
        useAutoFormatting: Bool
    ) {
        self.name = name
        self.cacheId = cacheId
        self.location = location
        self.fieldNames = fieldNames
        self.rowFieldCount = rowFieldCount
        self.colFieldCount = colFieldCount
        self.dataFieldCount = dataFieldCount
        self.useAutoFormatting = useAutoFormatting
    }
}

/// Parser for pivot table XML files (/xl/pivotTables/pivotTableN.xml)
public enum PivotTableParser {
    /// Parse a pivot table XML data
    public static func parse(data: Data) throws -> PivotTableData {
        let delegate = _PivotTableParserDelegate()
        let parser = XMLParser(data: data)
        
        // Must use shouldProcessNamespaces = true to get proper callback signatures
        parser.shouldProcessNamespaces = true
        parser.delegate = delegate
        
        let parseResult = parser.parse()
        
        if !parseResult {
            throw CuneiformError.malformedXML(part: "pivotTable", detail: "Failed to parse pivot table XML: \(parser.parserError?.localizedDescription ?? "unknown error")")
        }
        
        guard let result = delegate.result else {
            throw CuneiformError.malformedXML(part: "pivotTable", detail: "Invalid pivot table structure")
        }
        
        return result
    }
}

// MARK: - Private Delegate

private final class _PivotTableParserDelegate: NSObject, XMLParserDelegate {
    var result: PivotTableData?
    
    private var name: String?
    private var cacheId: Int?
    private var location: String?
    private var fieldNames: [String] = []
    private var rowFieldCount: Int = 0
    private var colFieldCount: Int = 0
    private var dataFieldCount: Int = 0
    private var useAutoFormatting: Bool = false
    
    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        // Handle namespaced elements - e.g., "p:pivotTableDefinition" or just "pivotTableDefinition"
        let localName = elementName.split(separator: ":").last.map(String.init) ?? elementName
        
        switch localName {
        case "pivotTableDefinition":
            name = attributeDict["name"]
            if let cacheIdStr = attributeDict["cacheId"], let cacheId = Int(cacheIdStr) {
                self.cacheId = cacheId
            }
            useAutoFormatting = attributeDict["useAutoFormatting"] == "1"
            
        case "location":
            location = attributeDict["ref"]
            
        case "rowFields":
            // Count from the count attribute
            if let countStr = attributeDict["count"], let count = Int(countStr) {
                rowFieldCount = count
            }
            
        case "colFields":
            // Count from the count attribute
            if let countStr = attributeDict["count"], let count = Int(countStr) {
                colFieldCount = count
            }
            
        case "dataFields":
            // Count from the count attribute
            if let countStr = attributeDict["count"], let count = Int(countStr) {
                dataFieldCount = count
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
        let localName = elementName.split(separator: ":").last.map(String.init) ?? elementName
        
        if localName == "pivotTableDefinition" {
            // Construct result at end of document - name and cacheId are required
            if let name = name, let cacheId = cacheId {
                result = PivotTableData(
                    name: name,
                    cacheId: cacheId,
                    location: location ?? "",
                    fieldNames: fieldNames,
                    rowFieldCount: rowFieldCount,
                    colFieldCount: colFieldCount,
                    dataFieldCount: dataFieldCount,
                    useAutoFormatting: useAutoFormatting
                )
            }
        }
    }
}
