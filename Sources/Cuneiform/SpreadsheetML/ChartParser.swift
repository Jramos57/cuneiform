import Foundation

/// Chart type classification
public enum ChartType: String, Sendable {
    case columnChart = "col"
    case barChart = "bar"
    case lineChart = "line"
    case areaChart = "area"
    case pieChart = "pie"
    case doughnutChart = "doughnut"
    case radarChart = "radar"
    case scatterChart = "scatter"
    case bubbleChart = "bubble"
    case stockChart = "stock"
    case surfaceChart = "surface"
    case unknown = "unknown"
}

/// Parsed chart metadata
///
/// Represents a single `<c:chartSpace>` entry from `/xl/charts/chartN.xml`.
/// Charts contain series data, titles, and styling information. For now, this parser
/// extracts basic metadata: chart type, title, and series count. Full chart
/// rendering requires additional styling and color information.
public struct ChartData: Sendable, Equatable {
    /// Chart type (column, bar, line, pie, etc.)
    public let type: ChartType

    /// Chart title (if present)
    public let title: String?

    /// Number of data series in the chart
    public let seriesCount: Int

    /// Cell range or reference for data source (if extractable)
    public let dataRange: String?

    public init(type: ChartType, title: String? = nil, seriesCount: Int = 0, dataRange: String? = nil) {
        self.type = type
        self.title = title
        self.seriesCount = seriesCount
        self.dataRange = dataRange
    }
}

/// Parser for chart XML
public enum ChartParser {
    /// Parse chart data from XML
    public static func parse(data: Data) throws(CuneiformError) -> ChartData {
        let delegate = _ChartParser()
        let xml = XMLParser(data: data)
        xml.delegate = delegate
        xml.shouldProcessNamespaces = false

        if !xml.parse() {
            if let err = delegate.error { throw err }
            throw CuneiformError.malformedXML(
                part: "/xl/charts/chart.xml",
                detail: xml.parserError?.localizedDescription ?? "Unknown error"
            )
        }

        return ChartData(
            type: delegate.chartType,
            title: delegate.chartTitle,
            seriesCount: delegate.seriesCount,
            dataRange: delegate.dataRange
        )
    }
}

// MARK: - XML Delegate

final class _ChartParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    fileprivate var error: CuneiformError?

    private(set) var chartType: ChartType = .unknown
    private(set) var chartTitle: String?
    private(set) var seriesCount: Int = 0
    private(set) var dataRange: String?

    private var inTitle = false
    private var titleBuffer = ""
    private var inPlotArea = false
    private var elementStack: [String] = []

    func parser(
        _ parser: XMLParser,
        didStartElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?,
        attributes attributeDict: [String: String]
    ) {
        let localName = elementName.split(separator: ":").last.map(String.init) ?? elementName
        elementStack.append(localName)
        
        // Extract chart type from plot area children
        if inPlotArea {
            // Check for chart type variations
            let typeChecks = [
                ("colChart", ChartType.columnChart),
                ("col3DChart", ChartType.columnChart),
                ("barChart", ChartType.barChart),
                ("bar3DChart", ChartType.barChart),
                ("lineChart", ChartType.lineChart),
                ("line3DChart", ChartType.lineChart),
                ("areaChart", ChartType.areaChart),
                ("area3DChart", ChartType.areaChart),
                ("pieChart", ChartType.pieChart),
                ("pie3DChart", ChartType.pieChart),
                ("doughnutChart", ChartType.doughnutChart),
                ("radarChart", ChartType.radarChart),
                ("scatterChart", ChartType.scatterChart),
                ("bubbleChart", ChartType.bubbleChart),
                ("stockChart", ChartType.stockChart),
                ("surfaceChart", ChartType.surfaceChart),
                ("surface3DChart", ChartType.surfaceChart),
            ]
            
            for (name, type) in typeChecks {
                if localName == name {
                    chartType = type
                    break
                }
            }
            
            if localName == "ser" {
                seriesCount += 1
            }
        }

        switch localName {
        case "plotArea":
            inPlotArea = true

        case "title":
            inTitle = true
            titleBuffer.removeAll(keepingCapacity: true)

        case "cat":
            // Simple extraction of category data range (e.g., from Excel formula reference)
            if let strRef = attributeDict["val"] {
                dataRange = strRef
            }

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        // Collect all characters when in title element
        if inTitle {
            let trimmed = string.trimmingCharacters(in: .whitespaces)
            if !trimmed.isEmpty {
                titleBuffer += trimmed
            }
        }
    }

    func parser(
        _ parser: XMLParser,
        didEndElement elementName: String,
        namespaceURI: String?,
        qualifiedName qName: String?
    ) {
        let localName = elementName.split(separator: ":").last.map(String.init) ?? elementName
        
        if !elementStack.isEmpty {
            _ = elementStack.popLast()
        }
        
        if localName == "plotArea" {
            inPlotArea = false
        }
        
        if localName == "title" {
            inTitle = false
            if !titleBuffer.isEmpty {
                chartTitle = titleBuffer
            }
        }
    }
}
