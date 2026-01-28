import Foundation

/// Builder for creating Excel chart XML
///
/// Generates DrawingML chart markup following ECMA-376 and ISO/IEC 29500.
/// Charts consist of:
/// - chartSpace (root element in /xl/charts/chartN.xml)
/// - drawing (relationship to worksheet via /xl/drawings/drawingN.xml)
/// - Anchor positioning for the chart on the worksheet
public struct ChartBuilder: Sendable {
    /// Chart configuration
    public struct ChartConfig: Sendable {
        public let type: ChartType
        public let title: String?
        public let dataRange: String
        public let categoryRange: String?
        public let seriesNames: [String]
        public let showLegend: Bool
        public let showDataLabels: Bool
        
        public init(
            type: ChartType,
            title: String? = nil,
            dataRange: String,
            categoryRange: String? = nil,
            seriesNames: [String] = [],
            showLegend: Bool = true,
            showDataLabels: Bool = false
        ) {
            self.type = type
            self.title = title
            self.dataRange = dataRange
            self.categoryRange = categoryRange
            self.seriesNames = seriesNames
            self.showLegend = showLegend
            self.showDataLabels = showDataLabels
        }
    }
    
    /// Anchor position for chart placement
    public struct ChartAnchor: Sendable {
        public let fromCol: Int
        public let fromRow: Int
        public let toCol: Int
        public let toRow: Int
        public let fromColOff: Int
        public let fromRowOff: Int
        public let toColOff: Int
        public let toRowOff: Int
        
        public init(
            fromCol: Int,
            fromRow: Int,
            toCol: Int,
            toRow: Int,
            fromColOff: Int = 0,
            fromRowOff: Int = 0,
            toColOff: Int = 0,
            toRowOff: Int = 0
        ) {
            self.fromCol = fromCol
            self.fromRow = fromRow
            self.toCol = toCol
            self.toRow = toRow
            self.fromColOff = fromColOff
            self.fromRowOff = fromRowOff
            self.toColOff = toColOff
            self.toRowOff = toRowOff
        }
        
        /// Convenience initializer for cell range (e.g., "E2:L15")
        public static func fromRange(_ range: String) -> ChartAnchor? {
            let parts = range.split(separator: ":")
            guard parts.count == 2,
                  let from = CellReference(String(parts[0])),
                  let to = CellReference(String(parts[1])) else {
                return nil
            }
            
            return ChartAnchor(
                fromCol: from.columnIndex,
                fromRow: from.row - 1, // 0-indexed for drawing
                toCol: to.columnIndex,
                toRow: to.row - 1
            )
        }
    }
    
    private var config: ChartConfig
    private var anchor: ChartAnchor
    private var chartId: Int
    
    public init(config: ChartConfig, anchor: ChartAnchor, chartId: Int = 1) {
        self.config = config
        self.anchor = anchor
        self.chartId = chartId
    }
    
    // MARK: - Build Chart XML
    
    /// Build the chart XML for /xl/charts/chartN.xml
    public func buildChartXML() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:date1904 val="0"/>
          <c:lang val="en-US"/>
          <c:roundedCorners val="0"/>
          <c:chart>
        """
        
        // Title
        if let title = config.title {
            xml += """
            
                <c:title>
                  <c:tx>
                    <c:rich>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:r>
                          <a:t>\(escapeXML(title))</a:t>
                        </a:r>
                      </a:p>
                    </c:rich>
                  </c:tx>
                  <c:overlay val="0"/>
                </c:title>
            """
        }
        
        // Plot area with chart type
        xml += """
        
            <c:plotArea>
              <c:layout/>
        """
        
        // Add the specific chart type element
        xml += buildChartTypeElement()
        
        xml += """
        
            </c:plotArea>
        """
        
        // Legend
        if config.showLegend {
            xml += """
            
                <c:legend>
                  <c:legendPos val="r"/>
                  <c:overlay val="0"/>
                </c:legend>
            """
        }
        
        xml += """
        
            <c:plotVisOnly val="1"/>
          </c:chart>
        </c:chartSpace>
        """
        
        return xml.data(using: .utf8)!
    }
    
    /// Build the chart type-specific element
    private func buildChartTypeElement() -> String {
        switch config.type {
        case .columnChart:
            return buildBarOrColumnChart(isBar: false)
        case .barChart:
            return buildBarOrColumnChart(isBar: true)
        case .lineChart:
            return buildLineChart()
        case .pieChart:
            return buildPieChart()
        case .areaChart:
            return buildAreaChart()
        case .scatterChart:
            return buildScatterChart()
        default:
            return buildBarOrColumnChart(isBar: false) // Default to column
        }
    }
    
    private func buildBarOrColumnChart(isBar: Bool) -> String {
        let element = isBar ? "barChart" : "colChart"
        let direction = isBar ? "bar" : "col"
        
        var xml = """
        
              <c:\(element)>
                <c:barDir val="\(direction)"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
        """
        
        // Add series
        xml += buildSeries()
        
        xml += """
        
                <c:axId val="1"/>
                <c:axId val="2"/>
              </c:\(element)>
              <c:catAx>
                <c:axId val="1"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="2"/>
              </c:catAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:crossAx val="1"/>
              </c:valAx>
        """
        
        return xml
    }
    
    private func buildLineChart() -> String {
        var xml = """
        
              <c:lineChart>
                <c:grouping val="standard"/>
                <c:varyColors val="0"/>
        """
        
        xml += buildSeries()
        
        xml += """
        
                <c:axId val="1"/>
                <c:axId val="2"/>
              </c:lineChart>
              <c:catAx>
                <c:axId val="1"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="2"/>
              </c:catAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:crossAx val="1"/>
              </c:valAx>
        """
        
        return xml
    }
    
    private func buildPieChart() -> String {
        var xml = """
        
              <c:pieChart>
                <c:varyColors val="1"/>
        """
        
        xml += buildSeries()
        
        xml += """
        
              </c:pieChart>
        """
        
        return xml
    }
    
    private func buildAreaChart() -> String {
        var xml = """
        
              <c:areaChart>
                <c:grouping val="standard"/>
                <c:varyColors val="0"/>
        """
        
        xml += buildSeries()
        
        xml += """
        
                <c:axId val="1"/>
                <c:axId val="2"/>
              </c:areaChart>
              <c:catAx>
                <c:axId val="1"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="2"/>
              </c:catAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:crossAx val="1"/>
              </c:valAx>
        """
        
        return xml
    }
    
    private func buildScatterChart() -> String {
        var xml = """
        
              <c:scatterChart>
                <c:scatterStyle val="lineMarker"/>
                <c:varyColors val="0"/>
        """
        
        xml += buildSeries()
        
        xml += """
        
                <c:axId val="1"/>
                <c:axId val="2"/>
              </c:scatterChart>
              <c:valAx>
                <c:axId val="1"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="2"/>
              </c:valAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:crossAx val="1"/>
              </c:valAx>
        """
        
        return xml
    }
    
    private func buildSeries() -> String {
        var xml = ""
        
        // For now, create a single series from the data range
        let seriesName = config.seriesNames.first ?? "Series 1"
        
        xml += """
        
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:tx>
                    <c:v>\(escapeXML(seriesName))</c:v>
                  </c:tx>
        """
        
        // Category reference
        if let catRange = config.categoryRange {
            xml += """
            
                  <c:cat>
                    <c:strRef>
                      <c:f>\(escapeXML(catRange))</c:f>
                    </c:strRef>
                  </c:cat>
            """
        }
        
        // Values reference
        xml += """
        
                  <c:val>
                    <c:numRef>
                      <c:f>\(escapeXML(config.dataRange))</c:f>
                    </c:numRef>
                  </c:val>
                </c:ser>
        """
        
        return xml
    }
    
    // MARK: - Build Drawing XML
    
    /// Build the drawing XML for /xl/drawings/drawingN.xml
    public func buildDrawingXML(chartRelId: String = "rId1") -> Data {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <xdr:twoCellAnchor>
            <xdr:from>
              <xdr:col>\(anchor.fromCol)</xdr:col>
              <xdr:colOff>\(anchor.fromColOff)</xdr:colOff>
              <xdr:row>\(anchor.fromRow)</xdr:row>
              <xdr:rowOff>\(anchor.fromRowOff)</xdr:rowOff>
            </xdr:from>
            <xdr:to>
              <xdr:col>\(anchor.toCol)</xdr:col>
              <xdr:colOff>\(anchor.toColOff)</xdr:colOff>
              <xdr:row>\(anchor.toRow)</xdr:row>
              <xdr:rowOff>\(anchor.toRowOff)</xdr:rowOff>
            </xdr:to>
            <xdr:graphicFrame macro="">
              <xdr:nvGraphicFramePr>
                <xdr:cNvPr id="\(chartId)" name="Chart \(chartId)"/>
                <xdr:cNvGraphicFramePr/>
              </xdr:nvGraphicFramePr>
              <xdr:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="0" cy="0"/>
              </xdr:xfrm>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="\(chartRelId)"/>
                </a:graphicData>
              </a:graphic>
            </xdr:graphicFrame>
            <xdr:clientData/>
          </xdr:twoCellAnchor>
        </xdr:wsDr>
        """
        
        return xml.data(using: .utf8)!
    }
    
    // MARK: - Build Relationships
    
    /// Build drawing relationships XML for /xl/drawings/_rels/drawingN.xml.rels
    public func buildDrawingRelsXML(chartPath: String = "../charts/chart1.xml") -> Data {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="\(chartPath)"/>
        </Relationships>
        """
        
        return xml.data(using: .utf8)!
    }
    
    // MARK: - Helper
    
    private func escapeXML(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

// MARK: - Drawing Builder

/// Builder for worksheet drawing parts (combines multiple charts/images)
public struct DrawingBuilder: Sendable {
    private var charts: [(config: ChartBuilder.ChartConfig, anchor: ChartBuilder.ChartAnchor)] = []
    private var images: [ImageConfig] = []
    
    public init() {}
    
    /// Add a chart to the drawing
    public mutating func addChart(config: ChartBuilder.ChartConfig, anchor: ChartBuilder.ChartAnchor) {
        charts.append((config, anchor))
    }
    
    /// Add an image to the drawing
    public mutating func addImage(_ config: ImageConfig) {
        images.append(config)
    }
    
    /// Number of charts added
    public var chartCount: Int { charts.count }
    
    /// Number of images added
    public var imageCount: Int { images.count }
    
    /// Build chart XML for the specified index
    public func buildChartXML(at index: Int) -> Data? {
        guard index < charts.count else { return nil }
        let (config, anchor) = charts[index]
        let builder = ChartBuilder(config: config, anchor: anchor, chartId: index + 1)
        return builder.buildChartXML()
    }
    
    /// Get image data for the specified index
    public func imageData(at index: Int) -> Data? {
        guard index < images.count else { return nil }
        return images[index].data
    }
    
    /// Get image config for the specified index
    public func image(at index: Int) -> ImageConfig? {
        guard index < images.count else { return nil }
        return images[index]
    }
    
    /// Build combined drawing XML with all charts and images
    public func buildDrawingXML() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """
        
        var objectId = 1
        
        // Add charts
        for (index, (_, anchor)) in charts.enumerated() {
            let relId = "rId\(index + 1)"
            
            xml += """
            
              <xdr:twoCellAnchor>
                <xdr:from>
                  <xdr:col>\(anchor.fromCol)</xdr:col>
                  <xdr:colOff>\(anchor.fromColOff)</xdr:colOff>
                  <xdr:row>\(anchor.fromRow)</xdr:row>
                  <xdr:rowOff>\(anchor.fromRowOff)</xdr:rowOff>
                </xdr:from>
                <xdr:to>
                  <xdr:col>\(anchor.toCol)</xdr:col>
                  <xdr:colOff>\(anchor.toColOff)</xdr:colOff>
                  <xdr:row>\(anchor.toRow)</xdr:row>
                  <xdr:rowOff>\(anchor.toRowOff)</xdr:rowOff>
                </xdr:to>
                <xdr:graphicFrame macro="">
                  <xdr:nvGraphicFramePr>
                    <xdr:cNvPr id="\(objectId)" name="Chart \(objectId)"/>
                    <xdr:cNvGraphicFramePr/>
                  </xdr:nvGraphicFramePr>
                  <xdr:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                  </xdr:xfrm>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="\(relId)"/>
                    </a:graphicData>
                  </a:graphic>
                </xdr:graphicFrame>
                <xdr:clientData/>
              </xdr:twoCellAnchor>
            """
            objectId += 1
        }
        
        // Add images
        for (index, image) in images.enumerated() {
            let relId = "rId\(charts.count + index + 1)"
            let anchor = image.anchor
            
            xml += """
            
              <xdr:twoCellAnchor editAs="oneCell">
                <xdr:from>
                  <xdr:col>\(anchor.fromCol)</xdr:col>
                  <xdr:colOff>\(anchor.fromColOff)</xdr:colOff>
                  <xdr:row>\(anchor.fromRow)</xdr:row>
                  <xdr:rowOff>\(anchor.fromRowOff)</xdr:rowOff>
                </xdr:from>
                <xdr:to>
                  <xdr:col>\(anchor.toCol)</xdr:col>
                  <xdr:colOff>\(anchor.toColOff)</xdr:colOff>
                  <xdr:row>\(anchor.toRow)</xdr:row>
                  <xdr:rowOff>\(anchor.toRowOff)</xdr:rowOff>
                </xdr:to>
                <xdr:pic>
                  <xdr:nvPicPr>
                    <xdr:cNvPr id="\(objectId)" name="\(escapeXML(image.name))"/>
                    <xdr:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </xdr:cNvPicPr>
                  </xdr:nvPicPr>
                  <xdr:blipFill>
                    <a:blip r:embed="\(relId)"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </xdr:blipFill>
                  <xdr:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="\(image.widthEMU)" cy="\(image.heightEMU)"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                  </xdr:spPr>
                </xdr:pic>
                <xdr:clientData/>
              </xdr:twoCellAnchor>
            """
            objectId += 1
        }
        
        xml += "\n</xdr:wsDr>"
        
        return xml.data(using: .utf8)!
    }
    
    /// Build drawing relationships XML
    public func buildDrawingRelsXML() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
        
        // Chart relationships
        for index in 0..<charts.count {
            let relId = index + 1
            xml += """
            
              <Relationship Id="rId\(relId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart\(relId).xml"/>
            """
        }
        
        // Image relationships
        for index in 0..<images.count {
            let relId = charts.count + index + 1
            let image = images[index]
            xml += """
            
              <Relationship Id="rId\(relId)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image\(index + 1).\(image.format.fileExtension)"/>
            """
        }
        
        xml += "\n</Relationships>"
        
        return xml.data(using: .utf8)!
    }
    
    private func escapeXML(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

// MARK: - Image Support

/// Image format for embedding
public enum ImageFormat: String, Sendable {
    case png
    case jpeg
    case gif
    case bmp
    case tiff
    case svg
    case emf
    case wmf
    
    /// MIME type for content types
    public var mimeType: String {
        switch self {
        case .png: return "image/png"
        case .jpeg: return "image/jpeg"
        case .gif: return "image/gif"
        case .bmp: return "image/bmp"
        case .tiff: return "image/tiff"
        case .svg: return "image/svg+xml"
        case .emf: return "image/x-emf"
        case .wmf: return "image/x-wmf"
        }
    }
    
    /// File extension
    public var fileExtension: String {
        switch self {
        case .png: return "png"
        case .jpeg: return "jpeg"
        case .gif: return "gif"
        case .bmp: return "bmp"
        case .tiff: return "tiff"
        case .svg: return "svg"
        case .emf: return "emf"
        case .wmf: return "wmf"
        }
    }
    
    /// Detect format from data magic bytes
    public static func detect(from data: Data) -> ImageFormat? {
        guard data.count >= 4 else { return nil }
        let bytes = [UInt8](data.prefix(8))
        
        // PNG: 89 50 4E 47 0D 0A 1A 0A
        if bytes.prefix(4) == [0x89, 0x50, 0x4E, 0x47] {
            return .png
        }
        
        // JPEG: FF D8 FF
        if bytes.prefix(3) == [0xFF, 0xD8, 0xFF] {
            return .jpeg
        }
        
        // GIF: 47 49 46 38
        if bytes.prefix(4) == [0x47, 0x49, 0x46, 0x38] {
            return .gif
        }
        
        // BMP: 42 4D
        if bytes.prefix(2) == [0x42, 0x4D] {
            return .bmp
        }
        
        // TIFF: 49 49 2A 00 or 4D 4D 00 2A
        if bytes.prefix(4) == [0x49, 0x49, 0x2A, 0x00] ||
           bytes.prefix(4) == [0x4D, 0x4D, 0x00, 0x2A] {
            return .tiff
        }
        
        // SVG: Check for XML declaration or <svg tag (text-based)
        if let str = String(data: data.prefix(100), encoding: .utf8) {
            let normalized = str.lowercased().trimmingCharacters(in: .whitespacesAndNewlines)
            if normalized.contains("<svg") || normalized.hasPrefix("<?xml") && normalized.contains("svg") {
                return .svg
            }
        }
        
        // EMF (Enhanced Metafile): 01 00 00 00 (little-endian record type)
        // Note: EMF detection is complex; this is a simplified check
        if bytes.count >= 44 && bytes[0] == 0x01 && bytes[1] == 0x00 && 
           bytes[2] == 0x00 && bytes[3] == 0x00 {
            // Additional check: EMF signature at offset 40
            if data.count >= 44 {
                let signature = [UInt8](data[40..<44])
                if signature == [0x20, 0x45, 0x4D, 0x46] { // " EMF"
                    return .emf
                }
            }
        }
        
        // WMF (Windows Metafile): D7 CD C6 9A (placeable metafile header)
        // or 01 00 09 00 (standard metafile header)
        if bytes.prefix(4) == [0xD7, 0xCD, 0xC6, 0x9A] {
            return .wmf
        }
        if bytes.prefix(4) == [0x01, 0x00, 0x09, 0x00] {
            return .wmf
        }
        
        return nil
    }
}

/// Image configuration for embedding
public struct ImageConfig: Sendable {
    /// Image name (displayed in Excel)
    public let name: String
    
    /// Image data
    public let data: Data
    
    /// Image format
    public let format: ImageFormat
    
    /// Position and size anchor
    public let anchor: ChartBuilder.ChartAnchor
    
    /// Width in pixels
    public let widthPixels: Int
    
    /// Height in pixels
    public let heightPixels: Int
    
    /// Width in EMUs (English Metric Units, 914400 EMU = 1 inch)
    public var widthEMU: Int {
        // Assuming 96 DPI
        return widthPixels * 914400 / 96
    }
    
    /// Height in EMUs
    public var heightEMU: Int {
        return heightPixels * 914400 / 96
    }
    
    public init(
        name: String,
        data: Data,
        format: ImageFormat,
        anchor: ChartBuilder.ChartAnchor,
        widthPixels: Int,
        heightPixels: Int
    ) {
        self.name = name
        self.data = data
        self.format = format
        self.anchor = anchor
        self.widthPixels = widthPixels
        self.heightPixels = heightPixels
    }
    
    /// Create from data with auto-detected format
    public static func autoDetect(
        name: String,
        data: Data,
        anchor: ChartBuilder.ChartAnchor,
        widthPixels: Int,
        heightPixels: Int
    ) -> ImageConfig? {
        guard let format = ImageFormat.detect(from: data) else {
            return nil
        }
        return ImageConfig(
            name: name,
            data: data,
            format: format,
            anchor: anchor,
            widthPixels: widthPixels,
            heightPixels: heightPixels
        )
    }
}
