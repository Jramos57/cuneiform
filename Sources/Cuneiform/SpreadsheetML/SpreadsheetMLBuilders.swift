import Foundation

/// Builder for sharedStrings.xml
public struct SharedStringsBuilder {
    private var strings: [String] = []
    private var stringIndex: [String: Int] = [:]
    
    public init() {}
    
    /// Add a string and return its index
    @discardableResult
    public mutating func addString(_ text: String) -> Int {
        if let existing = stringIndex[text] {
            return existing
        }
        let index = strings.count
        strings.append(text)
        stringIndex[text] = index
        return index
    }
    
    /// Build the XML data
    public func build() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="\(strings.count)" uniqueCount="\(strings.count)">
        """
        
        for str in strings {
            let escaped = xmlEscape(str)
            xml += """
            
            <si><t>\(escaped)</t></si>
            """
        }
        
        xml += """
        
        </sst>
        """
        
        return xml.data(using: .utf8)!
    }
    
    private func xmlEscape(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

/// Builder for workbook.xml
public struct WorkbookBuilder {
    public struct SheetEntry {
        public let name: String
        public let sheetId: Int
        public let relationshipId: String
        
        public init(name: String, sheetId: Int, relationshipId: String) {
            self.name = name
            self.sheetId = sheetId
            self.relationshipId = relationshipId
        }
    }
    
    private var sheets: [SheetEntry] = []
    public struct DefinedName {
        public let name: String
        public let refersTo: String
        public init(name: String, refersTo: String) {
            self.name = name
            self.refersTo = refersTo
        }
    }
    private var definedNames: [DefinedName] = []
    private var workbookProtection: WorkbookProtection?
    
    public init() {}
    
    /// Add a sheet
    public mutating func addSheet(name: String, sheetId: Int, relationshipId: String) {
        sheets.append(SheetEntry(name: name, sheetId: sheetId, relationshipId: relationshipId))
    }

    /// Add a defined name (named range)
    public mutating func addDefinedName(name: String, refersTo: String) {
        definedNames.append(DefinedName(name: name, refersTo: refersTo))
    }

    /// Set workbook protection
    public mutating func setProtection(_ protection: WorkbookProtection) {
        self.workbookProtection = protection
    }
    
    /// Build the XML data
    public func build() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        """
        
        // Add workbook protection if present
        if let protection = workbookProtection {
            xml += "\n<workbookProtection"
            if protection.structureProtected {
                xml += " sheet=\"1\""
            }
            if protection.windowsProtected {
                xml += " windows=\"1\""
            }
            if let passwordHash = protection.passwordHash {
                xml += " password=\"\(xmlEscape(passwordHash))\""
            }
            xml += "/>"
        }
        
        xml += """
        
        <sheets>
        """
        
        for sheet in sheets {
            xml += """
            
            <sheet name="\(xmlEscape(sheet.name))" sheetId="\(sheet.sheetId)" r:id="\(sheet.relationshipId)"/>
            """
        }
        
        xml += """
        
        </sheets>
        """

        if !definedNames.isEmpty {
            xml += """
            
            <definedNames>
            """
            for dn in definedNames {
                xml += """
                
                <definedName name="\(xmlEscape(dn.name))">\(xmlEscape(dn.refersTo))</definedName>
                """
            }
            xml += """
            
            </definedNames>
            """
        }

        xml += """
        
        </workbook>
        """
        
        return xml.data(using: .utf8)!
    }
    
    private func xmlEscape(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

/// Builder for worksheet XML
public struct WorksheetBuilder {
    public struct CellEntry {
        public let reference: CellReference
        public let value: WritableCellValue
        public let styleIndex: Int  // 0 = default (no style)
        
        public init(reference: CellReference, value: WritableCellValue, styleIndex: Int = 0) {
            self.reference = reference
            self.value = value
            self.styleIndex = styleIndex
        }
    }
    
    /// Cell value types that can be written
    public enum WritableCellValue {
        case text(String)
        case number(Double)
        case boolean(Bool)
        case formula(String, cachedValue: Double?)
    }
    
    private var cells: [CellEntry] = []
    private var sharedStringsBuilder: SharedStringsBuilder?
    private var mergedRanges: [String] = []

    /// Data validation rule
    public struct DataValidation {
        public enum Kind: String {
            case list
            case whole
            case decimal
            case date
            case custom
        }
        public let type: Kind
        public let allowBlank: Bool
        public let sqref: String
        public let formula1: String?
        public let formula2: String?
        /// Optional operator (e.g., "between", "greaterThanOrEqual")
        public let op: String?

        public init(type: Kind, allowBlank: Bool = true, sqref: String, formula1: String? = nil, formula2: String? = nil, op: String? = nil) {
            self.type = type
            self.allowBlank = allowBlank
            self.sqref = sqref
            self.formula1 = formula1
            self.formula2 = formula2
            self.op = op
        }
    }

    private var dataValidations: [DataValidation] = []
    
    /// Hyperlink entry for write-side emission
    public struct HyperlinkEntry {
        public let ref: CellReference
        public let display: String?
        public let tooltip: String?
        public let location: String?
        public let externalURL: String?
        public let relationshipId: String?
    }

    private var hyperlinks: [HyperlinkEntry] = []
    private var nextHyperlinkRelId: Int = 1
    private var protection: WorksheetData.Protection?
    
    public init(sharedStringsBuilder: SharedStringsBuilder? = nil) {
        self.sharedStringsBuilder = sharedStringsBuilder
    }
    
    /// Add a cell
    public mutating func addCell(at reference: CellReference, value: WritableCellValue, styleIndex: Int = 0) {
        cells.append(CellEntry(reference: reference, value: value, styleIndex: styleIndex))
    }

    /// Add a merged cell range (e.g., "A1:B1")
    public mutating func addMergeCell(_ range: String) {
        mergedRanges.append(range)
    }

    /// Add a data validation rule
    public mutating func addDataValidation(_ dv: DataValidation) {
        dataValidations.append(dv)
    }
    
    /// Add an external hyperlink (writes a `<hyperlink r:id=...>` and requires a worksheet `.rels` entry)
    public mutating func addHyperlinkExternal(at reference: CellReference, url: String, display: String? = nil, tooltip: String? = nil) {
        let rid = "rIdHL\(nextHyperlinkRelId)"
        nextHyperlinkRelId += 1
        hyperlinks.append(HyperlinkEntry(ref: reference, display: display, tooltip: tooltip, location: nil, externalURL: url, relationshipId: rid))
    }

    /// Add an internal hyperlink (writes a `<hyperlink location=...>`)
    public mutating func addHyperlinkInternal(at reference: CellReference, location: String, display: String? = nil, tooltip: String? = nil) {
        hyperlinks.append(HyperlinkEntry(ref: reference, display: display, tooltip: tooltip, location: location, externalURL: nil, relationshipId: nil))
    }

    /// Set sheet protection with all flags
    public mutating func setProtection(_ protection: WorksheetData.Protection) {
        self.protection = protection
    }
    
    /// Build the XML data
    public func build() -> Data {
        // Group cells by row
        var rowMap: [Int: [CellEntry]] = [:]
        for cell in cells {
            rowMap[cell.reference.row, default: []].append(cell)
        }
        
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetData>
        """
        
        for row in rowMap.keys.sorted() {
            xml += """
            
            <row r="\(row)">
            """
            
            let rowCells = rowMap[row]!.sorted { $0.reference.columnIndex < $1.reference.columnIndex }
            for cell in rowCells {
                xml += buildCellXML(cell)
            }
            
            xml += """
            
            </row>
            """
        }
        
        xml += """

        </sheetData>
        """

        // Merged cells
        if !mergedRanges.isEmpty {
            xml += """
            
            <mergeCells count="\(mergedRanges.count)">
            """
            for r in mergedRanges {
                xml += """
                
                <mergeCell ref="\(r)"/>
                """
            }
            xml += """
            
            </mergeCells>
            """
        }

        // Data validations (minimal subset)
        if !dataValidations.isEmpty {
            xml += """
            
            <dataValidations count="\(dataValidations.count)">
            """
            for dv in dataValidations {
                let allowBlankAttr = dv.allowBlank ? " allowBlank=\"1\"" : ""
                    let opAttr = dv.op != nil ? " operator=\"\(dv.op!)\"" : ""
                xml += """
                
                <dataValidation type="\(dv.type.rawValue)"\(allowBlankAttr)\(opAttr) sqref="\(dv.sqref)">
                """
                if let f1 = dv.formula1 {
                    xml += "\n<formula1>\(xmlEscape(f1))</formula1>"
                }
                if let f2 = dv.formula2 {
                    xml += "\n<formula2>\(xmlEscape(f2))</formula2>"
                }
                xml += """
                
                </dataValidation>
                """
            }
            xml += """
            
            </dataValidations>
            """
        }

        // Hyperlinks
        if !hyperlinks.isEmpty {
            xml += """
            
            <hyperlinks>
            """
            for h in hyperlinks {
                var attrs = " ref=\"\(h.ref)\""
                if let rid = h.relationshipId { attrs += " r:id=\"\(rid)\"" }
                if let display = h.display { attrs += " display=\"\(xmlEscape(display))\"" }
                if let tooltip = h.tooltip { attrs += " tooltip=\"\(xmlEscape(tooltip))\"" }
                if let loc = h.location { attrs += " location=\"\(xmlEscape(loc))\"" }
                xml += """
                
                <hyperlink\(attrs)/>
                """
            }
            xml += """
            
            </hyperlinks>
            """
        }

        if let legacyRid = legacyDrawingRelId {
            xml += """

            <legacyDrawing r:id="\(legacyRid)"/>
            """
        }

        if let prot = protection {
            xml += """

            <sheetProtection sheet="\(prot.sheet ? "1" : "0")" content="\(prot.content ? "1" : "0")" objects="\(prot.objects ? "1" : "0")" scenarios="\(prot.scenarios ? "1" : "0")" formatCells="\(prot.formatCells ? "0" : "1")" formatColumns="\(prot.formatColumns ? "0" : "1")" formatRows="\(prot.formatRows ? "0" : "1")" insertColumns="\(prot.insertColumns ? "0" : "1")" insertRows="\(prot.insertRows ? "0" : "1")" insertHyperlinks="\(prot.insertHyperlinks ? "0" : "1")" deleteColumns="\(prot.deleteColumns ? "0" : "1")" deleteRows="\(prot.deleteRows ? "0" : "1")" selectLockedCells="\(prot.selectLockedCells ? "0" : "1")" selectUnlockedCells="\(prot.selectUnlockedCells ? "0" : "1")" sort="\(prot.sort ? "0" : "1")" autoFilter="\(prot.autoFilter ? "0" : "1")" pivotTables="\(prot.pivotTables ? "0" : "1")"\(prot.passwordHash != nil ? " password=\"\(xmlEscape(prot.passwordHash!))\"" : "")/>
            """
        }

        xml += """
        
        </worksheet>
        """
        
        return xml.data(using: .utf8)!
    }
    
    /// Relationships to emit for external hyperlinks in this worksheet
    public func hyperlinkRelationships() -> [(id: String, target: String)] {
        hyperlinks.compactMap { h in
            if let rid = h.relationshipId, let url = h.externalURL { return (rid, url) }
            return nil
        }
    }

    // MARK: - Legacy Drawing (comments VML)

    private var legacyDrawingRelId: String?

    public mutating func setLegacyDrawingRelationship(id: String) {
        legacyDrawingRelId = id
    }
    
    private func buildCellXML(_ cell: CellEntry) -> String {
        let ref = "\(cell.reference.column)\(cell.reference.row)"
        let styleAttr = cell.styleIndex > 0 ? " s=\"\(cell.styleIndex)\"" : ""
        
        switch cell.value {
        case .text(let str):
            // For simplicity, always use inline strings (could use shared strings)
            let escaped = xmlEscape(str)
            return """
            
            <c r="\(ref)" t="str"\(styleAttr)><v>\(escaped)</v></c>
            """
            
        case .number(let num):
            return """
            
            <c r="\(ref)"\(styleAttr)><v>\(num)</v></c>
            """
            
        case .boolean(let bool):
            return """
            
            <c r="\(ref)" t="b"\(styleAttr)><v>\(bool ? "1" : "0")</v></c>
            """
            
        case .formula(let formula, let cached):
            if let cached = cached {
                return """
                
                <c r="\(ref)"\(styleAttr)><f>\(xmlEscape(formula))</f><v>\(cached)</v></c>
                """
            } else {
                return """
                
                <c r="\(ref)"\(styleAttr)><f>\(xmlEscape(formula))</f></c>
                """
            }
        }
    }
    
    private func xmlEscape(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

/// Builder for styles.xml
public struct StylesBuilder {
    /// Number format definition
    public struct NumberFormat {
        public let id: Int
        public let code: String
        
        public init(id: Int, code: String) {
            self.id = id
            self.code = code
        }
    }
    
    /// Font definition (simplified)
    public struct Font {
        public let bold: Bool
        public let italic: Bool
        public let size: Int?
        public let color: String?
        
        public init(bold: Bool = false, italic: Bool = false, size: Int? = nil, color: String? = nil) {
            self.bold = bold
            self.italic = italic
            self.size = size
            self.color = color
        }
    }
    
    /// Fill definition (simplified: color only)
    public struct Fill {
        public let patternType: String
        public let fgColor: String?
        
        public init(patternType: String = "solid", fgColor: String? = nil) {
            self.patternType = patternType
            self.fgColor = fgColor
        }
    }
    
    /// Border definition (simplified)
    public struct Border {
        public let left: (style: String, color: String)?
        public let right: (style: String, color: String)?
        public let top: (style: String, color: String)?
        public let bottom: (style: String, color: String)?
        
        public init(
            left: (String, String)? = nil,
            right: (String, String)? = nil,
            top: (String, String)? = nil,
            bottom: (String, String)? = nil
        ) {
            self.left = left
            self.right = right
            self.top = top
            self.bottom = bottom
        }
    }
    
    /// Cell format (xf) definition
    public struct CellFormat {
        public let numFmtId: Int
        public let fontId: Int
        public let fillId: Int
        public let borderId: Int
        public let horizontalAlignment: String?
        public let verticalAlignment: String?
        
        public init(
            numFmtId: Int = 0,
            fontId: Int = 0,
            fillId: Int = 0,
            borderId: Int = 0,
            horizontalAlignment: String? = nil,
            verticalAlignment: String? = nil
        ) {
            self.numFmtId = numFmtId
            self.fontId = fontId
            self.fillId = fillId
            self.borderId = borderId
            self.horizontalAlignment = horizontalAlignment
            self.verticalAlignment = verticalAlignment
        }
    }
    
    private var numberFormats: [NumberFormat] = []
    private var fonts: [Font] = []
    private var fills: [Fill] = []
    private var borders: [Border] = []
    private var cellFormats: [CellFormat] = []
    
    public init() {
        // Built-in defaults (Excel standard)
        fonts.append(Font())  // Default font at index 0
        fills.append(Fill(patternType: "none"))  // Default fill at index 0
        fills.append(Fill(patternType: "gray125"))  // Built-in fill at index 1
        borders.append(Border())  // Default border at index 0
        // Default cell format (index 0) with numFmtId 0
        cellFormats.append(CellFormat(numFmtId: 0, fontId: 0, fillId: 0, borderId: 0))
    }
    
    /// Add or get a number format ID
    @discardableResult
    public mutating func addNumberFormat(_ code: String) -> Int {
        // Check if already exists
        for (idx, nf) in numberFormats.enumerated() {
            if nf.code == code {
                return 164 + idx  // Custom formats start at 164
            }
        }
        numberFormats.append(NumberFormat(id: 164 + numberFormats.count, code: code))
        return 164 + numberFormats.count - 1
    }
    
    /// Add or get a font ID
    @discardableResult
    public mutating func addFont(_ font: Font) -> Int {
        // Check if already exists
        for (idx, f) in fonts.enumerated() {
            if f.bold == font.bold && f.italic == font.italic && f.size == font.size && f.color == font.color {
                return idx
            }
        }
        fonts.append(font)
        return fonts.count - 1
    }
    
    /// Add or get a fill ID
    @discardableResult
    public mutating func addFill(_ fill: Fill) -> Int {
        // Check if already exists
        for (idx, f) in fills.enumerated() {
            if f.patternType == fill.patternType && f.fgColor == fill.fgColor {
                return idx
            }
        }
        fills.append(fill)
        return fills.count - 1
    }
    
    /// Add or get a border ID
    @discardableResult
    public mutating func addBorder(_ border: Border) -> Int {
        borders.append(border)
        return borders.count - 1
    }
    
    /// Add a cell format and return its index
    @discardableResult
    public mutating func addCellFormat(_ format: CellFormat) -> Int {
        cellFormats.append(format)
        return cellFormats.count - 1
    }
    
    /// Convenience: create a date format cell style and return its index
    @discardableResult
    public mutating func addDateFormat(_ dateFormatCode: String = "yyyy-mm-dd") -> Int {
        let numFmtId = addNumberFormat(dateFormatCode)
        let format = CellFormat(numFmtId: numFmtId)
        return addCellFormat(format)
    }
    
    /// Add a high-level cell style and return its index
    @discardableResult
    public mutating func addCellStyle(_ style: CellStyle) -> Int {
        // Resolve style components into builder indices
        let numFmtId = style.numberFormat.map { $0.id } ?? 0
        
        let fontId: Int
        if let font = style.font {
            let legacyFont = Font(
                bold: font.bold,
                italic: font.italic,
                size: font.size.map { Int($0) },
                color: font.color?.rgbHex
            )
            fontId = addFont(legacyFont)
        } else {
            fontId = 0
        }
        
        let fillId: Int
        if let fill = style.fill {
            let legacyFill = Fill(
                patternType: fill.pattern.rawValue,
                fgColor: fill.foregroundColor?.rgbHex
            )
            fillId = addFill(legacyFill)
        } else {
            fillId = 0
        }
        
        let borderId: Int
        if let border = style.border {
            let legacyBorder = Border(
                left: border.left.map { ($0.style.rawValue, $0.color?.rgbHex ?? "FF000000") },
                right: border.right.map { ($0.style.rawValue, $0.color?.rgbHex ?? "FF000000") },
                top: border.top.map { ($0.style.rawValue, $0.color?.rgbHex ?? "FF000000") },
                bottom: border.bottom.map { ($0.style.rawValue, $0.color?.rgbHex ?? "FF000000") }
            )
            borderId = addBorder(legacyBorder)
        } else {
            borderId = 0
        }
        
        let format = CellFormat(
            numFmtId: numFmtId,
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            horizontalAlignment: style.alignment?.horizontal?.rawValue,
            verticalAlignment: style.alignment?.vertical?.rawValue
        )
        return addCellFormat(format)
    }
    
    /// Build the styles.xml data
    public func build() -> Data {
        var xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        """
        
        // Number formats
        if !numberFormats.isEmpty {
            xml += """
            
            <numFmts count="\(numberFormats.count)">
            """
            for nf in numberFormats {
                xml += """
                
                <numFmt numFmtId="\(nf.id)" formatCode="\(xmlEscape(nf.code))"/>
                """
            }
            xml += """
            
            </numFmts>
            """
        }
        
        // Fonts
        xml += """
        
        <fonts count="\(fonts.count)">
        """
        for font in fonts {
            xml += """
            
            <font>
            """
            if font.bold {
                xml += "\n<b/>"
            }
            if font.italic {
                xml += "\n<i/>"
            }
            if let size = font.size {
                xml += "\n<sz val=\"\(size)\"/>"
            }
            if let color = font.color {
                xml += "\n<color rgb=\"\(color)\"/>"
            }
            xml += """
            
            </font>
            """
        }
        xml += """
        
        </fonts>
        """
        
        // Fills
        xml += """
        
        <fills count="\(fills.count)">
        """
        for fill in fills {
            if fill.patternType == "none" || fill.patternType == "gray125" {
                xml += """
                
                <fill><patternFill patternType="\(fill.patternType)"/></fill>
                """
            } else {
                xml += """
                
                <fill><patternFill patternType="\(fill.patternType)">
                """
                if let fgColor = fill.fgColor {
                    xml += "\n<fgColor rgb=\"\(fgColor)\"/>"
                }
                xml += """
                
                </patternFill></fill>
                """
            }
        }
        xml += """
        
        </fills>
        """
        
        // Borders
        xml += """
        
        <borders count="\(borders.count)">
        """
        for border in borders {
            xml += """
            
            <border>
            """
            // Left border
            if let (style, color) = border.left {
                xml += "\n<left style=\"\(style)\"><color rgb=\"\(color)\"/></left>"
            } else {
                xml += "\n<left/>"
            }
            // Right border
            if let (style, color) = border.right {
                xml += "\n<right style=\"\(style)\"><color rgb=\"\(color)\"/></right>"
            } else {
                xml += "\n<right/>"
            }
            // Top border
            if let (style, color) = border.top {
                xml += "\n<top style=\"\(style)\"><color rgb=\"\(color)\"/></top>"
            } else {
                xml += "\n<top/>"
            }
            // Bottom border
            if let (style, color) = border.bottom {
                xml += "\n<bottom style=\"\(style)\"><color rgb=\"\(color)\"/></bottom>"
            } else {
                xml += "\n<bottom/>"
            }
            // Diagonal (always empty for now)
            xml += "\n<diagonal/>"
            xml += """
            
            </border>
            """
        }
        xml += """
        
        </borders>
        """
        
        // Cell formats (xf)
        xml += """
        
        <cellXfs count="\(cellFormats.count)">
        """
        for format in cellFormats {
            var xfAttrs = "numFmtId=\"\(format.numFmtId)\" fontId=\"\(format.fontId)\" fillId=\"\(format.fillId)\" borderId=\"\(format.borderId)\""
            if format.numFmtId > 0 || format.fontId > 0 || format.fillId > 0 || format.borderId > 0 {
                xfAttrs += " applyNumberFormat=\"1\""
            }
            if format.fontId > 0 {
                xfAttrs += " applyFont=\"1\""
            }
            if format.fillId > 0 {
                xfAttrs += " applyFill=\"1\""
            }
            if format.borderId > 0 {
                xfAttrs += " applyBorder=\"1\""
            }
            if format.horizontalAlignment != nil || format.verticalAlignment != nil {
                xfAttrs += " applyAlignment=\"1\""
            }
            
            xml += """
            
            <xf \(xfAttrs)>
            """
            
            let ha = format.horizontalAlignment ?? (format.numFmtId > 0 ? "left" : nil)
            let va = format.verticalAlignment ?? "center"
            if ha != nil {
                let h = ha ?? "left"
                xml += "\n<alignment horizontal=\"\(h)\" vertical=\"\(va)\"/>"
            }
            
            xml += """
            
            </xf>
            """
        }
        xml += """
        
        </cellXfs>
        """
        
        xml += """
        
        </styleSheet>
        """
        
        return xml.data(using: .utf8)!
    }
    
    private func xmlEscape(_ text: String) -> String {
        text
            .replacingOccurrences(of: "&", with: "&amp;")
            .replacingOccurrences(of: "<", with: "&lt;")
            .replacingOccurrences(of: ">", with: "&gt;")
            .replacingOccurrences(of: "\"", with: "&quot;")
    }
}
