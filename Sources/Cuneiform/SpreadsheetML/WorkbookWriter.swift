import Foundation

/// Workbook protection options for structure and windows
public struct WorkbookProtectionOptions: Sendable {
    public var structure: Bool = false
    public var windows: Bool = false

    public init(structure: Bool = false, windows: Bool = false) {
        self.structure = structure
        self.windows = windows
    }

    /// Default: no protection
    public static let `default` = WorkbookProtectionOptions()

    /// Strict: protect both structure and windows
    public static let strict = WorkbookProtectionOptions(structure: true, windows: true)

    /// Structure only: prevent sheet insertion/deletion
    public static let structureOnly = WorkbookProtectionOptions(structure: true, windows: false)

    fileprivate func toProtection(passwordHash: String?) -> WorkbookProtection {
        WorkbookProtection(
            structureProtected: structure,
            windowsProtected: windows,
            passwordHash: passwordHash
        )
    }
}

/// Sheet protection options for customizing which operations are allowed
public struct SheetProtectionOptions: Sendable {
    public var formatCells: Bool = true
    public var formatColumns: Bool = true
    public var formatRows: Bool = true
    public var insertColumns: Bool = true
    public var insertRows: Bool = true
    public var insertHyperlinks: Bool = true
    public var deleteColumns: Bool = true
    public var deleteRows: Bool = true
    public var selectLockedCells: Bool = true
    public var selectUnlockedCells: Bool = true
    public var sort: Bool = true
    public var autoFilter: Bool = true
    public var pivotTables: Bool = true

    public init(
        formatCells: Bool = true,
        formatColumns: Bool = true,
        formatRows: Bool = true,
        insertColumns: Bool = true,
        insertRows: Bool = true,
        insertHyperlinks: Bool = true,
        deleteColumns: Bool = true,
        deleteRows: Bool = true,
        selectLockedCells: Bool = true,
        selectUnlockedCells: Bool = true,
        sort: Bool = true,
        autoFilter: Bool = true,
        pivotTables: Bool = true
    ) {
        self.formatCells = formatCells
        self.formatColumns = formatColumns
        self.formatRows = formatRows
        self.insertColumns = insertColumns
        self.insertRows = insertRows
        self.insertHyperlinks = insertHyperlinks
        self.deleteColumns = deleteColumns
        self.deleteRows = deleteRows
        self.selectLockedCells = selectLockedCells
        self.selectUnlockedCells = selectUnlockedCells
        self.sort = sort
        self.autoFilter = autoFilter
        self.pivotTables = pivotTables
    }

    /// Default: allow all operations (locked content, no protection)
    public static let `default` = SheetProtectionOptions()

    /// Strict: prevent all modifications
    public static let strict = SheetProtectionOptions(
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertRows: false,
        insertHyperlinks: false,
        deleteColumns: false,
        deleteRows: false,
        selectLockedCells: true,
        selectUnlockedCells: true,
        sort: false,
        autoFilter: false,
        pivotTables: false
    )

    /// Readonly: allow only viewing and navigating
    public static let readonly = SheetProtectionOptions(
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertRows: false,
        insertHyperlinks: false,
        deleteColumns: false,
        deleteRows: false,
        selectLockedCells: true,
        selectUnlockedCells: false,
        sort: false,
        autoFilter: false,
        pivotTables: false
    )

    fileprivate func toProtection(passwordHash: String?) -> WorksheetData.Protection {
        WorksheetData.Protection(
            sheet: true,
            content: true,
            objects: false,
            scenarios: false,
            formatCells: formatCells,
            formatColumns: formatColumns,
            formatRows: formatRows,
            insertColumns: insertColumns,
            insertRows: insertRows,
            insertHyperlinks: insertHyperlinks,
            deleteColumns: deleteColumns,
            deleteRows: deleteRows,
            selectLockedCells: selectLockedCells,
            selectUnlockedCells: selectUnlockedCells,
            sort: sort,
            autoFilter: autoFilter,
            pivotTables: pivotTables,
            passwordHash: passwordHash
        )
    }
}

/// High-level API for creating .xlsx workbooks
public struct WorkbookWriter {
    public struct SheetWriter {
        private var builder: WorksheetBuilder
        private var commentsBuilder: CommentsBuilder?
        public let name: String
        
        init(name: String) {
            self.name = name
            self.builder = WorksheetBuilder()
        }
        
        /// Write a value to a cell
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: CellReference) {
            builder.addCell(at: reference, value: value)
        }
        
        /// Write a value to a cell by string reference
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: String) {
            guard let ref = CellReference(reference) else { return }
            write(value, to: ref)
        }
        
        /// Write a value with a style index to a cell
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: CellReference, styleIndex: Int) {
            builder.addCell(at: reference, value: value, styleIndex: styleIndex)
        }
        
        /// Write a value with a style index by string reference
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: String, styleIndex: Int) {
            guard let ref = CellReference(reference) else { return }
            write(value, to: ref, styleIndex: styleIndex)
        }
        
        /// Convenience: write a string
        public mutating func writeText(_ text: String, to reference: CellReference) {
            write(.text(text), to: reference)
        }
        
        /// Convenience: write a string with style
        public mutating func writeText(_ text: String, to reference: CellReference, styleIndex: Int) {
            write(.text(text), to: reference, styleIndex: styleIndex)
        }
        
        /// Convenience: write a number
        public mutating func writeNumber(_ number: Double, to reference: CellReference) {
            write(.number(number), to: reference)
        }
        
        /// Convenience: write a number with style
        public mutating func writeNumber(_ number: Double, to reference: CellReference, styleIndex: Int) {
            write(.number(number), to: reference, styleIndex: styleIndex)
        }
        
        /// Convenience: write a boolean
        public mutating func writeBoolean(_ bool: Bool, to reference: CellReference) {
            write(.boolean(bool), to: reference)
        }
        
        /// Convenience: write a boolean with style
        public mutating func writeBoolean(_ bool: Bool, to reference: CellReference, styleIndex: Int) {
            write(.boolean(bool), to: reference, styleIndex: styleIndex)
        }
        
        /// Convenience: write a formula
        public mutating func writeFormula(_ formula: String, cachedValue: Double? = nil, to reference: CellReference) {
            write(.formula(formula, cachedValue: cachedValue), to: reference)
        }
        
        /// Convenience: write a formula with style
        public mutating func writeFormula(_ formula: String, cachedValue: Double? = nil, to reference: CellReference, styleIndex: Int) {
            write(.formula(formula, cachedValue: cachedValue), to: reference, styleIndex: styleIndex)
        }

        /// Merge a single range (e.g., "A1:B1")
        public mutating func mergeCells(_ range: String) {
            builder.addMergeCell(range)
        }

        /// Merge multiple ranges
        public mutating func mergeCells(_ ranges: [String]) {
            for r in ranges { builder.addMergeCell(r) }
        }

        /// Add a data validation rule
        public mutating func addDataValidation(_ dv: WorksheetBuilder.DataValidation) {
            builder.addDataValidation(dv)
        }

        /// Add a conditional formatting entry for a range (sqref)
        public mutating func addConditionalFormat(range: String, rule: WorksheetData.ConditionalRule) {
            let cf = WorksheetData.ConditionalFormat(range: range, rules: [rule])
            builder.addConditionalFormat(cf)
        }

        /// Add a conditional formatting entry with multiple rules
        public mutating func addConditionalFormat(range: String, rules: [WorksheetData.ConditionalRule]) {
            let cf = WorksheetData.ConditionalFormat(range: range, rules: rules)
            builder.addConditionalFormat(cf)
        }

        /// Set an auto filter range for column filtering (e.g., "A1:D100")
        public mutating func setAutoFilter(range: String) {
            builder.setAutoFilter(range: range)
        }

        /// Add an external hyperlink to a cell
        public mutating func addHyperlinkExternal(at reference: CellReference, url: String, display: String? = nil, tooltip: String? = nil) {
            builder.addHyperlinkExternal(at: reference, url: url, display: display, tooltip: tooltip)
        }

        /// Add an internal hyperlink to a cell
        public mutating func addHyperlinkInternal(at reference: CellReference, location: String, display: String? = nil, tooltip: String? = nil) {
            builder.addHyperlinkInternal(at: reference, location: location, display: display, tooltip: tooltip)
        }

        /// Add a cell comment (note) with optional author.
        public mutating func addComment(at reference: CellReference, text: String, author: String? = nil) {
            var cb = commentsBuilder ?? CommentsBuilder()
            cb.addComment(at: reference, text: text, author: author)
            commentsBuilder = cb
        }

        /// Protect this sheet with optional password and specified protection options
        public mutating func protectSheet(password: String? = nil, options: SheetProtectionOptions = .default) {
            let prot = options.toProtection(passwordHash: password)
            builder.setProtection(prot)
        }
        
        /// Add a table (Excel Table or ListObject) to this sheet
        public mutating func addTable(
            name: String,
            displayName: String? = nil,
            ref: String,
            columns: [(id: Int, name: String, totalsFunction: String?)] = [],
            headerRowCount: Int = 1,
            totalsRowCount: Int = 0,
            tableId: Int = 1
        ) {
            let display = displayName ?? name
            var tableBuilder = TableBuilder(id: tableId, displayName: display, name: name, ref: ref)
            tableBuilder.setRowCounts(header: headerRowCount, totals: totalsRowCount)
            
            for (id, colName, totalsFunc) in columns {
                tableBuilder.addColumn(id: id, name: colName, totalsRowFunction: totalsFunc)
            }
            
            builder.addTable(tableBuilder)
        }
        
        fileprivate func build() -> Data {
            builder.build()
        }
        
        fileprivate func tables() -> [TableBuilder] {
            builder.tablesArray
        }
        
        fileprivate func hyperlinkRelationships() -> [(id: String, target: String)] {
            builder.hyperlinkRelationships()
        }

        fileprivate var hasComments: Bool {
            commentsBuilder?.hasComments ?? false
        }

        fileprivate func buildCommentsPart() -> Data? {
            guard let cb = commentsBuilder, cb.hasComments else { return nil }
            return cb.build()
        }

        fileprivate func buildVMLCommentsPart() -> Data? {
            guard let cb = commentsBuilder, cb.hasComments else { return nil }
            let vmlBuilder = VMLCommentsBuilder(entries: cb.allEntries)
            return vmlBuilder.build()
        }

        fileprivate mutating func setLegacyDrawingRelationship(id: String) {
            builder.setLegacyDrawingRelationship(id: id)
        }
    }
    
    private var sheets: [SheetWriter] = []
    private var stylesBuilder: StylesBuilder
    private var namedRanges: [(name: String, refersTo: String)] = []
    private var workbookProtection: WorkbookProtection?
    
    public init() {
        self.stylesBuilder = StylesBuilder()
    }
    
    /// Add a new sheet
    public mutating func addSheet(named name: String) -> Int {
        sheets.append(SheetWriter(name: name))
        return sheets.count - 1
    }
    
    /// Get a sheet by index for writing
    public mutating func sheet(at index: Int) -> SheetWriter? {
        guard index >= 0, index < sheets.count else { return nil }
        return sheets[index]
    }
    
    /// Modify a sheet
    public mutating func modifySheet(at index: Int, _ modify: (inout SheetWriter) -> Void) {
        guard index >= 0, index < sheets.count else { return }
        modify(&sheets[index])
    }
    
    /// Access the styles builder to add custom styles
    public mutating func style(_ modify: (inout StylesBuilder) -> Void) {
        modify(&stylesBuilder)
    }

    /// Add a named range (definedName)
    public mutating func addNamedRange(name: String, refersTo: String) {
        namedRanges.append((name, refersTo))
    }

    /// Protect the workbook structure and/or windows with optional password
    public mutating func protectWorkbook(password: String? = nil, options: WorkbookProtectionOptions = .default) {
        let prot = options.toProtection(passwordHash: password)
        workbookProtection = prot
    }
    
    /// Save the workbook to a file
    public mutating func save(to url: URL) throws {
        let data = try buildData()
        try data.write(to: url)
    }
    
    /// Build the workbook data
    public mutating func buildData() throws -> Data {
        var zipWriter = ZipWriter()
        
        // Build content types
        var contentTypes = ContentTypesBuilder()
        contentTypes.addOverride(partName: "/xl/workbook.xml", 
                                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        contentTypes.addOverride(partName: "/xl/styles.xml",
                                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        for i in 1...sheets.count {
            contentTypes.addOverride(partName: "/xl/worksheets/sheet\(i).xml",
                                    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        }
        zipWriter.addFile(path: "[Content_Types].xml", data: contentTypes.build())
        
        // Build root relationships
        var rootRels = RelationshipsBuilder()
        rootRels.addRelationship(
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            target: "xl/workbook.xml",
            id: "rId1"
        )
        zipWriter.addFile(path: "_rels/.rels", data: rootRels.build())
        
        // Build workbook
        var workbookBuilder = WorkbookBuilder()
        var workbookRels = RelationshipsBuilder()
        
        for (index, sheet) in sheets.enumerated() {
            let sheetId = index + 1
            let relId = "rId\(sheetId)"
            workbookBuilder.addSheet(name: sheet.name, sheetId: sheetId, relationshipId: relId)
            workbookRels.addRelationship(
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                target: "worksheets/sheet\(sheetId).xml",
                id: relId
            )
        }
        
        // Add styles relationship
        workbookRels.addRelationship(
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            target: "styles.xml",
            id: "rId\(sheets.count + 1)"
        )

        // Add defined names
        for (name, refersTo) in namedRanges {
            workbookBuilder.addDefinedName(name: name, refersTo: refersTo)
        }

        // Add workbook protection
        if let protection = workbookProtection {
            workbookBuilder.setProtection(protection)
        }
        
        zipWriter.addFile(path: "xl/workbook.xml", data: workbookBuilder.build())
        zipWriter.addFile(path: "xl/_rels/workbook.xml.rels", data: workbookRels.build())
        
        // Build styles
        zipWriter.addFile(path: "xl/styles.xml", data: stylesBuilder.build())
        
        // Build worksheets
        var globalTableCount = 0  // Global counter for table numbering across all sheets
        
        for index in sheets.indices {
            let sheetNum = index + 1
            var sheet = sheets[index]
            var wsRels = RelationshipsBuilder()
            var hasWorksheetRels = false

            // External hyperlinks
            let hlRels = sheet.hyperlinkRelationships()
            for (id, target) in hlRels {
                wsRels.addRelationship(
                    type: RelationshipType.hyperlink.uri,
                    target: target,
                    id: id,
                    targetMode: "External"
                )
                hasWorksheetRels = true
            }

            // Comments part + VML drawing
            if sheet.hasComments, let commentsData = sheet.buildCommentsPart() {
                let commentsPath = "/xl/comments\(sheetNum).xml"
                contentTypes.addOverride(partName: commentsPath, contentType: ContentType.comments.value)
                zipWriter.addFile(path: String(commentsPath.dropFirst()), data: commentsData)
                wsRels.addRelationship(
                    type: RelationshipType.comments.uri,
                    target: "../comments\(sheetNum).xml"
                )
                hasWorksheetRels = true

                if let vmlData = sheet.buildVMLCommentsPart() {
                    let vmlPath = "/xl/drawings/vmlDrawing\(sheetNum).vml"
                    let vmlRelId = wsRels.addRelationship(
                        type: RelationshipType.vmlDrawing.uri,
                        target: "../drawings/vmlDrawing\(sheetNum).vml"
                    )
                    contentTypes.addOverride(partName: vmlPath, contentType: ContentType.vmlDrawing.value)
                    zipWriter.addFile(path: String(vmlPath.dropFirst()), data: vmlData)
                    sheet.setLegacyDrawingRelationship(id: vmlRelId)
                }
            }

            // Tables
            let sheetTables = sheet.tables()
            for (tableIndex, var tableBuilder) in sheetTables.enumerated() {
                globalTableCount += 1
                tableBuilder.setTableId(globalTableCount)  // Set globally unique table ID
                let tableRelId = tableIndex + 1  // Per-sheet sequential ID for relationship
                let tablePath = "/xl/tables/table\(globalTableCount).xml"
                contentTypes.addOverride(partName: tablePath, contentType: ContentType.table.value)
                zipWriter.addFile(path: String(tablePath.dropFirst()), data: tableBuilder.build())
                
                wsRels.addRelationship(
                    type: RelationshipType.table.uri,
                    target: "../tables/table\(globalTableCount).xml",
                    id: "rIdTable\(tableRelId)"
                )
                hasWorksheetRels = true
            }

            zipWriter.addFile(path: "xl/worksheets/sheet\(sheetNum).xml", data: sheet.build())

            if hasWorksheetRels {
                zipWriter.addFile(path: "xl/worksheets/_rels/sheet\(sheetNum).xml.rels", data: wsRels.build())
            }

            sheets[index] = sheet
        }
        
        return try zipWriter.write()
    }
}
