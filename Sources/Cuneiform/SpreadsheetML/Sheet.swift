/// A high-level view of a worksheet with resolved cell values.
public struct Sheet: Sendable {
    private let rawData: WorksheetData
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    private let commentsList: [Comment]
    private let chartsList: [ChartData]

    /// Dimension of the worksheet
    public var dimension: String? { rawData.dimension }

    /// Number of rows with data
    public var rowCount: Int { rawData.rows.count }

    /// Merged cell ranges
    public var mergedCells: [String] { rawData.mergedCells }

    /// Data validations defined in the worksheet
    public var dataValidations: [WorksheetData.DataValidation] { rawData.dataValidations }

    /// Hyperlinks defined in the worksheet
    public var hyperlinks: [WorksheetData.Hyperlink] { rawData.hyperlinks }

    /// Conditional formats defined in the worksheet
    public var conditionalFormats: [WorksheetData.ConditionalFormat] { rawData.conditionalFormats }

    /// AutoFilter configuration for column filtering (if set)
    public var autoFilter: WorksheetData.AutoFilter? { rawData.autoFilter }

    /// Comments (notes) defined in the worksheet
    public var comments: [Comment] { commentsList }

    /// Charts embedded in the worksheet
    public var charts: [ChartData] { chartsList }

    /// Sheet protection state (if any)
    public var protection: WorksheetData.Protection? { rawData.protection }

    /// Page setup configuration (if set)
    public var pageSetup: PageSetup? { rawData.pageSetup }

    /// Get data validations intersecting a given A1 range (e.g., "A1:C3").
    public func validations(for range: String) -> [WorksheetData.DataValidation] {
        guard let (start, end) = parseRange(range) else { return [] }
        return rawData.dataValidations.filter { dv in
            intersects(rangeList: dv.sqref, withStart: start, end: end)
        }
    }

    /// Get data validations that apply to a specific cell reference.
    public func validations(at ref: CellReference) -> [WorksheetData.DataValidation] {
        return rawData.dataValidations.filter { dv in
            contains(rangeList: dv.sqref, reference: ref)
        }
    }

    /// Get hyperlinks applied to a specific cell reference.
    public func hyperlinks(at ref: CellReference) -> [WorksheetData.Hyperlink] {
        rawData.hyperlinks.filter { $0.ref == ref }
    }

    /// Get hyperlinks applied to a specific cell by string reference.
    public func hyperlinks(at ref: String) -> [WorksheetData.Hyperlink] {
        guard let cellRef = CellReference(ref) else { return [] }
        return hyperlinks(at: cellRef)
    }

    /// Get comments applied to a specific cell reference.
    public func comments(at ref: CellReference) -> [Comment] {
        commentsList.filter { $0.ref == ref }
    }

    /// Get comments applied to a specific cell by string reference.
    public func comments(at ref: String) -> [Comment] {
        guard let cellRef = CellReference(ref) else { return [] }
        return comments(at: cellRef)
    }

    init(data: WorksheetData, sharedStrings: SharedStrings, styles: StylesInfo, comments: [Comment] = [], charts: [ChartData] = []) {
        self.rawData = data
        self.sharedStrings = sharedStrings
        self.styles = styles
        self.commentsList = comments
        self.chartsList = charts
    }

    /// Get resolved cell value by reference
    public func cell(at ref: CellReference) -> CellValue? {
        guard let rawCell = rawData.cell(at: ref) else { return nil }
        return resolve(rawCell)
    }

    /// Get resolved cell value by string reference (e.g., "A1")
    public func cell(at ref: String) -> CellValue? {
        guard let cellRef = CellReference(ref) else { return nil }
        return cell(at: cellRef)
    }

    /// Get rich text runs from a cell if present (formatted text)
    public func richText(at ref: CellReference) -> RichText? {
        guard let cellValue = cell(at: ref) else { return nil }
        if case .richText(let runs) = cellValue {
            return runs
        }
        return nil
    }

    /// Get rich text runs from a cell by string reference if present (formatted text)
    public func richText(at ref: String) -> RichText? {
        guard let cellRef = CellReference(ref) else { return nil }
        return richText(at: cellRef)
    }

    /// Get complete cell style for a cell reference
    public func cellStyle(at ref: CellReference) -> CellStyle? {
        guard let rawCell = rawData.cell(at: ref) else { return nil }
        guard let styleIndex = rawCell.styleIndex else { return nil }
        return styles.cellStyle(forStyleIndex: styleIndex)
    }

    /// Get complete cell style for a cell by string reference (e.g., "A1")
    public func cellStyle(at ref: String) -> CellStyle? {
        guard let cellRef = CellReference(ref) else { return nil }
        return cellStyle(at: cellRef)
    }

    /// Get all cells in a row
    public func row(_ index: Int) -> [CellValue] {
        guard let rawRow = rawData.rows.first(where: { $0.index == index }) else { return [] }
        return rawRow.cells.map { resolve($0) }
    }

    /// Iterate through all rows lazily without loading all data into memory at once.
    /// This is more memory-efficient for large spreadsheets.
    ///
    /// ```swift
    /// for row in sheet.rows() {
    ///     for (ref, value) in row {
    ///         print("\(ref): \(value)")
    ///     }
    /// }
    /// ```
    public func rows() -> RowIterator {
        RowIterator(rawData: rawData, sharedStrings: sharedStrings, styles: styles)
    }

    /// Get formula for a cell by reference (if present)
    public func formula(at ref: CellReference) -> CellFormula? {
        rawData.cell(at: ref)?.formula
    }

    /// Get formula for a cell by string reference (if present)
    public func formula(at ref: String) -> CellFormula? {
        guard let cellRef = CellReference(ref) else { return nil }
        return formula(at: cellRef)
    }
    
    // MARK: - Formula Evaluation (Phase 5)
    
    /// Evaluate a formula string in the context of this sheet
    /// - Parameter formulaString: The formula to evaluate (e.g., "=SUM(A1:A10)" or "A1+B1")
    /// - Returns: The evaluated result as a FormulaValue
    /// - Throws: FormulaError if parsing or evaluation fails
    public func evaluate(formula formulaString: String) throws -> FormulaValue {
        let parser = FormulaParser(formulaString)
        let expression = try parser.parse()
        
        let evaluator = FormulaEvaluator { @Sendable [self] ref in
            self.cell(at: ref)
        }
        
        return try evaluator.evaluate(expression)
    }

    // MARK: - Advanced Queries

    /// Get all cells in a rectangular range (e.g., "A1:C3")
    public func range(_ rangeString: String) -> [(reference: CellReference, value: CellValue?)] {
        guard let (start, end) = parseRange(rangeString) else { return [] }
        
        var results: [(CellReference, CellValue?)] = []
        for row in start.row...end.row {
            for col in start.columnIndex...end.columnIndex {
                let colLetter = columnLetterFrom(index: col)
                let ref = CellReference(column: colLetter, row: row)
                results.append((ref, cell(at: ref)))
            }
        }
        return results
    }

    /// Get all cells in a column by letter (e.g., "A")
    public func column(_ letter: String) -> [(row: Int, value: CellValue?)] {
        let upperLetter = letter.uppercased()
        var results: [(row: Int, value: CellValue?)] = []
        
        for rawRow in rawData.rows {
            for rawCell in rawRow.cells {
                if rawCell.reference.column == upperLetter {
                    results.append((row: rawCell.reference.row, value: resolve(rawCell)))
                }
            }
        }
        
        return results.sorted(by: { $0.row < $1.row })
    }

    /// Get all cells in a column by 0-based index (0 = A, 1 = B, etc.)
    public func column(at index: Int) -> [(row: Int, value: CellValue?)] {
        let letter = columnLetterFrom(index: index)
        return column(letter)
    }

    /// Get all rows that match a predicate
    public func rows(where predicate: ([(reference: CellReference, value: CellValue)]) -> Bool) -> [[(reference: CellReference, value: CellValue)]] {
        var matchingRows: [[(CellReference, CellValue)]] = []
        
        for rawRow in rawData.rows {
            let rowCells: [(CellReference, CellValue)] = rawRow.cells.map { cell in
                (cell.reference, resolve(cell))
            }
            if predicate(rowCells) {
                matchingRows.append(rowCells)
            }
        }
        
        return matchingRows
    }

    /// Find the first cell matching a predicate
    public func find(where predicate: (CellReference, CellValue) -> Bool) -> (reference: CellReference, value: CellValue)? {
        for rawRow in rawData.rows {
            for rawCell in rawRow.cells {
                let value = resolve(rawCell)
                if predicate(rawCell.reference, value) {
                    return (rawCell.reference, value)
                }
            }
        }
        return nil
    }

    /// Find all cells matching a predicate
    public func findAll(where predicate: (CellReference, CellValue) -> Bool) -> [(reference: CellReference, value: CellValue)] {
        var results: [(CellReference, CellValue)] = []
        
        for rawRow in rawData.rows {
            for rawCell in rawRow.cells {
                let value = resolve(rawCell)
                if predicate(rawCell.reference, value) {
                    results.append((rawCell.reference, value))
                }
            }
        }
        
        return results
    }

    // MARK: - Helper Methods

    /// Parse a range string like "A1:C3" into start and end references
    private func parseRange(_ rangeString: String) -> (start: CellReference, end: CellReference)? {
        let parts = rangeString.split(separator: ":")
        guard parts.count == 2,
              let start = CellReference(String(parts[0])),
              let end = CellReference(String(parts[1])) else {
            return nil
        }
        return (start, end)
    }

    /// Check if any of the ranges in a space-separated list of A1 ranges contains the given reference.
    private func contains(rangeList: String, reference ref: CellReference) -> Bool {
        for token in rangeList.split(separator: " ") {
            let s = String(token)
            if s.contains(":") {
                if let (start, end) = parseRange(s), within(ref, start: start, end: end) {
                    return true
                }
            } else if let single = CellReference(s), single == ref {
                return true
            }
        }
        return false
    }

    /// Check if any of the ranges in list intersects the given rectangular region.
    private func intersects(rangeList: String, withStart start: CellReference, end: CellReference) -> Bool {
        for token in rangeList.split(separator: " ") {
            let s = String(token)
            if s.contains(":") {
                if let (s2, e2) = parseRange(s), rangesIntersect(aStart: start, aEnd: end, bStart: s2, bEnd: e2) {
                    return true
                }
            } else if let single = CellReference(s), within(single, start: start, end: end) {
                return true
            }
        }
        return false
    }

    private func within(_ ref: CellReference, start: CellReference, end: CellReference) -> Bool {
        let minRow = min(start.row, end.row)
        let maxRow = max(start.row, end.row)
        let minCol = min(start.columnIndex, end.columnIndex)
        let maxCol = max(start.columnIndex, end.columnIndex)
        return ref.row >= minRow && ref.row <= maxRow && ref.columnIndex >= minCol && ref.columnIndex <= maxCol
    }

    private func rangesIntersect(aStart: CellReference, aEnd: CellReference, bStart: CellReference, bEnd: CellReference) -> Bool {
        let aMinRow = min(aStart.row, aEnd.row)
        let aMaxRow = max(aStart.row, aEnd.row)
        let aMinCol = min(aStart.columnIndex, aEnd.columnIndex)
        let aMaxCol = max(aStart.columnIndex, aEnd.columnIndex)

        let bMinRow = min(bStart.row, bEnd.row)
        let bMaxRow = max(bStart.row, bEnd.row)
        let bMinCol = min(bStart.columnIndex, bEnd.columnIndex)
        let bMaxCol = max(bStart.columnIndex, bEnd.columnIndex)

        let rowsOverlap = aMinRow <= bMaxRow && bMinRow <= aMaxRow
        let colsOverlap = aMinCol <= bMaxCol && bMinCol <= aMaxCol
        return rowsOverlap && colsOverlap
    }

    /// Convert a 0-based column index to column letter(s)
    private func columnLetterFrom(index: Int) -> String {
        var result = ""
        var num = index + 1
        
        while num > 0 {
            let remainder = (num - 1) % 26
            result = String(UnicodeScalar(65 + remainder)!) + result
            num = (num - 1) / 26
        }
        
        return result
    }

    /// Resolve a raw cell value using shared strings and style information
    private func resolve(_ cell: RawCell) -> CellValue {
        switch cell.value {
        case .sharedString(let index):
            // Check if this shared string has rich text formatting
            if let richRuns = sharedStrings.richText(at: index) {
                return .richText(richRuns)
            } else if let str = sharedStrings[index] {
                return .text(str)
            } else {
                return .error("Invalid shared string index: \(index)")
            }

        case .number(let n):
            // Check if this cell is a date based on style
            if let styleIndex = cell.styleIndex,
               styles.isDateFormat(styleIndex: styleIndex) {
                // Excel stores dates as numbers (days since 1900-01-01)
                // For now, return the numeric value; caller can convert if needed
                return .date(String(n))
            }
            return .number(n)

        case .boolean(let b):
            return .boolean(b)

        case .inlineString(let s):
            return .text(s)

        case .error(let e):
            return .error(e)

        case .date(let d):
            return .date(d)

        case .empty:
            return .empty
        }
    }
}

// MARK: - Row Iterator

/// Lazy iterator for streaming through worksheet rows without loading all data into memory.
public struct RowIterator: Sequence {
    private let rawData: WorksheetData
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    
    init(rawData: WorksheetData, sharedStrings: SharedStrings, styles: StylesInfo) {
        self.rawData = rawData
        self.sharedStrings = sharedStrings
        self.styles = styles
    }
    
    public func makeIterator() -> Iterator {
        Iterator(rows: rawData.rows, sharedStrings: sharedStrings, styles: styles)
    }
    
    public struct Iterator: IteratorProtocol {
        private var rows: [RawRow]
        private var currentIndex = 0
        private let sharedStrings: SharedStrings
        private let styles: StylesInfo
        
        init(rows: [RawRow], sharedStrings: SharedStrings, styles: StylesInfo) {
            self.rows = rows
            self.sharedStrings = sharedStrings
            self.styles = styles
        }
        
        public mutating func next() -> [(reference: CellReference, value: CellValue)]? {
            guard currentIndex < rows.count else { return nil }
            
            let rawRow = rows[currentIndex]
            currentIndex += 1
            
            return rawRow.cells.map { cell in
                (reference: cell.reference, value: resolve(cell))
            }
        }
        
        private func resolve(_ cell: RawCell) -> CellValue {
            switch cell.value {
            case .sharedString(let index):
                if let str = sharedStrings[index] {
                    return .text(str)
                } else {
                    return .error("Invalid shared string index: \(index)")
                }
                
            case .number(let n):
                if let styleIndex = cell.styleIndex,
                   isDateFormat(styleIndex: styleIndex) {
                    return .date(String(n))
                }
                return .number(n)
                
            case .boolean(let b):
                return .boolean(b)
                
            case .inlineString(let s):
                return .text(s)
                
            case .error(let e):
                return .error(e)
                
            case .date(let d):
                return .date(d)
                
            case .empty:
                return .empty
            }
        }
        
        private func isDateFormat(styleIndex: Int) -> Bool {
            return styles.isDateFormat(styleIndex: styleIndex)
        }
    }
}
