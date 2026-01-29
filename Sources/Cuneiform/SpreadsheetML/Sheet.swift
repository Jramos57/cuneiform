/// A high-level representation of an Excel worksheet with resolved cell values and formatting.
///
/// `Sheet` provides a convenient, type-safe interface for accessing and querying worksheet data.
/// It resolves shared strings, applies styles, and provides rich access to cell values, formulas,
/// hyperlinks, comments, and other worksheet features.
///
/// ## Overview
///
/// Use `Sheet` to read and analyze data from Excel worksheets. The sheet automatically resolves:
/// - Shared string references to their actual text values
/// - Rich text formatting in cells
/// - Date values based on cell styles
/// - Cell formulas and their evaluated results
/// - Hyperlinks, comments, and data validations
///
/// ## Reading Cell Values
///
/// Access individual cells using A1-style references:
///
/// ```swift
/// let workbook = try Workbook(path: "data.xlsx")
/// let sheet = workbook.sheets[0]
///
/// // Access a single cell
/// if let value = sheet.cell(at: "A1") {
///     print("Cell A1 contains: \(value)")
/// }
///
/// // Use CellReference for type safety
/// let ref = CellReference(column: "B", row: 2)
/// if let value = sheet.cell(at: ref) {
///     print("Cell B2 contains: \(value)")
/// }
/// ```
///
/// ## Iterating Rows
///
/// Efficiently iterate through worksheet rows without loading all data into memory:
///
/// ```swift
/// // Memory-efficient row iteration
/// for row in sheet.rows() {
///     for (reference, value) in row {
///         print("\(reference): \(value)")
///     }
/// }
///
/// // Access a specific row by index
/// let firstRow = sheet.row(1)
/// print("First row has \(firstRow.count) cells")
/// ```
///
/// ## Working with Ranges
///
/// Query rectangular cell ranges and columns:
///
/// ```swift
/// // Get all cells in a range
/// let range = sheet.range("A1:C10")
/// for (reference, value) in range {
///     if let value = value {
///         print("\(reference): \(value)")
///     }
/// }
///
/// // Get all cells in a column
/// let columnA = sheet.column("A")
/// for (row, value) in columnA {
///     print("Row \(row): \(value ?? .empty)")
/// }
/// ```
///
/// ## Filtering and Searching
///
/// Use predicates to find specific cells or rows:
///
/// ```swift
/// // Find the first cell containing a specific value
/// if let result = sheet.find(where: { ref, value in
///     if case .number(let num) = value {
///         return num > 100
///     }
///     return false
/// }) {
///     print("Found at \(result.reference): \(result.value)")
/// }
///
/// // Find all cells matching criteria
/// let matches = sheet.findAll(where: { ref, value in
///     if case .text(let str) = value {
///         return str.contains("error")
///     }
///     return false
/// })
///
/// // Filter rows by condition
/// let nonEmptyRows = sheet.rows(where: { cells in
///     cells.contains(where: { $0.value != .empty })
/// })
/// ```
///
/// ## Accessing Formulas
///
/// Read and evaluate cell formulas:
///
/// ```swift
/// // Get the formula in a cell
/// if let formula = sheet.formula(at: "D10") {
///     print("Formula: \(formula.expression)")
/// }
///
/// // Evaluate a formula
/// let result = try sheet.evaluate(formula: "=SUM(A1:A10)")
/// print("Result: \(result)")
/// ```
///
/// ## Rich Text and Formatting
///
/// Access rich text content and cell styles:
///
/// ```swift
/// // Get rich text formatting
/// if let richText = sheet.richText(at: "A1") {
///     for run in richText.runs {
///         print("Text: \(run.text), Font: \(run.font?.name ?? "default")")
///     }
/// }
///
/// // Get complete cell style
/// if let style = sheet.cellStyle(at: "B2") {
///     print("Font: \(style.font?.name ?? "default")")
///     print("Fill: \(style.fill?.description ?? "none")")
/// }
/// ```
///
/// ## Metadata and Features
///
/// Access worksheet metadata and advanced features:
///
/// ```swift
/// // Check for hyperlinks
/// let links = sheet.hyperlinks(at: "A1")
/// for link in links {
///     print("Link: \(link.target)")
/// }
///
/// // Get comments
/// let comments = sheet.comments(at: "C5")
/// for comment in comments {
///     print("Comment: \(comment.text)")
/// }
///
/// // Check data validations
/// let validations = sheet.validations(at: "E10")
/// for validation in validations {
///     print("Validation: \(validation.type)")
/// }
/// ```
///
/// ## Performance Considerations
///
/// - Use ``rows()`` for memory-efficient iteration over large worksheets
/// - Cell lookups by reference are O(log n) on average
/// - Range operations create arrays and may use more memory
/// - Use ``find(where:)`` instead of ``findAll(where:)`` when you only need the first match
///
/// ## Topics
///
/// ### Reading Cell Values
///
/// - ``cell(at:)-9p4pu``
/// - ``cell(at:)-4b5vl``
/// - ``row(_:)``
/// - ``rows()``
///
/// ### Querying Ranges
///
/// - ``range(_:)``
/// - ``column(_:)``
/// - ``column(at:)``
///
/// ### Searching and Filtering
///
/// - ``find(where:)``
/// - ``findAll(where:)``
/// - ``rows(where:)``
///
/// ### Formulas
///
/// - ``formula(at:)-5oj0d``
/// - ``formula(at:)-8aqpq``
/// - ``evaluate(formula:)``
///
/// ### Rich Text and Styles
///
/// - ``richText(at:)-3lnpg``
/// - ``richText(at:)-8y8ko``
/// - ``cellStyle(at:)-6j2xh``
/// - ``cellStyle(at:)-1c8t7``
///
/// ### Metadata and Features
///
/// - ``hyperlinks(at:)-9k73y``
/// - ``hyperlinks(at:)-4cxjd``
/// - ``comments(at:)-6f0tj``
/// - ``comments(at:)-2hrnz``
/// - ``validations(for:)``
/// - ``validations(at:)``
///
/// ### Worksheet Properties
///
/// - ``dimension``
/// - ``rowCount``
/// - ``mergedCells``
/// - ``autoFilter``
/// - ``protection``
/// - ``pageSetup``
///
/// ### Advanced Features
///
/// - ``dataValidations``
/// - ``hyperlinks``
/// - ``comments``
/// - ``conditionalFormats``
/// - ``charts``
///
public struct Sheet: Sendable {
    private let rawData: WorksheetData
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    private let commentsList: [Comment]
    private let chartsList: [ChartData]

    /// The dimension of the worksheet in A1 notation (e.g., "A1:Z100").
    ///
    /// This represents the used range of the worksheet. May be `nil` if the worksheet is empty
    /// or if dimension information is not available in the file.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let dimension = sheet.dimension {
    ///     print("Used range: \(dimension)")
    ///     // Output: "Used range: A1:Z100"
    /// }
    /// ```
    public var dimension: String? { rawData.dimension }

    /// The number of rows containing data in the worksheet.
    ///
    /// This count includes all rows with at least one cell, whether or not the cells contain values.
    /// Empty rows at the end of the worksheet are not included.
    ///
    /// ## Example
    ///
    /// ```swift
    /// print("Sheet has \(sheet.rowCount) rows with data")
    /// 
    /// // Iterate through all rows by index
    /// for i in 1...sheet.rowCount {
    ///     let row = sheet.row(i)
    ///     print("Row \(i): \(row.count) cells")
    /// }
    /// ```
    ///
    /// - Note: Row indices in Excel start at 1, not 0.
    public var rowCount: Int { rawData.rows.count }

    /// Merged cell ranges defined in the worksheet.
    ///
    /// Each string represents a merged cell range in A1 notation (e.g., "A1:B2").
    /// When cells are merged, only the top-left cell contains the value; other cells
    /// in the range are typically empty.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for mergedRange in sheet.mergedCells {
    ///     print("Merged range: \(mergedRange)")
    ///     
    ///     // Get the value from the merged range (top-left cell)
    ///     let cells = sheet.range(mergedRange)
    ///     if let first = cells.first {
    ///         print("Value: \(first.value ?? .empty)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``range(_:)``
    public var mergedCells: [String] { rawData.mergedCells }

    /// All data validation rules defined in the worksheet.
    ///
    /// Data validations restrict the type and range of values that users can enter into cells.
    /// Each validation includes the cell ranges it applies to and the validation criteria.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for validation in sheet.dataValidations {
    ///     print("Type: \(validation.type)")
    ///     print("Applies to: \(validation.sqref)")
    ///     if let formula1 = validation.formula1 {
    ///         print("Formula: \(formula1)")
    ///     }
    /// }
    /// ```
    ///
    /// To find validations for a specific cell or range, use ``validations(at:)``
    /// or ``validations(for:)`` instead.
    ///
    /// - SeeAlso: ``validations(at:)``
    /// - SeeAlso: ``validations(for:)``
    public var dataValidations: [WorksheetData.DataValidation] { rawData.dataValidations }

    /// All hyperlinks defined in the worksheet.
    ///
    /// Hyperlinks can point to external URLs, email addresses, or other locations within
    /// the workbook. Each hyperlink includes the cell reference it's attached to and the target.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for link in sheet.hyperlinks {
    ///     print("Cell: \(link.ref)")
    ///     print("Target: \(link.target)")
    ///     if let display = link.display {
    ///         print("Display text: \(display)")
    ///     }
    /// }
    /// ```
    ///
    /// To find hyperlinks for a specific cell, use ``hyperlinks(at:)-9k73y``
    /// or ``hyperlinks(at:)-4cxjd`` instead.
    ///
    /// - SeeAlso: ``hyperlinks(at:)-9k73y``
    /// - SeeAlso: ``hyperlinks(at:)-4cxjd``
    public var hyperlinks: [WorksheetData.Hyperlink] { rawData.hyperlinks }

    /// All conditional formatting rules defined in the worksheet.
    ///
    /// Conditional formats apply visual formatting (colors, icons, data bars) to cells
    /// based on their values or formulas. Each format includes the ranges it applies to
    /// and the formatting rules.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for format in sheet.conditionalFormats {
    ///     print("Applies to: \(format.sqref)")
    ///     for rule in format.rules {
    ///         print("Type: \(rule.type)")
    ///         if let formula = rule.formula {
    ///             print("Formula: \(formula)")
    ///         }
    ///     }
    /// }
    /// ```
    public var conditionalFormats: [WorksheetData.ConditionalFormat] { rawData.conditionalFormats }

    /// The AutoFilter configuration, if column filtering is enabled.
    ///
    /// AutoFilter allows users to filter and sort data in Excel. The configuration includes
    /// the range of cells that can be filtered and any active filter criteria.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let autoFilter = sheet.autoFilter {
    ///     print("Filter range: \(autoFilter.ref)")
    ///     
    ///     for filterColumn in autoFilter.filterColumns {
    ///         print("Column \(filterColumn.colId) has filters")
    ///     }
    /// }
    /// ```
    ///
    /// Returns `nil` if AutoFilter is not enabled for this worksheet.
    public var autoFilter: WorksheetData.AutoFilter? { rawData.autoFilter }

    /// All comments (notes) in the worksheet.
    ///
    /// Comments are notes attached to specific cells. They typically contain text, author
    /// information, and formatting. In Excel, comments appear as small indicators in cells
    /// and can be displayed on hover or permanently.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for comment in sheet.comments {
    ///     print("Cell: \(comment.ref)")
    ///     print("Author: \(comment.author ?? "Unknown")")
    ///     print("Text: \(comment.text)")
    /// }
    /// ```
    ///
    /// To find comments for a specific cell, use ``comments(at:)-6f0tj``
    /// or ``comments(at:)-2hrnz`` instead.
    ///
    /// - SeeAlso: ``comments(at:)-6f0tj``
    /// - SeeAlso: ``comments(at:)-2hrnz``
    public var comments: [Comment] { commentsList }

    /// All charts embedded in the worksheet.
    ///
    /// Charts are graphical representations of data. Each chart includes information about
    /// its type, data ranges, styling, and position in the worksheet.
    ///
    /// ## Example
    ///
    /// ```swift
    /// for chart in sheet.charts {
    ///     print("Chart type: \(chart.type)")
    ///     print("Title: \(chart.title ?? "Untitled")")
    ///     
    ///     for series in chart.series {
    ///         print("Series: \(series.name)")
    ///         print("Values: \(series.values)")
    ///     }
    /// }
    /// ```
    public var charts: [ChartData] { chartsList }

    /// The sheet protection settings, if the worksheet is protected.
    ///
    /// Sheet protection prevents users from making changes to the worksheet structure
    /// or specific cells. Protection can be password-protected and can allow certain
    /// operations while blocking others.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let protection = sheet.protection {
    ///     print("Sheet is protected")
    ///     
    ///     if protection.sheet {
    ///         print("Sheet structure is protected")
    ///     }
    ///     
    ///     if protection.selectLockedCells == true {
    ///         print("Users can select locked cells")
    ///     }
    /// }
    /// ```
    ///
    /// Returns `nil` if the worksheet is not protected.
    public var protection: WorksheetData.Protection? { rawData.protection }

    /// The page setup configuration for printing, if defined.
    ///
    /// Page setup controls how the worksheet appears when printed, including paper size,
    /// orientation, margins, headers, footers, and print scaling.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let pageSetup = sheet.pageSetup {
    ///     print("Orientation: \(pageSetup.orientation ?? "default")")
    ///     print("Paper size: \(pageSetup.paperSize ?? 0)")
    ///     
    ///     if let scale = pageSetup.scale {
    ///         print("Scale: \(scale)%")
    ///     }
    /// }
    /// ```
    ///
    /// Returns `nil` if no custom page setup is defined.
    public var pageSetup: PageSetup? { rawData.pageSetup }

    /// Returns all data validations that intersect the specified range.
    ///
    /// Use this method to find validation rules that apply to any cell within a rectangular range.
    /// The range is specified in A1 notation (e.g., "A1:C3").
    ///
    /// - Parameter range: The range to check in A1 notation (e.g., "A1:C10").
    /// - Returns: An array of data validations that intersect the specified range.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let validations = sheet.validations(for: "A1:C10")
    /// for validation in validations {
    ///     print("Type: \(validation.type)")
    ///     print("Applies to: \(validation.sqref)")
    ///     
    ///     if let errorMessage = validation.error {
    ///         print("Error message: \(errorMessage)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``validations(at:)``
    /// - SeeAlso: ``dataValidations``
    public func validations(for range: String) -> [WorksheetData.DataValidation] {
        guard let (start, end) = parseRange(range) else { return [] }
        return rawData.dataValidations.filter { dv in
            intersects(rangeList: dv.sqref, withStart: start, end: end)
        }
    }

    /// Returns all data validations that apply to a specific cell.
    ///
    /// Use this method to find validation rules for a single cell. A cell can have multiple
    /// validations if it falls within multiple validation ranges.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: An array of data validations that apply to the specified cell.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "A", row: 5)
    /// let validations = sheet.validations(at: cellRef)
    ///
    /// for validation in validations {
    ///     print("Validation type: \(validation.type)")
    ///     
    ///     // Check the allowed values
    ///     if let formula = validation.formula1 {
    ///         print("Allowed values: \(formula)")
    ///     }
    ///     
    ///     // Check for custom error messages
    ///     if validation.showErrorMessage == true,
    ///        let errorTitle = validation.errorTitle,
    ///        let errorMessage = validation.error {
    ///         print("Error: \(errorTitle) - \(errorMessage)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``validations(for:)``
    /// - SeeAlso: ``dataValidations``
    public func validations(at ref: CellReference) -> [WorksheetData.DataValidation] {
        return rawData.dataValidations.filter { dv in
            contains(rangeList: dv.sqref, reference: ref)
        }
    }

    /// Returns all hyperlinks attached to a specific cell.
    ///
    /// Use this method to find hyperlinks for a cell using a type-safe `CellReference`.
    /// A cell can theoretically have multiple hyperlinks, though this is uncommon.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: An array of hyperlinks attached to the specified cell.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "A", row: 1)
    /// let links = sheet.hyperlinks(at: cellRef)
    ///
    /// for link in links {
    ///     print("Link target: \(link.target)")
    ///     
    ///     if let display = link.display {
    ///         print("Display text: \(display)")
    ///     }
    ///     
    ///     if let tooltip = link.tooltip {
    ///         print("Tooltip: \(tooltip)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``hyperlinks(at:)-4cxjd``
    /// - SeeAlso: ``hyperlinks``
    public func hyperlinks(at ref: CellReference) -> [WorksheetData.Hyperlink] {
        rawData.hyperlinks.filter { $0.ref == ref }
    }

    /// Returns all hyperlinks attached to a specific cell by string reference.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "A1") instead of
    /// a `CellReference` object.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "B10").
    /// - Returns: An array of hyperlinks attached to the specified cell, or an empty array
    ///   if the reference is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let links = sheet.hyperlinks(at: "A1")
    ///
    /// if let firstLink = links.first {
    ///     print("Cell A1 links to: \(firstLink.target)")
    /// } else {
    ///     print("No hyperlinks at A1")
    /// }
    /// ```
    ///
    /// - SeeAlso: ``hyperlinks(at:)-9k73y``
    /// - SeeAlso: ``hyperlinks``
    public func hyperlinks(at ref: String) -> [WorksheetData.Hyperlink] {
        guard let cellRef = CellReference(ref) else { return [] }
        return hyperlinks(at: cellRef)
    }

    /// Returns all comments attached to a specific cell.
    ///
    /// Use this method to find comments for a cell using a type-safe `CellReference`.
    /// Comments are notes that can contain text, author information, and formatting.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: An array of comments attached to the specified cell.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "C", row: 5)
    /// let comments = sheet.comments(at: cellRef)
    ///
    /// for comment in comments {
    ///     if let author = comment.author {
    ///         print("Comment by \(author):")
    ///     }
    ///     print(comment.text)
    ///     
    ///     // Access rich text formatting if available
    ///     if let richText = comment.richText {
    ///         for run in richText.runs {
    ///             print("- \(run.text)")
    ///         }
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``comments(at:)-2hrnz``
    /// - SeeAlso: ``comments``
    public func comments(at ref: CellReference) -> [Comment] {
        commentsList.filter { $0.ref == ref }
    }

    /// Returns all comments attached to a specific cell by string reference.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "C5") instead of
    /// a `CellReference` object.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "C5").
    /// - Returns: An array of comments attached to the specified cell, or an empty array
    ///   if the reference is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let comments = sheet.comments(at: "C5")
    ///
    /// if comments.isEmpty {
    ///     print("No comments at C5")
    /// } else {
    ///     print("Found \(comments.count) comment(s):")
    ///     for comment in comments {
    ///         print("- \(comment.text)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``comments(at:)-6f0tj``
    /// - SeeAlso: ``comments``
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

    /// Returns the resolved cell value at the specified cell reference.
    ///
    /// This is the primary method for accessing cell values. It resolves shared strings,
    /// applies date formatting based on cell styles, and returns the appropriate `CellValue` type.
    ///
    /// - Parameter ref: The cell reference to read.
    /// - Returns: The resolved cell value, or `nil` if the cell does not exist.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "A", row: 1)
    ///
    /// if let value = sheet.cell(at: cellRef) {
    ///     switch value {
    ///     case .text(let str):
    ///         print("Text: \(str)")
    ///     case .number(let num):
    ///         print("Number: \(num)")
    ///     case .boolean(let bool):
    ///         print("Boolean: \(bool)")
    ///     case .date(let date):
    ///         print("Date: \(date)")
    ///     case .richText(let richText):
    ///         print("Rich text with \(richText.runs.count) runs")
    ///     case .error(let error):
    ///         print("Error: \(error)")
    ///     case .empty:
    ///         print("Cell is empty")
    ///     }
    /// } else {
    ///     print("Cell does not exist")
    /// }
    /// ```
    ///
    /// - Note: Returns `nil` for cells that don't exist in the worksheet, while `.empty`
    ///   represents cells that exist but have no value.
    ///
    /// - SeeAlso: ``cell(at:)-4b5vl``
    /// - SeeAlso: ``CellValue``
    public func cell(at ref: CellReference) -> CellValue? {
        guard let rawCell = rawData.cell(at: ref) else { return nil }
        return resolve(rawCell)
    }

    /// Returns the resolved cell value at the specified string reference.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "A1") instead of
    /// a `CellReference` object. This is useful when working with user input or dynamic references.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "Z100").
    /// - Returns: The resolved cell value, or `nil` if the cell does not exist or the reference
    ///   is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Simple access
    /// if let value = sheet.cell(at: "A1") {
    ///     print("A1 contains: \(value)")
    /// }
    ///
    /// // Dynamic references
    /// let columnLetter = "B"
    /// let rowNumber = 5
    /// let reference = "\(columnLetter)\(rowNumber)"
    ///
    /// if let value = sheet.cell(at: reference) {
    ///     print("\(reference) contains: \(value)")
    /// }
    ///
    /// // Working with formulas
    /// if let formula = sheet.formula(at: "C10"),
    ///    let value = sheet.cell(at: "C10") {
    ///     print("Formula: \(formula.expression)")
    ///     print("Result: \(value)")
    /// }
    /// ```
    ///
    /// - SeeAlso: ``cell(at:)-9p4pu``
    /// - SeeAlso: ``CellValue``
    public func cell(at ref: String) -> CellValue? {
        guard let cellRef = CellReference(ref) else { return nil }
        return cell(at: cellRef)
    }

    /// Returns the rich text formatting from a cell, if present.
    ///
    /// Rich text cells contain multiple formatted text runs, each with its own font, color,
    /// and style properties. Use this method to access the detailed formatting information.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: The rich text formatting, or `nil` if the cell doesn't exist, is empty,
    ///   or doesn't contain rich text.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "A", row: 1)
    ///
    /// if let richText = sheet.richText(at: cellRef) {
    ///     print("Cell contains \(richText.runs.count) text runs:")
    ///     
    ///     for run in richText.runs {
    ///         print("\nText: \(run.text)")
    ///         
    ///         if let font = run.font {
    ///             if let name = font.name {
    ///                 print("Font: \(name)")
    ///             }
    ///             if let size = font.size {
    ///                 print("Size: \(size)")
    ///             }
    ///             if let color = font.color {
    ///                 print("Color: \(color)")
    ///             }
    ///             if font.bold == true {
    ///                 print("Bold")
    ///             }
    ///             if font.italic == true {
    ///                 print("Italic")
    ///             }
    ///         }
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``richText(at:)-8y8ko``
    /// - SeeAlso: ``cell(at:)-9p4pu``
    /// - SeeAlso: ``RichText``
    public func richText(at ref: CellReference) -> RichText? {
        guard let cellValue = cell(at: ref) else { return nil }
        if case .richText(let runs) = cellValue {
            return runs
        }
        return nil
    }

    /// Returns the rich text formatting from a cell by string reference, if present.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "A1") instead of
    /// a `CellReference` object.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "B5").
    /// - Returns: The rich text formatting, or `nil` if the cell doesn't exist, is empty,
    ///   doesn't contain rich text, or the reference is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let richText = sheet.richText(at: "A1") {
    ///     // Extract plain text from all runs
    ///     let plainText = richText.runs.map { $0.text }.joined()
    ///     print("Plain text: \(plainText)")
    ///     
    ///     // Find bold text
    ///     let boldRuns = richText.runs.filter { $0.font?.bold == true }
    ///     print("Bold text: \(boldRuns.map { $0.text }.joined())")
    /// }
    /// ```
    ///
    /// - SeeAlso: ``richText(at:)-3lnpg``
    /// - SeeAlso: ``cell(at:)-4b5vl``
    /// - SeeAlso: ``RichText``
    public func richText(at ref: String) -> RichText? {
        guard let cellRef = CellReference(ref) else { return nil }
        return richText(at: cellRef)
    }

    /// Returns the complete style information for a cell.
    ///
    /// Cell styles include font, fill, border, alignment, and number format information.
    /// Use this method to access detailed formatting properties for a cell.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: The cell style, or `nil` if the cell doesn't exist or has no style applied.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "B", row: 2)
    ///
    /// if let style = sheet.cellStyle(at: cellRef) {
    ///     // Font information
    ///     if let font = style.font {
    ///         print("Font: \(font.name ?? "default")")
    ///         print("Size: \(font.size ?? 11)")
    ///         if font.bold == true {
    ///             print("Bold")
    ///         }
    ///     }
    ///     
    ///     // Fill/background color
    ///     if let fill = style.fill {
    ///         print("Background: \(fill)")
    ///     }
    ///     
    ///     // Borders
    ///     if let border = style.border {
    ///         if let left = border.left {
    ///             print("Left border: \(left.style ?? "none")")
    ///         }
    ///     }
    ///     
    ///     // Alignment
    ///     if let alignment = style.alignment {
    ///         print("H-Align: \(alignment.horizontal ?? "general")")
    ///         print("V-Align: \(alignment.vertical ?? "bottom")")
    ///     }
    ///     
    ///     // Number format
    ///     if let numFmt = style.numberFormat {
    ///         print("Format: \(numFmt.formatCode)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``cellStyle(at:)-1c8t7``
    /// - SeeAlso: ``CellStyle``
    public func cellStyle(at ref: CellReference) -> CellStyle? {
        guard let rawCell = rawData.cell(at: ref) else { return nil }
        guard let styleIndex = rawCell.styleIndex else { return nil }
        return styles.cellStyle(forStyleIndex: styleIndex)
    }

    /// Returns the complete style information for a cell by string reference.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "A1") instead of
    /// a `CellReference` object.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "Z100").
    /// - Returns: The cell style, or `nil` if the cell doesn't exist, has no style applied,
    ///   or the reference is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Check if a cell is formatted as currency
    /// if let style = sheet.cellStyle(at: "C5"),
    ///    let numFmt = style.numberFormat {
    ///     if numFmt.formatCode.contains("$") {
    ///         print("Cell C5 is formatted as currency")
    ///     }
    /// }
    ///
    /// // Find cells with specific font
    /// for row in 1...sheet.rowCount {
    ///     if let style = sheet.cellStyle(at: "A\(row)"),
    ///        let font = style.font,
    ///        font.name == "Arial" {
    ///         print("Row \(row) uses Arial font")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``cellStyle(at:)-6j2xh``
    /// - SeeAlso: ``CellStyle``
    public func cellStyle(at ref: String) -> CellStyle? {
        guard let cellRef = CellReference(ref) else { return nil }
        return cellStyle(at: cellRef)
    }

    /// Returns all cells in a specific row by index.
    ///
    /// Use this method to access all cells in a row at once. The row index is 1-based to match
    /// Excel's row numbering.
    ///
    /// - Parameter index: The row index (1-based).
    /// - Returns: An array of resolved cell values in the row, or an empty array if the row
    ///   doesn't exist or is empty.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Get the first row (usually headers)
    /// let headers = sheet.row(1)
    /// print("Column headers:")
    /// for (index, value) in headers.enumerated() {
    ///     print("Column \(index + 1): \(value)")
    /// }
    ///
    /// // Process all rows
    /// for rowIndex in 1...sheet.rowCount {
    ///     let cells = sheet.row(rowIndex)
    ///     print("Row \(rowIndex) has \(cells.count) cells")
    ///     
    ///     // Sum numeric values in the row
    ///     let sum = cells.reduce(0.0) { total, value in
    ///         if case .number(let num) = value {
    ///             return total + num
    ///         }
    ///         return total
    ///     }
    ///     print("Sum: \(sum)")
    /// }
    /// ```
    ///
    /// - Note: This method returns only the cell values, not their references. For cell references,
    ///   use ``rows()`` instead.
    /// - Note: Row indices in Excel start at 1, not 0.
    ///
    /// - SeeAlso: ``rows()``
    /// - SeeAlso: ``rowCount``
    public func row(_ index: Int) -> [CellValue] {
        guard let rawRow = rawData.rows.first(where: { $0.index == index }) else { return [] }
        return rawRow.cells.map { resolve($0) }
    }

    /// Returns an iterator for efficiently traversing all rows in the worksheet.
    ///
    /// This method provides a memory-efficient way to iterate through worksheet rows without
    /// loading all data into memory at once. Each iteration yields a row as an array of
    /// `(reference, value)` tuples, allowing you to access both the cell location and its value.
    ///
    /// - Returns: A lazy iterator that yields rows one at a time.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Basic iteration
    /// for row in sheet.rows() {
    ///     for (reference, value) in row {
    ///         print("\(reference): \(value)")
    ///     }
    /// }
    ///
    /// // Process specific columns
    /// for row in sheet.rows() {
    ///     // Find cells in column A
    ///     let columnA = row.first { $0.reference.column == "A" }
    ///     if let (ref, value) = columnA {
    ///         print("Column A at row \(ref.row): \(value)")
    ///     }
    /// }
    ///
    /// // Filter rows during iteration
    /// for row in sheet.rows() {
    ///     // Only process rows with data in column A
    ///     guard let firstCell = row.first(where: { $0.reference.column == "A" }),
    ///           case .text(let text) = firstCell.value,
    ///           !text.isEmpty else {
    ///         continue
    ///     }
    ///     
    ///     print("Processing row with A value: \(text)")
    /// }
    ///
    /// // Calculate statistics
    /// var totalCells = 0
    /// var nonEmptyCells = 0
    ///
    /// for row in sheet.rows() {
    ///     totalCells += row.count
    ///     nonEmptyCells += row.filter { $0.value != .empty }.count
    /// }
    ///
    /// print("Total cells: \(totalCells)")
    /// print("Non-empty cells: \(nonEmptyCells)")
    /// ```
    ///
    /// ## Performance
    ///
    /// This method is more memory-efficient than loading all rows at once because it processes
    /// one row at a time. This is particularly important for large spreadsheets with thousands
    /// of rows.
    ///
    /// - Note: The iterator yields rows in ascending order by row index.
    /// - Note: Empty rows (rows with no cells) are not included in the iteration.
    ///
    /// - SeeAlso: ``row(_:)``
    /// - SeeAlso: ``rows(where:)``
    /// - SeeAlso: ``RowIterator``
    public func rows() -> RowIterator {
        RowIterator(rawData: rawData, sharedStrings: sharedStrings, styles: styles)
    }

    /// Returns the formula for a cell, if present.
    ///
    /// Use this method to access the formula definition stored in a cell, including the formula
    /// expression and any array formula information.
    ///
    /// - Parameter ref: The cell reference to check.
    /// - Returns: The cell formula, or `nil` if the cell doesn't exist or doesn't contain a formula.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let cellRef = CellReference(column: "D", row: 10)
    ///
    /// if let formula = sheet.formula(at: cellRef) {
    ///     print("Formula: \(formula.expression)")
    ///     
    ///     // Check if it's an array formula
    ///     if formula.isArray {
    ///         print("This is an array formula")
    ///         if let ref = formula.ref {
    ///             print("Array range: \(ref)")
    ///         }
    ///     }
    ///     
    ///     // Get the evaluated result
    ///     if let value = sheet.cell(at: cellRef) {
    ///         print("Result: \(value)")
    ///     }
    /// }
    ///
    /// // Find all cells with formulas
    /// var formulaCells: [(CellReference, CellFormula)] = []
    /// for row in sheet.rows() {
    ///     for (ref, _) in row {
    ///         if let formula = sheet.formula(at: ref) {
    ///             formulaCells.append((ref, formula))
    ///         }
    ///     }
    /// }
    /// print("Found \(formulaCells.count) cells with formulas")
    /// ```
    ///
    /// - SeeAlso: ``formula(at:)-8aqpq``
    /// - SeeAlso: ``evaluate(formula:)``
    /// - SeeAlso: ``CellFormula``
    public func formula(at ref: CellReference) -> CellFormula? {
        rawData.cell(at: ref)?.formula
    }

    /// Returns the formula for a cell by string reference, if present.
    ///
    /// Convenience method that accepts an A1-style string reference (e.g., "D10") instead of
    /// a `CellReference` object.
    ///
    /// - Parameter ref: The cell reference as a string (e.g., "A1", "D10").
    /// - Returns: The cell formula, or `nil` if the cell doesn't exist, doesn't contain a formula,
    ///   or the reference is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Check for a SUM formula
    /// if let formula = sheet.formula(at: "D10") {
    ///     if formula.expression.contains("SUM") {
    ///         print("Cell D10 contains a SUM formula: \(formula.expression)")
    ///     }
    /// }
    ///
    /// // List all formulas in a column
    /// print("Formulas in column D:")
    /// for row in 1...sheet.rowCount {
    ///     let ref = "D\(row)"
    ///     if let formula = sheet.formula(at: ref) {
    ///         print("\(ref): \(formula.expression)")
    ///     }
    /// }
    /// ```
    ///
    /// - SeeAlso: ``formula(at:)-5oj0d``
    /// - SeeAlso: ``evaluate(formula:)``
    /// - SeeAlso: ``CellFormula``
    public func formula(at ref: String) -> CellFormula? {
        guard let cellRef = CellReference(ref) else { return nil }
        return formula(at: cellRef)
    }
    
    // MARK: - Formula Evaluation
    
    /// Evaluates a formula string in the context of this sheet.
    ///
    /// Use this method to evaluate custom formulas or re-evaluate existing cell formulas.
    /// The formula is parsed and evaluated using the current sheet's cell values as context.
    ///
    /// - Parameter formulaString: The formula to evaluate. Can start with "=" or not (e.g.,
    ///   "=SUM(A1:A10)" or "SUM(A1:A10)" or "A1+B1").
    /// - Returns: The evaluated result as a `FormulaValue`.
    /// - Throws: `FormulaError` if the formula cannot be parsed or evaluated.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Evaluate a simple formula
    /// do {
    ///     let result = try sheet.evaluate(formula: "=SUM(A1:A10)")
    ///     print("Sum result: \(result)")
    /// } catch {
    ///     print("Formula error: \(error)")
    /// }
    ///
    /// // Evaluate arithmetic expressions
    /// let sum = try sheet.evaluate(formula: "A1 + B1")
    /// let product = try sheet.evaluate(formula: "C5 * D5")
    ///
    /// // Use functions
    /// let average = try sheet.evaluate(formula: "AVERAGE(A1:A100)")
    /// let max = try sheet.evaluate(formula: "MAX(B:B)")
    /// let conditional = try sheet.evaluate(formula: "IF(A1>100, \"High\", \"Low\")")
    ///
    /// // Combine with cell formulas
    /// if let formula = sheet.formula(at: "E10") {
    ///     do {
    ///         let result = try sheet.evaluate(formula: formula.expression)
    ///         print("Formula \(formula.expression) evaluates to: \(result)")
    ///     } catch {
    ///         print("Could not evaluate: \(error)")
    ///     }
    /// }
    /// ```
    ///
    /// ## Supported Features
    ///
    /// The formula evaluator supports:
    /// - Basic arithmetic operators (+, -, *, /, ^)
    /// - Comparison operators (=, <>, <, >, <=, >=)
    /// - Cell references (A1, B2, etc.)
    /// - Range references (A1:C10)
    /// - Common functions (SUM, AVERAGE, COUNT, IF, etc.)
    /// - String operations and concatenation
    ///
    /// - Note: Not all Excel functions are implemented. Unsupported functions will throw an error.
    ///
    /// - SeeAlso: ``formula(at:)-5oj0d``
    /// - SeeAlso: ``formula(at:)-8aqpq``
    /// - SeeAlso: ``FormulaValue``
    /// - SeeAlso: ``FormulaError``
    public func evaluate(formula formulaString: String) throws -> FormulaValue {
        let parser = FormulaParser(formulaString)
        let expression = try parser.parse()
        
        let evaluator = FormulaEvaluator { @Sendable [self] ref in
            self.cell(at: ref)
        }
        
        return try evaluator.evaluate(expression)
    }

    // MARK: - Advanced Queries

    /// Returns all cells in a rectangular range.
    ///
    /// Use this method to access multiple cells at once within a rectangular region. The range
    /// is specified in A1 notation (e.g., "A1:C3" for a 3x3 grid).
    ///
    /// - Parameter rangeString: The range in A1 notation (e.g., "A1:C10", "B2:D5").
    /// - Returns: An array of tuples containing cell references and their values. Cells that
    ///   don't exist have `nil` values.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Get all cells in a range
    /// let cells = sheet.range("A1:C10")
    ///
    /// for (reference, value) in cells {
    ///     if let value = value {
    ///         print("\(reference): \(value)")
    ///     } else {
    ///         print("\(reference): (does not exist)")
    ///     }
    /// }
    ///
    /// // Calculate sum of numeric values in a range
    /// let range = sheet.range("B2:B100")
    /// let sum = range.reduce(0.0) { total, cell in
    ///     if let value = cell.value, case .number(let num) = value {
    ///         return total + num
    ///     }
    ///     return total
    /// }
    /// print("Sum: \(sum)")
    ///
    /// // Extract values from a specific column in the range
    /// let columnBValues = sheet.range("A1:C10")
    ///     .filter { $0.reference.column == "B" }
    ///     .compactMap { $0.value }
    ///
    /// // Process merged cell ranges
    /// for mergedRange in sheet.mergedCells {
    ///     let cells = sheet.range(mergedRange)
    ///     if let first = cells.first {
    ///         print("Merged range \(mergedRange) has value: \(first.value ?? .empty)")
    ///     }
    /// }
    /// ```
    ///
    /// ## Performance
    ///
    /// This method creates an array containing all cells in the range, which may use significant
    /// memory for large ranges. For memory-efficient processing of large areas, consider using
    /// ``rows()`` or ``column(_:)`` instead.
    ///
    /// - Note: If the range string is invalid, returns an empty array.
    ///
    /// - SeeAlso: ``column(_:)``
    /// - SeeAlso: ``column(at:)``
    /// - SeeAlso: ``rows()``
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

    /// Returns all cells in a column by letter.
    ///
    /// Use this method to access all cells in a specific column. The column is identified by
    /// its letter (A, B, C, etc.). Results are sorted by row number in ascending order.
    ///
    /// - Parameter letter: The column letter (case-insensitive, e.g., "A", "B", "AA").
    /// - Returns: An array of tuples containing row numbers and cell values. Cells that don't
    ///   exist have `nil` values.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Get all cells in column A
    /// let columnA = sheet.column("A")
    ///
    /// for (row, value) in columnA {
    ///     print("A\(row): \(value ?? .empty)")
    /// }
    ///
    /// // Find the maximum value in a column
    /// let numbers = sheet.column("B").compactMap { row, value -> Double? in
    ///     if let value = value, case .number(let num) = value {
    ///         return num
    ///     }
    ///     return nil
    /// }
    ///
    /// if let max = numbers.max() {
    ///     print("Maximum value in column B: \(max)")
    /// }
    ///
    /// // Get column headers (assuming row 1 contains headers)
    /// let headers = ["A", "B", "C", "D"].map { letter -> String in
    ///     let column = sheet.column(letter)
    ///     if let first = column.first(where: { $0.row == 1 }),
    ///        let value = first.value,
    ///        case .text(let text) = value {
    ///         return text
    ///     }
    ///     return ""
    /// }
    /// print("Headers: \(headers)")
    ///
    /// // Count non-empty cells in a column
    /// let nonEmptyCells = sheet.column("C")
    ///     .filter { $0.value != nil && $0.value != .empty }
    ///     .count
    /// print("Column C has \(nonEmptyCells) non-empty cells")
    /// ```
    ///
    /// ## Performance
    ///
    /// This method scans all rows to find cells in the specified column, then sorts the results.
    /// The time complexity is O(n log n) where n is the number of cells in the column.
    ///
    /// - Note: The column letter is case-insensitive.
    /// - Note: Empty cells in the middle of a column are not included in the results.
    ///
    /// - SeeAlso: ``column(at:)``
    /// - SeeAlso: ``range(_:)``
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

    /// Returns all cells in a column by zero-based index.
    ///
    /// Convenience method that accepts a zero-based column index (0 = A, 1 = B, 2 = C, etc.)
    /// instead of a column letter. This is useful when programmatically iterating through columns.
    ///
    /// - Parameter index: The zero-based column index (0 = A, 1 = B, 2 = C, etc.).
    /// - Returns: An array of tuples containing row numbers and cell values, sorted by row number.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Iterate through the first 5 columns
    /// for colIndex in 0..<5 {
    ///     let column = sheet.column(at: colIndex)
    ///     let columnLetter = ["A", "B", "C", "D", "E"][colIndex]
    ///     
    ///     print("Column \(columnLetter):")
    ///     for (row, value) in column {
    ///         print("  Row \(row): \(value ?? .empty)")
    ///     }
    /// }
    ///
    /// // Calculate column statistics
    /// for colIndex in 0..<10 {
    ///     let column = sheet.column(at: colIndex)
    ///     let numbers = column.compactMap { row, value -> Double? in
    ///         if let value = value, case .number(let num) = value {
    ///             return num
    ///         }
    ///         return nil
    ///     }
    ///     
    ///     if !numbers.isEmpty {
    ///         let sum = numbers.reduce(0, +)
    ///         let avg = sum / Double(numbers.count)
    ///         print("Column \(colIndex): sum=\(sum), avg=\(avg)")
    ///     }
    /// }
    ///
    /// // Compare two columns
    /// let col0 = sheet.column(at: 0) // Column A
    /// let col1 = sheet.column(at: 1) // Column B
    ///
    /// for row in 1...sheet.rowCount {
    ///     let val0 = col0.first { $0.row == row }?.value
    ///     let val1 = col1.first { $0.row == row }?.value
    ///     
    ///     if val0 != val1 {
    ///         print("Row \(row): A=\(val0 ?? .empty), B=\(val1 ?? .empty)")
    ///     }
    /// }
    /// ```
    ///
    /// - Note: Column indices are zero-based: 0 = A, 1 = B, 2 = C, 25 = Z, 26 = AA, etc.
    ///
    /// - SeeAlso: ``column(_:)``
    /// - SeeAlso: ``range(_:)``
    public func column(at index: Int) -> [(row: Int, value: CellValue?)] {
        let letter = columnLetterFrom(index: index)
        return column(letter)
    }

    /// Returns all rows that match a predicate.
    ///
    /// Use this method to filter rows based on custom criteria. The predicate receives an array
    /// of cells (as reference-value pairs) for each row and returns `true` to include the row
    /// in the results.
    ///
    /// - Parameter predicate: A closure that takes an array of cells in a row and returns `true`
    ///   if the row should be included in the results.
    /// - Returns: An array of matching rows, where each row is represented as an array of
    ///   cell reference-value pairs.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Find all rows where column A contains "Error"
    /// let errorRows = sheet.rows(where: { cells in
    ///     cells.contains { ref, value in
    ///         ref.column == "A" &&
    ///         (if case .text(let str) = value { str.contains("Error") } else { false })
    ///     }
    /// })
    ///
    /// print("Found \(errorRows.count) rows with errors")
    ///
    /// // Find rows where the sum of columns B and C exceeds 100
    /// let highValueRows = sheet.rows(where: { cells in
    ///     let bValue = cells.first { $0.reference.column == "B" }?.value
    ///     let cValue = cells.first { $0.reference.column == "C" }?.value
    ///     
    ///     guard case .number(let b) = bValue,
    ///           case .number(let c) = cValue else {
    ///         return false
    ///     }
    ///     
    ///     return b + c > 100
    /// })
    ///
    /// // Find non-empty rows (rows with at least one non-empty cell)
    /// let nonEmptyRows = sheet.rows(where: { cells in
    ///     cells.contains { _, value in value != .empty }
    /// })
    ///
    /// // Find rows with specific pattern across columns
    /// let patternRows = sheet.rows(where: { cells in
    ///     guard let aValue = cells.first(where: { $0.reference.column == "A" })?.value,
    ///           let bValue = cells.first(where: { $0.reference.column == "B" })?.value else {
    ///         return false
    ///     }
    ///     
    ///     // Looking for rows where A is text and B is a number
    ///     if case .text = aValue, case .number = bValue {
    ///         return true
    ///     }
    ///     return false
    /// })
    ///
    /// // Extract data from matching rows
    /// for row in errorRows {
    ///     let rowNumber = row.first?.reference.row ?? 0
    ///     print("Row \(rowNumber):")
    ///     
    ///     for (ref, value) in row {
    ///         print("  \(ref): \(value)")
    ///     }
    /// }
    /// ```
    ///
    /// ## Performance
    ///
    /// This method iterates through all rows and applies the predicate to each one. For very
    /// large worksheets, consider using ``rows()`` with manual filtering for better control
    /// over memory usage.
    ///
    /// - SeeAlso: ``rows()``
    /// - SeeAlso: ``find(where:)``
    /// - SeeAlso: ``findAll(where:)``
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

    /// Finds the first cell that matches a predicate.
    ///
    /// Use this method to search for a specific cell in the worksheet. The search proceeds
    /// row by row, cell by cell, and stops at the first match.
    ///
    /// - Parameter predicate: A closure that takes a cell reference and value, and returns `true`
    ///   if this is the cell you're looking for.
    /// - Returns: A tuple containing the cell reference and value of the first matching cell,
    ///   or `nil` if no cell matches.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Find the first cell containing a specific text
    /// if let result = sheet.find(where: { ref, value in
    ///     if case .text(let str) = value {
    ///         return str == "Total"
    ///     }
    ///     return false
    /// }) {
    ///     print("Found 'Total' at \(result.reference)")
    ///     
    ///     // Get the cell to the right (assuming it contains the total value)
    ///     let nextCol = CellReference(
    ///         column: String(UnicodeScalar(result.reference.column.first!.unicodeScalars.first!.value + 1)!),
    ///         row: result.reference.row
    ///     )
    ///     if let totalValue = sheet.cell(at: nextCol) {
    ///         print("Total value: \(totalValue)")
    ///     }
    /// }
    ///
    /// // Find the first cell with a value greater than 1000
    /// if let result = sheet.find(where: { ref, value in
    ///     if case .number(let num) = value {
    ///         return num > 1000
    ///     }
    ///     return false
    /// }) {
    ///     print("First value > 1000 is at \(result.reference): \(result.value)")
    /// }
    ///
    /// // Find the first error cell
    /// if let result = sheet.find(where: { ref, value in
    ///     if case .error = value {
    ///         return true
    ///     }
    ///     return false
    /// }) {
    ///     print("Error at \(result.reference): \(result.value)")
    /// }
    ///
    /// // Find a cell by reference pattern
    /// if let result = sheet.find(where: { ref, value in
    ///     // Looking for any cell in column A after row 10
    ///     ref.column == "A" && ref.row > 10
    /// }) {
    ///     print("First match: \(result.reference) = \(result.value)")
    /// }
    /// ```
    ///
    /// ## Performance
    ///
    /// This method stops at the first match, making it more efficient than ``findAll(where:)``
    /// when you only need one result. Use this instead of `findAll` followed by `.first`.
    ///
    /// - SeeAlso: ``findAll(where:)``
    /// - SeeAlso: ``rows(where:)``
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

    /// Finds all cells that match a predicate.
    ///
    /// Use this method to search for all cells in the worksheet that meet specific criteria.
    /// The search proceeds row by row, cell by cell, and collects all matches.
    ///
    /// - Parameter predicate: A closure that takes a cell reference and value, and returns `true`
    ///   for cells to include in the results.
    /// - Returns: An array of tuples containing the cell reference and value for each matching cell.
    ///
    /// ## Example
    ///
    /// ```swift
    /// // Find all cells containing numeric values
    /// let numberCells = sheet.findAll(where: { ref, value in
    ///     if case .number = value {
    ///         return true
    ///     }
    ///     return false
    /// })
    /// print("Found \(numberCells.count) cells with numbers")
    ///
    /// // Find all error cells
    /// let errors = sheet.findAll(where: { ref, value in
    ///     if case .error(let errorMsg) = value {
    ///         print("Error at \(ref): \(errorMsg)")
    ///         return true
    ///     }
    ///     return false
    /// })
    ///
    /// // Find all cells in a specific column with values > 100
    /// let highValues = sheet.findAll(where: { ref, value in
    ///     guard ref.column == "C" else { return false }
    ///     
    ///     if case .number(let num) = value {
    ///         return num > 100
    ///     }
    ///     return false
    /// })
    ///
    /// // Find all cells with specific text pattern
    /// let matches = sheet.findAll(where: { ref, value in
    ///     if case .text(let str) = value {
    ///         return str.lowercased().contains("important")
    ///     }
    ///     return false
    /// })
    ///
    /// for (ref, value) in matches {
    ///     print("\(ref): \(value)")
    /// }
    ///
    /// // Find cells with hyperlinks
    /// let cellsWithLinks = sheet.findAll(where: { ref, value in
    ///     !sheet.hyperlinks(at: ref).isEmpty
    /// })
    /// print("Found \(cellsWithLinks.count) cells with hyperlinks")
    ///
    /// // Statistical analysis
    /// let allNumbers = sheet.findAll(where: { ref, value in
    ///     if case .number = value { return true }
    ///     return false
    /// }).compactMap { ref, value -> Double? in
    ///     if case .number(let num) = value { return num }
    ///     return nil
    /// }
    ///
    /// if !allNumbers.isEmpty {
    ///     let sum = allNumbers.reduce(0, +)
    ///     let avg = sum / Double(allNumbers.count)
    ///     let min = allNumbers.min() ?? 0
    ///     let max = allNumbers.max() ?? 0
    ///     
    ///     print("Statistics:")
    ///     print("  Count: \(allNumbers.count)")
    ///     print("  Sum: \(sum)")
    ///     print("  Average: \(avg)")
    ///     print("  Min: \(min)")
    ///     print("  Max: \(max)")
    /// }
    /// ```
    ///
    /// ## Performance
    ///
    /// This method scans the entire worksheet and collects all matches. For large worksheets,
    /// this may use significant memory. If you only need the first match, use ``find(where:)``
    /// instead.
    ///
    /// - SeeAlso: ``find(where:)``
    /// - SeeAlso: ``rows(where:)``
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

/// A lazy iterator for efficiently traversing worksheet rows.
///
/// `RowIterator` provides memory-efficient iteration through worksheet rows without loading
/// all data into memory at once. This is particularly important for large spreadsheets with
/// thousands of rows.
///
/// Each iteration yields a row as an array of `(reference, value)` tuples, allowing you to
/// access both the cell location and its value.
///
/// ## Usage
///
/// You typically don't create `RowIterator` instances directly. Instead, use the ``Sheet/rows()``
/// method:
///
/// ```swift
/// for row in sheet.rows() {
///     for (reference, value) in row {
///         print("\(reference): \(value)")
///     }
/// }
/// ```
///
/// ## Performance
///
/// The iterator processes one row at a time, keeping memory usage constant regardless of
/// worksheet size. This is much more efficient than loading all rows into an array at once.
///
/// - SeeAlso: ``Sheet/rows()``
/// - SeeAlso: ``Sheet/row(_:)``
public struct RowIterator: Sequence {
    private let rawData: WorksheetData
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    
    init(rawData: WorksheetData, sharedStrings: SharedStrings, styles: StylesInfo) {
        self.rawData = rawData
        self.sharedStrings = sharedStrings
        self.styles = styles
    }
    
    /// Creates a new iterator for traversing the rows.
    ///
    /// - Returns: An iterator that yields rows one at a time.
    public func makeIterator() -> Iterator {
        Iterator(rows: rawData.rows, sharedStrings: sharedStrings, styles: styles)
    }
    
    /// The iterator implementation for row iteration.
    ///
    /// This type handles the actual iteration through rows, resolving cell values and
    /// maintaining the current position in the worksheet.
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
        
        /// Advances to the next row and returns it, or returns `nil` if at the end.
        ///
        /// - Returns: An array of cell reference-value pairs for the next row, or `nil`
        ///   if there are no more rows.
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
