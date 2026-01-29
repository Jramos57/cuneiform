import Foundation

/// Protection options that control structural modifications to a workbook.
///
/// Use `WorkbookProtectionOptions` to configure which workbook-level operations users can perform.
/// These options protect the overall workbook structure, preventing users from adding, deleting,
/// or reordering sheets.
///
/// ## Topics
///
/// ### Creating Protection Options
/// - ``init(structure:windows:)``
///
/// ### Configuration Properties
/// - ``structure``
/// - ``windows``
///
/// ### Predefined Options
/// - ``default``
/// - ``strict``
/// - ``structureOnly``
///
/// ## Example: Protecting Workbook Structure
///
/// ```swift
/// var workbook = WorkbookWriter()
/// 
/// // Prevent users from adding or deleting sheets
/// workbook.protectWorkbook(password: "secret", options: .structureOnly)
/// 
/// // Or use custom options
/// let options = WorkbookProtectionOptions(structure: true, windows: false)
/// workbook.protectWorkbook(password: "secret", options: options)
/// ```
public struct WorkbookProtectionOptions: Sendable {
    /// Whether to protect the workbook structure (prevents adding/deleting/reordering sheets).
    public var structure: Bool = false
    
    /// Whether to protect window positioning and sizing.
    public var windows: Bool = false

    /// Creates protection options for a workbook.
    ///
    /// - Parameters:
    ///   - structure: If `true`, prevents users from adding, deleting, or reordering sheets. Defaults to `false`.
    ///   - windows: If `true`, prevents users from moving or resizing workbook windows. Defaults to `false`.
    public init(structure: Bool = false, windows: Bool = false) {
        self.structure = structure
        self.windows = windows
    }

    /// No protection applied to the workbook.
    ///
    /// Use this when you want no structural restrictions on the workbook.
    public static let `default` = WorkbookProtectionOptions()

    /// Full protection for both structure and windows.
    ///
    /// Use this to prevent all structural modifications and window manipulation.
    public static let strict = WorkbookProtectionOptions(structure: true, windows: true)

    /// Protection for workbook structure only.
    ///
    /// Use this to prevent sheet insertion, deletion, and reordering while allowing window manipulation.
    public static let structureOnly = WorkbookProtectionOptions(structure: true, windows: false)

    fileprivate func toProtection(passwordHash: String?) -> WorkbookProtection {
        WorkbookProtection(
            structureProtected: structure,
            windowsProtected: windows,
            passwordHash: passwordHash
        )
    }
}

/// Protection options that control modifications to worksheet content and structure.
///
/// Use `SheetProtectionOptions` to configure granular permissions for worksheet operations.
/// Each property represents a specific capability that can be allowed or denied when
/// sheet protection is enabled.
///
/// ## Topics
///
/// ### Creating Protection Options
/// - ``init(formatCells:formatColumns:formatRows:insertColumns:insertRows:insertHyperlinks:deleteColumns:deleteRows:selectLockedCells:selectUnlockedCells:sort:autoFilter:pivotTables:)``
///
/// ### Formatting Permissions
/// - ``formatCells``
/// - ``formatColumns``
/// - ``formatRows``
///
/// ### Structural Permissions
/// - ``insertColumns``
/// - ``insertRows``
/// - ``deleteColumns``
/// - ``deleteRows``
///
/// ### Content Permissions
/// - ``insertHyperlinks``
/// - ``selectLockedCells``
/// - ``selectUnlockedCells``
///
/// ### Data Operations
/// - ``sort``
/// - ``autoFilter``
/// - ``pivotTables``
///
/// ### Predefined Options
/// - ``default``
/// - ``strict``
/// - ``readonly``
///
/// ## Example: Protecting a Sheet
///
/// ```swift
/// var workbook = WorkbookWriter()
/// let sheetIndex = workbook.addSheet(named: "Protected Data")
///
/// workbook.modifySheet(at: sheetIndex) { sheet in
///     sheet.writeText("Protected Content", to: "A1")
///     
///     // Allow viewing but prevent all modifications
///     sheet.protectSheet(password: "secret", options: .readonly)
///     
///     // Or use custom options
///     let options = SheetProtectionOptions(
///         formatCells: false,
///         insertRows: false,
///         deleteRows: false,
///         selectLockedCells: true,
///         selectUnlockedCells: true
///     )
///     sheet.protectSheet(password: "secret", options: options)
/// }
/// ```
public struct SheetProtectionOptions: Sendable {
    /// Whether users can format cells when protection is enabled.
    public var formatCells: Bool = true
    
    /// Whether users can format columns when protection is enabled.
    public var formatColumns: Bool = true
    
    /// Whether users can format rows when protection is enabled.
    public var formatRows: Bool = true
    
    /// Whether users can insert columns when protection is enabled.
    public var insertColumns: Bool = true
    
    /// Whether users can insert rows when protection is enabled.
    public var insertRows: Bool = true
    
    /// Whether users can insert hyperlinks when protection is enabled.
    public var insertHyperlinks: Bool = true
    
    /// Whether users can delete columns when protection is enabled.
    public var deleteColumns: Bool = true
    
    /// Whether users can delete rows when protection is enabled.
    public var deleteRows: Bool = true
    
    /// Whether users can select locked cells when protection is enabled.
    public var selectLockedCells: Bool = true
    
    /// Whether users can select unlocked cells when protection is enabled.
    public var selectUnlockedCells: Bool = true
    
    /// Whether users can sort data when protection is enabled.
    public var sort: Bool = true
    
    /// Whether users can use auto-filter when protection is enabled.
    public var autoFilter: Bool = true
    
    /// Whether users can modify pivot tables when protection is enabled.
    public var pivotTables: Bool = true

    /// Creates sheet protection options with specific permissions.
    ///
    /// - Parameters:
    ///   - formatCells: Allow cell formatting. Defaults to `true`.
    ///   - formatColumns: Allow column formatting. Defaults to `true`.
    ///   - formatRows: Allow row formatting. Defaults to `true`.
    ///   - insertColumns: Allow column insertion. Defaults to `true`.
    ///   - insertRows: Allow row insertion. Defaults to `true`.
    ///   - insertHyperlinks: Allow hyperlink insertion. Defaults to `true`.
    ///   - deleteColumns: Allow column deletion. Defaults to `true`.
    ///   - deleteRows: Allow row deletion. Defaults to `true`.
    ///   - selectLockedCells: Allow selecting locked cells. Defaults to `true`.
    ///   - selectUnlockedCells: Allow selecting unlocked cells. Defaults to `true`.
    ///   - sort: Allow data sorting. Defaults to `true`.
    ///   - autoFilter: Allow auto-filter usage. Defaults to `true`.
    ///   - pivotTables: Allow pivot table modifications. Defaults to `true`.
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

    /// All operations allowed (default protection).
    ///
    /// Use this when applying protection but allowing most user operations.
    public static let `default` = SheetProtectionOptions()

    /// All modifications prevented except selection.
    ///
    /// Use this for strictly controlled sheets where users should only view and select content.
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

    /// Read-only mode allowing only locked cell selection.
    ///
    /// Use this for completely read-only sheets where users can only select and view locked content.
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

/// A high-level API for creating Excel (.xlsx) workbooks using a builder pattern.
///
/// `WorkbookWriter` provides a fluent interface for creating Excel workbooks with multiple sheets,
/// styles, formulas, data validation, hyperlinks, and more. The builder pattern allows you to
/// construct complex workbooks step-by-step while maintaining type safety and clarity.
///
/// ## Overview
///
/// Create a workbook by initializing `WorkbookWriter`, adding sheets, writing data to cells,
/// applying styles, and saving the result to a file. The writer handles all the complexity
/// of the Excel file format (SpreadsheetML) internally.
///
/// ## Topics
///
/// ### Creating a Workbook
/// - ``init()``
///
/// ### Managing Sheets
/// - ``addSheet(named:)``
/// - ``sheet(at:)``
/// - ``modifySheet(at:_:)``
/// - ``SheetWriter``
///
/// ### Styling
/// - ``style(_:)``
///
/// ### Named Ranges
/// - ``addNamedRange(name:refersTo:)``
///
/// ### Protection
/// - ``protectWorkbook(password:options:)``
///
/// ### Saving
/// - ``save(to:)``
/// - ``buildData()``
///
/// ## Basic Example
///
/// ```swift
/// // Create a new workbook
/// var workbook = WorkbookWriter()
///
/// // Add a sheet
/// let salesIndex = workbook.addSheet(named: "Sales Data")
///
/// // Write data to cells
/// workbook.modifySheet(at: salesIndex) { sheet in
///     sheet.writeText("Product", to: "A1")
///     sheet.writeText("Revenue", to: "B1")
///     sheet.writeText("Widget", to: "A2")
///     sheet.writeNumber(1500.00, to: "B2")
/// }
///
/// // Save to file
/// try workbook.save(to: URL(fileURLWithPath: "sales.xlsx"))
/// ```
///
/// ## Advanced Example: Multi-Sheet Workbook with Formulas
///
/// ```swift
/// var workbook = WorkbookWriter()
///
/// // Configure custom styles
/// workbook.style { styles in
///     styles.addNumberFormat("$#,##0.00")  // Currency format
///     styles.addFont(name: "Arial", size: 12, bold: true)
///     styles.addFill(pattern: .solid, fgColor: "FFD9D9D9")
///     styles.addAlignment(horizontal: .center, vertical: .center)
///     // Create a header style combining these elements
/// }
///
/// // Add data sheet
/// let dataIndex = workbook.addSheet(named: "Data")
/// workbook.modifySheet(at: dataIndex) { sheet in
///     // Headers
///     sheet.writeText("Month", to: "A1", styleIndex: 1)
///     sheet.writeText("Sales", to: "B1", styleIndex: 1)
///     sheet.writeText("Expenses", to: "C1", styleIndex: 1)
///     
///     // Data rows
///     sheet.writeText("January", to: "A2")
///     sheet.writeNumber(50000, to: "B2", styleIndex: 2)
///     sheet.writeNumber(30000, to: "C2", styleIndex: 2)
///     
///     sheet.writeText("February", to: "A3")
///     sheet.writeNumber(55000, to: "B3", styleIndex: 2)
///     sheet.writeNumber(32000, to: "C3", styleIndex: 2)
/// }
///
/// // Add summary sheet with formulas
/// let summaryIndex = workbook.addSheet(named: "Summary")
/// workbook.modifySheet(at: summaryIndex) { sheet in
///     sheet.writeText("Total Sales", to: "A1")
///     sheet.writeFormula("SUM(Data!B:B)", to: "B1", styleIndex: 2)
///     
///     sheet.writeText("Total Expenses", to: "A2")
///     sheet.writeFormula("SUM(Data!C:C)", to: "B2", styleIndex: 2)
///     
///     sheet.writeText("Net Profit", to: "A3")
///     sheet.writeFormula("B1-B2", to: "B3", styleIndex: 2)
///     
///     // Add hyperlink to data sheet
///     sheet.addHyperlinkInternal(at: "A5", location: "Data!A1", display: "View Data")
/// }
///
/// // Add named range for easy reference
/// workbook.addNamedRange(name: "TotalSales", refersTo: "Summary!$B$1")
///
/// // Protect workbook structure
/// workbook.protectWorkbook(password: "secret", options: .structureOnly)
///
/// // Save
/// try workbook.save(to: URL(fileURLWithPath: "financial_report.xlsx"))
/// ```
///
/// ## Working with Data Validation and Conditional Formatting
///
/// ```swift
/// var workbook = WorkbookWriter()
/// let sheetIndex = workbook.addSheet(named: "Validated Data")
///
/// workbook.modifySheet(at: sheetIndex) { sheet in
///     // Add dropdown list validation
///     let validation = WorksheetBuilder.DataValidation(
///         type: .list,
///         formula1: "\"Apple,Orange,Banana\"",
///         sqref: "A2:A100"
///     )
///     sheet.addDataValidation(validation)
///     
///     // Add conditional formatting
///     let rule = WorksheetData.ConditionalRule(
///         type: .cellIs,
///         operator: "greaterThan",
///         formula: ["100"],
///         dxfId: nil,
///         priority: 1
///     )
///     sheet.addConditionalFormat(range: "B2:B100", rule: rule)
///     
///     // Merge cells for header
///     sheet.mergeCells("A1:C1")
///     sheet.writeText("Data Entry Form", to: "A1")
/// }
///
/// try workbook.save(to: URL(fileURLWithPath: "validated.xlsx"))
/// ```
public struct WorkbookWriter {
    /// A writer for modifying individual worksheet content.
    ///
    /// `SheetWriter` provides methods for writing data to cells, applying formulas,
    /// adding hyperlinks, configuring protection, and managing worksheet features
    /// like merged cells, data validation, and tables.
    ///
    /// ## Topics
    ///
    /// ### Writing Cell Values
    /// - ``write(_:to:)-9kbhz``
    /// - ``write(_:to:)-8h4lm``
    /// - ``write(_:to:styleIndex:)-7m8p3``
    /// - ``write(_:to:styleIndex:)-2xz9k``
    ///
    /// ### Writing Specific Types
    /// - ``writeText(_:to:)-6j2n5``
    /// - ``writeText(_:to:styleIndex:)``
    /// - ``writeNumber(_:to:)-4p7q1``
    /// - ``writeNumber(_:to:styleIndex:)``
    /// - ``writeBoolean(_:to:)-3k8m2``
    /// - ``writeBoolean(_:to:styleIndex:)``
    /// - ``writeFormula(_:cachedValue:to:)-5n9l4``
    /// - ``writeFormula(_:cachedValue:to:styleIndex:)``
    ///
    /// ### Cell Formatting
    /// - ``mergeCells(_:)-8j3k1``
    /// - ``mergeCells(_:)-7m2p4``
    ///
    /// ### Data Validation
    /// - ``addDataValidation(_:)``
    ///
    /// ### Conditional Formatting
    /// - ``addConditionalFormat(range:rule:)``
    /// - ``addConditionalFormat(range:rules:)``
    ///
    /// ### Filtering
    /// - ``setAutoFilter(range:)``
    ///
    /// ### Hyperlinks
    /// - ``addHyperlinkExternal(at:url:display:tooltip:)``
    /// - ``addHyperlinkInternal(at:location:display:tooltip:)``
    ///
    /// ### Comments
    /// - ``addComment(at:text:author:)``
    ///
    /// ### Protection
    /// - ``protectSheet(password:options:)``
    ///
    /// ### Page Setup
    /// - ``setPageSetup(_:)``
    ///
    /// ### Tables
    /// - ``addTable(name:displayName:ref:columns:headerRowCount:totalsRowCount:tableId:)``
    ///
    /// ### Properties
    /// - ``name``
    ///
    /// ## Example: Writing Various Data Types
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// let sheetIndex = workbook.addSheet(named: "Data")
    ///
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     // Write text
    ///     sheet.writeText("Product Name", to: "A1")
    ///     
    ///     // Write numbers
    ///     sheet.writeNumber(99.99, to: "B1")
    ///     
    ///     // Write boolean
    ///     sheet.writeBoolean(true, to: "C1")
    ///     
    ///     // Write formula with cached value
    ///     sheet.writeFormula("=B1*1.08", cachedValue: 107.99, to: "D1")
    ///     
    ///     // Write with custom style
    ///     sheet.writeText("Styled Text", to: "E1", styleIndex: 2)
    /// }
    /// ```
    ///
    /// ## Example: Formulas and Cell References
    ///
    /// ```swift
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     // Basic arithmetic
    ///     sheet.writeFormula("=A1+B1", to: "C1")
    ///     
    ///     // Sum range
    ///     sheet.writeFormula("=SUM(A1:A10)", to: "A11")
    ///     
    ///     // Cross-sheet reference
    ///     sheet.writeFormula("=Sheet2!A1*2", to: "D1")
    ///     
    ///     // Formula with cached value (displays immediately without recalculation)
    ///     sheet.writeFormula("=AVERAGE(A1:A10)", cachedValue: 42.5, to: "A12")
    /// }
    /// ```
    public struct SheetWriter {
        private var builder: WorksheetBuilder
        private var commentsBuilder: CommentsBuilder?
        
        /// The name of this worksheet.
        public let name: String
        
        init(name: String) {
            self.name = name
            self.builder = WorksheetBuilder()
        }

        init(name: String, sharedStringsBuilder: SharedStringsBuilder) {
            self.name = name
            self.builder = WorksheetBuilder(sharedStringsBuilder: sharedStringsBuilder)
        }
        
        /// Writes a value to a cell using a ``CellReference``.
        ///
        /// Use this method to write any supported cell value type to a specific cell location.
        /// The value type determines how the data is stored and displayed in Excel.
        ///
        /// - Parameters:
        ///   - value: The value to write. Can be text, number, boolean, or formula.
        ///   - reference: The cell location (e.g., `CellReference(row: 0, column: 0)` for A1).
        ///
        /// ## Example
        ///
        /// ```swift
        /// let cellRef = CellReference(row: 0, column: 0)  // A1
        /// sheet.write(.text("Hello"), to: cellRef)
        /// sheet.write(.number(42.0), to: cellRef)
        /// sheet.write(.formula("=SUM(A1:A10)"), to: cellRef)
        /// ```
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: CellReference) {
            builder.addCell(at: reference, value: value)
        }
        
        /// Writes a value to a cell using a string reference.
        ///
        /// A convenience method for writing values using Excel-style cell references like "A1", "B5", etc.
        ///
        /// - Parameters:
        ///   - value: The value to write. Can be text, number, boolean, or formula.
        ///   - reference: The cell location in A1 notation (e.g., "A1", "C10", "AA100").
        ///
        /// - Note: Invalid cell references are silently ignored.
        ///
        /// ## Example
        ///
        /// ```swift
        /// sheet.write(.text("Revenue"), to: "A1")
        /// sheet.write(.number(1500.00), to: "B1")
        /// sheet.write(.formula("=B1*0.08", cachedValue: 120.0), to: "C1")
        /// ```
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: String) {
            guard let ref = CellReference(reference) else { return }
            write(value, to: ref)
        }
        
        /// Writes a value to a cell with a specific style using a ``CellReference``.
        ///
        /// Apply formatting to cells by providing a style index from the workbook's style table.
        /// Styles control number formats, fonts, fills, borders, and alignment.
        ///
        /// - Parameters:
        ///   - value: The value to write. Can be text, number, boolean, or formula.
        ///   - reference: The cell location.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// ## Example
        ///
        /// ```swift
        /// // First, configure styles in the workbook
        /// workbook.style { styles in
        ///     // styleIndex 1 will be the header style
        ///     styles.addFont(name: "Arial", size: 12, bold: true)
        /// }
        ///
        /// // Then use the style when writing
        /// let cellRef = CellReference(row: 0, column: 0)
        /// sheet.write(.text("Header"), to: cellRef, styleIndex: 1)
        /// ```
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: CellReference, styleIndex: Int) {
            builder.addCell(at: reference, value: value, styleIndex: styleIndex)
        }
        
        /// Writes a value with a specific style using a string reference.
        ///
        /// A convenience method combining string-based cell references with style formatting.
        ///
        /// - Parameters:
        ///   - value: The value to write. Can be text, number, boolean, or formula.
        ///   - reference: The cell location in A1 notation.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// - Note: Invalid cell references are silently ignored.
        ///
        /// ## Example
        ///
        /// ```swift
        /// sheet.write(.text("Total"), to: "A10", styleIndex: 1)
        /// sheet.write(.number(5000), to: "B10", styleIndex: 2)
        /// ```
        public mutating func write(_ value: WorksheetBuilder.WritableCellValue, to reference: String, styleIndex: Int) {
            guard let ref = CellReference(reference) else { return }
            write(value, to: ref, styleIndex: styleIndex)
        }
        
        /// Writes a text string to a cell.
        ///
        /// A convenience method for writing text values without wrapping in `.text()`.
        ///
        /// - Parameters:
        ///   - text: The string to write to the cell.
        ///   - reference: The cell location.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let cellRef = CellReference(row: 0, column: 0)
        /// sheet.writeText("Product Name", to: cellRef)
        /// sheet.writeText("Widget", to: CellReference(row: 1, column: 0))
        /// ```
        public mutating func writeText(_ text: String, to reference: CellReference) {
            write(.text(text), to: reference)
        }
        
        /// Writes a text string to a cell with a specific style.
        ///
        /// - Parameters:
        ///   - text: The string to write to the cell.
        ///   - reference: The cell location.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// ## Example
        ///
        /// ```swift
        /// let headerRef = CellReference(row: 0, column: 0)
        /// sheet.writeText("Product Name", to: headerRef, styleIndex: 1)
        /// ```
        public mutating func writeText(_ text: String, to reference: CellReference, styleIndex: Int) {
            write(.text(text), to: reference, styleIndex: styleIndex)
        }
        
        /// Writes a numeric value to a cell.
        ///
        /// A convenience method for writing numbers without wrapping in `.number()`.
        /// Use style indexes to apply number formatting like currency, percentage, or decimal places.
        ///
        /// - Parameters:
        ///   - number: The numeric value to write.
        ///   - reference: The cell location.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let priceRef = CellReference(row: 1, column: 1)
        /// sheet.writeNumber(99.99, to: priceRef)
        /// sheet.writeNumber(1500.00, to: CellReference(row: 2, column: 1))
        /// ```
        public mutating func writeNumber(_ number: Double, to reference: CellReference) {
            write(.number(number), to: reference)
        }
        
        /// Writes a numeric value to a cell with a specific style.
        ///
        /// Apply number formatting by providing a style index that includes the desired number format.
        ///
        /// - Parameters:
        ///   - number: The numeric value to write.
        ///   - reference: The cell location.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// ## Example
        ///
        /// ```swift
        /// // Assuming styleIndex 2 has currency format "$#,##0.00"
        /// let priceRef = CellReference(row: 1, column: 1)
        /// sheet.writeNumber(1299.99, to: priceRef, styleIndex: 2)
        /// ```
        public mutating func writeNumber(_ number: Double, to reference: CellReference, styleIndex: Int) {
            write(.number(number), to: reference, styleIndex: styleIndex)
        }
        
        /// Writes a boolean value to a cell.
        ///
        /// A convenience method for writing boolean values without wrapping in `.boolean()`.
        /// Excel displays boolean values as TRUE or FALSE.
        ///
        /// - Parameters:
        ///   - bool: The boolean value to write.
        ///   - reference: The cell location.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let statusRef = CellReference(row: 1, column: 2)
        /// sheet.writeBoolean(true, to: statusRef)
        /// sheet.writeBoolean(false, to: CellReference(row: 2, column: 2))
        /// ```
        public mutating func writeBoolean(_ bool: Bool, to reference: CellReference) {
            write(.boolean(bool), to: reference)
        }
        
        /// Writes a boolean value to a cell with a specific style.
        ///
        /// - Parameters:
        ///   - bool: The boolean value to write.
        ///   - reference: The cell location.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// ## Example
        ///
        /// ```swift
        /// let statusRef = CellReference(row: 1, column: 2)
        /// sheet.writeBoolean(true, to: statusRef, styleIndex: 3)
        /// ```
        public mutating func writeBoolean(_ bool: Bool, to reference: CellReference, styleIndex: Int) {
            write(.boolean(bool), to: reference, styleIndex: styleIndex)
        }
        
        /// Writes a formula to a cell with an optional cached result value.
        ///
        /// Formulas begin with `=` and use Excel's formula syntax. Providing a cached value
        /// allows Excel to display a result immediately without recalculating, which is useful
        /// for complex or slow calculations.
        ///
        /// - Parameters:
        ///   - formula: The Excel formula string (e.g., "=SUM(A1:A10)", "=B2*1.08").
        ///   - cachedValue: An optional pre-calculated result to display. Defaults to `nil`.
        ///   - reference: The cell location.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let totalRef = CellReference(row: 10, column: 0)
        /// 
        /// // Simple formula
        /// sheet.writeFormula("=SUM(A1:A10)", to: totalRef)
        /// 
        /// // Formula with cached value
        /// sheet.writeFormula("=AVERAGE(A1:A10)", cachedValue: 42.5, to: totalRef)
        /// 
        /// // Cross-sheet reference
        /// sheet.writeFormula("=Sheet2!B5*2", to: totalRef)
        /// 
        /// // Complex nested formula
        /// sheet.writeFormula("=IF(A1>100, A1*0.9, A1)", to: totalRef)
        /// ```
        public mutating func writeFormula(_ formula: String, cachedValue: Double? = nil, to reference: CellReference) {
            write(.formula(formula, cachedValue: cachedValue), to: reference)
        }
        
        /// Writes a formula to a cell with a style and optional cached result value.
        ///
        /// Combine formula writing with style formatting for professional-looking calculated cells.
        ///
        /// - Parameters:
        ///   - formula: The Excel formula string.
        ///   - cachedValue: An optional pre-calculated result to display. Defaults to `nil`.
        ///   - reference: The cell location.
        ///   - styleIndex: The index of the style in the workbook's style table (0-based).
        ///
        /// ## Example
        ///
        /// ```swift
        /// // Assuming styleIndex 2 has currency format
        /// let totalRef = CellReference(row: 10, column: 1)
        /// sheet.writeFormula("=SUM(B1:B10)", cachedValue: 15000.00, to: totalRef, styleIndex: 2)
        /// ```
        public mutating func writeFormula(_ formula: String, cachedValue: Double? = nil, to reference: CellReference, styleIndex: Int) {
            write(.formula(formula, cachedValue: cachedValue), to: reference, styleIndex: styleIndex)
        }

        /// Merges a range of cells into a single display cell.
        ///
        /// Cell merging combines multiple cells into one larger cell, useful for headers and labels.
        /// Only the content of the top-left cell is preserved; other cells' content is discarded.
        ///
        /// - Parameter range: The cell range to merge in A1 notation (e.g., "A1:C1", "B2:D5").
        ///
        /// ## Example
        ///
        /// ```swift
        /// // Merge cells for a header spanning three columns
        /// sheet.mergeCells("A1:C1")
        /// sheet.writeText("Report Title", to: "A1")
        /// 
        /// // Merge a rectangular block
        /// sheet.mergeCells("E5:G8")
        /// ```
        public mutating func mergeCells(_ range: String) {
            builder.addMergeCell(range)
        }

        /// Merges multiple ranges of cells.
        ///
        /// A convenience method for merging multiple cell ranges at once.
        ///
        /// - Parameter ranges: An array of cell ranges to merge in A1 notation.
        ///
        /// ## Example
        ///
        /// ```swift
        /// sheet.mergeCells([
        ///     "A1:C1",  // Header row
        ///     "A2:A3",  // Left column label
        ///     "D5:E5"   // Another merged area
        /// ])
        /// ```
        public mutating func mergeCells(_ ranges: [String]) {
            for r in ranges { builder.addMergeCell(r) }
        }

        /// Adds a data validation rule to the worksheet.
        ///
        /// Data validation restricts the type or range of data users can enter in cells,
        /// such as dropdown lists, number ranges, or custom formulas.
        ///
        /// - Parameter dv: The data validation configuration to apply.
        ///
        /// ## Example: Dropdown List
        ///
        /// ```swift
        /// let validation = WorksheetBuilder.DataValidation(
        ///     type: .list,
        ///     formula1: "\"Red,Green,Blue\"",
        ///     sqref: "A2:A100"
        /// )
        /// sheet.addDataValidation(validation)
        /// ```
        ///
        /// ## Example: Number Range
        ///
        /// ```swift
        /// let validation = WorksheetBuilder.DataValidation(
        ///     type: .decimal,
        ///     operator: "between",
        ///     formula1: "0",
        ///     formula2: "100",
        ///     sqref: "B2:B100"
        /// )
        /// sheet.addDataValidation(validation)
        /// ```
        public mutating func addDataValidation(_ dv: WorksheetBuilder.DataValidation) {
            builder.addDataValidation(dv)
        }

        /// Adds conditional formatting with a single rule to a range.
        ///
        /// Conditional formatting automatically applies visual formatting based on cell values,
        /// such as highlighting cells that meet certain criteria.
        ///
        /// - Parameters:
        ///   - range: The cell range to apply formatting to (e.g., "A1:D10").
        ///   - rule: The conditional formatting rule to apply.
        ///
        /// ## Example: Highlight Values Greater Than 100
        ///
        /// ```swift
        /// let rule = WorksheetData.ConditionalRule(
        ///     type: .cellIs,
        ///     operator: "greaterThan",
        ///     formula: ["100"],
        ///     dxfId: nil,
        ///     priority: 1
        /// )
        /// sheet.addConditionalFormat(range: "B2:B100", rule: rule)
        /// ```
        public mutating func addConditionalFormat(range: String, rule: WorksheetData.ConditionalRule) {
            let cf = WorksheetData.ConditionalFormat(range: range, rules: [rule])
            builder.addConditionalFormat(cf)
        }

        /// Adds conditional formatting with multiple rules to a range.
        ///
        /// Apply multiple conditional formatting rules to the same range for complex highlighting scenarios.
        ///
        /// - Parameters:
        ///   - range: The cell range to apply formatting to (e.g., "A1:D10").
        ///   - rules: An array of conditional formatting rules to apply.
        ///
        /// ## Example: Multiple Conditions
        ///
        /// ```swift
        /// let rules = [
        ///     WorksheetData.ConditionalRule(
        ///         type: .cellIs,
        ///         operator: "greaterThan",
        ///         formula: ["100"],
        ///         dxfId: nil,
        ///         priority: 1
        ///     ),
        ///     WorksheetData.ConditionalRule(
        ///         type: .cellIs,
        ///         operator: "lessThan",
        ///         formula: ["0"],
        ///         dxfId: nil,
        ///         priority: 2
        ///     )
        /// ]
        /// sheet.addConditionalFormat(range: "B2:B100", rules: rules)
        /// ```
        public mutating func addConditionalFormat(range: String, rules: [WorksheetData.ConditionalRule]) {
            let cf = WorksheetData.ConditionalFormat(range: range, rules: rules)
            builder.addConditionalFormat(cf)
        }

        /// Sets an auto-filter range for column-based filtering.
        ///
        /// Auto-filters add dropdown buttons to column headers, allowing users to filter
        /// and sort data interactively.
        ///
        /// - Parameter range: The cell range to enable filtering on (e.g., "A1:D100").
        ///   The first row typically contains headers.
        ///
        /// ## Example
        ///
        /// ```swift
        /// // Add headers
        /// sheet.writeText("Name", to: "A1")
        /// sheet.writeText("Department", to: "B1")
        /// sheet.writeText("Salary", to: "C1")
        /// 
        /// // Add data rows...
        /// 
        /// // Enable filtering on all columns
        /// sheet.setAutoFilter(range: "A1:C100")
        /// ```
        public mutating func setAutoFilter(range: String) {
            builder.setAutoFilter(range: range)
        }

        /// Adds an external hyperlink to a cell.
        ///
        /// External hyperlinks navigate to URLs outside the workbook, such as websites,
        /// email addresses, or file paths.
        ///
        /// - Parameters:
        ///   - reference: The cell location to add the hyperlink to.
        ///   - url: The target URL (e.g., "https://example.com", "mailto:user@example.com").
        ///   - display: Optional display text for the hyperlink. If `nil`, uses the cell's value.
        ///   - tooltip: Optional tooltip text shown when hovering over the hyperlink.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let linkRef = CellReference(row: 5, column: 0)
        /// 
        /// // Simple web link
        /// sheet.addHyperlinkExternal(
        ///     at: linkRef,
        ///     url: "https://www.apple.com",
        ///     display: "Visit Apple"
        /// )
        /// 
        /// // Email link with tooltip
        /// sheet.addHyperlinkExternal(
        ///     at: CellReference(row: 6, column: 0),
        ///     url: "mailto:support@example.com",
        ///     display: "Contact Support",
        ///     tooltip: "Send us an email"
        /// )
        /// ```
        public mutating func addHyperlinkExternal(at reference: CellReference, url: String, display: String? = nil, tooltip: String? = nil) {
            builder.addHyperlinkExternal(at: reference, url: url, display: display, tooltip: tooltip)
        }

        /// Adds an internal hyperlink to a cell that navigates within the workbook.
        ///
        /// Internal hyperlinks navigate to other locations within the same workbook,
        /// such as specific cells or named ranges on other sheets.
        ///
        /// - Parameters:
        ///   - reference: The cell location to add the hyperlink to.
        ///   - location: The target location within the workbook (e.g., "Sheet2!A1", "Summary!$B$5").
        ///   - display: Optional display text for the hyperlink. If `nil`, uses the cell's value.
        ///   - tooltip: Optional tooltip text shown when hovering over the hyperlink.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let navRef = CellReference(row: 0, column: 0)
        /// 
        /// // Link to another sheet
        /// sheet.addHyperlinkInternal(
        ///     at: navRef,
        ///     location: "Summary!A1",
        ///     display: "Go to Summary"
        /// )
        /// 
        /// // Link to a specific cell with tooltip
        /// sheet.addHyperlinkInternal(
        ///     at: CellReference(row: 1, column: 0),
        ///     location: "Data!$C$10",
        ///     display: "View Details",
        ///     tooltip: "Jump to detailed analysis"
        /// )
        /// ```
        public mutating func addHyperlinkInternal(at reference: CellReference, location: String, display: String? = nil, tooltip: String? = nil) {
            builder.addHyperlinkInternal(at: reference, location: location, display: display, tooltip: tooltip)
        }

        /// Adds a comment (note) to a cell.
        ///
        /// Comments provide additional context or notes for cells without affecting the cell's value.
        /// They appear as small indicators in Excel that expand when hovered or clicked.
        ///
        /// - Parameters:
        ///   - reference: The cell location to add the comment to.
        ///   - text: The comment text content.
        ///   - author: Optional author name for the comment. Defaults to `nil`.
        ///
        /// ## Example
        ///
        /// ```swift
        /// let dataRef = CellReference(row: 5, column: 2)
        /// 
        /// // Simple comment
        /// sheet.addComment(
        ///     at: dataRef,
        ///     text: "This value is estimated"
        /// )
        /// 
        /// // Comment with author
        /// sheet.addComment(
        ///     at: CellReference(row: 10, column: 0),
        ///     text: "Review this calculation",
        ///     author: "John Doe"
        /// )
        /// ```
        public mutating func addComment(at reference: CellReference, text: String, author: String? = nil) {
            var cb = commentsBuilder ?? CommentsBuilder()
            cb.addComment(at: reference, text: text, author: author)
            commentsBuilder = cb
        }

        /// Protects the worksheet with optional password and specific protection options.
        ///
        /// Sheet protection prevents users from modifying cells or performing certain operations
        /// unless they provide the correct password. Use ``SheetProtectionOptions`` to configure
        /// which operations are allowed.
        ///
        /// - Parameters:
        ///   - password: Optional password required to unprotect the sheet. Defaults to `nil`.
        ///   - options: Protection settings controlling allowed operations. Defaults to `.default`.
        ///
        /// ## Example: Read-Only Sheet
        ///
        /// ```swift
        /// sheet.protectSheet(password: "secret", options: .readonly)
        /// ```
        ///
        /// ## Example: Custom Protection
        ///
        /// ```swift
        /// let options = SheetProtectionOptions(
        ///     formatCells: false,
        ///     insertRows: false,
        ///     deleteRows: false,
        ///     selectLockedCells: true,
        ///     selectUnlockedCells: true
        /// )
        /// sheet.protectSheet(password: "mypassword", options: options)
        /// ```
        public mutating func protectSheet(password: String? = nil, options: SheetProtectionOptions = .default) {
            let prot = options.toProtection(passwordHash: password)
            builder.setProtection(prot)
        }
        
        /// Sets page setup configuration for printing.
        ///
        /// Configure page orientation, paper size, margins, and other print settings
        /// to control how the worksheet appears when printed or exported to PDF.
        ///
        /// - Parameter setup: The page setup configuration to apply.
        ///
        /// ## Example
        ///
        /// ```swift
        /// var pageSetup = PageSetup()
        /// pageSetup.orientation = .landscape
        /// pageSetup.paperSize = .letter
        /// pageSetup.fitToPage = true
        /// sheet.setPageSetup(pageSetup)
        /// ```
        public mutating func setPageSetup(_ setup: PageSetup) {
            builder.setPageSetup(setup)
        }
        
        /// Adds an Excel table (structured data range) to the worksheet.
        ///
        /// Excel tables provide structured references, automatic formatting, and easy data management.
        /// Tables support column headers, totals rows, and automatic formula generation.
        ///
        /// - Parameters:
        ///   - name: The internal table name used in formulas.
        ///   - displayName: The name shown to users. Defaults to `name` if not provided.
        ///   - ref: The cell range for the table (e.g., "A1:D10").
        ///   - columns: Array of column definitions with IDs, names, and optional totals functions.
        ///   - headerRowCount: Number of header rows. Defaults to `1`.
        ///   - totalsRowCount: Number of totals rows. Defaults to `0`.
        ///   - tableId: Unique table identifier. Defaults to `1`.
        ///
        /// ## Example: Simple Table
        ///
        /// ```swift
        /// sheet.addTable(
        ///     name: "SalesTable",
        ///     ref: "A1:C10",
        ///     columns: [
        ///         (id: 1, name: "Product", totalsFunction: nil),
        ///         (id: 2, name: "Quantity", totalsFunction: nil),
        ///         (id: 3, name: "Revenue", totalsFunction: "sum")
        ///     ],
        ///     headerRowCount: 1,
        ///     totalsRowCount: 1
        /// )
        /// ```
        ///
        /// ## Example: Table with Totals
        ///
        /// ```swift
        /// sheet.addTable(
        ///     name: "InventoryTable",
        ///     displayName: "Current Inventory",
        ///     ref: "A1:E100",
        ///     columns: [
        ///         (id: 1, name: "SKU", totalsFunction: nil),
        ///         (id: 2, name: "Description", totalsFunction: nil),
        ///         (id: 3, name: "Quantity", totalsFunction: "sum"),
        ///         (id: 4, name: "Unit Price", totalsFunction: "average"),
        ///         (id: 5, name: "Total Value", totalsFunction: "sum")
        ///     ],
        ///     totalsRowCount: 1,
        ///     tableId: 1
        /// )
        /// ```
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
    private var sharedStringsBuilder: SharedStringsBuilder
    private var namedRanges: [(name: String, refersTo: String)] = []
    private var workbookProtection: WorkbookProtection?
    
    /// Creates a new workbook writer.
    ///
    /// Initialize a workbook writer to begin creating an Excel workbook. After initialization,
    /// add sheets, write data, apply styles, and save to a file.
    ///
    /// ## Example
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// let sheetIndex = workbook.addSheet(named: "Sheet1")
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     sheet.writeText("Hello, Excel!", to: "A1")
    /// }
    /// try workbook.save(to: URL(fileURLWithPath: "output.xlsx"))
    /// ```
    public init() {
        self.stylesBuilder = StylesBuilder()
        self.sharedStringsBuilder = SharedStringsBuilder()
    }
    
    /// Adds a new worksheet to the workbook.
    ///
    /// Creates a new sheet with the specified name and returns its index.
    /// Use the index with ``modifySheet(at:_:)`` to write data to the sheet.
    ///
    /// - Parameter name: The name of the new worksheet.
    /// - Returns: The zero-based index of the newly created sheet.
    ///
    /// ## Example
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// 
    /// let salesIndex = workbook.addSheet(named: "Sales")
    /// let expensesIndex = workbook.addSheet(named: "Expenses")
    /// let summaryIndex = workbook.addSheet(named: "Summary")
    /// 
    /// // Now modify each sheet using its index
    /// workbook.modifySheet(at: salesIndex) { sheet in
    ///     sheet.writeText("Sales Data", to: "A1")
    /// }
    /// ```
    public mutating func addSheet(named name: String) -> Int {
        sheets.append(SheetWriter(name: name, sharedStringsBuilder: sharedStringsBuilder))
        return sheets.count - 1
    }
    
    /// Retrieves a sheet writer by index.
    ///
    /// Returns an immutable copy of the sheet at the specified index.
    /// For modifying sheets, use ``modifySheet(at:_:)`` instead.
    ///
    /// - Parameter index: The zero-based index of the sheet.
    /// - Returns: The sheet writer at the index, or `nil` if the index is out of bounds.
    ///
    /// ## Example
    ///
    /// ```swift
    /// if let sheet = workbook.sheet(at: 0) {
    ///     print("Sheet name: \(sheet.name)")
    /// }
    /// ```
    public mutating func sheet(at index: Int) -> SheetWriter? {
        guard index >= 0, index < sheets.count else { return nil }
        return sheets[index]
    }
    
    /// Modifies a sheet using a closure.
    ///
    /// Use this method to write data, apply formatting, add formulas, and configure
    /// sheet-level features. The closure receives a mutable reference to the sheet.
    ///
    /// - Parameters:
    ///   - index: The zero-based index of the sheet to modify.
    ///   - modify: A closure that receives a mutable sheet writer for modification.
    ///
    /// ## Example
    ///
    /// ```swift
    /// let sheetIndex = workbook.addSheet(named: "Data")
    /// 
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     // Write headers
    ///     sheet.writeText("Name", to: "A1")
    ///     sheet.writeText("Age", to: "B1")
    ///     
    ///     // Write data
    ///     sheet.writeText("Alice", to: "A2")
    ///     sheet.writeNumber(30, to: "B2")
    ///     
    ///     // Add formula
    ///     sheet.writeFormula("=AVERAGE(B:B)", to: "B3")
    ///     
    ///     // Protect sheet
    ///     sheet.protectSheet(password: "secret", options: .readonly)
    /// }
    /// ```
    public mutating func modifySheet(at index: Int, _ modify: (inout SheetWriter) -> Void) {
        guard index >= 0, index < sheets.count else { return }
        modify(&sheets[index])
    }
    
    /// Modifies the workbook's style table.
    ///
    /// Use this method to add custom fonts, fills, borders, number formats, and alignments
    /// that can be referenced by style index when writing cells.
    ///
    /// - Parameter modify: A closure that receives a mutable styles builder for modification.
    ///
    /// ## Example: Creating Custom Styles
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// 
    /// workbook.style { styles in
    ///     // Add number format for currency
    ///     styles.addNumberFormat("$#,##0.00")
    ///     
    ///     // Add bold font
    ///     styles.addFont(name: "Arial", size: 12, bold: true)
    ///     
    ///     // Add fill with light gray background
    ///     styles.addFill(pattern: .solid, fgColor: "FFD9D9D9")
    ///     
    ///     // Add center alignment
    ///     styles.addAlignment(horizontal: .center, vertical: .center)
    ///     
    ///     // Add border
    ///     styles.addBorder(
    ///         left: .thin,
    ///         right: .thin,
    ///         top: .thin,
    ///         bottom: .thin
    ///     )
    /// }
    /// 
    /// // Use the styles when writing cells
    /// let sheetIndex = workbook.addSheet(named: "Styled")
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     sheet.writeNumber(1234.56, to: "A1", styleIndex: 1)  // With currency format
    ///     sheet.writeText("Header", to: "B1", styleIndex: 2)    // With bold font
    /// }
    /// ```
    public mutating func style(_ modify: (inout StylesBuilder) -> Void) {
        modify(&stylesBuilder)
    }

    /// Adds a named range (defined name) to the workbook.
    ///
    /// Named ranges provide meaningful names for cell references, making formulas
    /// more readable and maintainable. They can reference single cells, ranges,
    /// or even formulas.
    ///
    /// - Parameters:
    ///   - name: The name for the range (e.g., "TotalSales", "TaxRate").
    ///   - refersTo: The cell reference or formula (e.g., "Sheet1!$A$1", "Sheet1!$A$1:$C$10").
    ///
    /// ## Example: Single Cell Reference
    ///
    /// ```swift
    /// workbook.addNamedRange(name: "TaxRate", refersTo: "Settings!$B$1")
    /// 
    /// // Now you can use =TaxRate in formulas instead of =Settings!$B$1
    /// sheet.writeFormula("=Revenue*TaxRate", to: "C1")
    /// ```
    ///
    /// ## Example: Range Reference
    ///
    /// ```swift
    /// workbook.addNamedRange(name: "SalesData", refersTo: "Sales!$A$2:$D$100")
    /// 
    /// // Use in formulas
    /// sheet.writeFormula("=SUM(SalesData)", to: "E1")
    /// sheet.writeFormula("=AVERAGE(SalesData)", to: "E2")
    /// ```
    public mutating func addNamedRange(name: String, refersTo: String) {
        namedRanges.append((name, refersTo))
    }

    /// Protects the workbook structure and windows with optional password.
    ///
    /// Workbook protection prevents users from adding, deleting, or reordering sheets,
    /// and can also lock window positions and sizes.
    ///
    /// - Parameters:
    ///   - password: Optional password required to unprotect the workbook. Defaults to `nil`.
    ///   - options: Protection settings. Defaults to `.default`. Use predefined options like
    ///     `.structureOnly` or `.strict`, or create custom ``WorkbookProtectionOptions``.
    ///
    /// ## Example: Protect Structure Only
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// workbook.addSheet(named: "Data")
    /// workbook.addSheet(named: "Summary")
    /// 
    /// // Prevent users from adding, deleting, or reordering sheets
    /// workbook.protectWorkbook(password: "secret", options: .structureOnly)
    /// 
    /// try workbook.save(to: URL(fileURLWithPath: "protected.xlsx"))
    /// ```
    ///
    /// ## Example: Full Protection
    ///
    /// ```swift
    /// // Protect both structure and windows
    /// workbook.protectWorkbook(password: "mypassword", options: .strict)
    /// ```
    public mutating func protectWorkbook(password: String? = nil, options: WorkbookProtectionOptions = .default) {
        let prot = options.toProtection(passwordHash: password)
        workbookProtection = prot
    }
    
    /// Saves the workbook to a file.
    ///
    /// Generates the complete Excel (.xlsx) file and writes it to the specified URL.
    /// This is a convenience method that calls ``buildData()`` and writes the result.
    ///
    /// - Parameter url: The file URL where the workbook should be saved.
    /// - Throws: An error if the file cannot be written or the workbook structure is invalid.
    ///
    /// ## Example
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// let sheetIndex = workbook.addSheet(named: "Sheet1")
    /// 
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     sheet.writeText("Hello, Excel!", to: "A1")
    /// }
    /// 
    /// let fileURL = URL(fileURLWithPath: "output.xlsx")
    /// try workbook.save(to: fileURL)
    /// ```
    public mutating func save(to url: URL) throws {
        let data = try buildData()
        try data.write(to: url)
    }
    
    /// Builds the workbook as data without writing to a file.
    ///
    /// Generates the complete Excel (.xlsx) file format as a `Data` object.
    /// Use this when you need to send the workbook over a network, store it in memory,
    /// or perform additional processing before writing to disk.
    ///
    /// - Returns: The workbook data in Excel (.xlsx) format.
    /// - Throws: An error if the workbook structure is invalid or cannot be generated.
    ///
    /// ## Example: Save to File
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// let sheetIndex = workbook.addSheet(named: "Sheet1")
    /// 
    /// workbook.modifySheet(at: sheetIndex) { sheet in
    ///     sheet.writeText("Hello, Excel!", to: "A1")
    /// }
    /// 
    /// let data = try workbook.buildData()
    /// try data.write(to: URL(fileURLWithPath: "output.xlsx"))
    /// ```
    ///
    /// ## Example: HTTP Response
    ///
    /// ```swift
    /// var workbook = WorkbookWriter()
    /// // ... configure workbook ...
    /// 
    /// let data = try workbook.buildData()
    /// // Send data as HTTP response with content-type:
    /// // application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
    /// ```
    public mutating func buildData() throws -> Data {
        var zipWriter = ZipWriter()
        
        // Build content types (write at end once all parts are known)
        var contentTypes = ContentTypesBuilder()
        contentTypes.addOverride(partName: "/xl/workbook.xml", 
                                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        contentTypes.addOverride(partName: "/xl/styles.xml",
                                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        for i in 1...sheets.count {
            contentTypes.addOverride(partName: "/xl/worksheets/sheet\(i).xml",
                                    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        }
        
        // Build root relationships
        var rootRels = RelationshipsBuilder()
        rootRels.addRelationship(
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            target: "xl/workbook.xml",
            id: "rId1"
        )
        zipWriter.addFile(path: "_rels/.rels", data: rootRels.build())
        
        // Prepare workbook (write after sheets so sharedStrings relationship can be added if needed)
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
        
        // Build styles
        zipWriter.addFile(path: "xl/styles.xml", data: stylesBuilder.build())
        
        // Build worksheets first to populate shared strings
        
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
        // After building sheets, emit sharedStrings if present and update content types + workbook rels
        if sharedStringsBuilder.count > 0 {
            // Content type override for shared strings
            contentTypes.addOverride(partName: "/xl/sharedStrings.xml",
                                    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
            // Add sharedStrings relationship
            workbookRels.addRelationship(
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                target: "sharedStrings.xml",
                id: "rId\(sheets.count + 2)"
            )
            // Write sharedStrings part
            zipWriter.addFile(path: "xl/sharedStrings.xml", data: sharedStringsBuilder.build())
        }

        // Write workbook and relationships now that rels are complete
        zipWriter.addFile(path: "xl/workbook.xml", data: workbookBuilder.build())
        zipWriter.addFile(path: "xl/_rels/workbook.xml.rels", data: workbookRels.build())
        
        // Finally write content types with all overrides
        zipWriter.addFile(path: "[Content_Types].xml", data: contentTypes.build())

        return try zipWriter.write()
    }
}
