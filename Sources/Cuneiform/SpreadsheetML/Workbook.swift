import Foundation

/// A high-level representation of an Excel workbook (.xlsx file).
///
/// A workbook is the top-level container for spreadsheet data in the Office Open XML format.
/// It provides access to all worksheets, shared resources (strings and styles), metadata,
/// and structural elements like pivot tables and Excel tables.
///
/// ## Overview
///
/// Use ``Workbook`` to open and read Excel files. Once opened, you can access individual sheets
/// by name or index, retrieve workbook-level metadata, and work with structured data like
/// pivot tables and Excel tables.
///
/// The workbook lazily loads worksheet content, so opening a large file is fast even if it
/// contains many sheets. Individual sheets are only parsed when you access them through
/// ``sheet(named:)`` or ``sheet(at:)``.
///
/// ## Opening a Workbook
///
/// Create a workbook instance by opening an `.xlsx` file from disk:
///
/// ```swift
/// let url = URL(fileURLWithPath: "sales_data.xlsx")
/// let workbook = try Workbook.open(url: url)
/// print("Opened workbook with \(workbook.sheets.count) sheets")
/// ```
///
/// ## Accessing Sheets
///
/// Retrieve individual worksheets by name or by zero-based index:
///
/// ```swift
/// // Access by name
/// if let sheet = try workbook.sheet(named: "Sales") {
///     print("Found sheet: \(sheet.name)")
/// }
///
/// // Access by index
/// if let firstSheet = try workbook.sheet(at: 0) {
///     print("First sheet: \(firstSheet.name)")
/// }
/// ```
///
/// ## Working with Defined Names
///
/// Access workbook-level named ranges and formulas:
///
/// ```swift
/// if let salesRange = workbook.definedName("SalesRegion") {
///     print("Sales region formula: \(salesRange.refersTo)")
/// }
///
/// // Parse a named range into components
/// if let (sheet, range) = workbook.definedNameRange("SalesRegion") {
///     print("Sheet: \(sheet), Range: \(range)")
/// }
/// ```
///
/// ## Thread Safety
///
/// ``Workbook`` conforms to `Sendable` and can be safely shared across concurrency domains.
/// All internal state is immutable after initialization.
///
/// ## Topics
///
/// ### Opening Workbooks
///
/// - ``open(url:)``
///
/// ### Accessing Sheets
///
/// - ``sheets``
/// - ``sheet(named:)``
/// - ``sheet(at:)``
///
/// ### Workbook Metadata
///
/// - ``protection``
///
/// ### Structured Data
///
/// - ``pivotTables``
/// - ``tables``
///
/// ### Named Ranges
///
/// - ``definedNamesList``
/// - ``definedName(_:)``
/// - ``definedNameRange(_:)``
///
/// - SeeAlso: ``Sheet``, ``SheetInfo``, ``DefinedName``
public struct Workbook: Sendable {
    private let package: OPCPackage
    private let workbookInfo: WorkbookInfo
    private let sharedStrings: SharedStrings
    private let styles: StylesInfo
    private let definedNames: [DefinedName]
    private let pivotTablesList: [PivotTableData]
    private let tablesList: [TableData]

    /// All sheets in the workbook, in their original order.
    ///
    /// This property returns metadata about each sheet without loading its full content.
    /// Use ``sheet(named:)`` or ``sheet(at:)`` to load and access the actual sheet data.
    ///
    /// The returned ``SheetInfo`` objects contain basic metadata like sheet name, ID,
    /// visibility state, and type (worksheet, chart sheet, etc.).
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    /// for (index, sheetInfo) in workbook.sheets.enumerated() {
    ///     print("\(index): \(sheetInfo.name) - \(sheetInfo.state)")
    /// }
    /// ```
    ///
    /// - SeeAlso: ``sheet(named:)``, ``sheet(at:)``
    public var sheets: [SheetInfo] { workbookInfo.sheets }

    /// Workbook-level protection settings, if the workbook is protected.
    ///
    /// Workbook protection restricts structural changes like adding, deleting, or renaming sheets.
    /// It does not affect the ability to read data or modify cell content in unprotected sheets.
    ///
    /// Returns `nil` if the workbook is not protected.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    /// if let protection = workbook.protection {
    ///     print("Workbook is protected with settings: \(protection)")
    /// } else {
    ///     print("Workbook is not protected")
    /// }
    /// ```
    ///
    /// - Note: Sheet-level protection is independent of workbook protection and is accessed
    ///   through individual ``Sheet`` objects.
    ///
    /// - SeeAlso: ``WorkbookProtection``
    public var protection: WorkbookProtection? { workbookInfo.protection }

    /// All pivot tables defined in the workbook.
    ///
    /// Pivot tables are discovered by traversing worksheet relationships. Each pivot table
    /// summarizes and analyzes data from a source range or data model.
    ///
    /// This property returns all pivot tables found across all worksheets in the workbook.
    /// If no pivot tables exist or if they cannot be parsed, this returns an empty array.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    /// for pivotTable in workbook.pivotTables {
    ///     print("Pivot table: \(pivotTable.name)")
    ///     print("  Location: \(pivotTable.location)")
    /// }
    /// ```
    ///
    /// - Note: Pivot tables that fail to parse are silently skipped. This ensures that
    ///   workbooks with partially invalid pivot tables can still be opened.
    ///
    /// - SeeAlso: ``PivotTableData``
    public var pivotTables: [PivotTableData] { pivotTablesList }

    /// All Excel tables defined in the workbook.
    ///
    /// Excel tables (also known as structured references) are formatted ranges with headers,
    /// optional totals rows, and automatic filtering. They provide named access to columns
    /// and enable structured formulas.
    ///
    /// This property returns all tables found across all worksheets in the workbook.
    /// If no tables exist or if they cannot be parsed, this returns an empty array.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    /// for table in workbook.tables {
    ///     print("Table: \(table.name)")
    ///     print("  Range: \(table.reference)")
    ///     print("  Columns: \(table.columns.count)")
    /// }
    /// ```
    ///
    /// - Note: Tables that fail to parse are silently skipped. This ensures that
    ///   workbooks with partially invalid tables can still be opened.
    ///
    /// - SeeAlso: ``TableData``
    public var tables: [TableData] { tablesList }

    /// Opens an Excel workbook from a file URL.
    ///
    /// Use this method to open and read `.xlsx` files from disk. The workbook provides access
    /// to all sheets, shared resources, and metadata. The file is fully validated and parsed
    /// at open time, but individual worksheet content is loaded lazily on demand.
    ///
    /// This method parses the workbook structure, shared strings table, styles, and
    /// relationship metadata. It also discovers optional resources like pivot tables and
    /// Excel tables by traversing worksheet relationships.
    ///
    /// ```swift
    /// let url = URL(fileURLWithPath: "sales_report.xlsx")
    /// let workbook = try Workbook.open(url: url)
    ///
    /// print("Opened workbook with \(workbook.sheets.count) sheets")
    /// print("Contains \(workbook.pivotTables.count) pivot tables")
    /// print("Contains \(workbook.tables.count) Excel tables")
    ///
    /// // Access the first sheet
    /// if let sheet = try workbook.sheet(at: 0) {
    ///     print("First sheet name: \(sheet.name)")
    /// }
    /// ```
    ///
    /// The method performs the following operations:
    /// 1. Opens the Office Open XML package and validates its structure
    /// 2. Parses the main workbook XML part
    /// 3. Loads the shared strings table (if present)
    /// 4. Loads style definitions (if present)
    /// 5. Discovers pivot tables across all worksheets
    /// 6. Discovers Excel tables across all worksheets
    ///
    /// - Parameter url: The file URL pointing to an `.xlsx` file on disk
    /// - Returns: A fully initialized workbook with all sheets and resources loaded
    /// - Throws: ``CuneiformError`` if the file cannot be opened, the package structure
    ///   is invalid, or required XML parts are malformed
    ///
    /// - Important: The file must be a valid Office Open XML SpreadsheetML document.
    ///   Legacy `.xls` files (BIFF format) are not supported.
    ///
    /// - Note: Optional resources like pivot tables and Excel tables are loaded on a
    ///   best-effort basis. If these fail to parse, they are silently skipped to ensure
    ///   the workbook can still be opened.
    ///
    /// - SeeAlso: ``sheet(named:)``, ``sheet(at:)``
    public static func open(url: URL) throws(CuneiformError) -> Workbook {
        let pkg: OPCPackage
        do {
            pkg = try OPCPackage.open(url: url)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.invalidPackageStructure(reason: error.localizedDescription)
        }

        // Parse workbook
        let wbData: Data
        do {
            wbData = try pkg.readPart(.workbook)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: PartPath.workbook.value, detail: error.localizedDescription)
        }
        let wb = try WorkbookParser.parse(data: wbData)

        // Parse shared strings
        let ss: SharedStrings
        if pkg.partExists(.sharedStrings) {
            let ssData: Data
            do {
                ssData = try pkg.readPart(.sharedStrings)
            } catch let err as CuneiformError {
                throw err
            } catch {
                throw CuneiformError.malformedXML(part: PartPath.sharedStrings.value, detail: error.localizedDescription)
            }
            ss = try SharedStringsParser.parse(data: ssData)
        } else {
            ss = .empty
        }

        // Parse styles
        let st: StylesInfo
        if pkg.partExists(.styles) {
            let stData: Data
            do {
                stData = try pkg.readPart(.styles)
            } catch let err as CuneiformError {
                throw err
            } catch {
                throw CuneiformError.malformedXML(part: PartPath.styles.value, detail: error.localizedDescription)
            }
            st = try StylesParser.parse(data: stData)
        } else {
            st = .empty
        }

        // Discover pivot tables via worksheet relationships
        var pivotTables: [PivotTableData] = []
        do {
            var mutablePkg = pkg
            let wbRels = try mutablePkg.relationships(for: .workbook)
            
            // Pivot tables are referenced from individual worksheets, not from the workbook level
            // Iterate through all sheet relationships to find pivot table references
            for sheet in wb.sheets {
                guard let rel = wbRels[sheet.relationshipId] else {
                    continue
                }
                
                let sheetPath = rel.resolveTarget(relativeTo: .workbook)
                
                do {
                    let wsRels = try mutablePkg.relationships(for: sheetPath)
                    let ptRels = wsRels[.pivotTable]
                    
                    for ptRel in ptRels {
                        let ptPath = ptRel.resolveTarget(relativeTo: sheetPath)
                        let ptData = try mutablePkg.readPart(ptPath)
                        if let pt = try? PivotTableParser.parse(data: ptData) {
                            pivotTables.append(pt)
                        }
                    }
                } catch {
                    // Silently ignore errors for individual sheet pivot tables
                }
            }
        } catch {
            // Silently ignore pivot table loading errors; they're optional
        }

        // Discover tables via worksheet relationships
        var tables: [TableData] = []
        do {
            var mutablePkg = pkg
            let wbRels = try mutablePkg.relationships(for: .workbook)
            
            // Tables are referenced from individual worksheets
            for sheet in wb.sheets {
                guard let rel = wbRels[sheet.relationshipId] else {
                    continue
                }
                
                let sheetPath = rel.resolveTarget(relativeTo: .workbook)
                
                do {
                    let wsRels = try mutablePkg.relationships(for: sheetPath)
                    let tableRels = wsRels[.table]
                    
                    for (index, tableRel) in tableRels.enumerated() {
                        let tablePath = tableRel.resolveTarget(relativeTo: sheetPath)
                        let tableData = try mutablePkg.readPart(tablePath)
                        // Extract table ID from relationship ID (e.g., "rId2" -> 2)
                        let tableId = index + 1
                        if let table = try? TableParser.parse(data: tableData, id: tableId) {
                            tables.append(table)
                        }
                    }
                } catch {
                    // Silently ignore errors for individual sheet tables
                }
            }
        } catch {
            // Silently ignore table loading errors; they're optional
        }

        return Workbook(package: pkg, workbookInfo: wb, sharedStrings: ss, styles: st, pivotTables: pivotTables, tables: tables)
    }

    /// Retrieves a worksheet by its name.
    ///
    /// Use this method to load a specific worksheet when you know its name. The name
    /// comparison is case-sensitive and must match exactly.
    ///
    /// This method lazily loads the worksheet content, including cell data, formatting,
    /// comments, and charts. If the worksheet has not been accessed before, it will be
    /// parsed from the underlying package on first access.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    ///
    /// // Load a sheet by name
    /// if let sheet = try workbook.sheet(named: "Q1 Sales") {
    ///     print("Found sheet: \(sheet.name)")
    ///     print("Dimensions: \(sheet.dimension ?? "unknown")")
    ///
    ///     // Access cell data
    ///     if let cellValue = sheet.cell(row: 0, column: 0)?.value {
    ///         print("A1 value: \(cellValue)")
    ///     }
    /// } else {
    ///     print("Sheet 'Q1 Sales' not found")
    /// }
    /// ```
    ///
    /// - Parameter named: The exact name of the worksheet to retrieve
    /// - Returns: A ``Sheet`` instance if a worksheet with the given name exists,
    ///   or `nil` if no matching sheet is found
    /// - Throws: ``CuneiformError`` if the worksheet exists but cannot be loaded due to
    ///   malformed XML, missing relationships, or other structural issues
    ///
    /// - Note: Sheet names in Excel are unique within a workbook and typically have
    ///   length limits (31 characters in Excel). Names may contain spaces and special
    ///   characters except: `[ ] : * ? / \`
    ///
    /// - SeeAlso: ``sheet(at:)``, ``sheets``
    public func sheet(named: String) throws(CuneiformError) -> Sheet? {
        guard let sheetInfo = workbookInfo.sheet(named: named) else { return nil }
        return try loadSheet(sheetInfo)
    }

    /// Retrieves a worksheet by its zero-based index.
    ///
    /// Use this method to load worksheets by their position in the workbook. The index
    /// corresponds to the order in which sheets appear in the Excel UI, starting from 0.
    ///
    /// This method lazily loads the worksheet content, including cell data, formatting,
    /// comments, and charts. If the worksheet has not been accessed before, it will be
    /// parsed from the underlying package on first access.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    ///
    /// // Load the first sheet
    /// if let firstSheet = try workbook.sheet(at: 0) {
    ///     print("First sheet: \(firstSheet.name)")
    /// }
    ///
    /// // Iterate through all sheets
    /// for i in 0..<workbook.sheets.count {
    ///     if let sheet = try workbook.sheet(at: i) {
    ///         print("Sheet \(i): \(sheet.name) has \(sheet.rows.count) rows")
    ///     }
    /// }
    /// ```
    ///
    /// - Parameter index: The zero-based index of the worksheet to retrieve. Valid indices
    ///   range from 0 to `sheets.count - 1`
    /// - Returns: A ``Sheet`` instance if a worksheet exists at the given index,
    ///   or `nil` if the index is out of bounds
    /// - Throws: ``CuneiformError`` if the worksheet exists but cannot be loaded due to
    ///   malformed XML, missing relationships, or other structural issues
    ///
    /// - Important: The index is zero-based, unlike Excel's 1-based sheet numbering
    ///   in formulas.
    ///
    /// - SeeAlso: ``sheet(named:)``, ``sheets``
    public func sheet(at index: Int) throws(CuneiformError) -> Sheet? {
        guard index >= 0, index < sheets.count else { return nil }
        return try loadSheet(sheets[index])
    }

    /// Load a sheet by SheetInfo
    private func loadSheet(_ sheetInfo: SheetInfo) throws(CuneiformError) -> Sheet {
        let wbRels: Relationships
        do {
            var mutablePkg = package
            wbRels = try mutablePkg.relationships(for: .workbook)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: PartPath.workbookRelationships.value, detail: error.localizedDescription)
        }
        
        guard let rel = wbRels[sheetInfo.relationshipId] else {
            throw CuneiformError.missingRequiredElement(
                element: "Relationship",
                inPart: "Sheet '\(sheetInfo.name)' with ID '\(sheetInfo.relationshipId)'"
            )
        }

        let sheetPath = rel.resolveTarget(relativeTo: .workbook)
        let sheetData: Data
        do {
            var mutablePkg = package
            sheetData = try mutablePkg.readPart(sheetPath)
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: sheetPath.value, detail: error.localizedDescription)
        }
        let worksheet = try WorksheetParser.parse(data: sheetData)

        // Parse comments via worksheet relationships (if present)
        let comments: [Comment]
        do {
            var mutablePkg = package
            let wsRels = try mutablePkg.relationships(for: sheetPath)
            let commentRels = wsRels[.comments]
            if commentRels.isEmpty {
                comments = []
            } else {
                var acc: [Comment] = []
                for cRel in commentRels {
                    let cPath = cRel.resolveTarget(relativeTo: sheetPath)
                    let cData = try mutablePkg.readPart(cPath)
                    let parsed = try CommentsParser.parse(data: cData)
                    acc.append(contentsOf: parsed.comments)
                }
                comments = acc
            }
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: sheetPath.relationshipsPath.value, detail: error.localizedDescription)
        }

        // Parse charts via worksheet relationships (if present)
        let charts: [ChartData]
        do {
            var mutablePkg = package
            let wsRels = try mutablePkg.relationships(for: sheetPath)
            let chartRels = wsRels[.chart]
            if chartRels.isEmpty {
                charts = []
            } else {
                var acc: [ChartData] = []
                for cRel in chartRels {
                    let cPath = cRel.resolveTarget(relativeTo: sheetPath)
                    let cData = try mutablePkg.readPart(cPath)
                    let parsed = try ChartParser.parse(data: cData)
                    acc.append(parsed)
                }
                charts = acc
            }
        } catch let err as CuneiformError {
            throw err
        } catch {
            throw CuneiformError.malformedXML(part: sheetPath.relationshipsPath.value, detail: error.localizedDescription)
        }

        return Sheet(data: worksheet, sharedStrings: sharedStrings, styles: styles, comments: comments, charts: charts)
    }

    /// All defined names (named ranges and formulas) defined at the workbook level.
    ///
    /// Defined names provide human-readable identifiers for cell ranges, formulas, and constants.
    /// They can be used in formulas across the workbook to reference cells or values without
    /// using explicit cell addresses.
    ///
    /// This property returns all workbook-scoped defined names. Sheet-scoped names are not
    /// included in this list.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    ///
    /// // List all defined names
    /// for name in workbook.definedNamesList {
    ///     print("Name: \(name.name)")
    ///     print("  Refers to: \(name.refersTo)")
    ///     if let comment = name.comment {
    ///         print("  Comment: \(comment)")
    ///     }
    /// }
    ///
    /// // Count named ranges
    /// print("Total defined names: \(workbook.definedNamesList.count)")
    /// ```
    ///
    /// Common uses for defined names include:
    /// - Named ranges: `SalesData` → `Sheet1!$A$1:$D$100`
    /// - Named constants: `TaxRate` → `0.08`
    /// - Named formulas: `CurrentQuarter` → `QUARTER(TODAY())`
    ///
    /// - Note: This includes all types of defined names: ranges, constants, and formulas.
    ///   Use ``definedName(_:)`` to retrieve a specific name or ``definedNameRange(_:)``
    ///   to parse range references.
    ///
    /// - SeeAlso: ``definedName(_:)``, ``definedNameRange(_:)``, ``DefinedName``
    public var definedNamesList: [DefinedName] { definedNames }

    /// Retrieves a workbook-level defined name by its exact name.
    ///
    /// Use this method to look up a specific defined name when you know its identifier.
    /// The name comparison is case-sensitive and must match exactly as defined in the workbook.
    ///
    /// Defined names can represent cell ranges, formulas, or constants. This method only
    /// searches workbook-scoped names; sheet-scoped names are not included.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    ///
    /// // Look up a named range
    /// if let salesRange = workbook.definedName("SalesData") {
    ///     print("Sales data refers to: \(salesRange.refersTo)")
    ///     // Output: "Sheet1!$A$1:$D$100"
    /// }
    ///
    /// // Look up a named constant
    /// if let taxRate = workbook.definedName("TaxRate") {
    ///     print("Tax rate formula: \(taxRate.refersTo)")
    ///     // Output: "0.08"
    /// }
    ///
    /// // Check if a name exists
    /// if workbook.definedName("QuarterlyTotal") != nil {
    ///     print("QuarterlyTotal is defined")
    /// }
    /// ```
    ///
    /// - Parameter name: The exact name of the defined name to retrieve, case-sensitive
    /// - Returns: A ``DefinedName`` instance if a name with the exact identifier exists,
    ///   or `nil` if no matching name is found
    ///
    /// - Note: Excel allows defined names to contain letters, numbers, periods, and underscores,
    ///   but they must start with a letter, underscore, or backslash. Names cannot look like
    ///   cell references (e.g., `A1`, `XFD1048576`).
    ///
    /// - SeeAlso: ``definedNamesList``, ``definedNameRange(_:)``
    public func definedName(_ name: String) -> DefinedName? {
        definedNames.first { $0.name == name }
    }

    /// Parses a defined name's reference into sheet and range components.
    ///
    /// Use this method to extract the sheet name and cell range from a defined name that
    /// refers to a worksheet range. This is useful when you need to programmatically access
    /// the cells referenced by a named range.
    ///
    /// This method handles both simple sheet names and sheet names containing spaces (which
    /// are quoted in Excel). It parses references of the form `SheetName!$A$1:$B$10` or
    /// `'Sheet Name'!$A$1:$B$10`.
    ///
    /// ```swift
    /// let workbook = try Workbook.open(url: fileURL)
    ///
    /// // Parse a simple range reference
    /// if let (sheet, range) = workbook.definedNameRange("SalesData") {
    ///     print("Sheet: \(sheet)")    // "Sheet1"
    ///     print("Range: \(range)")     // "$A$1:$D$100"
    ///
    ///     // Use the parsed components to access the sheet
    ///     if let targetSheet = try workbook.sheet(named: sheet) {
    ///         print("Successfully loaded sheet for named range")
    ///     }
    /// }
    ///
    /// // Parse a range with a sheet name containing spaces
    /// if let (sheet, range) = workbook.definedNameRange("QuarterlySales") {
    ///     print("Sheet: \(sheet)")    // "Q1 Sales"
    ///     print("Range: \(range)")     // "$B$2:$E$50"
    /// }
    ///
    /// // Handle non-range defined names
    /// if workbook.definedNameRange("TaxRate") == nil {
    ///     print("TaxRate is not a range reference")
    /// }
    /// ```
    ///
    /// The method handles the following formats:
    /// - `Sheet1!$A$1:$B$10` → `("Sheet1", "$A$1:$B$10")`
    /// - `'My Sheet'!$A$1:$B$10` → `("My Sheet", "$A$1:$B$10")`
    /// - `Sheet1!$A$1` → `("Sheet1", "$A$1")`
    ///
    /// - Parameter name: The name of the defined name to parse
    /// - Returns: A tuple containing the sheet name and range string if the defined name
    ///   exists and refers to a worksheet range, or `nil` if the name doesn't exist or
    ///   doesn't reference a sheet range (e.g., constants or formulas)
    ///
    /// - Note: This method only parses the reference format. It does not validate that
    ///   the sheet exists or that the range is valid. Use ``sheet(named:)`` to verify
    ///   the sheet exists.
    ///
    /// - SeeAlso: ``definedName(_:)``, ``definedNamesList``
    public func definedNameRange(_ name: String) -> (sheet: String, range: String)? {
        guard let dn = definedName(name) else { return nil }
        // Expected patterns like Sheet1!$A$1:$A$10 or 'My Sheet'!$A$1:$B$2
        let refers = dn.refersTo.trimmingCharacters(in: .whitespacesAndNewlines)
        if refers.isEmpty { return nil }
        // If sheet has quotes, split on last !
        if let bangIndex = refers.lastIndex(of: "!") {
            let sheetName = String(refers[..<bangIndex]).trimmingCharacters(in: CharacterSet(charactersIn: "'"))
            let rangePart = String(refers[refers.index(after: bangIndex)...])
            return (sheet: sheetName, range: rangePart)
        }
        return nil
    }

    init(package: OPCPackage, workbookInfo: WorkbookInfo, sharedStrings: SharedStrings, styles: StylesInfo, pivotTables: [PivotTableData] = [], tables: [TableData] = []) {
        self.package = package
        self.workbookInfo = workbookInfo
        self.sharedStrings = sharedStrings
        self.styles = styles
        self.definedNames = workbookInfo.definedNames
        self.pivotTablesList = pivotTables
        self.tablesList = tables
    }
}
