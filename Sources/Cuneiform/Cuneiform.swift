/// # Cuneiform
///
/// Pure Swift library for reading and writing Office Open XML SpreadsheetML (.xlsx) files.
///
/// ## Overview
///
/// Cuneiform provides comprehensive support for working with Excel files programmatically,
/// following the ISO/IEC 29500 Office Open XML standard. The library offers both high-level
/// and low-level APIs for reading, writing, and querying spreadsheet data.
///
/// ## Topics
///
/// ### Reading Workbooks
///
/// - ``Workbook``
/// - ``Sheet``
/// - ``CellValue``
/// - ``CellReference``
/// - ``CellFormula``
///
/// ### Writing Workbooks
///
/// - ``WorkbookWriter``
///
/// ### Advanced Queries
///
/// Use powerful query methods on ``Sheet`` for filtering and searching:
/// - ``Sheet/range(_:)``
/// - ``Sheet/column(_:)``
/// - ``Sheet/rows(where:)``
/// - ``Sheet/find(where:)``
/// - ``Sheet/findAll(where:)``
///
/// ### Core Components
///
/// - ``OPCPackage``
/// - ``PartPath``
/// - ``ContentType``
/// - ``Relationship``
///
/// ### Error Handling
///
/// - ``CuneiformError``
///
/// ## Getting Started
///
/// ### Reading an Excel File
///
/// ```swift
/// import Cuneiform
///
/// // Open an .xlsx file
/// let workbook = try Workbook.open(url: fileURL)
///
/// // Access a sheet
/// if let sheet = try workbook.sheet(named: "Sales") {
///     // Read cell values
///     if let revenue = sheet.cell(at: "B2") {
///         print("Revenue: \(revenue)")
///     }
///     
///     // Iterate through rows efficiently
///     for row in sheet.rows() {
///         for (ref, value) in row {
///             print("\(ref): \(value)")
///         }
///     }
/// }
/// ```
///
/// ### Writing an Excel File
///
/// ```swift
/// import Cuneiform
///
/// var writer = WorkbookWriter()
/// let sheetIndex = writer.addSheet(named: "Report")
///
/// writer.modifySheet(at: sheetIndex) { sheet in
///     sheet.writeText("Product", to: "A1")
///     sheet.writeText("Sales", to: "B1")
///     sheet.writeNumber(1250.50, to: "B2")
///     sheet.writeFormula("SUM(B2:B10)", to: "B11")
/// }
///
/// try writer.save(to: outputURL)
/// ```
///
/// ### Advanced Queries
///
/// ```swift
/// // Find all cells matching criteria
/// let highValues = sheet.findAll { _, value in
///     if case .number(let n) = value {
///         return n > 1000
///     }
///     return false
/// }
///
/// // Filter rows by condition
/// let activeCustomers = sheet.rows { cells in
///     cells.contains { $0.value == .text("Active") }
/// }
///
/// // Access ranges
/// let salesData = sheet.range("B2:B10")
/// ```
///
/// ## Performance Considerations
///
/// - Use ``Sheet/rows()`` for memory-efficient streaming of large files
/// - Sheets are loaded lazily—only accessed sheets are parsed
/// - Write operations are buffered until ``WorkbookWriter/save(to:)``
///
/// ## Requirements
///
/// - Swift 6.0+
/// - macOS 13.0+ / iOS 16.0+ / tvOS 16.0+ / watchOS 9.0+ / visionOS 1.0+
//     let sheet = workbook.sheets[0]
//     let cell = sheet.cell(at: .row(1), .column(1))
//     print(cell.stringValue)
//
// For more information about the Office Open XML format, see:
// https://www.iso.org/standard/71691.html

/// Library version information
public enum Cuneiform {
    /// The library version
    public static let version = "0.1.0"

    /// The supported Office Open XML version
    public static let ooXmlVersion = "ISO/IEC 29500 Transitional"
}
