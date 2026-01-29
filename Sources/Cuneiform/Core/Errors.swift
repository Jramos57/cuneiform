/// Errors that can occur during Cuneiform operations.
///
/// `CuneiformError` represents all possible errors that can occur when working with
/// Cuneiform, from loading XLSX files to parsing their contents.
///
/// ## Overview
///
/// Cuneiform operations may fail for various reasons: invalid file formats, corrupted
/// ZIP archives, malformed XML, missing required data, or file system issues. This error
/// type categorizes all possible failure modes to help you diagnose and handle problems.
///
/// All errors conform to Swift's `Error` protocol and provide detailed, human-readable
/// descriptions through ``CustomStringConvertible``.
///
/// ## Error Handling
///
/// Use standard Swift error handling to catch and handle Cuneiform errors:
///
/// ```swift
/// do {
///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
///     // Work with workbook...
/// } catch let error as CuneiformError {
///     switch error {
///     case .fileNotFound(let path):
///         print("File not found: \(path)")
///     case .notAnXlsxFile(let path):
///         print("Not an Excel file: \(path)")
///     case .malformedXML(let part, let detail):
///         print("Corrupt file in \(part): \(detail)")
///     case .invalidCellReference(let ref):
///         print("Invalid reference: \(ref)")
///     default:
///         print("Error: \(error)")
///     }
/// } catch {
///     print("Unexpected error: \(error)")
/// }
/// ```
///
/// ## Error Categories
///
/// Errors are organized into logical categories:
///
/// ### Package Errors
/// Issues with the XLSX file structure or ZIP archive.
/// - ``invalidZipArchive(reason:)``
/// - ``missingPart(path:)``
/// - ``invalidContentType(path:expected:found:)``
/// - ``invalidPackageStructure(reason:)``
///
/// ### XML Parsing Errors
/// Problems parsing the XML content within the XLSX file.
/// - ``malformedXML(part:detail:)``
/// - ``missingRequiredElement(element:inPart:)``
/// - ``invalidAttributeValue(attribute:value:inPart:)``
///
/// ### Data Errors
/// Issues with cell references, indices, or data integrity.
/// - ``invalidCellReference(_:)``
/// - ``sharedStringIndexOutOfRange(index:count:)``
/// - ``styleIndexOutOfRange(index:count:)``
///
/// ### File Errors
/// File system access problems.
/// - ``fileNotFound(path:)``
/// - ``accessDenied(path:)``
/// - ``notAnXlsxFile(path:)``
///
/// ## Debugging Tips
///
/// When encountering errors:
///
/// 1. **Check the file format**: Ensure the file is a valid `.xlsx` file (not `.xls` or other formats)
/// 2. **Verify file integrity**: Try opening the file in Excel to check for corruption
/// 3. **Examine the error details**: Use the associated values for specific information
/// 4. **Check file permissions**: Ensure your app has read access to the file
///
/// ```swift
/// // Extract detailed information from errors
/// if case .malformedXML(let part, let detail) = error {
///     print("The file is corrupted in \(part)")
///     print("Details: \(detail)")
/// }
/// ```
///
/// ## Topics
///
/// ### Package Errors
///
/// - ``invalidZipArchive(reason:)``
/// - ``missingPart(path:)``
/// - ``invalidContentType(path:expected:found:)``
/// - ``invalidPackageStructure(reason:)``
///
/// ### XML Parsing Errors
///
/// - ``malformedXML(part:detail:)``
/// - ``missingRequiredElement(element:inPart:)``
/// - ``invalidAttributeValue(attribute:value:inPart:)``
///
/// ### Data Errors
///
/// - ``invalidCellReference(_:)``
/// - ``sharedStringIndexOutOfRange(index:count:)``
/// - ``styleIndexOutOfRange(index:count:)``
///
/// ### File Errors
///
/// - ``fileNotFound(path:)``
/// - ``accessDenied(path:)``
/// - ``notAnXlsxFile(path:)``
///
/// ### Descriptions
///
/// - ``description``
///
/// ## See Also
///
/// - ``Cuneiform/loadWorkbook(from:)``
public enum CuneiformError: Error, Sendable {
    // MARK: - Package Errors

    /// The file is not a valid ZIP archive.
    ///
    /// XLSX files are ZIP archives containing XML documents. This error occurs when the
    /// file cannot be opened as a ZIP archive, indicating it's either corrupted or not
    /// an XLSX file.
    ///
    /// ```swift
    /// // Example: Attempting to open a non-ZIP file
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: textFileURL)
    /// } catch CuneiformError.invalidZipArchive(let reason) {
    ///     print("Not a valid XLSX file: \(reason)")
    /// }
    /// ```
    ///
    /// - Parameter reason: A description of why the ZIP archive is invalid.
    case invalidZipArchive(reason: String)

    /// A required part is missing from the package.
    ///
    /// XLSX files contain multiple parts (XML documents) with specific paths. This error
    /// occurs when a required part, such as the workbook definition or a worksheet, is
    /// missing from the archive.
    ///
    /// Common missing parts include:
    /// - `/xl/workbook.xml` - The main workbook definition
    /// - `/xl/worksheets/sheet1.xml` - Worksheet data
    /// - `/_rels/.rels` - Package relationships
    ///
    /// ```swift
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.missingPart(let path) {
    ///     print("File is missing required data: \(path)")
    /// }
    /// ```
    ///
    /// - Parameter path: The path to the missing part within the XLSX archive.
    case missingPart(path: String)

    /// Content type mismatch.
    ///
    /// Each part in an XLSX file has an associated content type. This error occurs when
    /// a part's content type doesn't match what's expected for that file type.
    ///
    /// - Parameters:
    ///   - path: The path to the part with the incorrect content type.
    ///   - expected: The expected content type.
    ///   - found: The actual content type found.
    case invalidContentType(path: String, expected: String, found: String)

    /// The package structure is invalid.
    ///
    /// This error occurs when the XLSX file's internal structure doesn't conform to the
    /// Office Open XML specification, such as missing required relationships or incorrect
    /// folder organization.
    ///
    /// ```swift
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.invalidPackageStructure(let reason) {
    ///     print("File structure is invalid: \(reason)")
    /// }
    /// ```
    ///
    /// - Parameter reason: A description of the structural problem.
    case invalidPackageStructure(reason: String)

    // MARK: - XML Parsing Errors

    /// XML parsing failed.
    ///
    /// This error occurs when the XML content within an XLSX part cannot be parsed,
    /// usually due to malformed XML syntax or encoding issues.
    ///
    /// ```swift
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.malformedXML(let part, let detail) {
    ///     print("Corrupted XML in \(part): \(detail)")
    /// }
    /// ```
    ///
    /// - Parameters:
    ///   - part: The part of the XLSX file containing malformed XML.
    ///   - detail: Detailed information about the parsing failure.
    case malformedXML(part: String, detail: String)

    /// Required XML element is missing.
    ///
    /// This error occurs when the XML parser cannot find a required element in the
    /// document structure, indicating an incomplete or invalid XLSX file.
    ///
    /// ```swift
    /// // Example: Missing worksheet dimension element
    /// catch CuneiformError.missingRequiredElement(let element, let inPart) {
    ///     print("Missing <\(element)> in \(inPart)")
    /// }
    /// ```
    ///
    /// - Parameters:
    ///   - element: The name of the missing XML element.
    ///   - inPart: The part of the XLSX file where the element should exist.
    case missingRequiredElement(element: String, inPart: String)

    /// Invalid attribute value.
    ///
    /// This error occurs when an XML element has an attribute with an invalid or
    /// unexpected value.
    ///
    /// - Parameters:
    ///   - attribute: The name of the attribute with an invalid value.
    ///   - value: The invalid value that was found.
    ///   - inPart: The part of the XLSX file containing the invalid attribute.
    case invalidAttributeValue(attribute: String, value: String, inPart: String)

    // MARK: - Data Errors

    /// Invalid cell reference format.
    ///
    /// This error occurs when a cell reference string doesn't match the expected A1
    /// notation format (e.g., "A1", "Z100", "AA5"). This typically indicates data
    /// corruption or an unsupported reference format.
    ///
    /// Valid references consist of:
    /// - One or more letters for the column (A-Z, AA-ZZ, etc.)
    /// - One or more digits for the row (1-based)
    /// - Optional `$` symbols for absolute references
    ///
    /// ```swift
    /// // Examples of invalid references that would trigger this error:
    /// // - "" (empty)
    /// // - "123" (no column)
    /// // - "ABC" (no row)
    /// // - "A-5" (invalid characters)
    ///
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.invalidCellReference(let ref) {
    ///     print("Invalid cell reference: \(ref)")
    /// }
    /// ```
    ///
    /// - Parameter reference: The invalid cell reference string.
    ///
    /// - SeeAlso: ``CellReference/init(_:)``
    case invalidCellReference(String)

    /// Shared string index out of range.
    ///
    /// XLSX files store repeated text in a shared string table to save space. This error
    /// occurs when a cell references a shared string index that doesn't exist in the
    /// table, indicating file corruption.
    ///
    /// ```swift
    /// catch CuneiformError.sharedStringIndexOutOfRange(let index, let count) {
    ///     print("Cell references string \(index), but only \(count) strings exist")
    /// }
    /// ```
    ///
    /// - Parameters:
    ///   - index: The requested shared string index.
    ///   - count: The total number of shared strings available.
    case sharedStringIndexOutOfRange(index: Int, count: Int)

    /// Style index out of range.
    ///
    /// Cells can reference styles defined in the styles table. This error occurs when
    /// a cell references a style index that doesn't exist, indicating file corruption
    /// or an incomplete styles definition.
    ///
    /// ```swift
    /// catch CuneiformError.styleIndexOutOfRange(let index, let count) {
    ///     print("Cell uses style \(index), but only \(count) styles defined")
    /// }
    /// ```
    ///
    /// - Parameters:
    ///   - index: The requested style index.
    ///   - count: The total number of styles available.
    case styleIndexOutOfRange(index: Int, count: Int)

    // MARK: - File Errors

    /// File not found.
    ///
    /// The specified file path doesn't exist or cannot be found. Check that the path
    /// is correct and the file exists.
    ///
    /// ```swift
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.fileNotFound(let path) {
    ///     print("File does not exist: \(path)")
    ///     // Maybe prompt user to select a different file
    /// }
    /// ```
    ///
    /// - Parameter path: The path to the file that was not found.
    case fileNotFound(path: String)

    /// Access denied to file.
    ///
    /// The application doesn't have permission to read the file. This can occur due to
    /// file system permissions, sandboxing restrictions, or security settings.
    ///
    /// ```swift
    /// catch CuneiformError.accessDenied(let path) {
    ///     print("Cannot access file: \(path)")
    ///     // On macOS/iOS, may need to request user permission
    /// }
    /// ```
    ///
    /// - Parameter path: The path to the file that cannot be accessed.
    case accessDenied(path: String)

    /// Not an xlsx file.
    ///
    /// The file exists but is not a valid XLSX file. This could be:
    /// - A different file format (e.g., `.xls`, `.csv`, `.txt`)
    /// - A corrupted or incomplete file
    /// - A file with the wrong extension
    ///
    /// ```swift
    /// do {
    ///     let workbook = try Cuneiform.loadWorkbook(from: fileURL)
    /// } catch CuneiformError.notAnXlsxFile(let path) {
    ///     print("File is not an Excel .xlsx file: \(path)")
    ///     // Suggest converting to XLSX format
    /// }
    /// ```
    ///
    /// - Parameter path: The path to the file that is not an XLSX file.
    ///
    /// - Note: Cuneiform only supports the modern XLSX format (Office Open XML). Legacy
    ///   XLS files (binary format) are not supported.
    case notAnXlsxFile(path: String)
}

extension CuneiformError: CustomStringConvertible {
    /// A human-readable description of the error.
    ///
    /// Provides detailed information about what went wrong, including any relevant
    /// paths, values, or context that can help diagnose the issue.
    ///
    /// ```swift
    /// catch let error as CuneiformError {
    ///     print(error.description)
    ///     // Example: "Invalid cell reference: 'XYZ'"
    ///     // Example: "File not found: /path/to/file.xlsx"
    /// }
    /// ```
    public var description: String {
        switch self {
        case .invalidZipArchive(let reason):
            return "Invalid ZIP archive: \(reason)"
        case .missingPart(let path):
            return "Missing required part: \(path)"
        case .invalidContentType(let path, let expected, let found):
            return "Invalid content type for '\(path)': expected '\(expected)', found '\(found)'"
        case .invalidPackageStructure(let reason):
            return "Invalid package structure: \(reason)"
        case .malformedXML(let part, let detail):
            return "Malformed XML in '\(part)': \(detail)"
        case .missingRequiredElement(let element, let inPart):
            return "Missing required element '\(element)' in '\(inPart)'"
        case .invalidAttributeValue(let attribute, let value, let inPart):
            return "Invalid value '\(value)' for attribute '\(attribute)' in '\(inPart)'"
        case .invalidCellReference(let ref):
            return "Invalid cell reference: '\(ref)'"
        case .sharedStringIndexOutOfRange(let index, let count):
            return "Shared string index \(index) out of range (count: \(count))"
        case .styleIndexOutOfRange(let index, let count):
            return "Style index \(index) out of range (count: \(count))"
        case .fileNotFound(let path):
            return "File not found: \(path)"
        case .accessDenied(let path):
            return "Access denied: \(path)"
        case .notAnXlsxFile(let path):
            return "Not an xlsx file: \(path)"
        }
    }
}
