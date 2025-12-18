/// Errors that can occur during Cuneiform operations
public enum CuneiformError: Error, Sendable {
    // MARK: - Package Errors

    /// The file is not a valid ZIP archive
    case invalidZipArchive(reason: String)

    /// A required part is missing from the package
    case missingPart(path: String)

    /// Content type mismatch
    case invalidContentType(path: String, expected: String, found: String)

    /// The package structure is invalid
    case invalidPackageStructure(reason: String)

    // MARK: - XML Parsing Errors

    /// XML parsing failed
    case malformedXML(part: String, detail: String)

    /// Required XML element is missing
    case missingRequiredElement(element: String, inPart: String)

    /// Invalid attribute value
    case invalidAttributeValue(attribute: String, value: String, inPart: String)

    // MARK: - Data Errors

    /// Invalid cell reference format
    case invalidCellReference(String)

    /// Shared string index out of range
    case sharedStringIndexOutOfRange(index: Int, count: Int)

    /// Style index out of range
    case styleIndexOutOfRange(index: Int, count: Int)

    // MARK: - File Errors

    /// File not found
    case fileNotFound(path: String)

    /// Access denied to file
    case accessDenied(path: String)

    /// Not an xlsx file
    case notAnXlsxFile(path: String)
}

extension CuneiformError: CustomStringConvertible {
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
