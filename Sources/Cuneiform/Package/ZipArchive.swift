import Foundation
import Compression

/// A minimal ZIP archive reader for OPC packages
///
/// This implementation handles the subset of ZIP features used by Office Open XML:
/// - Deflate compression (method 8)
/// - Store (no compression, method 0)
/// - Standard ZIP headers
struct ZipArchive: Sendable {
    /// Entries in the archive keyed by path
    let entries: [String: ZipEntry]

    /// Read a ZIP archive from a file URL
    static func read(from url: URL) throws -> ZipArchive {
        let data = try Data(contentsOf: url)
        return try read(from: data)
    }

    /// Read a ZIP archive from data
    static func read(from data: Data) throws -> ZipArchive {
        var entries: [String: ZipEntry] = [:]

        // Find the End of Central Directory record
        guard let eocdOffset = findEndOfCentralDirectory(in: data) else {
            throw CuneiformError.invalidZipArchive(reason: "Cannot find End of Central Directory record")
        }

        // Parse EOCD to get central directory location
        let eocd = try parseEOCD(data: data, offset: eocdOffset)

        // Parse central directory entries
        var offset = eocd.centralDirectoryOffset
        for _ in 0..<eocd.entryCount {
            let (entry, nextOffset) = try parseCentralDirectoryEntry(data: data, offset: offset)
            entries[entry.path] = entry
            offset = nextOffset
        }

        return ZipArchive(entries: entries)
    }

    /// Extract the data for an entry
    func extractData(for entry: ZipEntry, from archiveData: Data) throws -> Data {
        // Read local file header to get actual data offset
        let localHeaderOffset = entry.localHeaderOffset

        guard archiveData.count > localHeaderOffset + 30 else {
            throw CuneiformError.invalidZipArchive(reason: "Local file header offset out of bounds")
        }

        // Parse local file header
        let signature = archiveData.readUInt32LE(at: localHeaderOffset)
        guard signature == 0x04034b50 else {
            throw CuneiformError.invalidZipArchive(reason: "Invalid local file header signature")
        }

        let fileNameLength = Int(archiveData.readUInt16LE(at: localHeaderOffset + 26))
        let extraFieldLength = Int(archiveData.readUInt16LE(at: localHeaderOffset + 28))
        let dataOffset = localHeaderOffset + 30 + fileNameLength + extraFieldLength

        guard archiveData.count >= dataOffset + entry.compressedSize else {
            throw CuneiformError.invalidZipArchive(reason: "Compressed data extends beyond archive")
        }

        let compressedData = archiveData.subdata(in: dataOffset..<(dataOffset + entry.compressedSize))

        switch entry.compressionMethod {
        case 0: // Store (no compression)
            return compressedData

        case 8: // Deflate
            return try decompressDeflate(compressedData, uncompressedSize: entry.uncompressedSize)

        default:
            throw CuneiformError.invalidZipArchive(reason: "Unsupported compression method: \(entry.compressionMethod)")
        }
    }
}

// MARK: - ZipEntry

struct ZipEntry: Sendable {
    let path: String
    let compressionMethod: UInt16
    let compressedSize: Int
    let uncompressedSize: Int
    let localHeaderOffset: Int
}

// MARK: - EOCD Parsing

private struct EOCD {
    let entryCount: Int
    let centralDirectoryOffset: Int
}

private func findEndOfCentralDirectory(in data: Data) -> Int? {
    // EOCD signature: 0x06054b50
    // Search backwards from end of file (EOCD is at most 65535 + 22 bytes from end)
    let signature: [UInt8] = [0x50, 0x4b, 0x05, 0x06]
    let searchLimit = min(data.count, 65557)
    let startIndex = data.count - searchLimit

    for i in stride(from: data.count - 22, through: startIndex, by: -1) {
        if data[i] == signature[0] &&
           data[i + 1] == signature[1] &&
           data[i + 2] == signature[2] &&
           data[i + 3] == signature[3] {
            return i
        }
    }
    return nil
}

private func parseEOCD(data: Data, offset: Int) throws -> EOCD {
    guard data.count >= offset + 22 else {
        throw CuneiformError.invalidZipArchive(reason: "EOCD record too short")
    }

    let entryCount = Int(data.readUInt16LE(at: offset + 10))
    let centralDirectoryOffset = Int(data.readUInt32LE(at: offset + 16))

    return EOCD(entryCount: entryCount, centralDirectoryOffset: centralDirectoryOffset)
}

// MARK: - Central Directory Parsing

private func parseCentralDirectoryEntry(data: Data, offset: Int) throws -> (ZipEntry, Int) {
    guard data.count >= offset + 46 else {
        throw CuneiformError.invalidZipArchive(reason: "Central directory entry too short")
    }

    let signature = data.readUInt32LE(at: offset)
    guard signature == 0x02014b50 else {
        throw CuneiformError.invalidZipArchive(reason: "Invalid central directory signature")
    }

    let compressionMethod = data.readUInt16LE(at: offset + 10)
    let compressedSize = Int(data.readUInt32LE(at: offset + 20))
    let uncompressedSize = Int(data.readUInt32LE(at: offset + 24))
    let fileNameLength = Int(data.readUInt16LE(at: offset + 28))
    let extraFieldLength = Int(data.readUInt16LE(at: offset + 30))
    let commentLength = Int(data.readUInt16LE(at: offset + 32))
    let localHeaderOffset = Int(data.readUInt32LE(at: offset + 42))

    guard data.count >= offset + 46 + fileNameLength else {
        throw CuneiformError.invalidZipArchive(reason: "File name extends beyond archive")
    }

    let fileNameData = data.subdata(in: (offset + 46)..<(offset + 46 + fileNameLength))
    guard let fileName = String(data: fileNameData, encoding: .utf8) else {
        throw CuneiformError.invalidZipArchive(reason: "Invalid file name encoding")
    }

    let entry = ZipEntry(
        path: fileName,
        compressionMethod: compressionMethod,
        compressedSize: compressedSize,
        uncompressedSize: uncompressedSize,
        localHeaderOffset: localHeaderOffset
    )

    let nextOffset = offset + 46 + fileNameLength + extraFieldLength + commentLength
    return (entry, nextOffset)
}

// MARK: - Decompression

private func decompressDeflate(_ compressedData: Data, uncompressedSize: Int) throws -> Data {
    // Use Apple's Compression framework with raw deflate (no zlib header)
    var decompressed = Data(count: uncompressedSize)
    let decompressedSize = decompressed.withUnsafeMutableBytes { destBuffer in
        compressedData.withUnsafeBytes { srcBuffer in
            compression_decode_buffer(
                destBuffer.bindMemory(to: UInt8.self).baseAddress!,
                uncompressedSize,
                srcBuffer.bindMemory(to: UInt8.self).baseAddress!,
                compressedData.count,
                nil,
                COMPRESSION_ZLIB
            )
        }
    }

    guard decompressedSize == uncompressedSize else {
        throw CuneiformError.invalidZipArchive(reason: "Decompression size mismatch: expected \(uncompressedSize), got \(decompressedSize)")
    }

    return decompressed
}

// MARK: - Data Extensions

extension Data {
    func readUInt16LE(at offset: Int) -> UInt16 {
        UInt16(self[offset]) | (UInt16(self[offset + 1]) << 8)
    }

    func readUInt32LE(at offset: Int) -> UInt32 {
        UInt32(self[offset]) |
        (UInt32(self[offset + 1]) << 8) |
        (UInt32(self[offset + 2]) << 16) |
        (UInt32(self[offset + 3]) << 24)
    }
}
