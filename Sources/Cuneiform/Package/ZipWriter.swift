import Foundation

/// Simple ZIP archive writer for creating .xlsx files
public struct ZipWriter {
    public struct Entry {
        public let path: String
        public let data: Data
        
        public init(path: String, data: Data) {
            self.path = path
            self.data = data
        }
    }
    
    private var entries: [Entry] = []
    
    public init() {}
    
    /// Add a file to the archive
    public mutating func addFile(path: String, data: Data) {
        entries.append(Entry(path: path, data: data))
    }
    
    /// Write the ZIP archive
    public func write() throws -> Data {
        // For simplicity, use the zip command-line tool via a temp directory
        // In production, you'd use a proper ZIP library
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
        
        try FileManager.default.createDirectory(at: tempDir, withIntermediateDirectories: true)
        
        defer {
            try? FileManager.default.removeItem(at: tempDir)
        }
        
        // Write all entries to temp directory
        for entry in entries {
            let url = tempDir.appendingPathComponent(entry.path)
            let dir = url.deletingLastPathComponent()
            try FileManager.default.createDirectory(at: dir, withIntermediateDirectories: true)
            try entry.data.write(to: url)
        }
        
        // Create ZIP using command-line tool
        let zipURL = tempDir.appendingPathComponent("output.zip")
        let process = Process()
        process.currentDirectoryURL = tempDir
        process.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
        process.arguments = ["-r", "-q", "output.zip"] + entries.map { $0.path }
        
        try process.run()
        process.waitUntilExit()
        
        guard process.terminationStatus == 0 else {
            throw CuneiformError.invalidPackageStructure(reason: "ZIP creation failed")
        }
        
        return try Data(contentsOf: zipURL)
    }
}
