import Foundation

/// Fit-to-pages configuration
public struct FitToPages: Sendable, Equatable {
    /// Pages wide
    public let width: Int
    /// Pages tall
    public let height: Int
    
    public init(width: Int, height: Int) {
        self.width = width
        self.height = height
    }
}

/// Page setup and print configuration for a worksheet
///
/// Represents `<pageSetup>` and `<pageMargins>` elements defining
/// how a worksheet should be printed, including orientation,
/// paper size, margins, and scaling.
public struct PageSetup: Sendable, Equatable {
    /// Paper size identifier (ECMA-376 §18.18.31)
    public enum PaperSize: Int, Sendable {
        case letter = 1           // 8.5 x 11 in
        case tabloid = 3          // 11 x 17 in
        case ledger = 4           // 17 x 11 in
        case legal = 5            // 8.5 x 14 in
        case statement = 6        // 5.5 x 8.5 in
        case executive = 7        // 7.25 x 10.5 in
        case a3 = 8               // 297 x 420 mm
        case a4 = 9               // 210 x 297 mm
        case a4Small = 10         // 210 x 297 mm
        case a5 = 11              // 148 x 210 mm
        case b4 = 12              // 250 x 353 mm
        case b5 = 13              // 176 x 250 mm
        case folio = 14           // 8.5 x 13 in
        case quarto = 15          // 215 x 275 mm
        case standard10 = 16      // 10 x 11 in
        case standard15 = 17      // 11 x 15 in
        case note = 18            // 8.5 x 11 in
        case number9Envelope = 19 // 3.875 x 8.875 in
        case number10Envelope = 20 // 4.125 x 9.5 in
        case number11Envelope = 21 // 4.5 x 10.375 in
        case number12Envelope = 22 // 4.75 x 11 in
        case number14Envelope = 23 // 5 x 11.5 in
        case cEnvelope = 24       // 162 x 229 mm
        case dEnvelope = 25       // 110 x 220 mm
        case eEnvelope = 26       // 110 x 330 mm
        case isoB4 = 33           // 250 x 353 mm
        case isoB5 = 34           // 176 x 250 mm
        case isoB6 = 35           // 125 x 176 mm
        case isoCEnvelope = 36    // 162 x 229 mm
        case isoDEnvelope = 37    // 110 x 220 mm
        case isoEEnvelope = 38    // 110 x 330 mm
        
        // Helper
        public static let `default`: PaperSize = .letter
    }
    
    /// Page orientation
    public enum Orientation: String, Sendable {
        case portrait
        case landscape
    }
    
    /// Page margins (in inches)
    public struct Margins: Sendable, Equatable {
        /// Left margin
        public let left: Double
        /// Right margin
        public let right: Double
        /// Top margin
        public let top: Double
        /// Bottom margin
        public let bottom: Double
        /// Header margin
        public let header: Double
        /// Footer margin
        public let footer: Double
        
        public init(left: Double = 0.75, right: Double = 0.75, 
                    top: Double = 1.0, bottom: Double = 1.0,
                    header: Double = 0.5, footer: Double = 0.5) {
            self.left = left
            self.right = right
            self.top = top
            self.bottom = bottom
            self.header = header
            self.footer = footer
        }
        
        /// Default margins
        public static let `default` = Margins()
    }
    
    /// Page orientation (portrait or landscape)
    public let orientation: Orientation
    /// Paper size
    public let paperSize: PaperSize
    /// Scale percentage (10-400, or nil for fit-to-page)
    public let scale: Int?
    /// Fit to pages wide/tall
    public let fitToPages: FitToPages?
    /// Print quality in DPI (default 300)
    public let printQuality: Int
    /// First page number (default 1)
    public let firstPageNumber: Int?
    /// Page margins
    public let margins: Margins
    
    public init(orientation: Orientation = .portrait,
                paperSize: PaperSize = .letter,
                scale: Int? = 100,
                fitToPages: FitToPages? = nil,
                printQuality: Int = 300,
                firstPageNumber: Int? = nil,
                margins: Margins = .default) {
        self.orientation = orientation
        self.paperSize = paperSize
        self.scale = scale
        self.fitToPages = fitToPages
        self.printQuality = printQuality
        self.firstPageNumber = firstPageNumber
        self.margins = margins
    }
    
    /// Create default page setup
    public static let `default` = PageSetup()
}

/// Print area configuration
public struct PrintArea: Sendable, Equatable {
    /// Print area ranges (e.g., "A1:D10")
    public let ranges: [String]
    
    public init(ranges: [String]) {
        self.ranges = ranges
    }
    
    public init(range: String) {
        self.ranges = [range]
    }
}

/// Print titles configuration
public struct PrintTitles: Sendable, Equatable {
    /// Rows to repeat at top of each page
    public let repeatRows: String?
    /// Columns to repeat at left of each page
    public let repeatColumns: String?
    
    public init(repeatRows: String? = nil, repeatColumns: String? = nil) {
        self.repeatRows = repeatRows
        self.repeatColumns = repeatColumns
    }
}
