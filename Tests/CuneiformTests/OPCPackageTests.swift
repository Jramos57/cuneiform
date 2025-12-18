import Testing
import Foundation
@testable import Cuneiform

@Suite struct OPCPackageTests {
    // Path to the sample xlsx file from the ISO spec materials
    static let sampleXlsxPath = "/Users/jonathan/Desktop/garden/vek/iEC 29500/ISO_IEC_29500-1_2016(en)_einsert/OfficeOpenXML-SpreadsheetMLStyles/PivotTableFormats.xlsx"

    @Test func openPackageFromFile() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        // Verify content types were parsed
        #expect(package.contentTypes.overrides.count > 0)

        // Verify root relationships were parsed
        #expect(package.rootRelationships.all.count > 0)
    }

    @Test func findMainDocument() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        let mainDoc = package.findMainDocument()
        #expect(mainDoc != nil)
        #expect(mainDoc?.type == .officeDocument)
        #expect(mainDoc?.target.contains("workbook") == true)
    }

    @Test func readWorkbookPart() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        let workbookData = try package.readPart(PartPath.workbook)
        let workbookString = String(data: workbookData, encoding: .utf8)

        #expect(workbookString != nil)
        #expect(workbookString?.contains("<workbook") == true)
        #expect(workbookString?.contains("<sheets>") == true)
    }

    @Test func partExists() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        #expect(package.partExists(PartPath.workbook) == true)
        #expect(package.partExists(PartPath("/nonexistent.xml")) == false)
    }

    @Test func contentTypeForPart() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        let workbookType = package.contentType(for: PartPath.workbook)
        #expect(workbookType == ContentType.workbook)
    }

    @Test func workbookRelationships() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        var package = try OPCPackage.open(url: url)

        let rels = try package.relationships(for: PartPath.workbook)

        // Should have worksheet relationships
        let worksheetRels = rels[.worksheet]
        #expect(worksheetRels.count > 0)

        // Should have styles relationship
        let styleRels = rels[.styles]
        #expect(styleRels.count == 1)
    }

    @Test func listParts() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        let parts = package.partPaths
        #expect(parts.count > 0)

        // Should contain workbook
        #expect(parts.contains { $0.value.contains("workbook.xml") })

        // Should contain worksheets
        #expect(parts.contains { $0.value.contains("sheet") })
    }

    @Test func missingPartThrows() throws {
        let url = URL(fileURLWithPath: Self.sampleXlsxPath)
        let package = try OPCPackage.open(url: url)

        #expect(throws: CuneiformError.self) {
            _ = try package.readPart(PartPath("/nonexistent/part.xml"))
        }
    }
}
