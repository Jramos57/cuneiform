import Testing
@testable import Cuneiform

// MARK: - PartPath Tests

@Suite struct PartPathTests {
    @Test func partPathInitialization() {
        let path1 = PartPath("/xl/workbook.xml")
        #expect(path1.value == "/xl/workbook.xml")

        // Should add leading slash
        let path2 = PartPath("xl/workbook.xml")
        #expect(path2.value == "/xl/workbook.xml")
    }

    @Test func partPathComponents() {
        let path = PartPath("/xl/worksheets/sheet1.xml")

        #expect(path.fileName == "sheet1.xml")
        #expect(path.directory == "/xl/worksheets")
        #expect(path.zipEntryPath == "xl/worksheets/sheet1.xml")
    }

    @Test func partPathRelationships() {
        let workbook = PartPath("/xl/workbook.xml")
        #expect(workbook.relationshipsPath.value == "/xl/_rels/workbook.xml.rels")

        let sheet = PartPath("/xl/worksheets/sheet1.xml")
        #expect(sheet.relationshipsPath.value == "/xl/worksheets/_rels/sheet1.xml.rels")

        let root = PartPath("/[Content_Types].xml")
        #expect(root.relationshipsPath.value == "/_rels/[Content_Types].xml.rels")
    }

    @Test func wellKnownPaths() {
        #expect(PartPath.contentTypes.value == "/[Content_Types].xml")
        #expect(PartPath.workbook.value == "/xl/workbook.xml")
        #expect(PartPath.sharedStrings.value == "/xl/sharedStrings.xml")
    }
}

// MARK: - ContentType Tests

@Suite struct ContentTypeTests {
    @Test func wellKnownContentTypes() {
        #expect(ContentType.workbook.value == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        #expect(ContentType.worksheet.value == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    }
}

// MARK: - Relationship Tests

@Suite struct RelationshipTests {
    @Test func relationshipTargetResolution() {
        let rel = Relationship(
            id: "rId1",
            type: .worksheet,
            target: "worksheets/sheet1.xml"
        )

        let resolved = rel.resolveTarget(relativeTo: PartPath("/xl/workbook.xml"))
        #expect(resolved.value == "/xl/worksheets/sheet1.xml")
    }

    @Test func relationshipAbsoluteTarget() {
        let rel = Relationship(
            id: "rId1",
            type: .officeDocument,
            target: "/xl/workbook.xml"
        )

        let resolved = rel.resolveTarget(relativeTo: PartPath("/_rels/.rels"))
        #expect(resolved.value == "/xl/workbook.xml")
    }
}

// MARK: - Relationships Collection Tests

@Suite struct RelationshipsCollectionTests {
    @Test func relationshipsById() {
        let rels = Relationships([
            Relationship(id: "rId1", type: .worksheet, target: "worksheets/sheet1.xml"),
            Relationship(id: "rId2", type: .worksheet, target: "worksheets/sheet2.xml"),
            Relationship(id: "rId3", type: .styles, target: "styles.xml")
        ])

        #expect(rels["rId1"]?.target == "worksheets/sheet1.xml")
        #expect(rels["rId2"]?.target == "worksheets/sheet2.xml")
        #expect(rels["rId3"]?.target == "styles.xml")
        #expect(rels["rId99"] == nil)
    }

    @Test func relationshipsByType() {
        let rels = Relationships([
            Relationship(id: "rId1", type: .worksheet, target: "worksheets/sheet1.xml"),
            Relationship(id: "rId2", type: .worksheet, target: "worksheets/sheet2.xml"),
            Relationship(id: "rId3", type: .styles, target: "styles.xml")
        ])

        #expect(rels[.worksheet].count == 2)
        #expect(rels[.styles].count == 1)
        #expect(rels[.sharedStrings].count == 0)
    }
}

// MARK: - Error Tests

@Suite struct ErrorTests {
    @Test func errorDescriptions() {
        let error1 = CuneiformError.missingPart(path: "/xl/workbook.xml")
        #expect(error1.description.contains("workbook.xml"))

        let error2 = CuneiformError.invalidCellReference("XYZ")
        #expect(error2.description.contains("XYZ"))
    }
}

// MARK: - Version Tests

@Suite struct VersionTests {
    @Test func libraryVersion() {
        #expect(Cuneiform.version == "0.1.0")
        #expect(Cuneiform.ooXmlVersion.contains("29500"))
    }
}
