import Foundation
import Testing
@testable import Cuneiform

@Suite struct CellFormulaTests {
    @Test func createFormula() {
        let formula = CellFormula("=SUM(A1:A10)")
        #expect(formula.formula == "=SUM(A1:A10)")
        #expect(!formula.isArrayFormula)
        #expect(formula.description == "=SUM(A1:A10)")
    }

    @Test func arrayFormula() {
        let formula = CellFormula("=TRANSPOSE(A1:B5)", isArrayFormula: true)
        #expect(formula.isArrayFormula)
        #expect(formula.description == "{==TRANSPOSE(A1:B5)}")
    }

    @Test func formulaEquatableHashable() {
        let f1 = CellFormula("=A1+B1")
        let f2 = CellFormula("=A1+B1")
        let f3 = CellFormula("=A1+B2")

        #expect(f1 == f2)
        #expect(f1 != f3)

        var set: Set<CellFormula> = []
        set.insert(f1)
        set.insert(f2)
        #expect(set.count == 1, "Duplicate formula should not be added to set")
        set.insert(f3)
        #expect(set.count == 2)
    }
}

@Suite struct WorksheetFormulaParsingTests {
    @Test func parseSimpleFormula() throws {
        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><f>=SUM(A2:A10)</f><v>0</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)

        guard let cell = ws.cell(at: "A1") else {
            #expect(false, "Cell A1 should exist")
            return
        }

        #expect(cell.formula != nil, "Cell should have formula")
        #expect(cell.formula?.formula == "=SUM(A2:A10)")
    }

    @Test func parseFormulaWithComplexExpression() throws {
        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><f>=IF(B1>0, B1*2, 0)</f><v>0</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)

        guard let cell = ws.cell(at: "A1") else {
            #expect(false)
            return
        }

        #expect(cell.formula?.formula == "=IF(B1>0, B1*2, 0)")
    }

    @Test func cellWithoutFormula() throws {
        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><v>42</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)

        guard let cell = ws.cell(at: "A1") else {
            #expect(false)
            return
        }

        #expect(cell.formula == nil, "Cell without <f> should have nil formula")
    }

    @Test func multipleFormulas() throws {
        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><f>=1+1</f><v>2</v></c>
                    <c r="B1"><f>=A1*2</f><v>4</v></c>
                    <c r="C1"><v>100</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let data = xml.data(using: .utf8)!
        let ws = try WorksheetParser.parse(data: data)

        let a1 = ws.cell(at: "A1")
        let b1 = ws.cell(at: "B1")
        let c1 = ws.cell(at: "C1")

        #expect(a1?.formula?.formula == "=1+1")
        #expect(b1?.formula?.formula == "=A1*2")
        #expect(c1?.formula == nil)
    }
}

@Suite struct SheetFormulaAccessTests {
    @Test func accessFormulaFromSheet() throws {
        let sharedStrings = SharedStrings.empty
        let styles = StylesInfo.empty

        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><f>=SUM(B1:B5)</f><v>15</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let wsData = try WorksheetParser.parse(data: xml.data(using: .utf8)!)
        let sheet = Sheet(data: wsData, sharedStrings: sharedStrings, styles: styles)

        let formula = sheet.formula(at: "A1")
        #expect(formula?.formula == "=SUM(B1:B5)")

        let formulaByCellRef: CellReference = "A1"
        let formulaByCellRefResult = sheet.formula(at: formulaByCellRef)
        #expect(formulaByCellRefResult?.formula == "=SUM(B1:B5)")
    }

    @Test func formulaNilForNonFormulaCell() throws {
        let sharedStrings = SharedStrings.empty
        let styles = StylesInfo.empty

        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><v>42</v></c>
                </row>
            </sheetData>
        </worksheet>
        """
        let wsData = try WorksheetParser.parse(data: xml.data(using: .utf8)!)
        let sheet = Sheet(data: wsData, sharedStrings: sharedStrings, styles: styles)

        let formula = sheet.formula(at: "A1")
        #expect(formula == nil)
    }

    @Test func formulaNilForMissingCell() throws {
        let sharedStrings = SharedStrings.empty
        let styles = StylesInfo.empty

        let xml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
            </sheetData>
        </worksheet>
        """
        let wsData = try WorksheetParser.parse(data: xml.data(using: .utf8)!)
        let sheet = Sheet(data: wsData, sharedStrings: sharedStrings, styles: styles)

        let formula = sheet.formula(at: "Z99")
        #expect(formula == nil)
    }
}
