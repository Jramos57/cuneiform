import Testing
@testable import Cuneiform

@Suite struct TableParserTests {
    // MARK: - Basic Table Parsing
    
    @Test func parseSimpleTable() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1" name="Table1" displayName="Table1" ref="A1:C3"
               headerRowCount="1" totalsRowCount="0">
            <tableColumns count="3">
                <tableColumn id="1" name="Column1"/>
                <tableColumn id="2" name="Column2"/>
                <tableColumn id="3" name="Column3"/>
            </tableColumns>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 1)
        
        #expect(table.id == 1)
        #expect(table.name == "Table1")
        #expect(table.displayName == "Table1")
        #expect(table.ref == "A1:C3")
        #expect(table.headerRowCount == 1)
        #expect(table.totalsRowCount == 0)
        #expect(table.columns.count == 3)
        #expect(table.columns[0].name == "Column1")
        #expect(table.columns[1].name == "Column2")
        #expect(table.columns[2].name == "Column3")
    }
    
    @Test func parseTableWithTotalsRow() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="2" name="SalesData" displayName="Sales Data" ref="A1:D100"
               headerRowCount="1" totalsRowCount="1">
            <tableColumns count="4">
                <tableColumn id="1" name="Date"/>
                <tableColumn id="2" name="Product"/>
                <tableColumn id="3" name="Quantity"/>
                <tableColumn id="4" name="Amount" totalsRowFunction="sum"/>
            </tableColumns>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 2)
        
        #expect(table.id == 2)
        #expect(table.name == "SalesData")
        #expect(table.displayName == "Sales Data")
        #expect(table.totalsRowCount == 1)
        #expect(table.columns.count == 4)
        #expect(table.columns[3].totalsRowFunction == "sum")
    }
    
    @Test func parseTableWithAutoFilter() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1" name="FilteredTable" displayName="Filtered Table" ref="A1:C50"
               headerRowCount="1" totalsRowCount="0">
            <tableColumns count="3">
                <tableColumn id="1" name="Name"/>
                <tableColumn id="2" name="Status"/>
                <tableColumn id="3" name="Score"/>
            </tableColumns>
            <autoFilter ref="A1:C50"/>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 1)
        
        #expect(table.autoFilter != nil)
        #expect(table.autoFilter?.ref == "A1:C50")
    }
    
    @Test func parseTableWithStyle() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1" name="StyledTable" displayName="Styled Table" ref="A1:B10"
               headerRowCount="1" totalsRowCount="0">
            <tableColumns count="2">
                <tableColumn id="1" name="Name"/>
                <tableColumn id="2" name="Value"/>
            </tableColumns>
            <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" 
                           showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 1)
        
        #expect(table.tableStyleInfo != nil)
        #expect(table.tableStyleInfo?.name == "TableStyleMedium2")
        #expect(table.tableStyleInfo?.showFirstColumn == false)
        #expect(table.tableStyleInfo?.showRowStripes == true)
    }
    
    @Test func parseTableMultipleTotalsRowFunctions() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="3" name="Stats" displayName="Statistics" ref="A1:D20"
               headerRowCount="1" totalsRowCount="1">
            <tableColumns count="4">
                <tableColumn id="1" name="Item"/>
                <tableColumn id="2" name="Count" totalsRowFunction="count"/>
                <tableColumn id="3" name="Average" totalsRowFunction="average"/>
                <tableColumn id="4" name="Total" totalsRowFunction="sum"/>
            </tableColumns>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 3)
        
        #expect(table.columns.count == 4)
        #expect(table.columns[0].totalsRowFunction == nil)
        #expect(table.columns[1].totalsRowFunction == "count")
        #expect(table.columns[2].totalsRowFunction == "average")
        #expect(table.columns[3].totalsRowFunction == "sum")
    }
    
    // MARK: - Edge Cases
    
    @Test func parseTableWithoutColumns() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1" name="EmptyTable" displayName="Empty Table" ref="A1:A1"
               headerRowCount="1" totalsRowCount="0">
            <tableColumns count="0"/>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 1)
        
        #expect(table.columns.isEmpty)
    }
    
    @Test func parseTableWithDefaultAttributes() throws {
        let xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1" name="Minimal" displayName="Minimal" ref="A1:B2">
            <tableColumns count="2">
                <tableColumn id="1" name="A"/>
                <tableColumn id="2" name="B"/>
            </tableColumns>
        </table>
        """
        
        let data = xml.data(using: .utf8)!
        let table = try TableParser.parse(data: data, id: 1)
        
        #expect(table.headerRowCount == 1)  // Default
        #expect(table.totalsRowCount == 0)  // Default
        #expect(table.autoFilter == nil)
        #expect(table.tableStyleInfo == nil)
    }
    
    // MARK: - Equatable Conformance
    
    @Test func tableColumnsEquatable() {
        let col1 = TableColumn(id: 1, name: "Name")
        let col2 = TableColumn(id: 1, name: "Name")
        let col3 = TableColumn(id: 1, name: "Different")
        
        #expect(col1 == col2)
        #expect(col1 != col3)
    }
    
    @Test func tableDataEquatable() {
        let table1 = TableData(
            id: 1, displayName: "T1", name: "Table1", ref: "A1:B10",
            headerRowCount: 1, totalsRowCount: 0,
            columns: [TableColumn(id: 1, name: "A")]
        )
        
        let table2 = TableData(
            id: 1, displayName: "T1", name: "Table1", ref: "A1:B10",
            headerRowCount: 1, totalsRowCount: 0,
            columns: [TableColumn(id: 1, name: "A")]
        )
        
        let table3 = TableData(
            id: 2, displayName: "T2", name: "Table2", ref: "A1:B10",
            headerRowCount: 1, totalsRowCount: 0,
            columns: [TableColumn(id: 1, name: "A")]
        )
        
        #expect(table1 == table2)
        #expect(table1 != table3)
    }
    
    // MARK: - Sendable Conformance
    
    @Test func tableDataSendable() async {
        let table = TableData(
            id: 1, displayName: "Async", name: "AsyncTable", ref: "A1:C5",
            columns: [TableColumn(id: 1, name: "Col")]
        )
        
        let result = await Task {
            table.name
        }.value
        
        #expect(result == "AsyncTable")
    }
}
