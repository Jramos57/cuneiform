import Testing
@testable import Cuneiform

struct ChartParserTests {
    @Test
    func parseColumnChart() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:colChart>
                        <c:ser/>
                        <c:ser/>
                    </c:colChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .columnChart)
        #expect(chart.seriesCount == 2)
    }

    @Test
    func parseLineChart() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:lineChart>
                        <c:ser/>
                    </c:lineChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .lineChart)
        #expect(chart.seriesCount == 1)
    }

    @Test
    func parsePieChart() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:pieChart>
                        <c:ser/>
                    </c:pieChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .pieChart)
        #expect(chart.seriesCount == 1)
    }

    @Test
    func parseAreaChart() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:areaChart>
                        <c:ser/>
                        <c:ser/>
                        <c:ser/>
                    </c:areaChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .areaChart)
        #expect(chart.seriesCount == 3)
    }

    @Test
    func parseBarChart() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:barChart>
                        <c:ser/>
                    </c:barChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .barChart)
        #expect(chart.seriesCount == 1)
    }

    @Test
    func unknownChartTypeDefaultsToUnknown() throws {
        let xml = """
            <?xml version="1.0"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:plotArea>
                    <c:unknownChart>
                        <c:ser/>
                    </c:unknownChart>
                </c:plotArea>
            </c:chartSpace>
            """

        let data = xml.data(using: .utf8)!
        let chart = try ChartParser.parse(data: data)

        #expect(chart.type == .unknown)
        #expect(chart.seriesCount == 1)
    }
}
