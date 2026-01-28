import Foundation
import Testing
@testable import Cuneiform

@Suite("Formula Functions - Extended (Phase 5)")
struct FormulaFunctionsExtendedTests {
    
    // MARK: - Helper
    
    func makeTestEvaluator(cells: [String: CellValue]) -> FormulaEvaluator {
        FormulaEvaluator { ref in
            let key = ref.description
            return cells[key]
        }
    }
    
    // MARK: - Logical Functions
    // Note: FormulaExpression doesn't have a boolean case - we use numbers (0=false, non-zero=true)
    
    @Test func evaluateAND_allTrue() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(1), .number(1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateAND_oneFalse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(0), .number(1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateAND_withNumbers() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(5), .number(-3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true)) // All non-zero = true
    }
    
    @Test func evaluateAND_withZero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AND", [.number(1), .number(0), .number(5)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false)) // Zero = false
    }
    
    @Test func evaluateOR_allFalse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("OR", [.number(0), .number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateOR_oneTrue() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("OR", [.number(0), .number(1), .number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateNOT_true() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.number(1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateNOT_false() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateNOT_number() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOT", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true)) // 0 = false, NOT false = true
    }
    
    // MARK: - String Functions
    
    @Test func evaluateLEN_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("Hello")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateLEN_empty() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    @Test func evaluateLEN_withSpaces() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEN", [.string("Hello World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(11))
    }
    
    @Test func evaluateUPPER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("UPPER", [.string("hello world")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("HELLO WORLD"))
    }
    
    @Test func evaluateLOWER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LOWER", [.string("HELLO WORLD")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("hello world"))
    }
    
    @Test func evaluateCONCAT_strings() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [.string("Hello"), .string(" "), .string("World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    @Test func evaluateCONCAT_mixed() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [.string("Value: "), .number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Value: 42.0"))
    }
    
    @Test func evaluateCONCAT_empty() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCAT", [])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string(""))
    }
    
    // MARK: - Math Functions
    
    @Test func evaluateROUND_noDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(3.7)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4))
    }
    
    @Test func evaluateROUND_twoDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(3.14159), .number(2)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3.14))
    }
    
    @Test func evaluateROUND_negativeDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ROUND", [.number(1234.5), .number(-2)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1200))
    }
    
    @Test func evaluateABS_positive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateABS_negative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(-42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateABS_zero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ABS", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    @Test func evaluateMEDIAN_oddCount() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(1), .number(3), .number(5), .number(7), .number(9)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateMEDIAN_evenCount() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(1), .number(2), .number(3), .number(4)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2.5))
    }
    
    @Test func evaluateMEDIAN_unsorted() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(5), .number(1), .number(9), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4)) // (3 + 5) / 2
    }
    
    @Test func evaluateMEDIAN_single() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MEDIAN", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    // MARK: - Integration Tests: Parser + New Functions
    
    @Test func parseAndEvaluateAND() throws {
        let cells: [String: CellValue] = [
            "A1": .boolean(true),
            "A2": .boolean(true),
            "A3": .boolean(false)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=AND(A1, A2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func parseAndEvaluateOR() throws {
        let cells: [String: CellValue] = [
            "A1": .boolean(false),
            "A2": .boolean(true),
            "A3": .boolean(false)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=OR(A1, A3, A2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func parseAndEvaluateUPPER() throws {
        let cells: [String: CellValue] = [
            "A1": .text("hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=UPPER(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("HELLO"))
    }
    
    @Test func parseAndEvaluateROUND() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3.14159)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ROUND(A1, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3.14))
    }
    
    @Test func parseAndEvaluateMEDIAN() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MEDIAN(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20))
    }
    
    // MARK: - Nested Function Tests
    
    @Test func nestedIF_AND() throws {
        let cells: [String: CellValue] = [
            "A1": .number(75),
            "B1": .number(80)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IF(AND(A1 >= 70, B1 >= 70), \"Pass\", \"Fail\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Pass"))
    }
    
    @Test func nestedROUND_SUM() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1.234),
            "A2": .number(2.567),
            "A3": .number(3.890)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ROUND(SUM(A1:A3), 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(7.7))
    }
    
    @Test func complexNested() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "B1": .number(20),
            "C1": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IF(OR(A1 > 15, B1 > 15), ROUND(AVERAGE(A1:C1), 0), 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20)) // B1 > 15, so ROUND(AVERAGE(10,20,30), 0) = 20
    }
    
    // MARK: - Date Functions (Tier 1)
    
    @Test func evaluateTODAY_returnsSerialNumber() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TODAY", [])
        let result = try evaluator.evaluate(expr)
        
        // TODAY should return a serial number > 44000 (dates after 2020)
        guard case .number(let serial) = result else {
            Issue.record("Expected number result")
            return
        }
        #expect(serial > 44000) // After 2020
        #expect(serial < 100000) // Reasonable upper bound
    }
    
    @Test func evaluateNOW_returnsSerialWithTime() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("NOW", [])
        let result = try evaluator.evaluate(expr)
        
        guard case .number(let serial) = result else {
            Issue.record("Expected number result")
            return
        }
        // NOW includes a fractional part for time
        #expect(serial > 44000)
        // The fractional part represents time of day
        let fractionalPart = serial - floor(serial)
        #expect(fractionalPart >= 0 && fractionalPart < 1)
    }
    
    @Test func evaluateDATE_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // DATE(2024, 1, 1) - verify via round-trip
        let expr = FormulaExpression.functionCall("DATE", [.number(2024), .number(1), .number(1)])
        let result = try evaluator.evaluate(expr)
        
        // Verify by extracting components
        guard case .number(let serial) = result else {
            Issue.record("Expected number result")
            return
        }
        let yearExpr = FormulaExpression.functionCall("YEAR", [.number(serial)])
        let monthExpr = FormulaExpression.functionCall("MONTH", [.number(serial)])
        let dayExpr = FormulaExpression.functionCall("DAY", [.number(serial)])
        
        #expect(try evaluator.evaluate(yearExpr) == .number(2024))
        #expect(try evaluator.evaluate(monthExpr) == .number(1))
        #expect(try evaluator.evaluate(dayExpr) == .number(1))
    }
    
    @Test func evaluateDATE_leapYear() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // DATE(2024, 2, 29) - 2024 is a leap year, verify via round-trip
        let expr = FormulaExpression.functionCall("DATE", [.number(2024), .number(2), .number(29)])
        let result = try evaluator.evaluate(expr)
        
        guard case .number(let serial) = result else {
            Issue.record("Expected number result")
            return
        }
        let yearExpr = FormulaExpression.functionCall("YEAR", [.number(serial)])
        let monthExpr = FormulaExpression.functionCall("MONTH", [.number(serial)])
        let dayExpr = FormulaExpression.functionCall("DAY", [.number(serial)])
        
        #expect(try evaluator.evaluate(yearExpr) == .number(2024))
        #expect(try evaluator.evaluate(monthExpr) == .number(2))
        #expect(try evaluator.evaluate(dayExpr) == .number(29))
    }
    
    @Test func evaluateYEAR_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Use DATE to create a known serial, then extract year
        let dateExpr = FormulaExpression.functionCall("DATE", [.number(2024), .number(7), .number(4)])
        guard case .number(let serial) = try evaluator.evaluate(dateExpr) else {
            Issue.record("Expected number result")
            return
        }
        
        let expr = FormulaExpression.functionCall("YEAR", [.number(serial)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2024))
    }
    
    @Test func evaluateMONTH_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Use DATE to create a known serial, then extract month
        let dateExpr = FormulaExpression.functionCall("DATE", [.number(2024), .number(2), .number(15)])
        guard case .number(let serial) = try evaluator.evaluate(dateExpr) else {
            Issue.record("Expected number result")
            return
        }
        
        let expr = FormulaExpression.functionCall("MONTH", [.number(serial)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))
    }
    
    @Test func evaluateDAY_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Use DATE to create a known serial, then extract day
        let dateExpr = FormulaExpression.functionCall("DATE", [.number(2024), .number(1), .number(15)])
        guard case .number(let serial) = try evaluator.evaluate(dateExpr) else {
            Issue.record("Expected number result")
            return
        }
        
        let expr = FormulaExpression.functionCall("DAY", [.number(serial)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(15))
    }
    
    @Test func evaluateDATE_YEAR_MONTH_DAY_roundTrip() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Create a date
        let dateExpr = FormulaExpression.functionCall("DATE", [.number(2024), .number(6), .number(15)])
        let dateResult = try evaluator.evaluate(dateExpr)
        
        guard case .number(let serial) = dateResult else {
            Issue.record("Expected number result")
            return
        }
        
        // Extract components
        let yearExpr = FormulaExpression.functionCall("YEAR", [.number(serial)])
        let monthExpr = FormulaExpression.functionCall("MONTH", [.number(serial)])
        let dayExpr = FormulaExpression.functionCall("DAY", [.number(serial)])
        
        #expect(try evaluator.evaluate(yearExpr) == .number(2024))
        #expect(try evaluator.evaluate(monthExpr) == .number(6))
        #expect(try evaluator.evaluate(dayExpr) == .number(15))
    }
    
    // MARK: - String Functions (Tier 1)
    
    @Test func evaluateLEFT_default() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEFT", [.string("Hello")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("H"))
    }
    
    @Test func evaluateLEFT_multiple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEFT", [.string("Hello World"), .number(5)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello"))
    }
    
    @Test func evaluateLEFT_exceedsLength() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("LEFT", [.string("Hi"), .number(10)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hi"))
    }
    
    @Test func evaluateRIGHT_default() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("RIGHT", [.string("Hello")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("o"))
    }
    
    @Test func evaluateRIGHT_multiple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("RIGHT", [.string("Hello World"), .number(5)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("World"))
    }
    
    @Test func evaluateMID_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MID", [.string("Hello World"), .number(7), .number(5)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("World"))
    }
    
    @Test func evaluateMID_startAtBeginning() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MID", [.string("Hello"), .number(1), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hel"))
    }
    
    @Test func evaluateMID_exceedsLength() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MID", [.string("Hello"), .number(3), .number(100)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("llo"))
    }
    
    @Test func evaluateTRIM_leadingTrailing() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TRIM", [.string("  Hello World  ")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    @Test func evaluateTRIM_multipleSpaces() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TRIM", [.string("Hello    World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    @Test func evaluateTRIM_noExtraSpaces() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TRIM", [.string("Hello World")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    // MARK: - Conditional Aggregate Functions (Tier 1)
    
    @Test func evaluateSUMIF_greaterThan() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10), "B1": .number(100),
            "A2": .number(20), "B2": .number(200),
            "A3": .number(30), "B3": .number(300)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sum B values where A > 15
        let expr = FormulaExpression.functionCall("SUMIF", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3)),
            .string(">15"),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(500)) // 200 + 300
    }
    
    @Test func evaluateSUMIF_equals() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"), "B1": .number(10),
            "A2": .text("Banana"), "B2": .number(20),
            "A3": .text("Apple"), "B3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sum B values where A = "Apple"
        let expr = FormulaExpression.functionCall("SUMIF", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3)),
            .string("Apple"),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(40)) // 10 + 30
    }
    
    @Test func evaluateSUMIF_wildcard() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"), "B1": .number(10),
            "A2": .text("Apricot"), "B2": .number(20),
            "A3": .text("Banana"), "B3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sum B values where A starts with "Ap"
        let expr = FormulaExpression.functionCall("SUMIF", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3)),
            .string("Ap*"),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(30)) // 10 + 20
    }
    
    @Test func evaluateCOUNTIF_greaterThan() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30),
            "A4": .number(15)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Count cells > 15
        let expr = FormulaExpression.functionCall("COUNTIF", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 4)),
            .string(">15")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2)) // 20 and 30
    }
    
    @Test func evaluateCOUNTIF_text() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Yes"),
            "A2": .text("No"),
            "A3": .text("Yes"),
            "A4": .text("Yes")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("COUNTIF", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 4)),
            .string("Yes")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))
    }
    
    @Test func evaluateIFERROR_noError() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("IFERROR", [.number(42), .string("Error occurred")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateIFERROR_withError() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Division by zero creates an error
        let divExpr = FormulaExpression.binaryOp(.divide, .number(1), .number(0))
        let expr = FormulaExpression.functionCall("IFERROR", [divExpr, .string("Cannot divide by zero")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Cannot divide by zero"))
    }
    
    @Test func evaluateIFERROR_withCellError() throws {
        let cells: [String: CellValue] = [
            "A1": .error("DIV/0")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("IFERROR", [
            .cellRef(CellReference(column: "A", row: 1)),
            .number(0)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    // MARK: - Additional Math Functions (Tier 1)
    
    @Test func evaluateINT_positive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("INT", [.number(8.9)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(8))
    }
    
    @Test func evaluateINT_negative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Excel's INT rounds toward negative infinity
        let expr = FormulaExpression.functionCall("INT", [.number(-8.1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(-9))
    }
    
    @Test func evaluateMOD_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MOD", [.number(10), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))
    }
    
    @Test func evaluateMOD_negative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Excel: MOD(-10, 3) = 2 (result has same sign as divisor)
        let expr = FormulaExpression.functionCall("MOD", [.number(-10), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))
    }
    
    @Test func evaluateMOD_divByZero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MOD", [.number(10), .number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("DIV/0"))
    }
    
    @Test func evaluateSQRT_positive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SQRT", [.number(16)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4))
    }
    
    @Test func evaluateSQRT_fractional() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SQRT", [.number(2)])
        let result = try evaluator.evaluate(expr)
        guard case .number(let value) = result else {
            Issue.record("Expected number result")
            return
        }
        #expect(abs(value - 1.41421356) < 0.0001)
    }
    
    @Test func evaluateSQRT_negative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SQRT", [.number(-4)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("NUM"))
    }
    
    @Test func evaluateSQRT_zero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SQRT", [.number(0)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    // MARK: - Parser + Tier 1 Integration Tests
    
    @Test func parseAndEvaluateLEFT() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello World")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LEFT(A1, 5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello"))
    }
    
    @Test func parseAndEvaluateMID() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello World")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MID(A1, 7, 5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("World"))
    }
    
    @Test func parseAndEvaluateSUMIF() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100), "B1": .text("East"),
            "A2": .number(200), "B2": .text("West"),
            "A3": .number(300), "B3": .text("East")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMIF(B1:B3, \"East\", A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(400)) // 100 + 300
    }
    
    @Test func parseAndEvaluateCOUNTIF() throws {
        let cells: [String: CellValue] = [
            "A1": .number(85),
            "A2": .number(72),
            "A3": .number(90),
            "A4": .number(68)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=COUNTIF(A1:A4, \">=80\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2)) // 85 and 90
    }
    
    @Test func parseAndEvaluateSQRT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(144)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SQRT(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(12))
    }
    
    @Test func nestedIFERROR_VLOOKUP() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"), "B1": .number(1.50),
            "A2": .text("Banana"), "B2": .number(0.75),
            "A3": .text("Cherry"), "B3": .number(3.00)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // VLOOKUP for non-existent item should return error, IFERROR catches it
        // Use 0 instead of FALSE (parser doesn't handle TRUE/FALSE literals)
        let parser = FormulaParser("=IFERROR(VLOOKUP(\"Orange\", A1:B3, 2, 0), 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    // MARK: - Tier 2: SUMIFS Tests
    
    @Test func evaluateSUMIFS_singleCriteria() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100), "B1": .text("East"),
            "A2": .number(200), "B2": .text("West"),
            "A3": .number(300), "B3": .text("East"),
            "A4": .number(400), "B4": .text("West")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("SUMIFS", [
            .range(CellReference("A1"), CellReference("A4")), // sum_range
            .range(CellReference("B1"), CellReference("B4")), // criteria_range1
            .string("East") // criteria1
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(400)) // 100 + 300
    }
    
    @Test func evaluateSUMIFS_multipleCriteria() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100), "B1": .text("East"), "C1": .number(2020),
            "A2": .number(200), "B2": .text("West"), "C2": .number(2020),
            "A3": .number(300), "B3": .text("East"), "C3": .number(2021),
            "A4": .number(400), "B4": .text("East"), "C4": .number(2020)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("SUMIFS", [
            .range(CellReference("A1"), CellReference("A4")), // sum_range
            .range(CellReference("B1"), CellReference("B4")), // criteria_range1
            .string("East"), // criteria1 - region
            .range(CellReference("C1"), CellReference("C4")), // criteria_range2
            .number(2020) // criteria2 - year
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(500)) // 100 + 400 (East AND 2020)
    }
    
    @Test func evaluateSUMIFS_numericCriteria() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10), "B1": .number(100),
            "A2": .number(20), "B2": .number(200),
            "A3": .number(30), "B3": .number(300),
            "A4": .number(15), "B4": .number(150)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("SUMIFS", [
            .range(CellReference("B1"), CellReference("B4")), // sum_range
            .range(CellReference("A1"), CellReference("A4")), // criteria_range
            .string(">15") // criteria - greater than 15
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(500)) // 200 + 300 (where A > 15)
    }
    
    // MARK: - Tier 2: COUNTIFS Tests
    
    @Test func evaluateCOUNTIFS_singleCriteria() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Yes"),
            "A2": .text("No"),
            "A3": .text("Yes"),
            "A4": .text("Yes")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("COUNTIFS", [
            .range(CellReference("A1"), CellReference("A4")),
            .string("Yes")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))
    }
    
    @Test func evaluateCOUNTIFS_multipleCriteria() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"), "B1": .number(10),
            "A2": .text("Banana"), "B2": .number(20),
            "A3": .text("Apple"), "B3": .number(30),
            "A4": .text("Apple"), "B4": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("COUNTIFS", [
            .range(CellReference("A1"), CellReference("A4")),
            .string("Apple"),
            .range(CellReference("B1"), CellReference("B4")),
            .string(">=10")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2)) // Apple with value >= 10: rows 1 and 3
    }
    
    // MARK: - Tier 2: AVERAGEIF Tests
    
    @Test func evaluateAVERAGEIF_text() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Pass"), "B1": .number(85),
            "A2": .text("Fail"), "B2": .number(45),
            "A3": .text("Pass"), "B3": .number(90),
            "A4": .text("Pass"), "B4": .number(75)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("AVERAGEIF", [
            .range(CellReference("A1"), CellReference("A4")),
            .string("Pass"),
            .range(CellReference("B1"), CellReference("B4"))
        ])
        let result = try evaluator.evaluate(expr)
        guard case .number(let avg) = result else {
            Issue.record("Expected number result")
            return
        }
        #expect(abs(avg - 83.333) < 0.01) // (85 + 90 + 75) / 3
    }
    
    @Test func evaluateAVERAGEIF_numeric() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100),
            "A2": .number(50),
            "A3": .number(150),
            "A4": .number(75)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("AVERAGEIF", [
            .range(CellReference("A1"), CellReference("A4")),
            .string(">60")
        ])
        let result = try evaluator.evaluate(expr)
        guard case .number(let avg) = result else {
            Issue.record("Expected number result")
            return
        }
        #expect(abs(avg - 108.333) < 0.01) // (100 + 150 + 75) / 3
    }
    
    @Test func evaluateAVERAGEIF_noMatches() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("AVERAGEIF", [
            .range(CellReference("A1"), CellReference("A2")),
            .string(">100")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("DIV/0"))
    }
    
    // MARK: - Tier 2: FIND Tests
    
    @Test func evaluateFIND_simple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("FIND", [
            .string("World"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(7))
    }
    
    @Test func evaluateFIND_caseSensitive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // FIND is case-sensitive, "world" should not match "World"
        let expr = FormulaExpression.functionCall("FIND", [
            .string("world"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("VALUE"))
    }
    
    @Test func evaluateFIND_withStartPos() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Find "o" starting from position 6
        let expr = FormulaExpression.functionCall("FIND", [
            .string("o"),
            .string("Hello World"),
            .number(6)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(8)) // The "o" in "World"
    }
    
    @Test func evaluateFIND_notFound() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("FIND", [
            .string("xyz"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("VALUE"))
    }
    
    // MARK: - Tier 2: SEARCH Tests
    
    @Test func evaluateSEARCH_caseInsensitive() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // SEARCH is case-insensitive
        let expr = FormulaExpression.functionCall("SEARCH", [
            .string("world"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(7))
    }
    
    @Test func evaluateSEARCH_withWildcard() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // * matches any characters
        let expr = FormulaExpression.functionCall("SEARCH", [
            .string("H*o"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))
    }
    
    @Test func evaluateSEARCH_questionMark() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // ? matches single character
        let expr = FormulaExpression.functionCall("SEARCH", [
            .string("H?llo"),
            .string("Hello World")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))
    }
    
    // MARK: - Tier 2: SUBSTITUTE Tests
    
    @Test func evaluateSUBSTITUTE_all() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SUBSTITUTE", [
            .string("Hello World World"),
            .string("World"),
            .string("Universe")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello Universe Universe"))
    }
    
    @Test func evaluateSUBSTITUTE_specific() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Replace only the 2nd occurrence
        let expr = FormulaExpression.functionCall("SUBSTITUTE", [
            .string("a-b-c-d-e"),
            .string("-"),
            .string("_"),
            .number(2)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("a-b_c-d-e"))
    }
    
    @Test func evaluateSUBSTITUTE_notFound() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SUBSTITUTE", [
            .string("Hello"),
            .string("xyz"),
            .string("abc")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello")) // No change
    }
    
    // MARK: - Tier 2: CONCATENATE Tests
    
    @Test func evaluateCONCATENATE_multiple() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCATENATE", [
            .string("Hello"),
            .string(" "),
            .string("World"),
            .string("!")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World!"))
    }
    
    @Test func evaluateCONCATENATE_withNumbers() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("CONCATENATE", [
            .string("Value: "),
            .number(42)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Value: 42.0"))
    }
    
    @Test func evaluateCONCATENATE_withCells() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello"),
            "B1": .text("World")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("CONCATENATE", [
            .cellRef(CellReference("A1")),
            .string(" "),
            .cellRef(CellReference("B1"))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    // MARK: - Tier 2: TEXT Tests
    
    @Test func evaluateTEXT_percentage() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TEXT", [
            .number(0.75),
            .string("0%")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("75%"))
    }
    
    @Test func evaluateTEXT_percentageDecimals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TEXT", [
            .number(0.1234),
            .string("0.00%")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("12.34%"))
    }
    
    @Test func evaluateTEXT_currency() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TEXT", [
            .number(1234.5),
            .string("$#,##0.00")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("$1,234.50"))
    }
    
    @Test func evaluateTEXT_date() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // First create a date serial for 2024-06-15
        let dateExpr = FormulaExpression.functionCall("DATE", [.number(2024), .number(6), .number(15)])
        guard case .number(let serial) = try evaluator.evaluate(dateExpr) else {
            Issue.record("Expected date serial")
            return
        }
        
        let expr = FormulaExpression.functionCall("TEXT", [
            .number(serial),
            .string("yyyy-mm-dd")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("2024-06-15"))
    }
    
    @Test func evaluateTEXT_number() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("TEXT", [
            .number(3.14159),
            .string("0.00")
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("3.14"))
    }
    
    // MARK: - Tier 2: Type Checking Functions Tests
    
    @Test func evaluateISBLANK_empty() throws {
        let cells: [String: CellValue] = [
            "A1": .empty
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("ISBLANK", [
            .cellRef(CellReference("A1"))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateISBLANK_notEmpty() throws {
        let cells: [String: CellValue] = [
            "A1": .number(0)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("ISBLANK", [
            .cellRef(CellReference("A1"))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateISBLANK_missingCell() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("ISBLANK", [
            .cellRef(CellReference("Z99"))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true)) // Missing cell = blank
    }
    
    @Test func evaluateISNUMBER_number() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ISNUMBER", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateISNUMBER_text() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ISNUMBER", [.string("42")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateISNUMBER_cell() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("ISNUMBER", [
            .cellRef(CellReference("A1"))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateISTEXT_text() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ISTEXT", [.string("Hello")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateISTEXT_number() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ISTEXT", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateISERROR_error() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        // Create an error via division by zero
        let divExpr = FormulaExpression.binaryOp(.divide, .number(1), .number(0))
        let expr = FormulaExpression.functionCall("ISERROR", [divExpr])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateISERROR_notError() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("ISERROR", [.number(42)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    // MARK: - Tier 2: Parser Integration Tests
    
    @Test func parseAndEvaluateSUMIFS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10), "B1": .text("X"),
            "A2": .number(20), "B2": .text("Y"),
            "A3": .number(30), "B3": .text("X")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMIFS(A1:A3, B1:B3, \"X\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(40)) // 10 + 30
    }
    
    @Test func parseAndEvaluateCOUNTIFS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(85), "B1": .text("Pass"),
            "A2": .number(45), "B2": .text("Fail"),
            "A3": .number(92), "B3": .text("Pass"),
            "A4": .number(78), "B4": .text("Pass")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=COUNTIFS(B1:B4, \"Pass\", A1:A4, \">=80\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2)) // 85 Pass and 92 Pass
    }
    
    @Test func parseAndEvaluateFIND() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello World")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=FIND(\"World\", A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(7))
    }
    
    @Test func parseAndEvaluateSUBSTITUTE() throws {
        let cells: [String: CellValue] = [
            "A1": .text("2024-01-15")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUBSTITUTE(A1, \"-\", \"/\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("2024/01/15"))
    }
    
    @Test func parseAndEvaluateTEXT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(0.85)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=TEXT(A1, \"0%\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("85%"))
    }
    
    @Test func parseAndEvaluateISNUMBER() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42),
            "B1": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ISNUMBER(A1)")
        let result1 = try evaluator.evaluate(try parser1.parse())
        #expect(result1 == .boolean(true))
        
        let parser2 = FormulaParser("=ISNUMBER(B1)")
        let result2 = try evaluator.evaluate(try parser2.parse())
        #expect(result2 == .boolean(false))
    }
    
    // MARK: - Complex Combined Tests
    
    @Test func nestedSUMIFS_IF() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100), "B1": .text("East"),
            "A2": .number(200), "B2": .text("West"),
            "A3": .number(300), "B3": .text("East")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // IF(SUMIFS(...) > 300, "High", "Low")
        let sumifs = FormulaExpression.functionCall("SUMIFS", [
            .range(CellReference("A1"), CellReference("A3")),
            .range(CellReference("B1"), CellReference("B3")),
            .string("East")
        ])
        let condition = FormulaExpression.binaryOp(.greaterThan, sumifs, .number(300))
        let expr = FormulaExpression.functionCall("IF", [condition, .string("High"), .string("Low")])
        
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("High")) // 100 + 300 = 400 > 300
    }
    
    @Test func textProcessingPipeline() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // SUBSTITUTE(UPPER("hello world"), " ", "-")
        let upperExpr = FormulaExpression.functionCall("UPPER", [.string("hello world")])
        let expr = FormulaExpression.functionCall("SUBSTITUTE", [
            upperExpr,
            .string(" "),
            .string("-")
        ])
        
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("HELLO-WORLD"))
    }
    
    @Test func conditionalTextFormatting() throws {
        let cells: [String: CellValue] = [
            "A1": .number(0.95)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // IF(A1 >= 0.9, CONCATENATE("Score: ", TEXT(A1, "0%")), "Low")
        let textExpr = FormulaExpression.functionCall("TEXT", [
            .cellRef(CellReference("A1")),
            .string("0%")
        ])
        let concatExpr = FormulaExpression.functionCall("CONCATENATE", [
            .string("Score: "),
            textExpr
        ])
        let condition = FormulaExpression.binaryOp(.greaterThanOrEqual, 
            .cellRef(CellReference("A1")), .number(0.9))
        let expr = FormulaExpression.functionCall("IF", [condition, concatExpr, .string("Low")])
        
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Score: 95%"))
    }
}
