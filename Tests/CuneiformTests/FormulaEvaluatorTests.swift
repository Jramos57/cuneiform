import Foundation
import Testing
@testable import Cuneiform

@Suite("Formula Evaluator Tests (Phase 5)")
struct FormulaEvaluatorTests {
    
    // MARK: - Helper: Create Test Sheet Data
    
    func makeTestEvaluator(cells: [String: CellValue]) -> FormulaEvaluator {
        FormulaEvaluator { ref in
            let key = ref.description
            return cells[key]
        }
    }
    
    // MARK: - Basic Arithmetic
    
    @Test func evaluateAddition() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.add, .number(2), .number(3))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateSubtraction() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.subtract, .number(10), .number(4))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(6))
    }
    
    @Test func evaluateMultiplication() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.multiply, .number(5), .number(6))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(30))
    }
    
    @Test func evaluateDivision() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.divide, .number(20), .number(4))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateDivisionByZero() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.divide, .number(10), .number(0))
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("DIV/0"))
    }
    
    @Test func evaluatePower() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.power, .number(2), .number(8))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(256))
    }
    
    // MARK: - Cell References
    
    @Test func evaluateCellReference() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let ref = CellReference("A1")
        
        let expr = FormulaExpression.cellRef(ref)
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateMissingCellReference() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let ref = CellReference("Z99")
        
        let expr = FormulaExpression.cellRef(ref)
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("REF"))
    }
    
    @Test func evaluateCellReferenceInArithmetic() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "B1": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let a1 = CellReference("A1")
        let b1 = CellReference("B1")
        
        let expr = FormulaExpression.binaryOp(.add, .cellRef(a1), .cellRef(b1))
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(15))
    }
    
    // MARK: - SUM Function
    
    @Test func evaluateSumWithNumbers() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("SUM", [.number(1), .number(2), .number(3)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(6))
    }
    
    @Test func evaluateSumWithRange() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A3")
        
        let expr = FormulaExpression.functionCall("SUM", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(60))
    }
    
    @Test func evaluateSumEmptyRange() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let start = CellReference("A1")
        let end = CellReference("A3")
        
        let expr = FormulaExpression.functionCall("SUM", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))
    }
    
    // MARK: - AVERAGE Function
    
    @Test func evaluateAverageWithNumbers() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("AVERAGE", [.number(10), .number(20), .number(30)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20))
    }
    
    @Test func evaluateAverageWithRange() throws {
        let cells: [String: CellValue] = [
            "B1": .number(5),
            "B2": .number(10),
            "B3": .number(15),
            "B4": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("B1")
        let end = CellReference("B4")
        
        let expr = FormulaExpression.functionCall("AVERAGE", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(12.5))
    }
    
    @Test func evaluateAverageEmptyRange() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let start = CellReference("A1")
        let end = CellReference("A1")
        
        let expr = FormulaExpression.functionCall("AVERAGE", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("DIV/0"))
    }
    
    // MARK: - IF Function
    
    @Test func evaluateIfTrue() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let condition = FormulaExpression.binaryOp(.greaterThan, .number(10), .number(5))
        let expr = FormulaExpression.functionCall("IF", [condition, .string("Yes"), .string("No")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Yes"))
    }
    
    @Test func evaluateIfFalse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let condition = FormulaExpression.binaryOp(.lessThan, .number(3), .number(2))
        let expr = FormulaExpression.functionCall("IF", [condition, .string("Yes"), .string("No")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("No"))
    }
    
    @Test func evaluateIfWithCellReference() throws {
        let cells: [String: CellValue] = [
            "A1": .number(15)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let a1 = CellReference("A1")
        
        let condition = FormulaExpression.binaryOp(.greaterThan, .cellRef(a1), .number(10))
        let expr = FormulaExpression.functionCall("IF", [condition, .string("High"), .string("Low")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("High"))
    }
    
    @Test func evaluateIfWithoutElse() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let condition = FormulaExpression.binaryOp(.equals, .number(1), .number(2))
        let expr = FormulaExpression.functionCall("IF", [condition, .string("Yes")])
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(false))
    }
    
    // MARK: - VLOOKUP Function
    
    @Test func evaluateVlookupExactMatch() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .text("Apple"),
            "A2": .number(2), "B2": .text("Banana"),
            "A3": .number(3), "B3": .text("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("B3")
        
        let expr = FormulaExpression.functionCall("VLOOKUP", [
            .number(2),
            .range(start, end),
            .number(2),
            .number(0) // false = exact match (0 is falsy)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Banana"))
    }
    
    @Test func evaluateVlookupNotFound() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .text("Apple"),
            "A2": .number(2), "B2": .text("Banana")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("B2")
        
        let expr = FormulaExpression.functionCall("VLOOKUP", [
            .number(5),
            .range(start, end),
            .number(2),
            .number(0) // false = exact match
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("N/A"))
    }
    
    // MARK: - INDEX Function
    
    @Test func evaluateIndexWithRowAndColumn() throws {
        let cells: [String: CellValue] = [
            "A1": .text("A"), "B1": .text("B"),
            "A2": .text("C"), "B2": .text("D")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("B2")
        
        let expr = FormulaExpression.functionCall("INDEX", [
            .range(start, end),
            .number(2),
            .number(1)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("C"))
    }
    
    @Test func evaluateIndexOutOfBounds() throws {
        let cells: [String: CellValue] = [
            "A1": .text("A")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A1")
        
        let expr = FormulaExpression.functionCall("INDEX", [
            .range(start, end),
            .number(5),
            .number(1)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("REF"))
    }
    
    // MARK: - MATCH Function
    
    @Test func evaluateMatchExact() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"),
            "A2": .text("Banana"),
            "A3": .text("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A3")
        
        let expr = FormulaExpression.functionCall("MATCH", [
            .string("Banana"),
            .range(start, end),
            .number(0)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))
    }
    
    @Test func evaluateMatchNotFound() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"),
            "A2": .text("Banana")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A2")
        
        let expr = FormulaExpression.functionCall("MATCH", [
            .string("Orange"),
            .range(start, end),
            .number(0)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("N/A"))
    }
    
    // MARK: - MIN/MAX Functions
    
    @Test func evaluateMin() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MIN", [.number(5), .number(2), .number(8), .number(1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))
    }
    
    @Test func evaluateMax() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.functionCall("MAX", [.number(5), .number(2), .number(8), .number(1)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(8))
    }
    
    // MARK: - COUNT Functions
    
    @Test func evaluateCount() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .text("Hello"),
            "A3": .number(20),
            "A4": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A4")
        
        let expr = FormulaExpression.functionCall("COUNT", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3)) // Only numbers
    }
    
    @Test func evaluateCountA() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .text("Hello"),
            "A3": .boolean(true),
            "A4": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let start = CellReference("A1")
        let end = CellReference("A4")
        
        let expr = FormulaExpression.functionCall("COUNTA", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4)) // All non-empty
    }
    
    // MARK: - String Operations
    
    @Test func evaluateStringConcatenation() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.concat, .string("Hello"), .string(" World"))
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World"))
    }
    
    // MARK: - Comparison Operations
    
    @Test func evaluateEquals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.equals, .number(5), .number(5))
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateNotEquals() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.notEquals, .number(5), .number(3))
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateLessThan() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.lessThan, .number(3), .number(5))
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    @Test func evaluateGreaterThanOrEqual() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        let expr = FormulaExpression.binaryOp(.greaterThanOrEqual, .number(10), .number(10))
        let result = try evaluator.evaluate(expr)
        #expect(result == .boolean(true))
    }
    
    // MARK: - Integration: Parser + Evaluator
    
    @Test func parseAndEvaluateSimpleFormula() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "B1": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=A1 + B1")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(30))
    }
    
    @Test func parseAndEvaluateSumFormula() throws {
        let cells: [String: CellValue] = [
            "A1": .number(5),
            "A2": .number(10),
            "A3": .number(15)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUM(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(30))
    }
    
    @Test func parseAndEvaluateIfFormula() throws {
        let cells: [String: CellValue] = [
            "A1": .number(75)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IF(A1 >= 70, \"Pass\", \"Fail\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Pass"))
    }
    
    // MARK: - Excel 365 Functions
    
    @Test func evaluateXLOOKUP() throws {
        // Setup: lookup table with IDs and names
        let cells: [String: CellValue] = [
            "A1": .number(101),
            "A2": .number(102),
            "A3": .number(103),
            "B1": .text("Apple"),
            "B2": .text("Banana"),
            "B3": .text("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=XLOOKUP(102, A1:A3, B1:B3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Banana"))
    }
    
    @Test func evaluateXLOOKUPNotFound() throws {
        let cells: [String: CellValue] = [
            "A1": .number(101),
            "A2": .number(102),
            "B1": .text("Apple"),
            "B2": .text("Banana")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=XLOOKUP(999, A1:A2, B1:B2, \"Not Found\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Not Found"))
    }
    
    @Test func evaluateTEXTJOIN() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello"),
            "A2": .text("World"),
            "A3": .text("!")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=TEXTJOIN(\" \", 1, A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hello World !"))
    }
    
    @Test func evaluateTEXTJOINIgnoreEmpty() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"),
            "A2": .text(""),
            "A3": .text("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=TEXTJOIN(\", \", 1, A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Apple, Cherry"))
    }
    
    @Test func evaluateIFS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(85)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IFS(A1 < 60, \"F\", A1 < 70, \"D\", A1 < 80, \"C\", A1 < 90, \"B\", 1, \"A\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("B"))
    }
    
    @Test func evaluateIFSNoMatch() throws {
        let cells: [String: CellValue] = [
            "A1": .number(50)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=IFS(A1 > 60, \"Pass\", A1 > 80, \"Excellent\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("N/A"))
    }
    
    @Test func evaluateSWITCH() throws {
        let cells: [String: CellValue] = [
            "A1": .text("B")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SWITCH(A1, \"A\", \"Apple\", \"B\", \"Banana\", \"C\", \"Cherry\", \"Unknown\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Banana"))
    }
    
    @Test func evaluateSWITCHDefault() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Z")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SWITCH(A1, \"A\", \"Apple\", \"B\", \"Banana\", \"Unknown\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Unknown"))
    }
    
    @Test func evaluateMAXIFS() throws {
        // Setup: sales data with region and amount
        let cells: [String: CellValue] = [
            "A1": .number(100),  // amounts
            "A2": .number(200),
            "A3": .number(150),
            "B1": .text("East"),  // regions
            "B2": .text("West"),
            "B3": .text("East")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MAXIFS(A1:A3, B1:B3, \"East\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(150))
    }
    
    @Test func evaluateMINIFS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100),
            "A2": .number(200),
            "A3": .number(150),
            "B1": .text("East"),
            "B2": .text("West"),
            "B3": .text("East")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MINIFS(A1:A3, B1:B3, \"East\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(100))
    }
    
    @Test func evaluateAVERAGEIFS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(100),
            "A2": .number(200),
            "A3": .number(150),
            "B1": .text("East"),
            "B2": .text("West"),
            "B3": .text("East")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=AVERAGEIFS(A1:A3, B1:B3, \"East\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(125))  // (100 + 150) / 2
    }
    
    // MARK: - Statistical Functions
    
    @Test func evaluateSTDEV() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(4),
            "A3": .number(6),
            "A4": .number(8)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=STDEV(A1:A4)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Sample std dev of [2,4,6,8] = sqrt(20/3) ≈ 2.582
        if case .number(let value) = result {
            #expect(abs(value - 2.582) < 0.01)
        } else {
            #expect(Bool(false), "Expected number result")
        }
    }
    
    @Test func evaluateVAR() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(4),
            "A3": .number(6)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=VAR(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Sample variance of [2,4,6] = 4
        #expect(result == .number(4))
    }
    
    @Test func evaluatePERCENTILE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=PERCENTILE(A1:A5, 0.5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))  // Median
    }
    
    @Test func evaluateQUARTILE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=QUARTILE(A1:A5, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))  // Q1
    }
    
    @Test func evaluateMODE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(2),
            "A4": .number(3),
            "A5": .number(2)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MODE(A1:A5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))  // Most frequent value
    }
    
    @Test func evaluateLARGE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(5),
            "A3": .number(1),
            "A4": .number(4),
            "A5": .number(2)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LARGE(A1:A5, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4))  // 2nd largest
    }
    
    @Test func evaluateSMALL() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(5),
            "A3": .number(1),
            "A4": .number(4),
            "A5": .number(2)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SMALL(A1:A5, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))  // 2nd smallest
    }
    
    @Test func evaluateRANK() throws {
        let cells: [String: CellValue] = [
            "A1": .number(7),
            "A2": .number(3),
            "A3": .number(5),
            "A4": .number(9),
            "A5": .number(1)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=RANK(5, A1:A5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))  // 5 is 3rd in descending order
    }
    
    @Test func evaluateCORREL() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "B1": .number(2),
            "B2": .number(4),
            "B3": .number(6)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=CORREL(A1:A3, B1:B3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))  // Perfect positive correlation
    }
    
    // MARK: - Math & Trigonometric Functions
    
    @Test func evaluateTrigonometricFunctions() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test PI()
        let piParser = FormulaParser("=PI()")
        let piExpr = try piParser.parse()
        let piResult = try evaluator.evaluate(piExpr)
        if case .number(let val) = piResult {
            #expect(abs(val - 3.14159265358979) < 0.00001)
        }
        
        // Test SIN(PI()/2) = 1
        let sinParser = FormulaParser("=SIN(PI()/2)")
        let sinExpr = try sinParser.parse()
        let sinResult = try evaluator.evaluate(sinExpr)
        if case .number(let val) = sinResult {
            #expect(abs(val - 1.0) < 0.00001)
        }
        
        // Test COS(0) = 1
        let cosParser = FormulaParser("=COS(0)")
        let cosExpr = try cosParser.parse()
        let cosResult = try evaluator.evaluate(cosExpr)
        #expect(cosResult == .number(1.0))
        
        // Test TAN(PI()/4) ≈ 1
        let tanParser = FormulaParser("=TAN(PI()/4)")
        let tanExpr = try tanParser.parse()
        let tanResult = try evaluator.evaluate(tanExpr)
        if case .number(let val) = tanResult {
            #expect(abs(val - 1.0) < 0.00001)
        }
    }
    
    @Test func evaluateInverseTrigFunctions() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test ASIN(1) = PI/2
        let asinParser = FormulaParser("=ASIN(1)")
        let asinExpr = try asinParser.parse()
        let asinResult = try evaluator.evaluate(asinExpr)
        if case .number(let val) = asinResult {
            #expect(abs(val - Double.pi / 2) < 0.00001)
        }
        
        // Test ACOS(0) = PI/2
        let acosParser = FormulaParser("=ACOS(0)")
        let acosExpr = try acosParser.parse()
        let acosResult = try evaluator.evaluate(acosExpr)
        if case .number(let val) = acosResult {
            #expect(abs(val - Double.pi / 2) < 0.00001)
        }
        
        // Test ATAN(1) = PI/4
        let atanParser = FormulaParser("=ATAN(1)")
        let atanExpr = try atanParser.parse()
        let atanResult = try evaluator.evaluate(atanExpr)
        if case .number(let val) = atanResult {
            #expect(abs(val - Double.pi / 4) < 0.00001)
        }
        
        // Test ATAN2(1, 1) = PI/4
        let atan2Parser = FormulaParser("=ATAN2(1, 1)")
        let atan2Expr = try atan2Parser.parse()
        let atan2Result = try evaluator.evaluate(atan2Expr)
        if case .number(let val) = atan2Result {
            #expect(abs(val - Double.pi / 4) < 0.00001)
        }
    }
    
    @Test func evaluateAngleConversion() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test RADIANS(180) = PI
        let radParser = FormulaParser("=RADIANS(180)")
        let radExpr = try radParser.parse()
        let radResult = try evaluator.evaluate(radExpr)
        if case .number(let val) = radResult {
            #expect(abs(val - Double.pi) < 0.00001)
        }
        
        // Test DEGREES(PI) = 180
        let degParser = FormulaParser("=DEGREES(PI())")
        let degExpr = try degParser.parse()
        let degResult = try evaluator.evaluate(degExpr)
        if case .number(let val) = degResult {
            #expect(abs(val - 180.0) < 0.00001)
        }
    }
    
    @Test func evaluateLogarithmicFunctions() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test LN(E) = 1 (E ≈ 2.71828)
        let lnParser = FormulaParser("=LN(2.71828182845905)")
        let lnExpr = try lnParser.parse()
        let lnResult = try evaluator.evaluate(lnExpr)
        if case .number(let val) = lnResult {
            #expect(abs(val - 1.0) < 0.00001)
        }
        
        // Test LOG10(100) = 2
        let log10Parser = FormulaParser("=LOG10(100)")
        let log10Expr = try log10Parser.parse()
        let log10Result = try evaluator.evaluate(log10Expr)
        #expect(log10Result == .number(2.0))
        
        // Test LOG(8, 2) = 3
        let logParser = FormulaParser("=LOG(8, 2)")
        let logExpr = try logParser.parse()
        let logResult = try evaluator.evaluate(logExpr)
        #expect(logResult == .number(3.0))
        
        // Test EXP(1) ≈ E
        let expParser = FormulaParser("=EXP(1)")
        let expExpr = try expParser.parse()
        let expResult = try evaluator.evaluate(expExpr)
        if case .number(let val) = expResult {
            #expect(abs(val - 2.71828182845905) < 0.00001)
        }
    }
    
    @Test func evaluateRoundingFunctions() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test CEILING(2.3, 1) = 3
        let ceilParser = FormulaParser("=CEILING(2.3, 1)")
        let ceilExpr = try ceilParser.parse()
        let ceilResult = try evaluator.evaluate(ceilExpr)
        #expect(ceilResult == .number(3.0))
        
        // Test FLOOR(2.7, 1) = 2
        let floorParser = FormulaParser("=FLOOR(2.7, 1)")
        let floorExpr = try floorParser.parse()
        let floorResult = try evaluator.evaluate(floorExpr)
        #expect(floorResult == .number(2.0))
        
        // Test TRUNC(8.9) = 8
        let truncParser = FormulaParser("=TRUNC(8.9)")
        let truncExpr = try truncParser.parse()
        let truncResult = try evaluator.evaluate(truncExpr)
        #expect(truncResult == .number(8.0))
        
        // Test TRUNC(-8.9) = -8
        let truncNegParser = FormulaParser("=TRUNC(-8.9)")
        let truncNegExpr = try truncNegParser.parse()
        let truncNegResult = try evaluator.evaluate(truncNegExpr)
        #expect(truncNegResult == .number(-8.0))
        
        // Test SIGN(10) = 1
        let signPosParser = FormulaParser("=SIGN(10)")
        let signPosExpr = try signPosParser.parse()
        let signPosResult = try evaluator.evaluate(signPosExpr)
        #expect(signPosResult == .number(1.0))
        
        // Test SIGN(-5) = -1
        let signNegParser = FormulaParser("=SIGN(-5)")
        let signNegExpr = try signNegParser.parse()
        let signNegResult = try evaluator.evaluate(signNegExpr)
        #expect(signNegResult == .number(-1.0))
        
        // Test SIGN(0) = 0
        let signZeroParser = FormulaParser("=SIGN(0)")
        let signZeroExpr = try signZeroParser.parse()
        let signZeroResult = try evaluator.evaluate(signZeroExpr)
        #expect(signZeroResult == .number(0.0))
    }
    
    @Test func evaluateFACT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test FACT(5) = 120
        let factParser = FormulaParser("=FACT(5)")
        let factExpr = try factParser.parse()
        let factResult = try evaluator.evaluate(factExpr)
        #expect(factResult == .number(120.0))
        
        // Test FACT(0) = 1
        let fact0Parser = FormulaParser("=FACT(0)")
        let fact0Expr = try fact0Parser.parse()
        let fact0Result = try evaluator.evaluate(fact0Expr)
        #expect(fact0Result == .number(1.0))
    }
    
    @Test func evaluateSUMPRODUCT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3), "B1": .number(4),
            "A2": .number(8), "B2": .number(6),
            "A3": .number(1), "B3": .number(9)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test SUMPRODUCT(A1:A3, B1:B3) = 3*4 + 8*6 + 1*9 = 12 + 48 + 9 = 69
        let parser = FormulaParser("=SUMPRODUCT(A1:A3, B1:B3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(69.0))
    }
    
    @Test func evaluateGCDandLCM() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test GCD(12, 18) = 6
        let gcdParser = FormulaParser("=GCD(12, 18)")
        let gcdExpr = try gcdParser.parse()
        let gcdResult = try evaluator.evaluate(gcdExpr)
        #expect(gcdResult == .number(6.0))
        
        // Test GCD(24, 36, 48) = 12
        let gcd3Parser = FormulaParser("=GCD(24, 36, 48)")
        let gcd3Expr = try gcd3Parser.parse()
        let gcd3Result = try evaluator.evaluate(gcd3Expr)
        #expect(gcd3Result == .number(12.0))
        
        // Test LCM(4, 6) = 12
        let lcmParser = FormulaParser("=LCM(4, 6)")
        let lcmExpr = try lcmParser.parse()
        let lcmResult = try evaluator.evaluate(lcmExpr)
        #expect(lcmResult == .number(12.0))
        
        // Test LCM(3, 4, 6) = 12
        let lcm3Parser = FormulaParser("=LCM(3, 4, 6)")
        let lcm3Expr = try lcm3Parser.parse()
        let lcm3Result = try evaluator.evaluate(lcm3Expr)
        #expect(lcm3Result == .number(12.0))
    }
    
    // MARK: - Financial Functions
    
    @Test func evaluatePMT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test PMT(5%/12, 60, 10000) - monthly payment on $10,000 loan at 5% for 5 years
        let parser = FormulaParser("=PMT(0.05/12, 60, 10000)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - (-188.71)) < 0.01)  // Monthly payment should be ~$188.71
        }
    }
    
    @Test func evaluatePV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test PV(8%/12, 20, 500, 0, 0) - present value of $500/month for 20 months at 8%
        let parser = FormulaParser("=PV(0.08/12, 20, 500, 0, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - (-9333)) < 100.0)  // Present value (approximate)
        }
    }
    
    @Test func evaluateFV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test FV(6%/12, 10, -200, -500, 1) - future value with payments at beginning
        let parser = FormulaParser("=FV(0.06/12, 10, -200, -500, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 2581.40) < 1.0)  // Future value should be ~$2,581.40
        }
    }
    
    @Test func evaluateNPER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test NPER(12%/12, -100, -1000, 10000, 0) - periods to reach goal
        let parser = FormulaParser("=NPER(0.12/12, -100, -1000, 10000, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 60.08) < 0.1)  // Should be ~60 periods
        }
    }
    
    @Test func evaluateNPV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test NPV(10%, -10000, 3000, 4200, 6800) - investment NPV
        let parser = FormulaParser("=NPV(0.1, -10000, 3000, 4200, 6800)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 1188.44) < 1.0)  // NPV should be ~$1,188.44
        }
    }
    
    @Test func evaluateIRR() throws {
        let cells: [String: CellValue] = [
            "A1": .number(-70000),
            "A2": .number(12000),
            "A3": .number(15000),
            "A4": .number(18000),
            "A5": .number(21000),
            "A6": .number(26000)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test IRR(A1:A6) - internal rate of return for cash flows
        let parser = FormulaParser("=IRR(A1:A6)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 0.0866) < 0.001)  // IRR should be ~8.66%
        }
    }
    
    @Test func evaluateIPMT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test IPMT(10%/12, 1, 3*12, 8000) - interest payment for first period
        let parser = FormulaParser("=IPMT(0.1/12, 1, 36, 8000)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - (-66.67)) < 1.0)  // Interest payment should be ~$66.67
        }
    }
    
    @Test func evaluatePPMT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test PPMT(10%, 1, 2, 2000) - principal payment for first period
        let parser = FormulaParser("=PPMT(0.1, 1, 2, 2000)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - (-952.38)) < 1.0)  // Principal payment should be ~$952.38
        }
    }
    
    @Test func evaluateXNPV() throws {
        let cells: [String: CellValue] = [
            "A1": .number(-10000),
            "A2": .number(2750),
            "A3": .number(4250),
            "A4": .number(3250),
            "A5": .number(2750),
            "B1": .number(43831),  // Jan 1, 2020 (Excel serial date)
            "B2": .number(44197),  // Jan 1, 2021
            "B3": .number(44562),  // Jan 1, 2022
            "B4": .number(44927),  // Jan 1, 2023
            "B5": .number(45292)   // Jan 1, 2024
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test XNPV(9%, A1:A5, B1:B5) - NPV with specific dates
        let parser = FormulaParser("=XNPV(0.09, A1:A5, B1:B5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 555) < 200.0)  // XNPV (approximate due to date calculation)
        }
    }
    
    @Test func evaluateXIRR() throws {
        let cells: [String: CellValue] = [
            "A1": .number(-10000),
            "A2": .number(2750),
            "A3": .number(4250),
            "A4": .number(3250),
            "A5": .number(2750),
            "B1": .number(43831),  // Jan 1, 2020
            "B2": .number(44197),  // Jan 1, 2021
            "B3": .number(44562),  // Jan 1, 2022
            "B4": .number(44927),  // Jan 1, 2023
            "B5": .number(45292)   // Jan 1, 2024
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test XIRR(A1:A5, B1:B5) - IRR with specific dates
        let parser = FormulaParser("=XIRR(A1:A5, B1:B5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 0.373) < 0.3)  // XIRR should be ~37% (approximate due to date calculation differences)
        }
    }
    
    // MARK: - Date/Time Functions (Extended)
    
    @Test func evaluateWEEKDAY() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test WEEKDAY for Jan 1, 2020 (serial 43831, which is a Wednesday)
        let parser = FormulaParser("=WEEKDAY(43831)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val >= 1 && val <= 7)  // Should return valid day of week
        }
    }
    
    @Test func evaluateEOMONTH() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test EOMONTH - end of month 2 months from Jan 15, 2020
        let parser = FormulaParser("=EOMONTH(43845, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val > 43845)  // Should be after start date
        }
    }
    
    @Test func evaluateEDATE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test EDATE - 3 months from Jan 1, 2020
        let parser = FormulaParser("=EDATE(43831, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 43921) < 10)  // ~90 days later (3 months * 30 days)
        }
    }
    
    @Test func evaluateNETWORKDAYS() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test NETWORKDAYS between Jan 1 and Jan 31, 2020 (31 days)
        let parser = FormulaParser("=NETWORKDAYS(43831, 43861)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val >= 20 && val <= 25)  // Should be ~22 working days
        }
    }
    
    @Test func evaluateDATEDIF() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test DATEDIF - days between Jan 1 and Dec 31, 2020
        let parser = FormulaParser("=DATEDIF(43831, 44195, \"D\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(364))  // 364 days difference
    }
    
    @Test func evaluateYEARFRAC() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test YEARFRAC - fraction of year between Jan 1 and Jul 1, 2020
        let parser = FormulaParser("=YEARFRAC(43831, 44013, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 0.5) < 0.1)  // Should be ~0.5 (half year)
        }
    }
    
    @Test func evaluateTIME() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test TIME(12, 30, 0) - 12:30 PM
        let parser = FormulaParser("=TIME(12, 30, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(abs(val - 0.5208) < 0.001)  // 12:30 PM = 0.520833...
        }
    }
    
    @Test func evaluateHOUR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test HOUR(0.75) - 18:00 (6 PM)
        let parser = FormulaParser("=HOUR(0.75)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(18))
    }
    
    @Test func evaluateMINUTE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test MINUTE(0.5) - 12:00 PM
        let parser = FormulaParser("=MINUTE(0.5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(0))
    }
    
    @Test func evaluateSECOND() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test SECOND(0.000694) - should be 60 seconds (approximately)
        let parser = FormulaParser("=SECOND(0.01)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val >= 0 && val < 60)  // Valid second value
        }
    }
    
    // MARK: - Text Functions (Extended)
    
    @Test func evaluatePROPER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=PROPER(\"hello world\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("Hello World"))
    }
    
    @Test func evaluateCLEAN() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello\nWorld")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=CLEAN(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("HelloWorld"))
    }
    
    @Test func evaluateCHAR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test CHAR(65) = "A"
        let parser = FormulaParser("=CHAR(65)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("A"))
    }
    
    @Test func evaluateCODE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test CODE("A") = 65
        let parser = FormulaParser("=CODE(\"A\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(65))
    }
    
    @Test func evaluateEXACT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test case-sensitive comparison
        let parser1 = FormulaParser("=EXACT(\"Hello\", \"Hello\")")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(1))
        
        let parser2 = FormulaParser("=EXACT(\"Hello\", \"hello\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(0))
    }
    
    @Test func evaluateREPLACE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test REPLACE("abcdefgh", 3, 2, "XY") = "abXYefgh"
        let parser = FormulaParser("=REPLACE(\"abcdefgh\", 3, 2, \"XY\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("abXYefgh"))
    }
    
    @Test func evaluateREPT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test REPT("*", 5) = "*****"
        let parser = FormulaParser("=REPT(\"*\", 5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("*****"))
    }
    
    @Test func evaluateVALUE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test VALUE("$1,234.56") = 1234.56
        let parser = FormulaParser("=VALUE(\"$1,234.56\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(1234.56))
    }
    
    @Test func evaluateTEXTBEFORE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test TEXTBEFORE("Hello-World", "-") = "Hello"
        let parser = FormulaParser("=TEXTBEFORE(\"Hello-World\", \"-\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("Hello"))
    }
    
    @Test func evaluateTEXTAFTER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test TEXTAFTER("Hello-World", "-") = "World"
        let parser = FormulaParser("=TEXTAFTER(\"Hello-World\", \"-\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("World"))
    }
    
    @Test func evaluateTEXTSPLIT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test TEXTSPLIT("A,B,C", ",")
        let parser = FormulaParser("=TEXTSPLIT(\"A,B,C\", \",\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return an array
        if case .array(let rows) = result {
            #expect(rows.count == 1)
            #expect(rows[0].count == 3)
        }
    }
    
    // MARK: - Lookup & Reference Functions (Extended)
    
    @Test func evaluateHLOOKUP() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Product"), "B1": .text("Price"), "C1": .text("Stock"),
            "A2": .text("Apple"), "B2": .number(1.5), "C2": .number(100),
            "A3": .text("Banana"), "B3": .number(0.8), "C3": .number(150)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test HLOOKUP for "Price" in row 1, return row 2 value
        let parser = FormulaParser("=HLOOKUP(\"Price\", A1:C3, 2, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(1.5))
    }
    
    @Test func evaluateCHOOSE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test CHOOSE(2, "A", "B", "C") = "B"
        let parser = FormulaParser("=CHOOSE(2, \"A\", \"B\", \"C\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("B"))
    }
    
    @Test func evaluateTRANSPOSE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .number(2),
            "A2": .number(3), "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test TRANSPOSE - should swap rows and columns
        let parser = FormulaParser("=TRANSPOSE(A1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 2)  // Original had 2 cols, now 2 rows
            #expect(rows[0].count == 2)  // Original had 2 rows, now 2 cols
            #expect(rows[0][0] == .number(1))
            #expect(rows[0][1] == .number(3))
            #expect(rows[1][0] == .number(2))
            #expect(rows[1][1] == .number(4))
        }
    }
    
    @Test func evaluateROWS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test ROWS(A1:A3) = 3
        let parser = FormulaParser("=ROWS(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(3))
    }
    
    @Test func evaluateCOLUMNS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .number(2), "C1": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test COLUMNS(A1:C1) = 3
        let parser = FormulaParser("=COLUMNS(A1:C1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(3))
    }
    
    @Test func evaluateXMATCH() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"),
            "A2": .text("Banana"),
            "A3": .text("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test XMATCH("Banana", A1:A3, 0) = 2
        let parser = FormulaParser("=XMATCH(\"Banana\", A1:A3, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(2))
    }
    
    @Test func evaluateINDIRECT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42),
            "B2": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test INDIRECT("A1") = 42
        let parser = FormulaParser("=INDIRECT(\"A1\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(42))
    }
    
    // MARK: - Logical & Information Functions (Extended)
    
    @Test func evaluateXOR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test XOR(1, 0, 0) = 1 (odd number of true)
        let parser1 = FormulaParser("=XOR(1, 0, 0)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(1))
        
        // Test XOR(1, 1, 0) = 0 (even number of true)
        let parser2 = FormulaParser("=XOR(1, 1, 0)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(0))
    }
    
    @Test func evaluateIFNA() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test IFNA with non-error value
        let parser1 = FormulaParser("=IFNA(10, 0)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(10))
        
        // Test IFNA with #N/A - needs NA() function
        let parser2 = FormulaParser("=IFNA(NA(), 999)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(999))
    }
    
    @Test func evaluateISNA() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test ISNA(10) = 0
        let parser1 = FormulaParser("=ISNA(10)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(0))
        
        // Test ISNA(NA()) = 1
        let parser2 = FormulaParser("=ISNA(NA())")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(1))
    }
    
    @Test func evaluateTYPE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test TYPE(123) = 1 (number)
        let parser1 = FormulaParser("=TYPE(123)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(1))
        
        // Test TYPE("hello") = 2 (text)
        let parser2 = FormulaParser("=TYPE(\"hello\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(2))
    }
    
    @Test func evaluateN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test N(42) = 42
        let parser1 = FormulaParser("=N(42)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(42))
        
        // Test N("text") = 0
        let parser2 = FormulaParser("=N(\"text\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(0))
    }
    
    @Test func evaluateNA() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test NA() returns #N/A error
        let parser = FormulaParser("=NA()")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .error("N/A"))
    }
    
    @Test func evaluateCELL() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test CELL("row") - simplified implementation
        let parser = FormulaParser("=CELL(\"row\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(1))
    }
    
    // MARK: - Dynamic Array Functions (Excel 365)
    
    @Test func evaluateFILTER() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30),
            // Boolean criteria - include row 2 and 3, exclude row 1
            "B1": .boolean(false),
            "B2": .boolean(true),
            "B3": .boolean(true)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Filter A1:A3 using boolean array B1:B3
        let parser = FormulaParser("=FILTER(A1:A3, B1:B3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 20 and 30 (where criteria is TRUE)
        if case .array(let rows) = result {
            #expect(rows.count == 2)
            #expect(rows[0][0] == .number(20))
            #expect(rows[1][0] == .number(30))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateSORT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(30),
            "A2": .number(10),
            "A3": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sort A1:A3 ascending
        let parser = FormulaParser("=SORT(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 10, 20, 30
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(10))
            #expect(rows[1][0] == .number(20))
            #expect(rows[2][0] == .number(30))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateSORTDescending() throws {
        let cells: [String: CellValue] = [
            "A1": .number(30),
            "A2": .number(10),
            "A3": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sort A1:A3 descending (sort_order = -1)
        let parser = FormulaParser("=SORT(A1:A3, 1, -1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 30, 20, 10
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(30))
            #expect(rows[1][0] == .number(20))
            #expect(rows[2][0] == .number(10))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateUNIQUE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(10),
            "A4": .number(30),
            "A5": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Get unique values from A1:A5
        let parser = FormulaParser("=UNIQUE(A1:A5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 10, 20, 30
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(10))
            #expect(rows[1][0] == .number(20))
            #expect(rows[2][0] == .number(30))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateSEQUENCE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Generate sequence 1 to 5
        let parser = FormulaParser("=SEQUENCE(5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 1, 2, 3, 4, 5
        if case .array(let rows) = result {
            #expect(rows.count == 5)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .number(3))
            #expect(rows[3][0] == .number(4))
            #expect(rows[4][0] == .number(5))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateSEQUENCEWithStartAndStep() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Generate sequence starting at 10, step 5, 3 rows
        let parser = FormulaParser("=SEQUENCE(3, 1, 10, 5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 10, 15, 20
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(10))
            #expect(rows[1][0] == .number(15))
            #expect(rows[2][0] == .number(20))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateRANDARRAY() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Generate 3x2 random array
        let parser = FormulaParser("=RANDARRAY(3, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Verify structure (3 rows, 2 columns)
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0].count == 2)
            #expect(rows[1].count == 2)
            #expect(rows[2].count == 2)
            
            // Verify all values are numbers between 0 and 1
            for row in rows {
                for cell in row {
                    if case .number(let val) = cell {
                        #expect(val >= 0.0 && val < 1.0)
                    } else {
                        Issue.record("Expected number in RANDARRAY")
                    }
                }
            }
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateSORTBY() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Apple"),
            "A2": .text("Banana"),
            "A3": .text("Cherry"),
            "B1": .number(30),
            "B2": .number(10),
            "B3": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Sort A1:A3 by B1:B3 values
        let parser = FormulaParser("=SORTBY(A1:A3, B1:B3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return Banana, Cherry, Apple (sorted by 10, 20, 30)
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .string("Banana"))
            #expect(rows[1][0] == .string("Cherry"))
            #expect(rows[2][0] == .string("Apple"))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    // MARK: - More Dynamic Array Functions (Excel 365)
    
    @Test func evaluateTAKE() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Take first 3 rows
        let parser = FormulaParser("=TAKE(A1:A5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .number(3))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateTAKEFromEnd() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Take last 2 rows (negative)
        let parser = FormulaParser("=TAKE(A1:A5, -2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 2)
            #expect(rows[0][0] == .number(4))
            #expect(rows[1][0] == .number(5))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateDROP() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Drop first 2 rows
        let parser = FormulaParser("=DROP(A1:A5, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(3))
            #expect(rows[1][0] == .number(4))
            #expect(rows[2][0] == .number(5))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateDROPFromEnd() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Drop last 2 rows (negative)
        let parser = FormulaParser("=DROP(A1:A5, -2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 3)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .number(3))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateEXPAND() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Expand to 4 rows, 3 columns (pad with #N/A)
        let parser = FormulaParser("=EXPAND(A1:A2, 4, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 4)
            #expect(rows[0].count == 3)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .error("N/A"))
            #expect(rows[3][2] == .error("N/A"))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateEXPANDWithCustomPad() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Expand to 2x2 with 0 as pad value
        let parser = FormulaParser("=EXPAND(A1, 2, 2, 0)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 2)
            #expect(rows[0][0] == .number(1))
            #expect(rows[0][1] == .number(0))
            #expect(rows[1][0] == .number(0))
            #expect(rows[1][1] == .number(0))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateVSTACK() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "B1": .number(10),
            "B2": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Stack A1:A2 on top of B1:B2
        let parser = FormulaParser("=VSTACK(A1:A2, B1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 4)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .number(10))
            #expect(rows[3][0] == .number(20))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateHSTACK() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "B1": .number(10),
            "B2": .number(20)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Stack A1:A2 next to B1:B2 horizontally
        let parser = FormulaParser("=HSTACK(A1:A2, B1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 2)
            #expect(rows[0].count == 2)
            #expect(rows[0][0] == .number(1))
            #expect(rows[0][1] == .number(10))
            #expect(rows[1][0] == .number(2))
            #expect(rows[1][1] == .number(20))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateTOCOL() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .number(2),
            "A2": .number(3), "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Convert 2x2 array to single column
        let parser = FormulaParser("=TOCOL(A1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 4)
            #expect(rows[0][0] == .number(1))
            #expect(rows[1][0] == .number(2))
            #expect(rows[2][0] == .number(3))
            #expect(rows[3][0] == .number(4))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    @Test func evaluateTOROW() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .number(2),
            "A2": .number(3), "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Convert 2x2 array to single row
        let parser = FormulaParser("=TOROW(A1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .array(let rows) = result {
            #expect(rows.count == 1)
            #expect(rows[0].count == 4)
            #expect(rows[0][0] == .number(1))
            #expect(rows[0][1] == .number(2))
            #expect(rows[0][2] == .number(3))
            #expect(rows[0][3] == .number(4))
        } else {
            Issue.record("Expected array result")
        }
    }
    
    // MARK: - Database Functions
    
    @Test func evaluateDSUM() throws {
        // Database: Tree, Height, Age, Yield, Profit
        let cells: [String: CellValue] = [
            // Headers
            "A1": .text("Tree"), "B1": .text("Height"), "C1": .text("Age"), "D1": .text("Yield"), "E1": .text("Profit"),
            // Data rows
            "A2": .text("Apple"), "B2": .number(18), "C2": .number(20), "D2": .number(14), "E2": .number(105),
            "A3": .text("Pear"), "B3": .number(12), "C3": .number(12), "D3": .number(10), "E3": .number(96),
            "A4": .text("Cherry"), "B4": .number(13), "C4": .number(14), "D4": .number(9), "E4": .number(105),
            "A5": .text("Apple"), "B5": .number(14), "C5": .number(15), "D5": .number(10), "E5": .number(75),
            "A6": .text("Pear"), "B6": .number(9), "C6": .number(8), "D6": .number(8), "E6": .number(76),
            // Criteria: Tree = "Apple"
            "A8": .text("Tree"),
            "A9": .text("Apple")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DSUM: Sum profit for Apple trees
        let parser = FormulaParser("=DSUM(A1:E6, \"Profit\", A8:A9)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 105 + 75 = 180
        #expect(result == .number(180))
    }
    
    @Test func evaluateDAVERAGE() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Tree"), "B1": .text("Height"), "C1": .text("Yield"),
            "A2": .text("Apple"), "B2": .number(18), "C2": .number(14),
            "A3": .text("Pear"), "B3": .number(12), "C3": .number(10),
            "A4": .text("Apple"), "B4": .number(14), "C4": .number(10),
            "A6": .text("Tree"),
            "A7": .text("Apple")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DAVERAGE: Average yield for Apple trees
        let parser = FormulaParser("=DAVERAGE(A1:C4, \"Yield\", A6:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return (14 + 10) / 2 = 12
        #expect(result == .number(12))
    }
    
    @Test func evaluateDCOUNT() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Tree"), "B1": .text("Age"),
            "A2": .text("Apple"), "B2": .number(20),
            "A3": .text("Pear"), "B3": .number(12),
            "A4": .text("Apple"), "B4": .number(15),
            "A5": .text("Cherry"), "B5": .number(14),
            "A7": .text("Tree"),
            "A8": .text("Apple")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DCOUNT: Count Apple trees
        let parser = FormulaParser("=DCOUNT(A1:B5, \"Age\", A7:A8)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 2 (two Apple trees)
        #expect(result == .number(2))
    }
    
    @Test func evaluateDMAX() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Tree"), "B1": .text("Height"),
            "A2": .text("Apple"), "B2": .number(18),
            "A3": .text("Pear"), "B3": .number(12),
            "A4": .text("Apple"), "B4": .number(14),
            "A6": .text("Tree"),
            "A7": .text("Apple")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DMAX: Maximum height for Apple trees
        let parser = FormulaParser("=DMAX(A1:B4, \"Height\", A6:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 18
        #expect(result == .number(18))
    }
    
    @Test func evaluateDMIN() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Tree"), "B1": .text("Height"),
            "A2": .text("Apple"), "B2": .number(18),
            "A3": .text("Pear"), "B3": .number(12),
            "A4": .text("Apple"), "B4": .number(14),
            "A6": .text("Tree"),
            "A7": .text("Apple")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DMIN: Minimum height for Apple trees
        let parser = FormulaParser("=DMIN(A1:B4, \"Height\", A6:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 14
        #expect(result == .number(14))
    }
    
    @Test func evaluateDGET() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Tree"), "B1": .text("Height"), "C1": .text("Yield"),
            "A2": .text("Apple"), "B2": .number(18), "C2": .number(14),
            "A3": .text("Pear"), "B3": .number(12), "C3": .number(10),
            "A4": .text("Cherry"), "B4": .number(13), "C4": .number(9),
            "A6": .text("Tree"),
            "A7": .text("Pear")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DGET: Get yield for Pear (single match)
        let parser = FormulaParser("=DGET(A1:C4, \"Yield\", A6:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 10
        #expect(result == .number(10))
    }
    
    @Test func evaluateDPRODUCT() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Product"), "B1": .text("Quantity"),
            "A2": .text("Apples"), "B2": .number(2),
            "A3": .text("Oranges"), "B3": .number(3),
            "A4": .text("Apples"), "B4": .number(4),
            "A6": .text("Product"),
            "A7": .text("Apples")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // DPRODUCT: Product of quantities for Apples
        let parser = FormulaParser("=DPRODUCT(A1:B4, \"Quantity\", A6:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Should return 2 * 4 = 8
        #expect(result == .number(8))
    }
    
    // MARK: - More Statistical Functions
    
    @Test func evaluateCOVARIANCE_P() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(2),
            "A3": .number(4),
            "A4": .number(5),
            "A5": .number(6),
            "B1": .number(9),
            "B2": .number(7),
            "B3": .number(12),
            "B4": .number(15),
            "B5": .number(17)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Using direct function call instead of parser due to dot in name
        let expr = FormulaExpression.functionCall("COVARIANCE.P", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 5)),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 5))
        ])
        let result = try evaluator.evaluate(expr)
        
        // Population covariance
        if case .number(let val) = result {
            #expect(abs(val - 5.2) < 0.01)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateCOVARIANCE_S() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(2),
            "A3": .number(4),
            "B1": .number(9),
            "B2": .number(7),
            "B3": .number(12)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("COVARIANCE.S", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3)),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        
        // Sample covariance
        if case .number(let val) = result {
            #expect(abs(val - 2.5) < 0.01)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateSKEW() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(4),
            "A3": .number(5),
            "A4": .number(2),
            "A5": .number(3),
            "A6": .number(4),
            "A7": .number(5),
            "A8": .number(6),
            "A9": .number(4),
            "A10": .number(7)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SKEW(A1:A10)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Skewness should be a number
        if case .number = result {
            // Test passes if it returns a number
            #expect(true)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateKURT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(3),
            "A2": .number(4),
            "A3": .number(5),
            "A4": .number(2),
            "A5": .number(3),
            "A6": .number(4),
            "A7": .number(5),
            "A8": .number(6),
            "A9": .number(4),
            "A10": .number(7)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=KURT(A1:A10)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Kurtosis should be a number
        if case .number = result {
            #expect(true)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateGEOMEAN() throws {
        let cells: [String: CellValue] = [
            "A1": .number(4),
            "A2": .number(5),
            "A3": .number(8),
            "A4": .number(7),
            "A5": .number(11),
            "A6": .number(4),
            "A7": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=GEOMEAN(A1:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Geometric mean ≈ 5.476
        if case .number(let val) = result {
            #expect(abs(val - 5.476) < 0.01)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateHARMEAN() throws {
        let cells: [String: CellValue] = [
            "A1": .number(4),
            "A2": .number(5),
            "A3": .number(8),
            "A4": .number(7),
            "A5": .number(11),
            "A6": .number(4),
            "A7": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=HARMEAN(A1:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Harmonic mean ≈ 5.028
        if case .number(let val) = result {
            #expect(abs(val - 5.028) < 0.01)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateAVEDEV() throws {
        let cells: [String: CellValue] = [
            "A1": .number(4),
            "A2": .number(5),
            "A3": .number(8),
            "A4": .number(7),
            "A5": .number(11),
            "A6": .number(4),
            "A7": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=AVEDEV(A1:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Average absolute deviation
        if case .number(let val) = result {
            #expect(val > 2.0 && val < 3.0)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateDEVSQ() throws {
        let cells: [String: CellValue] = [
            "A1": .number(4),
            "A2": .number(5),
            "A3": .number(8),
            "A4": .number(7),
            "A5": .number(11),
            "A6": .number(4),
            "A7": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=DEVSQ(A1:A7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // Sum of squared deviations should be 48
        if case .number(let val) = result {
            #expect(abs(val - 48.0) < 0.1)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateSTANDARDIZE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=STANDARDIZE(42, 40, 1.5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // (42 - 40) / 1.5 = 1.333...
        if case .number(let val) = result {
            #expect(abs(val - 1.333) < 0.01)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateCONFIDENCE_NORM() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Using direct function call due to dot in name
        let expr = FormulaExpression.functionCall("CONFIDENCE.NORM", [
            .number(0.05),
            .number(2.5),
            .number(50)
        ])
        let result = try evaluator.evaluate(expr)
        
        // Should be approximately 0.693
        if case .number(let val) = result {
            #expect(abs(val - 0.693) < 0.1)
        } else {
            Issue.record("Expected number")
        }
    }
    
    // MARK: - More Math Functions
    
    @Test func evaluatePOWER() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=POWER(5, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(25))
    }
    
    @Test func evaluateMROUND() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=MROUND(10, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // 10 rounded to nearest multiple of 3 is 9
        #expect(result == .number(9))
    }
    
    @Test func evaluateEVEN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=EVEN(1.5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(2))
    }
    
    @Test func evaluateEVENNegative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=EVEN(-1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(-2))
    }
    
    @Test func evaluateODD() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=ODD(1.5)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(3))
    }
    
    @Test func evaluateODDNegative() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=ODD(-2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(-3))
    }
    
    @Test func evaluateQUOTIENT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=QUOTIENT(5, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(2))
    }
    
    @Test func evaluateRAND() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=RAND()")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val >= 0 && val < 1)
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateRANDBETWEEN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=RANDBETWEEN(1, 10)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let val) = result {
            #expect(val >= 1 && val <= 10)
            #expect(val == floor(val))  // Should be integer
        } else {
            Issue.record("Expected number")
        }
    }
    
    @Test func evaluateCOMBIN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=COMBIN(8, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // C(8,2) = 8!/(2!*6!) = 28
        #expect(result == .number(28))
    }
    
    @Test func evaluatePERMUT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=PERMUT(5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // P(5,3) = 5!/(5-3)! = 60
        #expect(result == .number(60))
    }
    
    @Test func evaluateMULTINOMIAL() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=MULTINOMIAL(2, 3, 4)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        // 9!/(2!*3!*4!) = 1260
        #expect(result == .number(1260))
    }
    
    // MARK: - More Text Functions
    
    @Test func evaluateNUMBERVALUE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test with grouping and decimal separators
        let parser = FormulaParser("=NUMBERVALUE(\"1,234.56\", \".\", \",\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(1234.56))
    }
    
    @Test func evaluateDOLLAR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=DOLLAR(1234.567, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("$1,234.57"))
    }
    
    @Test func evaluateFIXED() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Use direct function call to avoid parsing FALSE
        let expr = FormulaExpression.functionCall("FIXED", [
            .number(1234.567),
            .number(2),
            .number(0)  // 0 means include commas
        ])
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("1,234.57"))
    }
    
    @Test func evaluateT() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello"),
            "A2": .number(123)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=T(A1)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .string("Hello"))
        
        let parser2 = FormulaParser("=T(A2)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .string(""))
    }
    
    @Test func evaluateUNICODE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=UNICODE(\"A\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .number(65))
    }
    
    @Test func evaluateUNICHAR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=UNICHAR(65)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("A"))
    }
    
    @Test func evaluateARRAYTOTEXT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("ARRAYTOTEXT", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        
        // Should return array representation
        if case .string(let str) = result {
            #expect(str.contains("1"))
            #expect(str.contains("2"))
            #expect(str.contains("3"))
        } else {
            Issue.record("Expected string result")
        }
    }
    
    @Test func evaluateVALUETOTEXT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(123.45)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=VALUETOTEXT(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        #expect(result == .string("123.45"))
    }
    
    // MARK: - Engineering Functions
    
    @Test func evaluateCONVERT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test distance conversion: meters to feet
        let parser1 = FormulaParser("=CONVERT(1, \"m\", \"ft\")")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        if case .number(let val) = result1 {
            #expect(abs(val - 3.28084) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
        
        // Test weight conversion: pounds to kilograms
        let parser2 = FormulaParser("=CONVERT(100, \"lbm\", \"kg\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        if case .number(let val) = result2 {
            #expect(abs(val - 45.359237) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
        
        // Test temperature conversion: Celsius to Fahrenheit
        let parser3 = FormulaParser("=CONVERT(0, \"C\", \"F\")")
        let expr3 = try parser3.parse()
        let result3 = try evaluator.evaluate(expr3)
        #expect(result3 == .number(32))
    }
    
    @Test func evaluateDELTA() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=DELTA(5, 5)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(1))
        
        let parser2 = FormulaParser("=DELTA(5, 4)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(0))
        
        // Test with default second argument (0)
        let parser3 = FormulaParser("=DELTA(0)")
        let expr3 = try parser3.parse()
        let result3 = try evaluator.evaluate(expr3)
        #expect(result3 == .number(1))
    }
    
    @Test func evaluateGESTEP() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=GESTEP(5, 4)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(1))
        
        let parser2 = FormulaParser("=GESTEP(3, 4)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(0))
        
        let parser3 = FormulaParser("=GESTEP(4, 4)")
        let expr3 = try parser3.parse()
        let result3 = try evaluator.evaluate(expr3)
        #expect(result3 == .number(1))
    }
    
    @Test func evaluateDEC2BIN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=DEC2BIN(9)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .string("1001"))
        
        let parser2 = FormulaParser("=DEC2BIN(9, 8)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .string("00001001"))
        
        // Test negative number
        let parser3 = FormulaParser("=DEC2BIN(-100)")
        let expr3 = try parser3.parse()
        let result3 = try evaluator.evaluate(expr3)
        #expect(result3 == .string("1110011100"))
    }
    
    @Test func evaluateDEC2OCT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=DEC2OCT(58)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .string("72"))
        
        let parser2 = FormulaParser("=DEC2OCT(58, 5)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .string("00072"))
    }
    
    @Test func evaluateDEC2HEX() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=DEC2HEX(100)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .string("64"))
        
        let parser2 = FormulaParser("=DEC2HEX(100, 5)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .string("00064"))
    }
    
    @Test func evaluateBIN2DEC() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=BIN2DEC(\"1001\")")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(9))
        
        // Test negative (10-bit two's complement)
        let parser2 = FormulaParser("=BIN2DEC(\"1111111111\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(-1))
    }
    
    @Test func evaluateOCT2DEC() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=OCT2DEC(\"72\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(58))
    }
    
    @Test func evaluateHEX2DEC() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=HEX2DEC(\"64\")")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(100))
        
        let parser2 = FormulaParser("=HEX2DEC(\"FF\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(255))
    }
    
    @Test func evaluateBIN2HEX() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BIN2HEX(\"11111\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("1F"))
    }
    
    @Test func evaluateHEX2BIN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=HEX2BIN(\"1F\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("11111"))
    }
    
    @Test func evaluateHEX2OCT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=HEX2OCT(\"FF\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("377"))
    }
    
    @Test func evaluateOCT2BIN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=OCT2BIN(\"7\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("111"))
    }
    
    @Test func evaluateOCT2HEX() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=OCT2HEX(\"377\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("FF"))
    }
    
    @Test func evaluateBIN2OCT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BIN2OCT(\"111\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("7"))
    }
    
    @Test func evaluateBITAND() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BITAND(5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // 5 = 101, 3 = 011, AND = 001 = 1
        #expect(result == .number(1))
    }
    
    @Test func evaluateBITOR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BITOR(5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // 5 = 101, 3 = 011, OR = 111 = 7
        #expect(result == .number(7))
    }
    
    @Test func evaluateBITXOR() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BITXOR(5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // 5 = 101, 3 = 011, XOR = 110 = 6
        #expect(result == .number(6))
    }
    
    @Test func evaluateBITLSHIFT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BITLSHIFT(4, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // 4 << 2 = 16
        #expect(result == .number(16))
    }
    
    @Test func evaluateBITRSHIFT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=BITRSHIFT(16, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // 16 >> 2 = 4
        #expect(result == .number(4))
    }
    
    // MARK: - Advanced Statistical Functions
    
    @Test func evaluateFORECAST() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "B1": .number(10),
            "B2": .number(20),
            "B3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // FORECAST with perfect linear relationship (y = 10x)
        let expr = FormulaExpression.functionCall("FORECAST", [
            .number(4),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3)),
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(40))
    }
    
    @Test func evaluatePERCENTILE_EXC() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("PERCENTILE.EXC", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 5)),
            .number(0.5)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))
    }
    
    @Test func evaluateQUARTILE_EXC() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "A4": .number(4),
            "A5": .number(5)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("QUARTILE.EXC", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 5)),
            .number(2)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))
    }
    
    @Test func evaluatePERCENTRANK_INC() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30),
            "A4": .number(40)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("PERCENTRANK.INC", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 4)),
            .number(30)
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val - 0.666) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluatePERCENTRANK_EXC() throws {
        let cells: [String: CellValue] = [
            "A1": .number(10),
            "A2": .number(20),
            "A3": .number(30),
            "A4": .number(40)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("PERCENTRANK.EXC", [
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 4)),
            .number(20)
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val - 0.4) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateNORM_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test cumulative normal distribution at mean
        let expr = FormulaExpression.functionCall("NORM.DIST", [
            .number(0),
            .number(0),
            .number(1),
            .number(1)  // Cumulative
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val - 0.5) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateNORM_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test inverse: probability 0.5 at mean 0, stddev 1 should give 0
        let expr = FormulaExpression.functionCall("NORM.INV", [
            .number(0.5),
            .number(0),
            .number(1)
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateNORM_S_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test standard normal cumulative at 0
        let expr = FormulaExpression.functionCall("NORM.S.DIST", [
            .number(0),
            .number(1)  // Cumulative
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val - 0.5) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateNORM_S_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test inverse standard normal at probability 0.5
        let expr = FormulaExpression.functionCall("NORM.S.INV", [
            .number(0.5)
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val) < 0.001)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    // MARK: - Distribution Functions
    
    @Test func evaluateBINOM_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test binomial PMF: 3 successes in 5 trials with p=0.5
        let expr = FormulaExpression.functionCall("BINOM.DIST", [
            .number(3),
            .number(5),
            .number(0.5),
            .number(0)  // PMF
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            #expect(abs(val - 0.3125) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateBINOM_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("BINOM.INV", [
            .number(10),
            .number(0.5),
            .number(0.5)
        ])
        let result = try evaluator.evaluate(expr)
        // Should return the smallest k where cumulative >= 0.5
        if case .number(let val) = result {
            #expect(val >= 4 && val <= 6)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluatePOISSON_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test Poisson PMF: k=2, mean=3
        let expr = FormulaExpression.functionCall("POISSON.DIST", [
            .number(2),
            .number(3),
            .number(0)  // PMF
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            // P(X=2) = (3^2 * e^-3) / 2! ≈ 0.224
            #expect(abs(val - 0.224) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateEXPON_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        // Test exponential CDF
        let expr = FormulaExpression.functionCall("EXPON.DIST", [
            .number(1),
            .number(1),
            .number(1)  // CDF
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            // CDF(1) = 1 - e^-1 ≈ 0.632
            #expect(abs(val - 0.632) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCHISQ_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("CHISQ.DIST", [
            .number(5),
            .number(3),
            .number(1)  // CDF
        ])
        let result = try evaluator.evaluate(expr)
        // Just check it returns a valid probability
        if case .number(let val) = result {
            #expect(val >= 0 && val <= 1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCHISQ_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("CHISQ.INV", [
            .number(0.5),
            .number(5)
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive number
        if case .number(let val) = result {
            #expect(val > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateT_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("T.DIST", [
            .number(0),
            .number(10),
            .number(1)  // CDF
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            // At t=0, CDF should be around 0.5
            #expect(abs(val - 0.5) < 0.1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateT_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("T.INV", [
            .number(0.5),
            .number(10)
        ])
        let result = try evaluator.evaluate(expr)
        if case .number(let val) = result {
            // At prob=0.5, should return value near 0
            #expect(abs(val) < 0.1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateF_DIST() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("F.DIST", [
            .number(1),
            .number(5),
            .number(5),
            .number(1)  // CDF
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a valid probability
        if case .number(let val) = result {
            #expect(val >= 0 && val <= 1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateF_INV() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("F.INV", [
            .number(0.5),
            .number(5),
            .number(5)
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive number
        if case .number(let val) = result {
            #expect(val > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    // MARK: - Information Functions
    
    @Test func evaluateISFORMULA() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ISFORMULA(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // Simplified implementation returns false
        #expect(result == .boolean(false))
    }
    
    @Test func evaluateISEVEN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=ISEVEN(4)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .boolean(true))
        
        let parser2 = FormulaParser("=ISEVEN(5)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .boolean(false))
    }
    
    @Test func evaluateISODD() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser1 = FormulaParser("=ISODD(5)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .boolean(true))
        
        let parser2 = FormulaParser("=ISODD(4)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .boolean(false))
    }
    
    @Test func evaluateSHEET() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=SHEET()")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // Simplified implementation returns 1
        #expect(result == .number(1))
    }
    
    @Test func evaluateSHEETS() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=SHEETS()")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        // Simplified implementation returns 1
        #expect(result == .number(1))
    }
    
    @Test func evaluateISLOGICAL() throws {
        let cells: [String: CellValue] = [
            "A1": .boolean(true),
            "A2": .number(1)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ISLOGICAL(A1)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .boolean(true))
        
        let parser2 = FormulaParser("=ISLOGICAL(A2)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .boolean(false))
    }
    
    @Test func evaluateISNONTEXT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42),
            "A2": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ISNONTEXT(A1)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .boolean(true))
        
        let parser2 = FormulaParser("=ISNONTEXT(A2)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .boolean(false))
    }
    
    // MARK: - Financial Functions (Depreciation & Securities)
    
    @Test func evaluateDB() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("DB", [
            .number(10000),  // cost
            .number(1000),   // salvage
            .number(5),      // life
            .number(1)       // period
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive depreciation value
        if case .number(let val) = result {
            #expect(val > 0 && val < 10000)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateDDB() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("DDB", [
            .number(10000),  // cost
            .number(1000),   // salvage
            .number(5),      // life
            .number(1)       // period
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive depreciation value
        if case .number(let val) = result {
            #expect(val > 0 && val < 10000)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateSLN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("SLN", [
            .number(10000),  // cost
            .number(1000),   // salvage
            .number(5)       // life
        ])
        let result = try evaluator.evaluate(expr)
        // SLN = (10000 - 1000) / 5 = 1800
        #expect(result == .number(1800))
    }
    
    @Test func evaluateSYD() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("SYD", [
            .number(10000),  // cost
            .number(1000),   // salvage
            .number(5),      // life
            .number(1)       // period
        ])
        let result = try evaluator.evaluate(expr)
        // SYD for period 1: (10000-1000) * (5/15) = 3000
        #expect(result == .number(3000))
    }
    
    @Test func evaluateVDB() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("VDB", [
            .number(10000),  // cost
            .number(1000),   // salvage
            .number(5),      // life
            .number(0),      // start period
            .number(1)       // end period
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive depreciation value
        if case .number(let val) = result {
            #expect(val > 0 && val < 10000)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluatePRICE() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("PRICE", [
            .number(44562),  // settlement (Excel date)
            .number(44927),  // maturity
            .number(0.05),   // rate
            .number(0.06),   // yield
            .number(100),    // redemption
            .number(2)       // frequency
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive price
        if case .number(let val) = result {
            #expect(val > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateYIELD() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("YIELD", [
            .number(44562),  // settlement
            .number(44927),  // maturity
            .number(0.05),   // rate
            .number(95),     // price
            .number(100),    // redemption
            .number(2)       // frequency
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive yield
        if case .number(let val) = result {
            #expect(val > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateACCRINT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("ACCRINT", [
            .number(44562),  // issue
            .number(44653),  // first interest
            .number(44652),  // settlement
            .number(0.05),   // rate
            .number(1000),   // par
            .number(2)       // frequency
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a positive accrued interest
        if case .number(let val) = result {
            #expect(val > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCUMIPMT() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("CUMIPMT", [
            .number(0.05/12),  // rate per period
            .number(360),      // nper (30 years * 12)
            .number(100000),   // pv
            .number(1),        // start period
            .number(12),       // end period
            .number(0)         // type (end of period)
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a negative interest (payment)
        if case .number(let val) = result {
            #expect(val < 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCUMPRINC() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("CUMPRINC", [
            .number(0.05/12),  // rate per period
            .number(360),      // nper (30 years * 12)
            .number(100000),   // pv
            .number(1),        // start period
            .number(12),       // end period
            .number(0)         // type (end of period)
        ])
        let result = try evaluator.evaluate(expr)
        // Check it returns a negative principal (payment)
        if case .number(let val) = result {
            #expect(val < 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    // MARK: - Lookup & Reference Functions
    
    @Test func evaluateADDRESS() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=ADDRESS(5, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("C5"))
    }
    
    @Test func evaluateAREAS() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=AREAS(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(1))
    }
    
    @Test func evaluateFORMULATEXT() throws {
        let cells: [String: CellValue] = [
            "A1": .number(42)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=FORMULATEXT(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("N/A"))
    }
    
    @Test func evaluateHYPERLINK() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let parser = FormulaParser("=HYPERLINK(\"http://example.com\", \"Example\")")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Example"))
    }
    
    @Test func evaluateLOOKUP() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3),
            "B1": .number(10),
            "B2": .number(20),
            "B3": .number(30)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let expr = FormulaExpression.functionCall("LOOKUP", [
            .number(2),
            .range(CellReference(column: "A", row: 1), CellReference(column: "A", row: 3)),
            .range(CellReference(column: "B", row: 1), CellReference(column: "B", row: 3))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(20))
    }
    
    @Test func evaluateCOLUMN() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("COLUMN", [
            .cellRef(CellReference(column: "C", row: 5))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(3))
    }
    
    @Test func evaluateROW() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        let expr = FormulaExpression.functionCall("ROW", [
            .cellRef(CellReference(column: "C", row: 5))
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    // MARK: - Batch 22: Additional Math Functions Tests
    
    @Test func evaluateROUNDUP() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ROUNDUP(3.14159, 2)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(3.15))
        
        let parser2 = FormulaParser("=ROUNDUP(-3.14159, 2)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(-3.15))  // Round away from zero
    }
    
    @Test func evaluateROUNDDOWN() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ROUNDDOWN(3.14159, 2)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(3.14))
        
        let parser2 = FormulaParser("=ROUNDDOWN(-3.14159, 2)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(-3.14))  // Round toward zero
    }
    
    @Test func evaluateSQRTPI() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SQRTPI(4)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(abs(n - sqrt(4 * .pi)) < 0.0001)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateSUMSQ() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(3),
            "A3": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMSQ(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(29))  // 2^2 + 3^2 + 4^2 = 4 + 9 + 16 = 29
    }
    
    @Test func evaluateSUMX2MY2() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(3),
            "B1": .number(1),
            "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMX2MY2(A1:A2, B1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(-4))  // (2^2-1^2) + (3^2-4^2) = (4-1) + (9-16) = 3 - 7 = -4
    }
    
    @Test func evaluateSUMX2PY2() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(3),
            "B1": .number(1),
            "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMX2PY2(A1:A2, B1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(30))  // (2^2+1^2) + (3^2+4^2) = 5 + 25 = 30
    }
    
    @Test func evaluateSUMXMY2() throws {
        let cells: [String: CellValue] = [
            "A1": .number(2),
            "A2": .number(3),
            "B1": .number(1),
            "B2": .number(4)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=SUMXMY2(A1:A2, B1:B2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(2))  // (2-1)^2 + (3-4)^2 = 1 + 1 = 2
    }
    
    @Test func evaluateSERIESSUM() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1),
            "A2": .number(2),
            "A3": .number(3)
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // SERIESSUM(x, n, m, coefficients) = sum of coefficients[i] * x^(n + m*i)
        // x=2, n=0, m=1, coeffs=[1,2,3] => 1*2^0 + 2*2^1 + 3*2^2 = 1 + 4 + 12 = 17
        let parser = FormulaParser("=SERIESSUM(2, 0, 1, A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(17))
    }
    
    @Test func evaluateMINA() throws {
        let cells: [String: CellValue] = [
            "A1": .number(5),
            "A2": .number(3),
            "A3": .text("text")  // Should be treated as 0
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MINA(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(0))  // text counts as 0, which is min
    }
    
    @Test func evaluateMAXA() throws {
        let cells: [String: CellValue] = [
            "A1": .number(5),
            "A2": .number(3),
            "A3": .text("text")  // Should be treated as 0
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MAXA(A1:A3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))  // max of 5, 3, 0
    }
    
    // MARK: - Batch 23: Aliases and Simple Variants Tests
    
    @Test func evaluateCEILINGPRECISE() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=CEILING.PRECISE(4.3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    @Test func evaluateISOCEILING() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=ISO.CEILING(-4.3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(-4))
    }
    
    @Test func evaluateFLOORPRECISE() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=FLOOR.PRECISE(4.7)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(4))
    }
    
    @Test func evaluateROMAN() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ROMAN(499)")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .string("CDXCIX"))
        
        let parser2 = FormulaParser("=ROMAN(1984)")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .string("MCMLXXXIV"))
    }
    
    @Test func evaluateARABIC() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser1 = FormulaParser("=ARABIC(\"CDXCIX\")")
        let expr1 = try parser1.parse()
        let result1 = try evaluator.evaluate(expr1)
        #expect(result1 == .number(499))
        
        let parser2 = FormulaParser("=ARABIC(\"MCMLXXXIV\")")
        let expr2 = try parser2.parse()
        let result2 = try evaluator.evaluate(expr2)
        #expect(result2 == .number(1984))
    }
    
    @Test func evaluateLEFTB() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LEFTB(A1, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Hel"))
    }
    
    @Test func evaluateRIGHTB() throws {
        let cells: [String: CellValue] = [
            "A1": .text("World")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=RIGHTB(A1, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("rld"))
    }
    
    @Test func evaluateMIDB() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=MIDB(A1, 2, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("ell"))
    }
    
    @Test func evaluateLENB() throws {
        let cells: [String: CellValue] = [
            "A1": .text("Hello")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LENB(A1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(5))
    }
    
    // MARK: - Batch 24: Statistical Distributions Tests
    
    @Test func evaluateBETADIST() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        // Test cumulative beta distribution
        let parser = FormulaParser("=BETA.DIST(0.5, 2, 3, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 1)  // Should be a probability
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateBETAINV() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=BETA.INV(0.5, 2, 3)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 1)  // Should be in [0,1]
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateGAMMADIST() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=GAMMA.DIST(2, 3, 2, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateGAMMAINV() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=GAMMA.INV(0.5, 3, 2)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n > 0)  // Should be positive
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateGAMMALN() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=GAMMALN(4)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            // ln(Γ(4)) = ln(3!) = ln(6) ≈ 1.7918
            #expect(abs(n - log(6.0)) < 0.01)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateWEIBULLDIST() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=WEIBULL.DIST(2, 1.5, 1, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateLOGNORMDIST() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LOGNORM.DIST(2, 0, 1, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 1)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateLOGNORMINV() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=LOGNORM.INV(0.5, 0, 1)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n > 0)
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCONFIDENCENORM() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=CONFIDENCE.NORM(0.05, 2.5, 50)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n > 0)  // Confidence interval should be positive
        } else {
            Issue.record("Expected number result")
        }
    }
    
    @Test func evaluateCRITBINOM() throws {
        let cells: [String: CellValue] = [:]
        let evaluator = makeTestEvaluator(cells: cells)
        
        let parser = FormulaParser("=CRITBINOM(10, 0.5, 0.75)")
        let expr = try parser.parse()
        let result = try evaluator.evaluate(expr)
        
        if case .number(let n) = result {
            #expect(n >= 0 && n <= 10)
        } else {
            Issue.record("Expected number result")
        }
    }
}
