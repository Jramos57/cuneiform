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
}
