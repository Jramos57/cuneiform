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
        
        guard let ref = CellReference("A1") else {
            #expect(Bool(false), "Failed to create cell reference")
            return
        }
        
        let expr = FormulaExpression.cellRef(ref)
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(42))
    }
    
    @Test func evaluateMissingCellReference() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        guard let ref = CellReference("Z99") else {
            #expect(Bool(false), "Failed to create cell reference")
            return
        }
        
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
        
        guard let a1 = CellReference("A1"), let b1 = CellReference("B1") else {
            #expect(Bool(false), "Failed to create cell references")
            return
        }
        
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
        
        guard let start = CellReference("A1"), let end = CellReference("A3") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
        let expr = FormulaExpression.functionCall("SUM", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(60))
    }
    
    @Test func evaluateSumEmptyRange() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        guard let start = CellReference("A1"), let end = CellReference("A3") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
        
        guard let start = CellReference("B1"), let end = CellReference("B4") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
        let expr = FormulaExpression.functionCall("AVERAGE", [.range(start, end)])
        let result = try evaluator.evaluate(expr)
        #expect(result == .number(12.5))
    }
    
    @Test func evaluateAverageEmptyRange() throws {
        let evaluator = makeTestEvaluator(cells: [:])
        
        guard let start = CellReference("A1"), let end = CellReference("A1") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
        
        guard let a1 = CellReference("A1") else {
            #expect(Bool(false), "Failed to create cell reference")
            return
        }
        
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
            "A1": .number(1), "B1": .string("Apple"),
            "A2": .number(2), "B2": .string("Banana"),
            "A3": .number(3), "B3": .string("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("B3") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
        let expr = FormulaExpression.functionCall("VLOOKUP", [
            .number(2),
            .range(start, end),
            .number(2),
            .boolean(false)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .string("Banana"))
    }
    
    @Test func evaluateVlookupNotFound() throws {
        let cells: [String: CellValue] = [
            "A1": .number(1), "B1": .string("Apple"),
            "A2": .number(2), "B2": .string("Banana")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("B2") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
        let expr = FormulaExpression.functionCall("VLOOKUP", [
            .number(5),
            .range(start, end),
            .number(2),
            .boolean(false)
        ])
        let result = try evaluator.evaluate(expr)
        #expect(result == .error("N/A"))
    }
    
    // MARK: - INDEX Function
    
    @Test func evaluateIndexWithRowAndColumn() throws {
        let cells: [String: CellValue] = [
            "A1": .string("A"), "B1": .string("B"),
            "A2": .string("C"), "B2": .string("D")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("B2") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
            "A1": .string("A")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("A1") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
            "A1": .string("Apple"),
            "A2": .string("Banana"),
            "A3": .string("Cherry")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("A3") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
            "A1": .string("Apple"),
            "A2": .string("Banana")
        ]
        let evaluator = makeTestEvaluator(cells: cells)
        
        guard let start = CellReference("A1"), let end = CellReference("A2") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
        
        guard let start = CellReference("A1"), let end = CellReference("A4") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
        
        guard let start = CellReference("A1"), let end = CellReference("A4") else {
            #expect(Bool(false), "Failed to create range")
            return
        }
        
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
}
