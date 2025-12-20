import Foundation
import Testing
@testable import Cuneiform

@Suite("Formula Parser Tests (Phase 5)")
struct FormulaParserTests {
    
    // MARK: - Basic Literals
    
    @Test func parseNumber() throws {
        let parser = FormulaParser("42")
        let expr = try parser.parse()
        #expect(expr == .number(42))
    }
    
    @Test func parseDecimal() throws {
        let parser = FormulaParser("3.14159")
        let expr = try parser.parse()
        #expect(expr == .number(3.14159))
    }
    
    @Test func parseString() throws {
        let parser = FormulaParser("\"Hello World\"")
        let expr = try parser.parse()
        #expect(expr == .string("Hello World"))
    }
    
    // MARK: - Cell References
    
    @Test func parseCellReference() throws {
        let parser = FormulaParser("A1")
        let expr = try parser.parse()
        guard case .cellRef(let ref) = expr else {
            #expect(Bool(false), "Expected cellRef")
            return
        }
        #expect(ref.column == "A")
        #expect(ref.row == 1)
    }
    
    @Test func parseAbsoluteCellReference() throws {
        let parser = FormulaParser("$B$5")
        let expr = try parser.parse()
        guard case .cellRef(let ref) = expr else {
            #expect(Bool(false), "Expected cellRef")
            return
        }
        #expect(ref.column == "B")
        #expect(ref.row == 5)
    }
    
    @Test func parseMixedCellReference() throws {
        let parser = FormulaParser("$C10")
        let expr = try parser.parse()
        guard case .cellRef = expr else {
            #expect(Bool(false), "Expected cellRef")
            return
        }
    }
    
    @Test func parseRange() throws {
        let parser = FormulaParser("A1:B10")
        let expr = try parser.parse()
        guard case .range(let start, let end) = expr else {
            #expect(Bool(false), "Expected range")
            return
        }
        #expect(start.column == "A")
        #expect(start.row == 1)
        #expect(end.column == "B")
        #expect(end.row == 10)
    }
    
    // MARK: - Arithmetic Operations
    
    @Test func parseAddition() throws {
        let parser = FormulaParser("2 + 3")
        let expr = try parser.parse()
        guard case .binaryOp(.add, .number(2), .number(3)) = expr else {
            #expect(Bool(false), "Expected addition")
            return
        }
    }
    
    @Test func parseSubtraction() throws {
        let parser = FormulaParser("10 - 4")
        let expr = try parser.parse()
        guard case .binaryOp(.subtract, .number(10), .number(4)) = expr else {
            #expect(Bool(false), "Expected subtraction")
            return
        }
    }
    
    @Test func parseMultiplication() throws {
        let parser = FormulaParser("5 * 6")
        let expr = try parser.parse()
        guard case .binaryOp(.multiply, .number(5), .number(6)) = expr else {
            #expect(Bool(false), "Expected multiplication")
            return
        }
    }
    
    @Test func parseDivision() throws {
        let parser = FormulaParser("20 / 4")
        let expr = try parser.parse()
        guard case .binaryOp(.divide, .number(20), .number(4)) = expr else {
            #expect(Bool(false), "Expected division")
            return
        }
    }
    
    @Test func parsePower() throws {
        let parser = FormulaParser("2 ^ 8")
        let expr = try parser.parse()
        guard case .binaryOp(.power, .number(2), .number(8)) = expr else {
            #expect(Bool(false), "Expected power")
            return
        }
    }
    
    // MARK: - Operator Precedence
    
    @Test func parseMultiplicationBeforeAddition() throws {
        let parser = FormulaParser("2 + 3 * 4")
        let expr = try parser.parse()
        
        // Should parse as 2 + (3 * 4)
        guard case .binaryOp(.add, .number(2), let right) = expr else {
            #expect(Bool(false), "Expected addition at top level")
            return
        }
        
        guard case .binaryOp(.multiply, .number(3), .number(4)) = right else {
            #expect(Bool(false), "Expected multiplication on right side")
            return
        }
    }
    
    @Test func parsePowerBeforeMultiplication() throws {
        let parser = FormulaParser("2 * 3 ^ 2")
        let expr = try parser.parse()
        
        // Should parse as 2 * (3 ^ 2)
        guard case .binaryOp(.multiply, .number(2), let right) = expr else {
            #expect(Bool(false), "Expected multiplication at top level")
            return
        }
        
        guard case .binaryOp(.power, .number(3), .number(2)) = right else {
            #expect(Bool(false), "Expected power on right side")
            return
        }
    }
    
    @Test func parseParenthesesOverridePrecedence() throws {
        let parser = FormulaParser("(2 + 3) * 4")
        let expr = try parser.parse()
        
        // Should parse as (2 + 3) * 4
        guard case .binaryOp(.multiply, let left, .number(4)) = expr else {
            #expect(Bool(false), "Expected multiplication at top level")
            return
        }
        
        guard case .binaryOp(.add, .number(2), .number(3)) = left else {
            #expect(Bool(false), "Expected addition in parentheses")
            return
        }
    }
    
    // MARK: - Comparison Operators
    
    @Test func parseEquals() throws {
        let parser = FormulaParser("A1 = 5")
        let expr = try parser.parse()
        guard case .binaryOp(.equals, .cellRef, .number(5)) = expr else {
            #expect(Bool(false), "Expected equals comparison")
            return
        }
    }
    
    @Test func parseNotEquals() throws {
        let parser = FormulaParser("A1 <> 0")
        let expr = try parser.parse()
        guard case .binaryOp(.notEquals, .cellRef, .number(0)) = expr else {
            #expect(Bool(false), "Expected not-equals comparison")
            return
        }
    }
    
    @Test func parseLessThan() throws {
        let parser = FormulaParser("B2 < 10")
        let expr = try parser.parse()
        guard case .binaryOp(.lessThan, .cellRef, .number(10)) = expr else {
            #expect(Bool(false), "Expected less-than comparison")
            return
        }
    }
    
    @Test func parseGreaterThanOrEqual() throws {
        let parser = FormulaParser("C3 >= 20")
        let expr = try parser.parse()
        guard case .binaryOp(.greaterThanOrEqual, .cellRef, .number(20)) = expr else {
            #expect(Bool(false), "Expected greater-than-or-equal comparison")
            return
        }
    }
    
    // MARK: - Function Calls
    
    @Test func parseFunctionNoArgs() throws {
        let parser = FormulaParser("TODAY()")
        let expr = try parser.parse()
        guard case .functionCall(let name, let args) = expr else {
            #expect(Bool(false), "Expected function call")
            return
        }
        #expect(name == "TODAY")
        #expect(args.isEmpty)
    }
    
    @Test func parseFunctionOneArg() throws {
        let parser = FormulaParser("SUM(A1:A10)")
        let expr = try parser.parse()
        guard case .functionCall(let name, let args) = expr else {
            #expect(Bool(false), "Expected function call")
            return
        }
        #expect(name == "SUM")
        #expect(args.count == 1)
        guard case .range = args[0] else {
            #expect(Bool(false), "Expected range argument")
            return
        }
    }
    
    @Test func parseFunctionMultipleArgs() throws {
        let parser = FormulaParser("IF(A1 > 5, \"High\", \"Low\")")
        let expr = try parser.parse()
        guard case .functionCall(let name, let args) = expr else {
            #expect(Bool(false), "Expected function call")
            return
        }
        #expect(name == "IF")
        #expect(args.count == 3)
    }
    
    @Test func parseNestedFunctions() throws {
        let parser = FormulaParser("SUM(A1, AVERAGE(B1:B5))")
        let expr = try parser.parse()
        guard case .functionCall(let name, let args) = expr else {
            #expect(Bool(false), "Expected function call")
            return
        }
        #expect(name == "SUM")
        #expect(args.count == 2)
        
        guard case .functionCall(let innerName, _) = args[1] else {
            #expect(Bool(false), "Expected nested function")
            return
        }
        #expect(innerName == "AVERAGE")
    }
    
    // MARK: - Complex Formulas
    
    @Test func parseComplexFormula() throws {
        let parser = FormulaParser("=SUM(A1:A10) / COUNT(A1:A10)")
        let expr = try parser.parse()
        
        guard case .binaryOp(.divide, let left, let right) = expr else {
            #expect(Bool(false), "Expected division")
            return
        }
        
        guard case .functionCall("SUM", _) = left else {
            #expect(Bool(false), "Expected SUM on left")
            return
        }
        
        guard case .functionCall("COUNT", _) = right else {
            #expect(Bool(false), "Expected COUNT on right")
            return
        }
    }
    
    @Test func parseCellReferencesInArithmetic() throws {
        let parser = FormulaParser("A1 + B2 * C3")
        let expr = try parser.parse()
        
        // Should parse as A1 + (B2 * C3)
        guard case .binaryOp(.add, .cellRef, let right) = expr else {
            #expect(Bool(false), "Expected addition at top")
            return
        }
        
        guard case .binaryOp(.multiply, .cellRef, .cellRef) = right else {
            #expect(Bool(false), "Expected multiplication on right")
            return
        }
    }
    
    @Test func parseStringConcatenation() throws {
        let parser = FormulaParser("\"Hello\" & \" \" & \"World\"")
        let expr = try parser.parse()
        
        // Should parse as ("Hello" & " ") & "World"
        guard case .binaryOp(.concat, let left, .string("World")) = expr else {
            #expect(Bool(false), "Expected concat at top")
            return
        }
        
        guard case .binaryOp(.concat, .string("Hello"), .string(" ")) = left else {
            #expect(Bool(false), "Expected concat on left")
            return
        }
    }
    
    // MARK: - Edge Cases
    
    @Test func parseLeadingEquals() throws {
        let parser = FormulaParser("=42")
        let expr = try parser.parse()
        #expect(expr == .number(42))
    }
    
    @Test func parseWhitespace() throws {
        let parser = FormulaParser("  2   +   3  ")
        let expr = try parser.parse()
        guard case .binaryOp(.add, .number(2), .number(3)) = expr else {
            #expect(Bool(false), "Expected addition")
            return
        }
    }
    
    @Test func parseUnaryMinus() throws {
        let parser = FormulaParser("-5")
        let expr = try parser.parse()
        
        // Unary minus is represented as -1 * 5
        guard case .binaryOp(.multiply, .number(-1), .number(5)) = expr else {
            #expect(Bool(false), "Expected unary minus")
            return
        }
    }
}
