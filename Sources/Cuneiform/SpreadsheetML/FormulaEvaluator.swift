import Foundation

/// Result of evaluating a formula expression
public enum FormulaValue: Sendable, Equatable {
    case number(Double)
    case string(String)
    case boolean(Bool)
    case error(String)
    case array([[FormulaValue]])
    
    public var asDouble: Double? {
        switch self {
        case .number(let n): return n
        case .boolean(let b): return b ? 1.0 : 0.0
        case .string(let s): return Double(s)
        default: return nil
        }
    }
    
    public var asString: String {
        switch self {
        case .number(let n): return String(n)
        case .string(let s): return s
        case .boolean(let b): return b ? "TRUE" : "FALSE"
        case .error(let e): return "#\(e)!"
        case .array: return "#ARRAY!"
        }
    }
    
    public var asBoolean: Bool? {
        switch self {
        case .boolean(let b): return b
        case .number(let n): return n != 0
        case .string(let s): return s.uppercased() == "TRUE"
        default: return nil
        }
    }
}

/// Evaluates parsed formula expressions in the context of a worksheet
public struct FormulaEvaluator: Sendable {
    private let cellResolver: @Sendable (CellReference) -> CellValue?
    
    /// Create evaluator with a cell resolver function
    public init(cellResolver: @escaping @Sendable (CellReference) -> CellValue?) {
        self.cellResolver = cellResolver
    }
    
    /// Evaluate a formula expression
    public func evaluate(_ expression: FormulaExpression) throws -> FormulaValue {
        switch expression {
        case .number(let value):
            return .number(value)
            
        case .string(let value):
            return .string(value)
            
        case .cellRef(let ref):
            guard let cell = cellResolver(ref) else {
                return .error("REF")
            }
            return cellValueToFormulaValue(cell)
            
        case .range(let start, let end):
            // Convert range to array
            var values: [[FormulaValue]] = []
            for row in start.row...end.row {
                var rowValues: [FormulaValue] = []
                // Iterate through column indices
                for col in start.columnIndex...end.columnIndex {
                    // Convert column index back to letter
                    let colLetter = columnLetterFrom(index: col)
                    let ref = CellReference(column: colLetter, row: row)
                    if let cell = cellResolver(ref) {
                        rowValues.append(cellValueToFormulaValue(cell))
                    } else {
                        rowValues.append(.number(0))
                    }
                }
                values.append(rowValues)
            }
            return .array(values)
            
        case .binaryOp(let op, let left, let right):
            return try evaluateBinaryOp(op, left, right)
            
        case .functionCall(let name, let args):
            return try evaluateFunction(name, args)
            
        case .error(let message):
            return .error(message)
        }
    }
    
    // MARK: - Binary Operations
    
    private func evaluateBinaryOp(_ op: FormulaExpression.BinaryOperator, _ left: FormulaExpression, _ right: FormulaExpression) throws -> FormulaValue {
        let leftVal = try evaluate(left)
        let rightVal = try evaluate(right)
        
        switch op {
        case .add:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l + r)
            
        case .subtract:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l - r)
            
        case .multiply:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(l * r)
            
        case .divide:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            guard r != 0 else {
                return .error("DIV/0")
            }
            return .number(l / r)
            
        case .power:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .number(pow(l, r))
            
        case .concat:
            return .string(leftVal.asString + rightVal.asString)
            
        case .equals:
            return .boolean(leftVal == rightVal)
            
        case .notEquals:
            return .boolean(leftVal != rightVal)
            
        case .lessThan:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l < r)
            
        case .greaterThan:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l > r)
            
        case .lessThanOrEqual:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l <= r)
            
        case .greaterThanOrEqual:
            guard let l = leftVal.asDouble, let r = rightVal.asDouble else {
                return .error("VALUE")
            }
            return .boolean(l >= r)
        }
    }
    
    // MARK: - Function Evaluation
    
    private func evaluateFunction(_ name: String, _ args: [FormulaExpression]) throws -> FormulaValue {
        let upperName = name.uppercased()
        
        switch upperName {
        case "SUM":
            return try evaluateSUM(args)
        case "AVERAGE":
            return try evaluateAVERAGE(args)
        case "IF":
            return try evaluateIF(args)
        case "VLOOKUP":
            return try evaluateVLOOKUP(args)
        case "INDEX":
            return try evaluateINDEX(args)
        case "MATCH":
            return try evaluateMATCH(args)
        case "MIN":
            return try evaluateMIN(args)
        case "MAX":
            return try evaluateMAX(args)
        case "COUNT":
            return try evaluateCOUNT(args)
        case "COUNTA":
            return try evaluateCOUNTA(args)
        case "AND":
            return try evaluateAND(args)
        case "OR":
            return try evaluateOR(args)
        case "NOT":
            return try evaluateNOT(args)
        case "LEN":
            return try evaluateLEN(args)
        case "UPPER":
            return try evaluateUPPER(args)
        case "LOWER":
            return try evaluateLOWER(args)
        case "CONCAT":
            return try evaluateCONCAT(args)
        case "ROUND":
            return try evaluateROUND(args)
        case "ABS":
            return try evaluateABS(args)
        case "MEDIAN":
            return try evaluateMEDIAN(args)
        default:
            return .error("NAME")
        }
    }
    
    // MARK: - Core Functions
    
    private func evaluateSUM(_ args: [FormulaExpression]) throws -> FormulaValue {
        var sum: Double = 0
        for arg in args {
            let val = try evaluate(arg)
            sum += flattenToNumbers(val).reduce(0, +)
        }
        return .number(sum)
    }
    
    private func evaluateAVERAGE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var sum: Double = 0
        var count = 0
        
        for arg in args {
            let val = try evaluate(arg)
            let numbers = flattenToNumbers(val)
            sum += numbers.reduce(0, +)
            count += numbers.count
        }
        
        guard count > 0 else {
            return .error("DIV/0")
        }
        
        return .number(sum / Double(count))
    }
    
    private func evaluateIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "IF", expected: 3, got: args.count)
        }
        
        let condition = try evaluate(args[0])
        guard let isTrue = condition.asBoolean else {
            return .error("VALUE")
        }
        
        if isTrue {
            return try evaluate(args[1])
        } else if args.count == 3 {
            return try evaluate(args[2])
        } else {
            return .boolean(false)
        }
    }
    
    private func evaluateVLOOKUP(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "VLOOKUP", expected: 4, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let tableArray = try evaluate(args[1])
        let colIndex = try evaluate(args[2])
        let rangeLookup = args.count == 4 ? try evaluate(args[3]) : .boolean(true)
        
        guard case .array(let table) = tableArray else {
            return .error("VALUE")
        }
        
        guard let colNum = colIndex.asDouble, colNum >= 1, Int(colNum) <= (table.first?.count ?? 0) else {
            return .error("VALUE")
        }
        
        let col = Int(colNum) - 1
        let exactMatch = !(rangeLookup.asBoolean ?? true)
        
        // Search first column for lookup value
        for row in table {
            guard !row.isEmpty else { continue }
            
            if exactMatch {
                if row[0] == lookupValue {
                    return row[col]
                }
            } else {
                // Approximate match (assumes sorted)
                if let rowVal = row[0].asDouble, let lookupVal = lookupValue.asDouble {
                    if rowVal <= lookupVal {
                        return row[col]
                    }
                }
            }
        }
        
        return .error("N/A")
    }
    
    private func evaluateINDEX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "INDEX", expected: 3, got: args.count)
        }
        
        let arrayVal = try evaluate(args[0])
        let rowNum = try evaluate(args[1])
        
        guard case .array(let array) = arrayVal else {
            return .error("VALUE")
        }
        
        guard let row = rowNum.asDouble, row >= 1, Int(row) <= array.count else {
            return .error("REF")
        }
        
        let rowIndex = Int(row) - 1
        
        if args.count == 3 {
            let colNum = try evaluate(args[2])
            guard let col = colNum.asDouble, col >= 1, Int(col) <= array[rowIndex].count else {
                return .error("REF")
            }
            let colIndex = Int(col) - 1
            return array[rowIndex][colIndex]
        } else {
            // Return entire row as array if no column specified
            return .array([array[rowIndex]])
        }
    }
    
    private func evaluateMATCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "MATCH", expected: 3, got: args.count)
        }
        
        let lookupValue = try evaluate(args[0])
        let lookupArray = try evaluate(args[1])
        let matchType = args.count == 3 ? try evaluate(args[2]) : .number(1)
        
        guard case .array(let array) = lookupArray else {
            return .error("VALUE")
        }
        
        let matchMode = Int(matchType.asDouble ?? 1)
        
        // Flatten array to 1D
        let values = array.flatMap { $0 }
        
        for (index, value) in values.enumerated() {
            if matchMode == 0 {
                // Exact match
                if value == lookupValue {
                    return .number(Double(index + 1))
                }
            } else if matchMode == 1 {
                // Less than or equal (assumes sorted ascending)
                if let valNum = value.asDouble, let lookupNum = lookupValue.asDouble {
                    if valNum <= lookupNum {
                        return .number(Double(index + 1))
                    }
                }
            }
        }
        
        return .error("N/A")
    }
    
    private func evaluateMIN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard let min = numbers.min() else {
            return .error("VALUE")
        }
        
        return .number(min)
    }
    
    private func evaluateMAX(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard let max = numbers.max() else {
            return .error("VALUE")
        }
        
        return .number(max)
    }
    
    private func evaluateCOUNT(_ args: [FormulaExpression]) throws -> FormulaValue {
        var count = 0
        for arg in args {
            let val = try evaluate(arg)
            count += flattenToNumbers(val).count
        }
        return .number(Double(count))
    }
    
    private func evaluateCOUNTA(_ args: [FormulaExpression]) throws -> FormulaValue {
        var count = 0
        for arg in args {
            let val = try evaluate(arg)
            count += flattenToAll(val).count
        }
        return .number(Double(count))
    }
    
    // MARK: - Logical Functions
    
    private func evaluateAND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        for arg in args {
            let val = try evaluate(arg)
            let bools = flattenToBooleans(val)
            for b in bools {
                if !b { return .boolean(false) }
            }
        }
        return .boolean(true)
    }
    
    private func evaluateOR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        for arg in args {
            let val = try evaluate(arg)
            let bools = flattenToBooleans(val)
            for b in bools {
                if b { return .boolean(true) }
            }
        }
        return .boolean(false)
    }
    
    private func evaluateNOT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "NOT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let bool = val.asBoolean else {
            return .error("VALUE")
        }
        return .boolean(!bool)
    }
    
    // MARK: - String Functions
    
    private func evaluateLEN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LEN", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        return .number(Double(str.count))
    }
    
    private func evaluateUPPER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "UPPER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        return .string(val.asString.uppercased())
    }
    
    private func evaluateLOWER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "LOWER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        return .string(val.asString.lowercased())
    }
    
    private func evaluateCONCAT(_ args: [FormulaExpression]) throws -> FormulaValue {
        var result = ""
        for arg in args {
            let val = try evaluate(arg)
            result += val.asString
        }
        return .string(result)
    }
    
    // MARK: - Math Functions
    
    private func evaluateROUND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "ROUND", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        let digits = args.count == 2 ? Int(try evaluate(args[1]).asDouble ?? 0) : 0
        let multiplier = pow(10.0, Double(digits))
        return .number((num * multiplier).rounded() / multiplier)
    }
    
    private func evaluateABS(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ABS", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        return .number(abs(num))
    }
    
    private func evaluateMEDIAN(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard !args.isEmpty else {
            return .error("VALUE")
        }
        
        var numbers: [Double] = []
        for arg in args {
            let val = try evaluate(arg)
            numbers.append(contentsOf: flattenToNumbers(val))
        }
        
        guard !numbers.isEmpty else {
            return .error("VALUE")
        }
        
        let sorted = numbers.sorted()
        let count = sorted.count
        
        if count % 2 == 0 {
            return .number((sorted[count / 2 - 1] + sorted[count / 2]) / 2.0)
        } else {
            return .number(sorted[count / 2])
        }
    }
    
    // MARK: - Helpers
    
    private func cellValueToFormulaValue(_ cell: CellValue) -> FormulaValue {
        switch cell {
        case .empty:
            return .number(0)
        case .number(let n):
            return .number(n)
        case .text(let s):
            return .string(s)
        case .boolean(let b):
            return .boolean(b)
        case .date(let d):
            // Date is stored as string in ISO format; return as string
            return .string(d)
        case .error(let e):
            return .error(e)
        case .richText(let rt):
            return .string(rt.plainText)
        }
    }
    
    private func flattenToNumbers(_ value: FormulaValue) -> [Double] {
        switch value {
        case .number(let n):
            return [n]
        case .boolean(let b):
            return [b ? 1.0 : 0.0]
        case .array(let arr):
            return arr.flatMap { row in
                row.compactMap { $0.asDouble }
            }
        default:
            return []
        }
    }
    
    private func flattenToAll(_ value: FormulaValue) -> [FormulaValue] {
        switch value {
        case .array(let arr):
            return arr.flatMap { $0 }
        case .error:
            return []
        default:
            return [value]
        }
    }
    
    private func flattenToBooleans(_ value: FormulaValue) -> [Bool] {
        switch value {
        case .boolean(let b):
            return [b]
        case .number(let n):
            return [n != 0]
        case .array(let arr):
            return arr.flatMap { row in
                row.compactMap { $0.asBoolean }
            }
        default:
            return []
        }
    }
}

// MARK: - Helper Functions

/// Convert 0-based column index to letter (0 = A, 1 = B, ..., 25 = Z, 26 = AA)
private func columnLetterFrom(index: Int) -> String {
    var col = index + 1
    var result = ""
    
    while col > 0 {
        col -= 1
        let remainder = col % 26
        result = String(UnicodeScalar(UInt8(65 + remainder))) + result
        col /= 26
    }
    
    return result
}
