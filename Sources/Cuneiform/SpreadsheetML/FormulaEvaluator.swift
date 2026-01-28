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
                        // Skip empty cells - matches Excel behavior where AVERAGE ignores empty cells
                        if case .empty = cell {
                            continue
                        }
                        rowValues.append(cellValueToFormulaValue(cell))
                    }
                    // Don't add anything for missing cells (skip them)
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
        // Date functions
        case "TODAY":
            return evaluateTODAY(args)
        case "NOW":
            return evaluateNOW(args)
        case "DATE":
            return try evaluateDATE(args)
        case "YEAR":
            return try evaluateYEAR(args)
        case "MONTH":
            return try evaluateMONTH(args)
        case "DAY":
            return try evaluateDAY(args)
        // String functions
        case "LEFT":
            return try evaluateLEFT(args)
        case "RIGHT":
            return try evaluateRIGHT(args)
        case "MID":
            return try evaluateMID(args)
        case "TRIM":
            return try evaluateTRIM(args)
        // Conditional aggregates
        case "SUMIF":
            return try evaluateSUMIF(args)
        case "COUNTIF":
            return try evaluateCOUNTIF(args)
        case "IFERROR":
            return try evaluateIFERROR(args)
        // Additional math
        case "INT":
            return try evaluateINT(args)
        case "MOD":
            return try evaluateMOD(args)
        case "SQRT":
            return try evaluateSQRT(args)
        // Tier 2: Multi-criteria aggregates
        case "SUMIFS":
            return try evaluateSUMIFS(args)
        case "COUNTIFS":
            return try evaluateCOUNTIFS(args)
        case "AVERAGEIF":
            return try evaluateAVERAGEIF(args)
        // Tier 2: String search/manipulation
        case "FIND":
            return try evaluateFIND(args)
        case "SEARCH":
            return try evaluateSEARCH(args)
        case "SUBSTITUTE":
            return try evaluateSUBSTITUTE(args)
        case "CONCATENATE":
            return try evaluateCONCATENATE(args)
        case "TEXT":
            return try evaluateTEXT(args)
        // Tier 2: Type checking
        case "ISBLANK":
            return try evaluateISBLANK(args)
        case "ISNUMBER":
            return try evaluateISNUMBER(args)
        case "ISTEXT":
            return try evaluateISTEXT(args)
        case "ISERROR":
            return try evaluateISERROR(args)
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
    
    // MARK: - Date Functions
    
    /// Returns today's date as an Excel serial number
    private func evaluateTODAY(_ args: [FormulaExpression]) -> FormulaValue {
        // Excel serial date: days since 1899-12-30 (accounting for the 1900 leap year bug)
        let now = Date()
        let calendar = Calendar(identifier: .gregorian)
        let components = calendar.dateComponents([.year, .month, .day], from: now)
        return dateToSerial(year: components.year!, month: components.month!, day: components.day!)
    }
    
    /// Returns current date and time as an Excel serial number
    private func evaluateNOW(_ args: [FormulaExpression]) -> FormulaValue {
        let now = Date()
        let calendar = Calendar(identifier: .gregorian)
        let components = calendar.dateComponents([.year, .month, .day, .hour, .minute, .second], from: now)
        
        // Get the date portion
        guard case .number(let datePart) = dateToSerial(year: components.year!, month: components.month!, day: components.day!) else {
            return .error("VALUE")
        }
        
        // Add time as fraction of day
        let timeFraction = Double(components.hour!) / 24.0 +
                          Double(components.minute!) / 1440.0 +
                          Double(components.second!) / 86400.0
        
        return .number(datePart + timeFraction)
    }
    
    /// Creates a date serial number from year, month, day
    private func evaluateDATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "DATE", expected: 3, got: args.count)
        }
        
        let yearVal = try evaluate(args[0])
        let monthVal = try evaluate(args[1])
        let dayVal = try evaluate(args[2])
        
        guard let year = yearVal.asDouble,
              let month = monthVal.asDouble,
              let day = dayVal.asDouble else {
            return .error("VALUE")
        }
        
        return dateToSerial(year: Int(year), month: Int(month), day: Int(day))
    }
    
    /// Extracts year from a date serial number
    private func evaluateYEAR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "YEAR", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (year, _, _) = serialToDate(serial: serial)
        return .number(Double(year))
    }
    
    /// Extracts month from a date serial number
    private func evaluateMONTH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "MONTH", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (_, month, _) = serialToDate(serial: serial)
        return .number(Double(month))
    }
    
    /// Extracts day from a date serial number
    private func evaluateDAY(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "DAY", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let serial = val.asDouble else {
            return .error("VALUE")
        }
        
        let (_, _, day) = serialToDate(serial: serial)
        return .number(Double(day))
    }
    
    /// Convert year/month/day to Excel serial number
    private func dateToSerial(year: Int, month: Int, day: Int) -> FormulaValue {
        // Handle year values 0-99 as 1900-1999, 100-9999 as-is (Excel behavior)
        var adjustedYear = year
        if year >= 0 && year <= 99 {
            adjustedYear = 1900 + year
        }
        
        // Handle month overflow/underflow
        var adjMonth = month
        var adjYear = adjustedYear
        while adjMonth > 12 {
            adjMonth -= 12
            adjYear += 1
        }
        while adjMonth < 1 {
            adjMonth += 12
            adjYear -= 1
        }
        
        // Create date components
        var components = DateComponents()
        components.year = adjYear
        components.month = adjMonth
        components.day = day
        
        let calendar = Calendar(identifier: .gregorian)
        guard let date = calendar.date(from: components) else {
            return .error("VALUE")
        }
        
        // Excel epoch: 1899-12-30 (due to the 1900 leap year bug)
        var epochComponents = DateComponents()
        epochComponents.year = 1899
        epochComponents.month = 12
        epochComponents.day = 30
        
        guard let epoch = calendar.date(from: epochComponents) else {
            return .error("VALUE")
        }
        
        let days = calendar.dateComponents([.day], from: epoch, to: date).day!
        
        // Excel incorrectly treats 1900 as a leap year, so dates after Feb 28, 1900 need +1
        if days > 59 {
            return .number(Double(days + 1))
        }
        return .number(Double(days))
    }
    
    /// Convert Excel serial number to year/month/day
    private func serialToDate(serial: Double) -> (year: Int, month: Int, day: Int) {
        var adjustedSerial = Int(serial)
        
        // Account for Excel's 1900 leap year bug
        if adjustedSerial > 60 {
            adjustedSerial -= 1
        }
        
        var components = DateComponents()
        components.year = 1899
        components.month = 12
        components.day = 30
        
        let calendar = Calendar(identifier: .gregorian)
        guard let epoch = calendar.date(from: components),
              let date = calendar.date(byAdding: .day, value: adjustedSerial, to: epoch) else {
            return (1900, 1, 1)
        }
        
        let resultComponents = calendar.dateComponents([.year, .month, .day], from: date)
        return (resultComponents.year!, resultComponents.month!, resultComponents.day!)
    }
    
    // MARK: - Extended String Functions
    
    /// Returns leftmost characters from a string
    private func evaluateLEFT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "LEFT", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        let numChars: Int
        if args.count == 2 {
            guard let n = try evaluate(args[1]).asDouble, n >= 0 else {
                return .error("VALUE")
            }
            numChars = Int(n)
        } else {
            numChars = 1
        }
        
        if numChars >= str.count {
            return .string(str)
        }
        
        let endIndex = str.index(str.startIndex, offsetBy: numChars)
        return .string(String(str[..<endIndex]))
    }
    
    /// Returns rightmost characters from a string
    private func evaluateRIGHT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 1 && args.count <= 2 else {
            throw FormulaError.invalidArgumentCount(function: "RIGHT", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        let numChars: Int
        if args.count == 2 {
            guard let n = try evaluate(args[1]).asDouble, n >= 0 else {
                return .error("VALUE")
            }
            numChars = Int(n)
        } else {
            numChars = 1
        }
        
        if numChars >= str.count {
            return .string(str)
        }
        
        let startIndex = str.index(str.endIndex, offsetBy: -numChars)
        return .string(String(str[startIndex...]))
    }
    
    /// Returns characters from the middle of a string
    private func evaluateMID(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 3 else {
            throw FormulaError.invalidArgumentCount(function: "MID", expected: 3, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        guard let startNum = try evaluate(args[1]).asDouble,
              let numChars = try evaluate(args[2]).asDouble,
              startNum >= 1, numChars >= 0 else {
            return .error("VALUE")
        }
        
        let startIndex = Int(startNum) - 1 // Excel is 1-based
        let length = Int(numChars)
        
        guard startIndex < str.count else {
            return .string("")
        }
        
        let start = str.index(str.startIndex, offsetBy: startIndex)
        let end = str.index(start, offsetBy: min(length, str.count - startIndex))
        return .string(String(str[start..<end]))
    }
    
    /// Removes leading and trailing spaces, and reduces internal spaces to single spaces
    private func evaluateTRIM(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "TRIM", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        let str = val.asString
        
        // Excel's TRIM removes leading/trailing spaces and collapses multiple spaces to single
        let components = str.split(separator: " ", omittingEmptySubsequences: true)
        return .string(components.joined(separator: " "))
    }
    
    // MARK: - Conditional Aggregate Functions
    
    /// Sum cells that match a criteria
    private func evaluateSUMIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "SUMIF", expected: 3, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        // If sum_range is provided, use it; otherwise use the criteria range
        let sumRangeVal: FormulaValue
        if args.count == 3 {
            sumRangeVal = try evaluate(args[2])
        } else {
            sumRangeVal = rangeVal
        }
        
        guard case .array(let criteriaRange) = rangeVal else {
            // Single cell
            if matchesCriteria(rangeVal, criteriaVal) {
                return sumRangeVal.asDouble.map { .number($0) } ?? .number(0)
            }
            return .number(0)
        }
        
        let sumRange: [[FormulaValue]]
        if case .array(let arr) = sumRangeVal {
            sumRange = arr
        } else {
            sumRange = [[sumRangeVal]]
        }
        
        var sum: Double = 0
        for (rowIdx, row) in criteriaRange.enumerated() {
            for (colIdx, cell) in row.enumerated() {
                if matchesCriteria(cell, criteriaVal) {
                    // Get corresponding value from sum range
                    if rowIdx < sumRange.count && colIdx < sumRange[rowIdx].count {
                        if let num = sumRange[rowIdx][colIdx].asDouble {
                            sum += num
                        }
                    }
                }
            }
        }
        
        return .number(sum)
    }
    
    /// Count cells that match a criteria
    private func evaluateCOUNTIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "COUNTIF", expected: 2, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        guard case .array(let range) = rangeVal else {
            // Single cell
            return .number(matchesCriteria(rangeVal, criteriaVal) ? 1 : 0)
        }
        
        var count = 0
        for row in range {
            for cell in row {
                if matchesCriteria(cell, criteriaVal) {
                    count += 1
                }
            }
        }
        
        return .number(Double(count))
    }
    
    /// Returns value if no error, otherwise returns alternate value
    private func evaluateIFERROR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "IFERROR", expected: 2, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        if case .error = val {
            return try evaluate(args[1])
        }
        
        return val
    }
    
    /// Helper to check if a value matches a criteria
    private func matchesCriteria(_ value: FormulaValue, _ criteria: FormulaValue) -> Bool {
        let criteriaStr = criteria.asString
        
        // Check for comparison operators in criteria string
        if criteriaStr.hasPrefix(">=") {
            guard let threshold = Double(String(criteriaStr.dropFirst(2))),
                  let val = value.asDouble else { return false }
            return val >= threshold
        } else if criteriaStr.hasPrefix("<=") {
            guard let threshold = Double(String(criteriaStr.dropFirst(2))),
                  let val = value.asDouble else { return false }
            return val <= threshold
        } else if criteriaStr.hasPrefix("<>") {
            let compareStr = String(criteriaStr.dropFirst(2))
            if let threshold = Double(compareStr), let val = value.asDouble {
                return val != threshold
            }
            return value.asString != compareStr
        } else if criteriaStr.hasPrefix(">") {
            guard let threshold = Double(String(criteriaStr.dropFirst(1))),
                  let val = value.asDouble else { return false }
            return val > threshold
        } else if criteriaStr.hasPrefix("<") {
            guard let threshold = Double(String(criteriaStr.dropFirst(1))),
                  let val = value.asDouble else { return false }
            return val < threshold
        } else if criteriaStr.hasPrefix("=") {
            let compareStr = String(criteriaStr.dropFirst(1))
            if let threshold = Double(compareStr), let val = value.asDouble {
                return val == threshold
            }
            return value.asString == compareStr
        } else if criteriaStr.contains("*") || criteriaStr.contains("?") {
            // Wildcard matching
            return wildcardMatch(value.asString, pattern: criteriaStr)
        } else {
            // Direct comparison
            if let criteriaNum = criteria.asDouble, let valNum = value.asDouble {
                return valNum == criteriaNum
            }
            return value.asString == criteriaStr
        }
    }
    
    /// Simple wildcard matching (* and ?)
    private func wildcardMatch(_ string: String, pattern: String) -> Bool {
        let regexPattern = "^" + NSRegularExpression.escapedPattern(for: pattern)
            .replacingOccurrences(of: "\\*", with: ".*")
            .replacingOccurrences(of: "\\?", with: ".") + "$"
        
        guard let regex = try? NSRegularExpression(pattern: regexPattern, options: .caseInsensitive) else {
            return false
        }
        
        let range = NSRange(string.startIndex..., in: string)
        return regex.firstMatch(in: string, options: [], range: range) != nil
    }
    
    // MARK: - Additional Math Functions
    
    /// Truncates a number to an integer
    private func evaluateINT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "INT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        // Excel's INT rounds toward negative infinity
        return .number(floor(num))
    }
    
    /// Returns the remainder after division
    private func evaluateMOD(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "MOD", expected: 2, got: args.count)
        }
        
        let numVal = try evaluate(args[0])
        let divisorVal = try evaluate(args[1])
        
        guard let num = numVal.asDouble, let divisor = divisorVal.asDouble else {
            return .error("VALUE")
        }
        
        guard divisor != 0 else {
            return .error("DIV/0")
        }
        
        // Excel's MOD: result has the same sign as divisor
        let result = num - divisor * floor(num / divisor)
        return .number(result)
    }
    
    /// Returns the square root of a number
    private func evaluateSQRT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SQRT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        guard let num = val.asDouble else {
            return .error("VALUE")
        }
        
        guard num >= 0 else {
            return .error("NUM")
        }
        
        return .number(sqrt(num))
    }
    
    // MARK: - Tier 2: Multi-Criteria Aggregate Functions
    
    /// SUMIFS - Sum cells that match multiple criteria
    private func evaluateSUMIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
        guard args.count >= 3 && args.count % 2 == 1 else {
            throw FormulaError.invalidArgumentCount(function: "SUMIFS", expected: 3, got: args.count)
        }
        
        let sumRangeVal = try evaluate(args[0])
        let sumRange: [[FormulaValue]]
        if case .array(let arr) = sumRangeVal {
            sumRange = arr
        } else {
            sumRange = [[sumRangeVal]]
        }
        
        // Collect criteria pairs
        var criteriaPairs: [(range: [[FormulaValue]], criteria: FormulaValue)] = []
        var i = 1
        while i < args.count {
            let rangeVal = try evaluate(args[i])
            let criteriaVal = try evaluate(args[i + 1])
            
            let range: [[FormulaValue]]
            if case .array(let arr) = rangeVal {
                range = arr
            } else {
                range = [[rangeVal]]
            }
            criteriaPairs.append((range: range, criteria: criteriaVal))
            i += 2
        }
        
        var sum: Double = 0
        
        // Iterate through sum range
        for (rowIdx, row) in sumRange.enumerated() {
            for (colIdx, _) in row.enumerated() {
                // Check all criteria
                var allMatch = true
                for pair in criteriaPairs {
                    if rowIdx < pair.range.count && colIdx < pair.range[rowIdx].count {
                        if !matchesCriteria(pair.range[rowIdx][colIdx], pair.criteria) {
                            allMatch = false
                            break
                        }
                    } else {
                        allMatch = false
                        break
                    }
                }
                
                if allMatch {
                    if rowIdx < sumRange.count && colIdx < sumRange[rowIdx].count {
                        if let num = sumRange[rowIdx][colIdx].asDouble {
                            sum += num
                        }
                    }
                }
            }
        }
        
        return .number(sum)
    }
    
    /// COUNTIFS - Count cells that match multiple criteria
    private func evaluateCOUNTIFS(_ args: [FormulaExpression]) throws -> FormulaValue {
        // COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
        guard args.count >= 2 && args.count % 2 == 0 else {
            throw FormulaError.invalidArgumentCount(function: "COUNTIFS", expected: 2, got: args.count)
        }
        
        // Get first range to determine dimensions
        let firstRangeVal = try evaluate(args[0])
        let firstRange: [[FormulaValue]]
        if case .array(let arr) = firstRangeVal {
            firstRange = arr
        } else {
            firstRange = [[firstRangeVal]]
        }
        
        // Collect all criteria pairs
        var criteriaPairs: [(range: [[FormulaValue]], criteria: FormulaValue)] = []
        var i = 0
        while i < args.count {
            let rangeVal = try evaluate(args[i])
            let criteriaVal = try evaluate(args[i + 1])
            
            let range: [[FormulaValue]]
            if case .array(let arr) = rangeVal {
                range = arr
            } else {
                range = [[rangeVal]]
            }
            criteriaPairs.append((range: range, criteria: criteriaVal))
            i += 2
        }
        
        var count = 0
        
        // Iterate through first range dimensions
        for rowIdx in 0..<firstRange.count {
            for colIdx in 0..<(firstRange.first?.count ?? 0) {
                var allMatch = true
                for pair in criteriaPairs {
                    if rowIdx < pair.range.count && colIdx < pair.range[rowIdx].count {
                        if !matchesCriteria(pair.range[rowIdx][colIdx], pair.criteria) {
                            allMatch = false
                            break
                        }
                    } else {
                        allMatch = false
                        break
                    }
                }
                
                if allMatch {
                    count += 1
                }
            }
        }
        
        return .number(Double(count))
    }
    
    /// AVERAGEIF - Average cells that match a criteria
    private func evaluateAVERAGEIF(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "AVERAGEIF", expected: 3, got: args.count)
        }
        
        let rangeVal = try evaluate(args[0])
        let criteriaVal = try evaluate(args[1])
        
        // If average_range is provided, use it; otherwise use the criteria range
        let avgRangeVal: FormulaValue
        if args.count == 3 {
            avgRangeVal = try evaluate(args[2])
        } else {
            avgRangeVal = rangeVal
        }
        
        guard case .array(let criteriaRange) = rangeVal else {
            // Single cell
            if matchesCriteria(rangeVal, criteriaVal) {
                return avgRangeVal.asDouble.map { .number($0) } ?? .error("DIV/0")
            }
            return .error("DIV/0")
        }
        
        let avgRange: [[FormulaValue]]
        if case .array(let arr) = avgRangeVal {
            avgRange = arr
        } else {
            avgRange = [[avgRangeVal]]
        }
        
        var sum: Double = 0
        var count = 0
        
        for (rowIdx, row) in criteriaRange.enumerated() {
            for (colIdx, cell) in row.enumerated() {
                if matchesCriteria(cell, criteriaVal) {
                    if rowIdx < avgRange.count && colIdx < avgRange[rowIdx].count {
                        if let num = avgRange[rowIdx][colIdx].asDouble {
                            sum += num
                            count += 1
                        }
                    }
                }
            }
        }
        
        guard count > 0 else {
            return .error("DIV/0")
        }
        
        return .number(sum / Double(count))
    }
    
    // MARK: - Tier 2: String Search & Manipulation
    
    /// FIND - Find text within text (case-sensitive)
    private func evaluateFIND(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "FIND", expected: 2, got: args.count)
        }
        
        let findTextVal = try evaluate(args[0])
        let withinTextVal = try evaluate(args[1])
        
        let findText = findTextVal.asString
        let withinText = withinTextVal.asString
        
        let startNum: Int
        if args.count == 3 {
            guard let n = try evaluate(args[2]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            startNum = Int(n)
        } else {
            startNum = 1
        }
        
        guard startNum <= withinText.count else {
            return .error("VALUE")
        }
        
        let searchStartIndex = withinText.index(withinText.startIndex, offsetBy: startNum - 1)
        let searchRange = searchStartIndex..<withinText.endIndex
        
        if let range = withinText.range(of: findText, range: searchRange) {
            let position = withinText.distance(from: withinText.startIndex, to: range.lowerBound) + 1
            return .number(Double(position))
        }
        
        return .error("VALUE")
    }
    
    /// SEARCH - Find text within text (case-insensitive, supports wildcards)
    private func evaluateSEARCH(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 2 && args.count <= 3 else {
            throw FormulaError.invalidArgumentCount(function: "SEARCH", expected: 2, got: args.count)
        }
        
        let findTextVal = try evaluate(args[0])
        let withinTextVal = try evaluate(args[1])
        
        let findText = findTextVal.asString
        let withinText = withinTextVal.asString
        
        let startNum: Int
        if args.count == 3 {
            guard let n = try evaluate(args[2]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            startNum = Int(n)
        } else {
            startNum = 1
        }
        
        guard startNum <= withinText.count else {
            return .error("VALUE")
        }
        
        let searchStartIndex = withinText.index(withinText.startIndex, offsetBy: startNum - 1)
        let searchString = String(withinText[searchStartIndex...])
        
        // Handle wildcards in find text
        if findText.contains("*") || findText.contains("?") {
            let pattern = "^(.*?)(" + NSRegularExpression.escapedPattern(for: findText)
                .replacingOccurrences(of: "\\*", with: ".*")
                .replacingOccurrences(of: "\\?", with: ".") + ")"
            
            if let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
               let match = regex.firstMatch(in: searchString, options: [], range: NSRange(searchString.startIndex..., in: searchString)) {
                let position = startNum + match.range(at: 1).length
                return .number(Double(position))
            }
        } else {
            // Simple case-insensitive search
            if let range = searchString.range(of: findText, options: .caseInsensitive) {
                let position = startNum + searchString.distance(from: searchString.startIndex, to: range.lowerBound)
                return .number(Double(position))
            }
        }
        
        return .error("VALUE")
    }
    
    /// SUBSTITUTE - Replace occurrences of text
    private func evaluateSUBSTITUTE(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count >= 3 && args.count <= 4 else {
            throw FormulaError.invalidArgumentCount(function: "SUBSTITUTE", expected: 3, got: args.count)
        }
        
        let textVal = try evaluate(args[0])
        let oldTextVal = try evaluate(args[1])
        let newTextVal = try evaluate(args[2])
        
        var text = textVal.asString
        let oldText = oldTextVal.asString
        let newText = newTextVal.asString
        
        if args.count == 4 {
            // Replace only the nth occurrence
            guard let n = try evaluate(args[3]).asDouble, n >= 1 else {
                return .error("VALUE")
            }
            let instanceNum = Int(n)
            
            var occurrence = 0
            var searchStart = text.startIndex
            
            while let range = text.range(of: oldText, range: searchStart..<text.endIndex) {
                occurrence += 1
                if occurrence == instanceNum {
                    text.replaceSubrange(range, with: newText)
                    break
                }
                searchStart = range.upperBound
            }
        } else {
            // Replace all occurrences
            text = text.replacingOccurrences(of: oldText, with: newText)
        }
        
        return .string(text)
    }
    
    /// CONCATENATE - Join multiple text strings
    private func evaluateCONCATENATE(_ args: [FormulaExpression]) throws -> FormulaValue {
        var result = ""
        for arg in args {
            let val = try evaluate(arg)
            result += val.asString
        }
        return .string(result)
    }
    
    /// TEXT - Format a number as text with a format string
    private func evaluateTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 2 else {
            throw FormulaError.invalidArgumentCount(function: "TEXT", expected: 2, got: args.count)
        }
        
        let valueArg = try evaluate(args[0])
        let formatArg = try evaluate(args[1])
        
        guard let num = valueArg.asDouble else {
            return .string(valueArg.asString)
        }
        
        let formatStr = formatArg.asString
        
        // Handle common Excel format patterns
        let result: String
        
        if formatStr.lowercased().contains("0.00%") || formatStr == "0%" {
            // Percentage format
            let percentage = num * 100
            let decimals = formatStr.contains(".00") ? 2 : 0
            result = String(format: "%.\(decimals)f%%", percentage)
        } else if formatStr.contains("$") || formatStr.lowercased().contains("#,##0") {
            // Currency/number format with thousands separator
            let formatter = NumberFormatter()
            formatter.numberStyle = .decimal
            formatter.groupingSeparator = ","
            
            // Count decimal places
            if let decimalPart = formatStr.split(separator: ".").last {
                formatter.minimumFractionDigits = decimalPart.filter { $0 == "0" }.count
                formatter.maximumFractionDigits = decimalPart.filter { $0 == "0" }.count
            } else {
                formatter.minimumFractionDigits = 0
                formatter.maximumFractionDigits = 0
            }
            
            let formatted = formatter.string(from: NSNumber(value: num)) ?? String(num)
            if formatStr.contains("$") {
                result = "$" + formatted
            } else {
                result = formatted
            }
        } else if formatStr.lowercased().contains("yyyy") || formatStr.lowercased().contains("mm") || formatStr.lowercased().contains("dd") {
            // Date format - convert serial to date string
            let (year, month, day) = serialToDate(serial: num)
            var output = formatStr
            output = output.replacingOccurrences(of: "yyyy", with: String(format: "%04d", year))
            output = output.replacingOccurrences(of: "YYYY", with: String(format: "%04d", year))
            output = output.replacingOccurrences(of: "yy", with: String(format: "%02d", year % 100))
            output = output.replacingOccurrences(of: "YY", with: String(format: "%02d", year % 100))
            output = output.replacingOccurrences(of: "mm", with: String(format: "%02d", month))
            output = output.replacingOccurrences(of: "MM", with: String(format: "%02d", month))
            output = output.replacingOccurrences(of: "m", with: String(month))
            output = output.replacingOccurrences(of: "M", with: String(month))
            output = output.replacingOccurrences(of: "dd", with: String(format: "%02d", day))
            output = output.replacingOccurrences(of: "DD", with: String(format: "%02d", day))
            output = output.replacingOccurrences(of: "d", with: String(day))
            output = output.replacingOccurrences(of: "D", with: String(day))
            result = output
        } else if formatStr.contains("0") || formatStr.contains("#") {
            // General number format
            let decimalCount = formatStr.split(separator: ".").last?.filter { $0 == "0" }.count ?? 0
            result = String(format: "%.\(decimalCount)f", num)
        } else {
            result = String(num)
        }
        
        return .string(result)
    }
    
    // MARK: - Tier 2: Type Checking Functions
    
    /// ISBLANK - Check if a cell is blank
    private func evaluateISBLANK(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISBLANK", expected: 1, got: args.count)
        }
        
        // For cell references, check if cell is empty
        if case .cellRef(let ref) = args[0] {
            if let value = cellResolver(ref) {
                switch value {
                case .empty:
                    return .boolean(true)
                default:
                    return .boolean(false)
                }
            }
            // Cell not found = blank
            return .boolean(true)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .string(let s):
            return .boolean(s.isEmpty)
        case .error:
            return .boolean(false)
        default:
            return .boolean(false)
        }
    }
    
    /// ISNUMBER - Check if a value is a number
    private func evaluateISNUMBER(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISNUMBER", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .number:
            return .boolean(true)
        default:
            return .boolean(false)
        }
    }
    
    /// ISTEXT - Check if a value is text
    private func evaluateISTEXT(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISTEXT", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .string:
            return .boolean(true)
        default:
            return .boolean(false)
        }
    }
    
    /// ISERROR - Check if a value is an error
    private func evaluateISERROR(_ args: [FormulaExpression]) throws -> FormulaValue {
        guard args.count == 1 else {
            throw FormulaError.invalidArgumentCount(function: "ISERROR", expected: 1, got: args.count)
        }
        
        let val = try evaluate(args[0])
        
        switch val {
        case .error:
            return .boolean(true)
        default:
            return .boolean(false)
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
