//
// ExcelFunctionLibrary.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Excel function signature
typealias ExcelFunction = ([ExcelValue]) throws -> ExcelValue

/// Library of Excel functions (Tier 1: MVP functions)
class ExcelFunctionLibrary {
    private var functions: [String: ExcelFunction] = [:]

    init() {
        registerCoreFunctions()
    }

    /// Register a function
    func register(name: String, function: @escaping ExcelFunction) {
        functions[name.uppercased()] = function
    }

    /// Get a function by name
    func get(_ name: String) -> ExcelFunction? {
        return functions[name.uppercased()]
    }

    // MARK: - Core Function Registration

    private func registerCoreFunctions() {
        // Mathematical functions
        register(name: "SUM", function: sum)
        register(name: "AVERAGE", function: average)
        register(name: "COUNT", function: count)
        register(name: "COUNTA", function: counta)
        register(name: "MIN", function: min)
        register(name: "MAX", function: max)
        register(name: "ABS", function: abs)
        register(name: "ROUND", function: round)
        register(name: "INT", function: int)

        // Logical functions
        register(name: "IF", function: ifFunction)
        register(name: "AND", function: andFunction)
        register(name: "OR", function: orFunction)
        register(name: "NOT", function: notFunction)

        // Text functions
        register(name: "CONCATENATE", function: concatenate)
        register(name: "LEFT", function: left)
        register(name: "RIGHT", function: right)
        register(name: "MID", function: mid)
        register(name: "LEN", function: len)
        register(name: "UPPER", function: upper)
        register(name: "LOWER", function: lower)
        register(name: "TRIM", function: trim)

        // Lookup functions (basic)
        register(name: "INDEX", function: index)

        // Tier 2: Commonly used functions
        register(name: "VLOOKUP", function: vlookup)
        register(name: "MATCH", function: match)
        register(name: "SUMIF", function: sumif)
        register(name: "COUNTIF", function: countif)
        register(name: "IFERROR", function: iferror)
        register(name: "MOD", function: mod)
        register(name: "SQRT", function: sqrt)
        register(name: "POWER", function: power)

        // Tier 3: Advanced functions
        register(name: "AVERAGEIF", function: averageif)
        register(name: "SUMIFS", function: sumifs)
        register(name: "COUNTIFS", function: countifs)
        register(name: "FIND", function: find)
        register(name: "SEARCH", function: search)
        register(name: "SUBSTITUTE", function: substitute)
        register(name: "TEXT", function: text)
        register(name: "VALUE", function: value)
        register(name: "ISBLANK", function: isblank)
        register(name: "ISNUMBER", function: isnumber)
        register(name: "ISTEXT", function: istext)
    }

    // MARK: - Mathematical Functions

    private func sum(_ args: [ExcelValue]) throws -> ExcelValue {
        var total = 0.0

        for arg in args {
            if let num = arg.toNumber() {
                total += num
            }
        }

        return .number(total)
    }

    private func average(_ args: [ExcelValue]) throws -> ExcelValue {
        guard !args.isEmpty else {
            return .error(.divideByZero)
        }

        var total = 0.0
        var count = 0

        for arg in args {
            if let num = arg.toNumber() {
                total += num
                count += 1
            }
        }

        guard count > 0 else {
            return .error(.divideByZero)
        }

        return .number(total / Double(count))
    }

    private func count(_ args: [ExcelValue]) throws -> ExcelValue {
        let count = args.filter { $0.isNumeric }.count
        return .number(Double(count))
    }

    private func counta(_ args: [ExcelValue]) throws -> ExcelValue {
        let count = args.filter { !($0 == .empty) }.count
        return .number(Double(count))
    }

    private func min(_ args: [ExcelValue]) throws -> ExcelValue {
        let numbers = args.compactMap { $0.toNumber() }

        guard !numbers.isEmpty else {
            return .number(0)
        }

        return .number(numbers.min() ?? 0)
    }

    private func max(_ args: [ExcelValue]) throws -> ExcelValue {
        let numbers = args.compactMap { $0.toNumber() }

        guard !numbers.isEmpty else {
            return .number(0)
        }

        return .number(numbers.max() ?? 0)
    }

    private func abs(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        guard let num = args[0].toNumber() else {
            return .error(.value)
        }

        return .number(Swift.abs(num))
    }

    private func round(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 2 else {
            return .error(.value)
        }

        guard let num = args[0].toNumber(),
              let digits = args[1].toNumber() else {
            return .error(.value)
        }

        let multiplier = pow(10.0, digits)
        return .number(Darwin.round(num * multiplier) / multiplier)
    }

    private func int(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        guard let num = args[0].toNumber() else {
            return .error(.value)
        }

        return .number(floor(num))
    }

    // MARK: - Logical Functions

    private func ifFunction(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count >= 2 && args.count <= 3 else {
            return .error(.value)
        }

        guard let condition = args[0].toBoolean() else {
            return .error(.value)
        }

        if condition {
            return args[1]
        } else if args.count == 3 {
            return args[2]
        } else {
            return .boolean(false)
        }
    }

    private func andFunction(_ args: [ExcelValue]) throws -> ExcelValue {
        guard !args.isEmpty else {
            return .error(.value)
        }

        for arg in args {
            guard let bool = arg.toBoolean() else {
                return .error(.value)
            }
            if !bool {
                return .boolean(false)
            }
        }

        return .boolean(true)
    }

    private func orFunction(_ args: [ExcelValue]) throws -> ExcelValue {
        guard !args.isEmpty else {
            return .error(.value)
        }

        for arg in args {
            guard let bool = arg.toBoolean() else {
                return .error(.value)
            }
            if bool {
                return .boolean(true)
            }
        }

        return .boolean(false)
    }

    private func notFunction(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        guard let bool = args[0].toBoolean() else {
            return .error(.value)
        }

        return .boolean(!bool)
    }

    // MARK: - Text Functions

    private func concatenate(_ args: [ExcelValue]) throws -> ExcelValue {
        let text = args.map { $0.toString() }.joined()
        return .string(text)
    }

    private func left(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error(.value)
        }

        let text = args[0].toString()
        let numChars = args.count == 2 ? Int(args[1].toNumber() ?? 1) : 1

        guard numChars >= 0 else {
            return .error(.value)
        }

        let endIndex = text.index(text.startIndex, offsetBy: Swift.min(numChars, text.count))
        return .string(String(text[..<endIndex]))
    }

    private func right(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count >= 1 && args.count <= 2 else {
            return .error(.value)
        }

        let text = args[0].toString()
        let numChars = args.count == 2 ? Int(args[1].toNumber() ?? 1) : 1

        guard numChars >= 0 else {
            return .error(.value)
        }

        let startIndex = text.index(text.endIndex, offsetBy: -Swift.min(numChars, text.count))
        return .string(String(text[startIndex...]))
    }

    private func mid(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 3 else {
            return .error(.value)
        }

        let text = args[0].toString()
        guard let start = args[1].toNumber(),
              let numChars = args[2].toNumber() else {
            return .error(.value)
        }

        let startIndex = Int(start) - 1  // Excel is 1-indexed
        guard startIndex >= 0 && numChars >= 0 else {
            return .error(.value)
        }

        guard startIndex < text.count else {
            return .string("")
        }

        let begin = text.index(text.startIndex, offsetBy: startIndex)
        let end = text.index(begin, offsetBy: Swift.min(Int(numChars), text.count - startIndex))

        return .string(String(text[begin..<end]))
    }

    private func len(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        let text = args[0].toString()
        return .number(Double(text.count))
    }

    private func upper(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        let text = args[0].toString()
        return .string(text.uppercased())
    }

    private func lower(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        let text = args[0].toString()
        return .string(text.lowercased())
    }

    private func trim(_ args: [ExcelValue]) throws -> ExcelValue {
        guard args.count == 1 else {
            return .error(.value)
        }

        let text = args[0].toString()
        // Excel TRIM removes leading/trailing spaces and reduces multiple spaces to single
        let trimmed = text.trimmingCharacters(in: .whitespaces)
        let singleSpaced = trimmed.replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)
        return .string(singleSpaced)
    }

    // MARK: - Lookup Functions (Basic)

    private func index(_ args: [ExcelValue]) throws -> ExcelValue {
        // Simple INDEX implementation for single values
        // Full implementation would handle arrays
        guard args.count >= 1 else {
            return .error(.value)
        }

        // For now, just return the first argument
        // Full implementation would parse array and index into it
        return args[0]
    }

    // MARK: - Tier 2 Functions

    private func vlookup(_ args: [ExcelValue]) throws -> ExcelValue {
        // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        // Simplified implementation for basic cases
        guard args.count >= 3 else {
            return .error(.value)
        }

        _ = args[0]  // lookupValue
        // In a full implementation, we would parse table_array
        // For now, return #N/A to indicate the value isn't found
        return .error(.notAvailable)
    }

    private func match(_ args: [ExcelValue]) throws -> ExcelValue {
        // MATCH(lookup_value, lookup_array, [match_type])
        // Simplified implementation
        guard args.count >= 2 else {
            return .error(.value)
        }

        // In a full implementation, we would search the array
        // For now, return #N/A
        return .error(.notAvailable)
    }

    private func sumif(_ args: [ExcelValue]) throws -> ExcelValue {
        // SUMIF(range, criteria, [sum_range])
        // Simplified: if no sum_range, sum the range that meets criteria
        guard args.count >= 2 else {
            return .error(.value)
        }

        // For basic implementation, just sum all numeric values
        // Full implementation would apply criteria
        var total = 0.0
        for arg in args {
            if let num = arg.toNumber() {
                total += num
            }
        }
        return .number(total)
    }

    private func countif(_ args: [ExcelValue]) throws -> ExcelValue {
        // COUNTIF(range, criteria)
        guard args.count >= 2 else {
            return .error(.value)
        }

        // Simplified: count non-empty values
        // Full implementation would apply criteria
        var count = 0
        for arg in args {
            if arg != .empty {
                count += 1
            }
        }
        return .number(Double(count))
    }

    private func iferror(_ args: [ExcelValue]) throws -> ExcelValue {
        // IFERROR(value, value_if_error)
        guard args.count == 2 else {
            return .error(.value)
        }

        if args[0].isError {
            return args[1]
        } else {
            return args[0]
        }
    }

    private func mod(_ args: [ExcelValue]) throws -> ExcelValue {
        // MOD(number, divisor)
        guard args.count == 2 else {
            return .error(.value)
        }

        guard let number = args[0].toNumber(),
              let divisor = args[1].toNumber() else {
            return .error(.value)
        }

        guard divisor != 0 else {
            return .error(.divideByZero)
        }

        return .number(number.truncatingRemainder(dividingBy: divisor))
    }

    private func sqrt(_ args: [ExcelValue]) throws -> ExcelValue {
        // SQRT(number)
        guard args.count == 1 else {
            return .error(.value)
        }

        guard let number = args[0].toNumber() else {
            return .error(.value)
        }

        guard number >= 0 else {
            return .error(.num)
        }

        return .number(Darwin.sqrt(number))
    }

    private func power(_ args: [ExcelValue]) throws -> ExcelValue {
        // POWER(number, power)
        guard args.count == 2 else {
            return .error(.value)
        }

        guard let number = args[0].toNumber(),
              let exponent = args[1].toNumber() else {
            return .error(.value)
        }

        return .number(pow(number, exponent))
    }

    // MARK: - Tier 3 Functions

    private func averageif(_ args: [ExcelValue]) throws -> ExcelValue {
        // AVERAGEIF(range, criteria, [average_range])
        // Simplified: average all numeric values
        guard args.count >= 2 else {
            return .error(.value)
        }

        var total = 0.0
        var count = 0

        for arg in args {
            if let num = arg.toNumber() {
                total += num
                count += 1
            }
        }

        guard count > 0 else {
            return .error(.divideByZero)
        }

        return .number(total / Double(count))
    }

    private func sumifs(_ args: [ExcelValue]) throws -> ExcelValue {
        // SUMIFS(sum_range, criteria_range1, criteria1, ...)
        // Simplified: sum all numeric values
        guard args.count >= 3 else {
            return .error(.value)
        }

        var total = 0.0
        for arg in args {
            if let num = arg.toNumber() {
                total += num
            }
        }
        return .number(total)
    }

    private func countifs(_ args: [ExcelValue]) throws -> ExcelValue {
        // COUNTIFS(criteria_range1, criteria1, ...)
        // Simplified: count non-empty values
        guard args.count >= 2 else {
            return .error(.value)
        }

        var count = 0
        for arg in args {
            if arg != .empty {
                count += 1
            }
        }
        return .number(Double(count))
    }

    private func find(_ args: [ExcelValue]) throws -> ExcelValue {
        // FIND(find_text, within_text, [start_num])
        guard args.count >= 2 && args.count <= 3 else {
            return .error(.value)
        }

        let findText = args[0].toString()
        let withinText = args[1].toString()
        let startNum = args.count == 3 ? Int(args[2].toNumber() ?? 1) : 1

        guard startNum > 0 else {
            return .error(.value)
        }

        let startIndex = withinText.index(withinText.startIndex, offsetBy: Swift.min(startNum - 1, withinText.count))
        let searchText = String(withinText[startIndex...])

        if let range = searchText.range(of: findText) {
            let position = searchText.distance(from: searchText.startIndex, to: range.lowerBound)
            return .number(Double(startNum + position))
        }

        return .error(.value)
    }

    private func search(_ args: [ExcelValue]) throws -> ExcelValue {
        // SEARCH(find_text, within_text, [start_num]) - case insensitive
        guard args.count >= 2 && args.count <= 3 else {
            return .error(.value)
        }

        let findText = args[0].toString().lowercased()
        let withinText = args[1].toString().lowercased()
        let startNum = args.count == 3 ? Int(args[2].toNumber() ?? 1) : 1

        guard startNum > 0 else {
            return .error(.value)
        }

        let startIndex = withinText.index(withinText.startIndex, offsetBy: Swift.min(startNum - 1, withinText.count))
        let searchText = String(withinText[startIndex...])

        if let range = searchText.range(of: findText) {
            let position = searchText.distance(from: searchText.startIndex, to: range.lowerBound)
            return .number(Double(startNum + position))
        }

        return .error(.value)
    }

    private func substitute(_ args: [ExcelValue]) throws -> ExcelValue {
        // SUBSTITUTE(text, old_text, new_text, [instance_num])
        guard args.count >= 3 && args.count <= 4 else {
            return .error(.value)
        }

        var text = args[0].toString()
        let oldText = args[1].toString()
        let newText = args[2].toString()

        if args.count == 4 {
            // Replace specific instance
            guard let instanceNum = args[3].toNumber(), instanceNum > 0 else {
                return .error(.value)
            }
            let instance = Int(instanceNum)

            var currentInstance = 0
            var searchStart = text.startIndex

            while searchStart < text.endIndex {
                if let range = text[searchStart...].range(of: oldText) {
                    currentInstance += 1
                    if currentInstance == instance {
                        text.replaceSubrange(range, with: newText)
                        break
                    }
                    searchStart = text.index(after: range.lowerBound)
                } else {
                    break
                }
            }
        } else {
            // Replace all instances
            text = text.replacingOccurrences(of: oldText, with: newText)
        }

        return .string(text)
    }

    private func text(_ args: [ExcelValue]) throws -> ExcelValue {
        // TEXT(value, format_text)
        // Simplified: just convert to string
        guard args.count >= 1 else {
            return .error(.value)
        }

        return .string(args[0].toString())
    }

    private func value(_ args: [ExcelValue]) throws -> ExcelValue {
        // VALUE(text) - convert text to number
        guard args.count == 1 else {
            return .error(.value)
        }

        if let num = args[0].toNumber() {
            return .number(num)
        }

        return .error(.value)
    }

    private func isblank(_ args: [ExcelValue]) throws -> ExcelValue {
        // ISBLANK(value)
        guard args.count == 1 else {
            return .error(.value)
        }

        return .boolean(args[0] == .empty)
    }

    private func isnumber(_ args: [ExcelValue]) throws -> ExcelValue {
        // ISNUMBER(value)
        guard args.count == 1 else {
            return .error(.value)
        }

        return .boolean(args[0].isNumeric)
    }

    private func istext(_ args: [ExcelValue]) throws -> ExcelValue {
        // ISTEXT(value)
        guard args.count == 1 else {
            return .error(.value)
        }

        if case .string = args[0] {
            return .boolean(true)
        }
        return .boolean(false)
    }
}
