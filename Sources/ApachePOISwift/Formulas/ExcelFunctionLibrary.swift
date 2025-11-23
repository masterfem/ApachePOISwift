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
}
