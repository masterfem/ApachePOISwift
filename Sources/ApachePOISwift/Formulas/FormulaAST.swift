//
// FormulaAST.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Abstract Syntax Tree node for Excel formulas
public indirect enum FormulaAST: Equatable {
    // Literals
    case number(Double)
    case string(String)
    case boolean(Bool)
    case cellReference(String)
    case range(String)

    // Binary operations
    case add(FormulaAST, FormulaAST)
    case subtract(FormulaAST, FormulaAST)
    case multiply(FormulaAST, FormulaAST)
    case divide(FormulaAST, FormulaAST)
    case power(FormulaAST, FormulaAST)
    case concat(FormulaAST, FormulaAST)

    // Comparison operations
    case equal(FormulaAST, FormulaAST)
    case notEqual(FormulaAST, FormulaAST)
    case lessThan(FormulaAST, FormulaAST)
    case lessOrEqual(FormulaAST, FormulaAST)
    case greaterThan(FormulaAST, FormulaAST)
    case greaterOrEqual(FormulaAST, FormulaAST)

    // Unary operations
    case negate(FormulaAST)
    case positive(FormulaAST)

    // Function call
    case function(name: String, arguments: [FormulaAST])

    // MARK: - Debug

    var description: String {
        switch self {
        case .number(let value):
            return "\(value)"
        case .string(let value):
            return "\"\(value)\""
        case .boolean(let value):
            return value ? "TRUE" : "FALSE"
        case .cellReference(let ref):
            return ref
        case .range(let range):
            return range
        case .add(let left, let right):
            return "(\(left.description) + \(right.description))"
        case .subtract(let left, let right):
            return "(\(left.description) - \(right.description))"
        case .multiply(let left, let right):
            return "(\(left.description) * \(right.description))"
        case .divide(let left, let right):
            return "(\(left.description) / \(right.description))"
        case .power(let left, let right):
            return "(\(left.description) ^ \(right.description))"
        case .concat(let left, let right):
            return "(\(left.description) & \(right.description))"
        case .equal(let left, let right):
            return "(\(left.description) = \(right.description))"
        case .notEqual(let left, let right):
            return "(\(left.description) <> \(right.description))"
        case .lessThan(let left, let right):
            return "(\(left.description) < \(right.description))"
        case .lessOrEqual(let left, let right):
            return "(\(left.description) <= \(right.description))"
        case .greaterThan(let left, let right):
            return "(\(left.description) > \(right.description))"
        case .greaterOrEqual(let left, let right):
            return "(\(left.description) >= \(right.description))"
        case .negate(let expr):
            return "(-\(expr.description))"
        case .positive(let expr):
            return "(+\(expr.description))"
        case .function(let name, let args):
            let argStrs = args.map { $0.description }.joined(separator: ", ")
            return "\(name)(\(argStrs))"
        }
    }
}
