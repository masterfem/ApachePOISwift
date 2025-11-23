//
// ExcelValue.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a value in Excel formulas with automatic type coercion
public enum ExcelValue: Equatable {
    case number(Double)
    case string(String)
    case boolean(Bool)
    case error(ExcelError)
    case empty

    /// Excel errors
    public enum ExcelError: String, Equatable {
        case divideByZero = "#DIV/0!"
        case notAvailable = "#N/A"
        case name = "#NAME?"
        case null = "#NULL!"
        case num = "#NUM!"
        case ref = "#REF!"
        case value = "#VALUE!"
    }

    // MARK: - Type Coercion (Excel's loose typing rules)

    /// Convert to Double (for numeric operations)
    public func toNumber() -> Double? {
        switch self {
        case .number(let value):
            return value
        case .string(let str):
            // Excel tries to parse strings as numbers
            return Double(str)
        case .boolean(let bool):
            // TRUE = 1, FALSE = 0
            return bool ? 1.0 : 0.0
        case .empty:
            return 0.0
        case .error:
            return nil
        }
    }

    /// Convert to String (for text operations)
    public func toString() -> String {
        switch self {
        case .number(let value):
            // Format numbers like Excel does
            if value.truncatingRemainder(dividingBy: 1) == 0 {
                return String(Int(value))
            }
            return String(value)
        case .string(let str):
            return str
        case .boolean(let bool):
            return bool ? "TRUE" : "FALSE"
        case .empty:
            return ""
        case .error(let err):
            return err.rawValue
        }
    }

    /// Convert to Boolean (for logical operations)
    public func toBoolean() -> Bool? {
        switch self {
        case .number(let value):
            // Non-zero = TRUE
            return value != 0.0
        case .string(let str):
            // "TRUE" = true, "FALSE" = false (case insensitive)
            let upper = str.uppercased()
            if upper == "TRUE" { return true }
            if upper == "FALSE" { return false }
            return nil
        case .boolean(let bool):
            return bool
        case .empty:
            return false
        case .error:
            return nil
        }
    }

    /// Check if value is an error
    public var isError: Bool {
        if case .error = self { return true }
        return false
    }

    /// Check if value is numeric
    public var isNumeric: Bool {
        return toNumber() != nil
    }

    // MARK: - Comparison

    /// Compare values (Excel comparison rules)
    public func compare(_ other: ExcelValue) -> ComparisonResult? {
        // Errors cannot be compared
        if self.isError || other.isError {
            return nil
        }

        // Try numeric comparison first
        if let a = self.toNumber(), let b = other.toNumber() {
            if a < b { return .orderedAscending }
            if a > b { return .orderedDescending }
            return .orderedSame
        }

        // Fall back to string comparison
        let a = self.toString()
        let b = other.toString()
        return a.compare(b)
    }
}

// MARK: - CustomStringConvertible

extension ExcelValue: CustomStringConvertible {
    public var description: String {
        return toString()
    }
}
