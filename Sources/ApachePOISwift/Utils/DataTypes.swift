//
// DataTypes.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Excel cell type as specified in the XML
public enum CellType: String {
    /// Shared string (reference to sharedStrings.xml)
    case string = "s"

    /// Inline string (string stored directly in cell)
    case inlineString = "inlineStr"

    /// Numeric value (default if type not specified)
    case number = "n"

    /// Boolean value
    case boolean = "b"

    /// Error value
    case error = "e"

    /// Formula result stored as string
    case formula = "str"
}

/// Value stored in an Excel cell
public enum CellValue: Equatable {
    /// String value
    case string(String)

    /// Numeric value
    case number(Double)

    /// Boolean value
    case boolean(Bool)

    /// Date value (internally stored as number in Excel)
    case date(Date)

    /// Formula (the formula string, not the calculated result)
    case formula(String)

    /// Empty cell
    case empty

    /// Get the value as a String if possible
    public var stringValue: String? {
        switch self {
        case .string(let value):
            return value
        case .number(let value):
            return String(value)
        case .boolean(let value):
            return String(value)
        case .date(let value):
            return value.description
        case .formula(let value):
            return value
        case .empty:
            return nil
        }
    }

    /// Get the value as a Double if possible
    public var numberValue: Double? {
        switch self {
        case .number(let value):
            return value
        case .string(let value):
            return Double(value)
        case .boolean(let value):
            return value ? 1.0 : 0.0
        default:
            return nil
        }
    }

    /// Get the value as a Bool if possible
    public var boolValue: Bool? {
        switch self {
        case .boolean(let value):
            return value
        case .number(let value):
            return value != 0
        case .string(let value):
            return value.lowercased() == "true" || value == "1"
        default:
            return nil
        }
    }
}
