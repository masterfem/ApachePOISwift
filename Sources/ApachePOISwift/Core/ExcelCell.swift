//
// ExcelCell.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a single cell in an Excel worksheet
public class ExcelCell {
    /// Cell reference in A1 notation (e.g., "A1", "B2")
    public let reference: String

    /// Raw cell data from XML
    private var cellData: CellData

    /// Reference to parent workbook for shared strings lookup
    private weak var workbook: ExcelWorkbook?

    init(reference: String, cellData: CellData, workbook: ExcelWorkbook?) {
        self.reference = reference
        self.cellData = cellData
        self.workbook = workbook
    }

    /// The value of the cell
    public var value: CellValue {
        // If cell has a formula, return the formula
        if let formula = cellData.formula, !formula.isEmpty {
            return .formula(formula)
        }

        // If no value, cell is empty
        guard let rawValue = cellData.value, !rawValue.isEmpty else {
            return .empty
        }

        // Determine value based on cell type
        switch cellData.type {
        case .string:
            // Shared string - lookup in shared strings table
            if let index = Int(rawValue),
               let sharedStrings = workbook?.sharedStrings,
               index >= 0 && index < sharedStrings.count {
                return .string(sharedStrings[index])
            }
            return .empty

        case .inlineString:
            // Inline string (less common)
            return .string(rawValue)

        case .boolean:
            // Boolean: 0 = false, 1 = true
            return .boolean(rawValue == "1")

        case .number, .none:
            // Default is number
            if let number = Double(rawValue) {
                // Check if this might be a date (Excel dates are numbers >= 1)
                // For now, just return as number - date handling can be added later
                return .number(number)
            }
            return .empty

        case .error:
            // Error value - return as string for now
            return .string(rawValue)

        case .formula:
            // Formula result stored as string
            return .string(rawValue)
        }
    }

    /// The formula in the cell (if any)
    public var formula: String? {
        return cellData.formula
    }

    /// Whether the cell has a formula
    public var hasFormula: Bool {
        return cellData.formula != nil && !cellData.formula!.isEmpty
    }

    /// Whether the cell is empty
    public var isEmpty: Bool {
        if case .empty = value {
            return true
        }
        return false
    }

    /// Get the cell reference as a CellReference object
    public var cellReference: CellReference? {
        return try? CellReference(reference)
    }
}

extension ExcelCell: CustomStringConvertible {
    public var description: String {
        return "\(reference): \(value)"
    }
}
