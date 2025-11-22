//
// CellReference.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents an Excel cell reference in A1 notation
public struct CellReference {
    /// Zero-based column index (A=0, B=1, ..., Z=25, AA=26, etc.)
    public let column: Int

    /// Zero-based row index (Excel row 1 = 0, row 2 = 1, etc.)
    public let row: Int

    /// Whether the column is absolute (e.g., $A in $A1)
    public let absoluteColumn: Bool

    /// Whether the row is absolute (e.g., $1 in A$1)
    public let absoluteRow: Bool

    /// Initialize from A1 notation (e.g., "A1", "$B$2", "AA10")
    public init(_ reference: String) throws {
        let trimmed = reference.trimmingCharacters(in: .whitespaces)

        guard !trimmed.isEmpty else {
            throw ExcelError.invalidCellReference("Empty reference")
        }

        var currentIndex = trimmed.startIndex
        var absoluteCol = false
        var absoluteRow = false
        var columnLetters = ""
        var rowDigits = ""

        // Parse absolute column marker
        if trimmed[currentIndex] == "$" {
            absoluteCol = true
            currentIndex = trimmed.index(after: currentIndex)
        }

        // Parse column letters
        while currentIndex < trimmed.endIndex {
            let char = trimmed[currentIndex]
            if char.isLetter {
                columnLetters.append(char.uppercased())
                currentIndex = trimmed.index(after: currentIndex)
            } else {
                break
            }
        }

        guard !columnLetters.isEmpty else {
            throw ExcelError.invalidCellReference("No column letters found in '\(reference)'")
        }

        // Parse absolute row marker
        if currentIndex < trimmed.endIndex && trimmed[currentIndex] == "$" {
            absoluteRow = true
            currentIndex = trimmed.index(after: currentIndex)
        }

        // Parse row digits
        while currentIndex < trimmed.endIndex {
            let char = trimmed[currentIndex]
            if char.isNumber {
                rowDigits.append(char)
                currentIndex = trimmed.index(after: currentIndex)
            } else {
                break
            }
        }

        guard !rowDigits.isEmpty else {
            throw ExcelError.invalidCellReference("No row number found in '\(reference)'")
        }

        // Ensure we consumed the entire string
        guard currentIndex == trimmed.endIndex else {
            throw ExcelError.invalidCellReference("Invalid characters in '\(reference)'")
        }

        // Convert column letters to index
        let col = try Self.columnLettersToIndex(columnLetters)

        // Convert row number to zero-based index
        guard let rowNum = Int(rowDigits), rowNum > 0, rowNum <= 1048576 else {
            throw ExcelError.invalidCellReference("Invalid row number in '\(reference)'")
        }

        self.column = col
        self.row = rowNum - 1  // Convert to 0-based
        self.absoluteColumn = absoluteCol
        self.absoluteRow = absoluteRow
    }

    /// Initialize with explicit column and row indices
    public init(column: Int, row: Int, absoluteColumn: Bool = false, absoluteRow: Bool = false) throws {
        guard column >= 0 && column < 16384 else {
            throw ExcelError.invalidCellReference("Column \(column) out of range (0-16383)")
        }

        guard row >= 0 && row < 1048576 else {
            throw ExcelError.invalidCellReference("Row \(row) out of range (0-1048575)")
        }

        self.column = column
        self.row = row
        self.absoluteColumn = absoluteColumn
        self.absoluteRow = absoluteRow
    }

    /// Convert to Excel A1 notation
    public func toExcelNotation() -> String {
        let colStr = Self.columnIndexToLetters(column)
        let rowStr = String(row + 1)  // Convert to 1-based

        let colPrefix = absoluteColumn ? "$" : ""
        let rowPrefix = absoluteRow ? "$" : ""

        return "\(colPrefix)\(colStr)\(rowPrefix)\(rowStr)"
    }

    /// Convert column letters (A, B, ..., Z, AA, AB, ...) to zero-based index
    private static func columnLettersToIndex(_ letters: String) throws -> Int {
        guard !letters.isEmpty else {
            throw ExcelError.invalidCellReference("Empty column letters")
        }

        var index = 0
        for char in letters.uppercased() {
            guard let value = char.asciiValue, value >= 65 && value <= 90 else {
                throw ExcelError.invalidCellReference("Invalid column letter '\(char)'")
            }

            index = index * 26 + Int(value - 64)  // A=1, B=2, ..., Z=26
        }

        return index - 1  // Convert to 0-based
    }

    /// Convert zero-based column index to letters (0=A, 1=B, ..., 25=Z, 26=AA, ...)
    private static func columnIndexToLetters(_ index: Int) -> String {
        var col = index + 1  // Convert to 1-based
        var letters = ""

        while col > 0 {
            let remainder = (col - 1) % 26
            letters = String(UnicodeScalar(65 + remainder)!) + letters
            col = (col - 1) / 26
        }

        return letters
    }
}

extension CellReference: Equatable {
    public static func == (lhs: CellReference, rhs: CellReference) -> Bool {
        return lhs.column == rhs.column && lhs.row == rhs.row
    }
}

extension CellReference: CustomStringConvertible {
    public var description: String {
        return toExcelNotation()
    }
}
