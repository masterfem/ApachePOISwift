//
// ExcelSheet.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a worksheet in an Excel workbook
public class ExcelSheet {
    /// Name of the sheet
    public let name: String

    /// Sheet information from workbook.xml
    private let sheetInfo: SheetInfo

    /// Cell data parsed from the sheet XML
    private var cells: [String: CellData]

    /// Reference to parent workbook
    private weak var workbook: ExcelWorkbook?

    init(sheetInfo: SheetInfo, cells: [String: CellData], workbook: ExcelWorkbook?) {
        self.name = sheetInfo.name
        self.sheetInfo = sheetInfo
        self.cells = cells
        self.workbook = workbook
    }

    /// Get a cell by its reference (e.g., "A1", "B2")
    /// - Parameter reference: Cell reference in A1 notation
    /// - Returns: ExcelCell object
    /// - Throws: ExcelError if the reference is invalid
    public func cell(_ reference: String) throws -> ExcelCell {
        // Validate reference format
        _ = try CellReference(reference)

        // Get cell data (or create empty cell if not found)
        let cellData = cells[reference] ?? CellData(
            reference: reference,
            type: nil,
            value: nil,
            formula: nil
        )

        return ExcelCell(reference: reference, cellData: cellData, workbook: workbook)
    }

    /// Get a cell by column and row indices (zero-based)
    /// - Parameters:
    ///   - column: Column index (0 = A, 1 = B, etc.)
    ///   - row: Row index (0 = row 1, 1 = row 2, etc.)
    /// - Returns: ExcelCell object
    /// - Throws: ExcelError if the indices are invalid
    public func cell(column: Int, row: Int) throws -> ExcelCell {
        let cellRef = try CellReference(column: column, row: row)
        return try cell(cellRef.toExcelNotation())
    }

    /// Get all cells in the sheet
    /// - Returns: Array of ExcelCell objects
    public func allCells() -> [ExcelCell] {
        return cells.map { (reference, data) in
            ExcelCell(reference: reference, cellData: data, workbook: workbook)
        }
    }

    /// Get all non-empty cells in the sheet
    /// - Returns: Array of ExcelCell objects that have values
    public func nonEmptyCells() -> [ExcelCell] {
        return allCells().filter { !$0.isEmpty }
    }

    /// Get the number of cells with data in the sheet
    public var cellCount: Int {
        return cells.count
    }

    /// Get the sheet ID
    public var sheetId: String {
        return sheetInfo.sheetId
    }
}

extension ExcelSheet: CustomStringConvertible {
    public var description: String {
        return "Sheet '\(name)' (\(cellCount) cells)"
    }
}
