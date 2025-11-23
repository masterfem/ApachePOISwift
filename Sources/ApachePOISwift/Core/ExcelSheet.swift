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

    /// Merged cell ranges (e.g., ["A1:B2", "C3:D5"])
    internal var mergedCells: [String] = []

    /// Reference to parent workbook
    private weak var workbook: ExcelWorkbook?

    /// Cache of cell objects (reuse same object for same reference)
    private var cellCache: [String: ExcelCell] = [:]

    /// Whether this sheet has been modified
    internal var isModified: Bool = false

    init(sheetInfo: SheetInfo, cells: [String: CellData], mergedCells: [String] = [], workbook: ExcelWorkbook?) {
        self.name = sheetInfo.name
        self.sheetInfo = sheetInfo
        self.cells = cells
        self.mergedCells = mergedCells
        self.workbook = workbook
    }

    /// Get a cell by its reference (e.g., "A1", "B2")
    /// - Parameter reference: Cell reference in A1 notation
    /// - Returns: ExcelCell object
    /// - Throws: ExcelError if the reference is invalid
    public func cell(_ reference: String) throws -> ExcelCell {
        // Validate reference format
        _ = try CellReference(reference)

        // Return cached cell if exists (important for write operations)
        if let cachedCell = cellCache[reference] {
            return cachedCell
        }

        // Get cell data (or create empty cell if not found)
        let cellData = cells[reference] ?? CellData(
            reference: reference,
            type: nil,
            value: nil,
            formula: nil,
            styleIndex: nil
        )

        let cell = ExcelCell(reference: reference, cellData: cellData, workbook: workbook, sheet: self)
        cellCache[reference] = cell
        return cell
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

    // MARK: - Write Support (Phase 2)

    /// Mark this sheet as modified
    internal func markAsModified() {
        isModified = true
        workbook?.markAsModified()
    }

    /// Get all cells for XML writing (includes modified cells)
    internal func getAllCellsForWriting() -> [ExcelCell] {
        // Collect all cells from cache (which includes modifications)
        var allCells: [String: ExcelCell] = cellCache

        // Add cells from original data that aren't in cache
        for (reference, cellData) in cells {
            if allCells[reference] == nil {
                allCells[reference] = ExcelCell(reference: reference, cellData: cellData, workbook: workbook, sheet: self)
            }
        }

        return Array(allCells.values).sorted { $0.reference < $1.reference }
    }

    /// Get the sheet info (for workbook writing)
    internal func getSheetInfo() -> SheetInfo {
        return sheetInfo
    }
}

extension ExcelSheet: CustomStringConvertible {
    public var description: String {
        return "Sheet '\(name)' (\(cellCount) cells)"
    }
}
