//
// BasicReadExample.swift
// ApachePOISwift Examples
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation
import ApachePOISwift

// MARK: - Example 1: Open and Read Excel File

func example1_BasicRead() throws {
    // Open an Excel file
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)

    // Print workbook information
    print("Workbook: \(workbook)")
    print("Number of sheets: \(workbook.sheetCount)")
    print("Sheet names: \(workbook.sheetNames.joined(separator: ", "))")

    // Access first sheet
    let sheet = try workbook.sheet(at: 0)
    print("\nReading sheet: \(sheet.name)")

    // Read specific cells
    let cellA1 = try sheet.cell("A1")
    print("A1 = \(cellA1.value)")

    let cellB2 = try sheet.cell("B2")
    print("B2 = \(cellB2.value)")

    // Read cell by column/row index (0-based)
    let cellByIndex = try sheet.cell(column: 2, row: 3)  // C4
    print("C4 = \(cellByIndex.value)")
}

// MARK: - Example 2: Read All Non-Empty Cells

func example2_ReadAllCells() throws {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)

    let sheet = try workbook.sheet(named: "Sales")

    // Get all non-empty cells
    let nonEmptyCells = sheet.nonEmptyCells()

    print("Found \(nonEmptyCells.count) non-empty cells")

    for cell in nonEmptyCells {
        print("\(cell.reference): \(cell.value)")
    }
}

// MARK: - Example 3: Work with Different Cell Types

func example3_CellTypes() throws {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)
    let sheet = try workbook.sheet(at: 0)

    // Read a string value
    let stringCell = try sheet.cell("A1")
    if case .string(let value) = stringCell.value {
        print("String value: \(value)")
    }

    // Read a number value
    let numberCell = try sheet.cell("B1")
    if case .number(let value) = numberCell.value {
        print("Number value: \(value)")
    }

    // Read a boolean value
    let boolCell = try sheet.cell("C1")
    if case .boolean(let value) = boolCell.value {
        print("Boolean value: \(value)")
    }

    // Read a formula
    let formulaCell = try sheet.cell("D1")
    if formulaCell.hasFormula {
        print("Formula: \(formulaCell.formula ?? "")")
        print("Result: \(formulaCell.value)")
    }

    // Check if cell is empty
    let emptyCell = try sheet.cell("Z99")
    if emptyCell.isEmpty {
        print("Cell Z99 is empty")
    }
}

// MARK: - Example 4: Access Sheet by Name

func example4_AccessByName() throws {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)

    // Access sheet by name
    let salesSheet = try workbook.sheet(named: "Sales")
    let inventorySheet = try workbook.sheet(named: "Inventory")

    print("Sales sheet has \(salesSheet.cellCount) cells")
    print("Inventory sheet has \(inventorySheet.cellCount) cells")
}

// MARK: - Example 5: Check for VBA Macros

func example5_CheckMacros() throws {
    let xlsxFile = URL(fileURLWithPath: "/path/to/file.xlsx")
    let xlsmFile = URL(fileURLWithPath: "/path/to/file.xlsm")

    let workbook1 = try ExcelWorkbook(fileURL: xlsxFile)
    print("XLSX has macros: \(workbook1.hasVBAMacros)")  // false

    let workbook2 = try ExcelWorkbook(fileURL: xlsmFile)
    print("XLSM has macros: \(workbook2.hasVBAMacros)")  // true
}

// MARK: - Example 6: Iterate Through All Sheets

func example6_IterateSheets() throws {
    let fileURL = URL(fileURLWithPath: "/path/to/your/file.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)

    // Access all sheets
    for sheet in workbook.allSheets {
        print("\nSheet: \(sheet.name)")
        print("  Cell count: \(sheet.cellCount)")
        print("  Non-empty cells: \(sheet.nonEmptyCells().count)")

        // Print first 5 non-empty cells
        for cell in sheet.nonEmptyCells().prefix(5) {
            print("    \(cell.reference): \(cell.value)")
        }
    }
}

// MARK: - Example 7: Error Handling

func example7_ErrorHandling() {
    let fileURL = URL(fileURLWithPath: "/path/to/file.xlsx")

    do {
        let workbook = try ExcelWorkbook(fileURL: fileURL)
        let sheet = try workbook.sheet(named: "Sales")
        let cell = try sheet.cell("A1")
        print("Value: \(cell.value)")

    } catch ExcelError.fileNotFound(let url) {
        print("File not found: \(url.path)")

    } catch ExcelError.sheetNotFound(let name) {
        print("Sheet '\(name)' not found")

    } catch ExcelError.invalidCellReference(let reference) {
        print("Invalid cell reference: \(reference)")

    } catch {
        print("Error: \(error)")
    }
}

// MARK: - Example 8: Load from Bundle

func example8_LoadFromBundle() throws {
    // Load Excel file from app bundle
    let workbook = try ExcelWorkbook(
        name: "template",
        bundle: .main,
        extension: "xlsx"
    )

    print("Loaded template with \(workbook.sheetCount) sheets")
}

// MARK: - Example 9: Real-World Use Case - Extract Data

func example9_ExtractData() throws {
    let fileURL = URL(fileURLWithPath: "/path/to/sales_report.xlsx")
    let workbook = try ExcelWorkbook(fileURL: fileURL)
    let sheet = try workbook.sheet(named: "Monthly Sales")

    // Extract data from specific cells
    struct SalesData {
        let month: String
        let revenue: Double
        let units: Int
    }

    var salesRecords: [SalesData] = []

    // Assuming data starts at row 2 (row 1 is headers)
    for row in 1...12 {  // 12 months
        let monthCell = try sheet.cell(column: 0, row: row)  // Column A
        let revenueCell = try sheet.cell(column: 1, row: row)  // Column B
        let unitsCell = try sheet.cell(column: 2, row: row)  // Column C

        if case .string(let month) = monthCell.value,
           case .number(let revenue) = revenueCell.value,
           case .number(let units) = unitsCell.value {

            let record = SalesData(
                month: month,
                revenue: revenue,
                units: Int(units)
            )
            salesRecords.append(record)
        }
    }

    print("Extracted \(salesRecords.count) sales records")
    for record in salesRecords {
        print("\(record.month): $\(record.revenue) (\(record.units) units)")
    }
}

// MARK: - Run Examples

// Uncomment to run examples:
// try example1_BasicRead()
// try example2_ReadAllCells()
// try example3_CellTypes()
// try example4_AccessByName()
// try example5_CheckMacros()
// try example6_IterateSheets()
// example7_ErrorHandling()
// try example8_LoadFromBundle()
// try example9_ExtractData()
