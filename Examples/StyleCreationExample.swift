//
// StyleCreationExample.swift
// ApachePOISwift Examples
//
// Demonstrates how to create and apply cell styles programmatically
//

import Foundation
import ApachePOISwift

/// Example 1: Basic Font Styling
func example1_BasicFontStyling() throws {
    print("=== Example 1: Basic Font Styling ===")

    // Open template
    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Apply bold
    let cell1 = try sheet.cell("A1")
    cell1.setValue(.string("Bold Text"))
    cell1.makeBold()

    // Apply italic
    let cell2 = try sheet.cell("A2")
    cell2.setValue(.string("Italic Text"))
    cell2.makeItalic()

    // Apply custom font
    let cell3 = try sheet.cell("A3")
    cell3.setValue(.string("Custom Font"))
    cell3.applyFont(Font(
        name: "Arial",
        size: 16,
        bold: true,
        italic: true,
        color: "FFFF0000"  // Red
    ))

    // Save
    try workbook.save(to: URL(fileURLWithPath: "output_fonts.xlsx"))
    print("Saved to output_fonts.xlsx")
}

/// Example 2: Background Colors
func example2_BackgroundColors() throws {
    print("\n=== Example 2: Background Colors ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Red background
    let cell1 = try sheet.cell("B1")
    cell1.setValue(.string("Red"))
    cell1.setBackgroundColor("FFFF0000")

    // Green background
    let cell2 = try sheet.cell("B2")
    cell2.setValue(.string("Green"))
    cell2.setBackgroundColor("FF00FF00")

    // Blue background
    let cell3 = try sheet.cell("B3")
    cell3.setValue(.string("Blue"))
    cell3.setBackgroundColor("FF0000FF")

    // Custom fill with pattern
    let cell4 = try sheet.cell("B4")
    cell4.setValue(.string("Gray Fill"))
    cell4.applyFill(Fill(
        patternType: .solid,
        foregroundColor: "FFCCCCCC"
    ))

    try workbook.save(to: URL(fileURLWithPath: "output_colors.xlsx"))
    print("Saved to output_colors.xlsx")
}

/// Example 3: Borders
func example3_Borders() throws {
    print("\n=== Example 3: Borders ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Simple border (all sides)
    let cell1 = try sheet.cell("C1")
    cell1.setValue(.string("Bordered"))
    cell1.setBorder(style: .thin, color: "FF000000")

    // Custom border (different styles per edge)
    let cell2 = try sheet.cell("C2")
    cell2.setValue(.string("Custom Border"))
    cell2.applyBorder(Border(
        left: BorderEdge(style: .thin, color: "FF000000"),
        right: BorderEdge(style: .thick, color: "FFFF0000"),
        top: BorderEdge(style: .double, color: "FF0000FF"),
        bottom: BorderEdge(style: .thin, color: "FF00FF00")
    ))

    try workbook.save(to: URL(fileURLWithPath: "output_borders.xlsx"))
    print("Saved to output_borders.xlsx")
}

/// Example 4: Complete Cell Styling
func example4_CompleteStyling() throws {
    print("\n=== Example 4: Complete Cell Styling ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Apply everything at once
    let cell = try sheet.cell("D1")
    cell.setValue(.string("Fully Styled Cell"))
    cell.applyStyle(
        font: Font(
            name: "Times New Roman",
            size: 14,
            bold: true,
            color: "FFFFFFFF"  // White text
        ),
        fill: Fill(
            patternType: .solid,
            foregroundColor: "FF0070C0"  // Blue background
        ),
        border: Border(
            left: BorderEdge(style: .medium),
            right: BorderEdge(style: .medium),
            top: BorderEdge(style: .medium),
            bottom: BorderEdge(style: .medium)
        ),
        horizontalAlignment: .center,
        verticalAlignment: .center,
        wrapText: true
    )

    try workbook.save(to: URL(fileURLWithPath: "output_complete.xlsx"))
    print("Saved to output_complete.xlsx")
}

/// Example 5: Styled Table Header
func example5_StyledTableHeader() throws {
    print("\n=== Example 5: Styled Table Header ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Create header row
    let headers = ["Name", "Age", "Email", "Department", "Salary"]

    for (index, headerText) in headers.enumerated() {
        let cell = try sheet.cell(column: index, row: 0)
        cell.setValue(.string(headerText))

        // Apply header styling
        cell.makeBold()
        cell.setBackgroundColor("FF4472C4")  // Blue
        cell.applyFont(Font(
            name: "Calibri",
            size: 11,
            bold: true,
            color: "FFFFFFFF"  // White text
        ))
        cell.setBorder(style: .thin, color: "FFFFFFFF")

        // Center align
        cell.applyStyle(
            font: Font(name: "Calibri", size: 11, bold: true, color: "FFFFFFFF"),
            fill: Fill(patternType: .solid, foregroundColor: "FF4472C4"),
            border: Border(
                left: BorderEdge(style: .thin, color: "FFFFFFFF"),
                right: BorderEdge(style: .thin, color: "FFFFFFFF"),
                top: BorderEdge(style: .thin, color: "FFFFFFFF"),
                bottom: BorderEdge(style: .thin, color: "FFFFFFFF")
            ),
            horizontalAlignment: .center,
            verticalAlignment: .center
        )
    }

    // Add sample data
    let data = [
        ["John Doe", "30", "john@example.com", "Engineering", "75000"],
        ["Jane Smith", "28", "jane@example.com", "Marketing", "65000"],
        ["Bob Johnson", "35", "bob@example.com", "Sales", "70000"]
    ]

    for (rowIndex, rowData) in data.enumerated() {
        for (colIndex, value) in rowData.enumerated() {
            let cell = try sheet.cell(column: colIndex, row: rowIndex + 1)
            cell.setValue(.string(value))

            // Alternate row colors
            if rowIndex % 2 == 0 {
                cell.setBackgroundColor("FFD9E1F2")  // Light blue
            }
        }
    }

    try workbook.save(to: URL(fileURLWithPath: "output_table.xlsx"))
    print("Saved to output_table.xlsx")
}

/// Example 6: Number Formatting
func example6_NumberFormatting() throws {
    print("\n=== Example 6: Number Formatting ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Currency format
    let cell1 = try sheet.cell("E1")
    cell1.setValue(.number(1234.56))
    cell1.applyNumberFormat(NumberFormat(formatId: 44))  // Currency

    // Percentage format
    let cell2 = try sheet.cell("E2")
    cell2.setValue(.number(0.75))
    cell2.applyNumberFormat(NumberFormat(formatId: 10))  // Percentage

    // Date format
    let cell3 = try sheet.cell("E3")
    cell3.setValue(.date(Date()))
    cell3.applyNumberFormat(NumberFormat(formatId: 14))  // Date

    // Decimal format
    let cell4 = try sheet.cell("E4")
    cell4.setValue(.number(1234.5678))
    cell4.applyNumberFormat(NumberFormat(formatId: 2))  // 0.00

    try workbook.save(to: URL(fileURLWithPath: "output_numbers.xlsx"))
    print("Saved to output_numbers.xlsx")
}

/// Example 7: Marbar Report Styling (Real-World Use Case)
func example7_MarbarReportStyling() throws {
    print("\n=== Example 7: Marbar Report Styling ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "marbar_template.xlsm"))
    let sheet = try workbook.sheet(named: "GENERALES")

    // Style section headers
    let headerCell = try sheet.cell("A1")
    headerCell.makeBold()
    headerCell.setBackgroundColor("FF4472C4")
    headerCell.applyFont(Font(
        name: "Calibri",
        size: 12,
        bold: true,
        color: "FFFFFFFF"
    ))

    // Style data entry cells
    let dataCell = try sheet.cell("B5")
    dataCell.setValue(.string("PAD-123"))
    dataCell.applyStyle(
        font: Font(name: "Calibri", size: 11),
        fill: Fill(patternType: .solid, foregroundColor: "FFF2F2F2"),
        border: Border(
            left: BorderEdge(style: .thin),
            right: BorderEdge(style: .thin),
            top: BorderEdge(style: .thin),
            bottom: BorderEdge(style: .thin)
        )
    )

    // Highlight important values
    let importantCell = try sheet.cell("B10")
    importantCell.setValue(.number(123.45))
    importantCell.makeBold()
    importantCell.setBackgroundColor("FFFFFF00")  // Yellow highlight

    // Save with macros preserved
    try workbook.save(to: URL(fileURLWithPath: "marbar_styled.xlsm"))
    print("Saved to marbar_styled.xlsm with macros preserved!")
}

// Run all examples
func runAllExamples() {
    do {
        try example1_BasicFontStyling()
        try example2_BackgroundColors()
        try example3_Borders()
        try example4_CompleteStyling()
        try example5_StyledTableHeader()
        try example6_NumberFormatting()
        // try example7_MarbarReportStyling()  // Uncomment if you have the template

        print("\n✅ All examples completed successfully!")
    } catch {
        print("❌ Error: \(error)")
    }
}

// Uncomment to run:
// runAllExamples()
