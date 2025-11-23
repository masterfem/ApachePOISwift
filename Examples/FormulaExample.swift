//
// FormulaExample.swift
// ApachePOISwift Examples
//
// Demonstrates how to use formulas in Excel files
//

import Foundation
import ApachePOISwift

/// Example 1: Basic Arithmetic Formulas
func example1_BasicArithmetic() throws {
    print("=== Example 1: Basic Arithmetic ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Set some values
    try sheet.cell("A1").setValue(.number(10))
    try sheet.cell("A2").setValue(.number(20))
    try sheet.cell("A3").setValue(.number(30))

    // Simple addition
    try sheet.cell("B1").setFormula("=A1+A2")

    // Multiplication
    try sheet.cell("B2").setFormula("=A1*A2")

    // Complex expression
    try sheet.cell("B3").setFormula("=(A1+A2)*A3")

    try workbook.save(to: URL(fileURLWithPath: "output_arithmetic.xlsx"))
    print("Saved! Open in Excel to see calculated results")
}

/// Example 2: SUM and AVERAGE Functions
func example2_SUMAndAVERAGE() throws {
    print("\n=== Example 2: SUM and AVERAGE ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Create a sales data table
    try sheet.cell("A1").setValue(.string("Product"))
    try sheet.cell("B1").setValue(.string("Q1 Sales"))
    try sheet.cell("C1").setValue(.string("Q2 Sales"))
    try sheet.cell("D1").setValue(.string("Q3 Sales"))
    try sheet.cell("E1").setValue(.string("Q4 Sales"))
    try sheet.cell("F1").setValue(.string("Total"))
    try sheet.cell("G1").setValue(.string("Average"))

    // Product data
    let products = ["Widget A", "Widget B", "Widget C"]
    let sales: [[Double]] = [
        [1000, 1200, 1100, 1300],
        [800, 900, 850, 950],
        [1500, 1600, 1550, 1650]
    ]

    for (rowIndex, product) in products.enumerated() {
        let row = rowIndex + 1

        try sheet.cell(column: 0, row: row).setValue(.string(product))
        try sheet.cell(column: 1, row: row).setValue(.number(sales[rowIndex][0]))
        try sheet.cell(column: 2, row: row).setValue(.number(sales[rowIndex][1]))
        try sheet.cell(column: 3, row: row).setValue(.number(sales[rowIndex][2]))
        try sheet.cell(column: 4, row: row).setValue(.number(sales[rowIndex][3]))

        // Total: SUM of Q1-Q4
        try sheet.cell(column: 5, row: row).setFormula("=SUM(B\(row + 1):E\(row + 1))")

        // Average
        try sheet.cell(column: 6, row: row).setFormula("=AVERAGE(B\(row + 1):E\(row + 1))")
    }

    // Grand totals
    try sheet.cell("A5").setValue(.string("Grand Total"))
    try sheet.cell("F5").setFormula("=SUM(F2:F4)")
    try sheet.cell("G5").setFormula("=AVERAGE(G2:G4)")

    try workbook.save(to: URL(fileURLWithPath: "output_sum_average.xlsx"))
    print("Saved! Sales data with SUM and AVERAGE formulas")
}

/// Example 3: IF Statements (Conditional Logic)
func example3_IFStatements() throws {
    print("\n=== Example 3: IF Statements ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Student grades
    try sheet.cell("A1").setValue(.string("Student"))
    try sheet.cell("B1").setValue(.string("Score"))
    try sheet.cell("C1").setValue(.string("Grade"))
    try sheet.cell("D1").setValue(.string("Pass/Fail"))

    let students = ["Alice", "Bob", "Charlie", "Diana"]
    let scores = [95, 72, 68, 88]

    for (index, student) in students.enumerated() {
        let row = index + 1

        try sheet.cell(column: 0, row: row).setValue(.string(student))
        try sheet.cell(column: 1, row: row).setValue(.number(Double(scores[index])))

        // Letter grade (nested IF)
        try sheet.cell(column: 2, row: row).setFormula(
            "=IF(B\(row + 1)>=90,\"A\",IF(B\(row + 1)>=80,\"B\",IF(B\(row + 1)>=70,\"C\",\"F\")))"
        )

        // Pass/Fail
        try sheet.cell(column: 3, row: row).setFormula(
            "=IF(B\(row + 1)>=70,\"Pass\",\"Fail\")"
        )
    }

    try workbook.save(to: URL(fileURLWithPath: "output_if.xlsx"))
    print("Saved! Student grades with IF formulas")
}

/// Example 4: VLOOKUP (Lookup Tables)
func example4_VLOOKUP() throws {
    print("\n=== Example 4: VLOOKUP ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Create a lookup table in columns F-G
    try sheet.cell("F1").setValue(.string("Product Code"))
    try sheet.cell("G1").setValue(.string("Price"))

    let productCodes = ["A001", "A002", "A003", "A004"]
    let prices = [10.99, 25.50, 15.75, 30.00]

    for (index, code) in productCodes.enumerated() {
        try sheet.cell(column: 5, row: index + 1).setValue(.string(code))
        try sheet.cell(column: 6, row: index + 1).setValue(.number(prices[index]))
    }

    // Order form
    try sheet.cell("A1").setValue(.string("Order"))
    try sheet.cell("B1").setValue(.string("Product Code"))
    try sheet.cell("C1").setValue(.string("Quantity"))
    try sheet.cell("D1").setValue(.string("Price"))
    try sheet.cell("E1").setValue(.string("Total"))

    let orders = [
        ("A002", 3),
        ("A001", 5),
        ("A004", 2)
    ]

    for (index, order) in orders.enumerated() {
        let row = index + 1

        try sheet.cell(column: 0, row: row).setValue(.number(Double(index + 1)))
        try sheet.cell(column: 1, row: row).setValue(.string(order.0))
        try sheet.cell(column: 2, row: row).setValue(.number(Double(order.1)))

        // VLOOKUP to get price
        try sheet.cell(column: 3, row: row).setFormula(
            "=VLOOKUP(B\(row + 1),$F$2:$G$5,2,FALSE)"
        )

        // Calculate total
        try sheet.cell(column: 4, row: row).setFormula(
            "=C\(row + 1)*D\(row + 1)"
        )
    }

    try workbook.save(to: URL(fileURLWithPath: "output_vlookup.xlsx"))
    print("Saved! Order form with VLOOKUP formulas")
}

/// Example 5: Absolute vs Relative References
func example5_AbsoluteReferences() throws {
    print("\n=== Example 5: Absolute vs Relative References ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Tax rate in A1
    try sheet.cell("A1").setValue(.string("Tax Rate:"))
    try sheet.cell("B1").setValue(.number(0.08))  // 8%

    // Price table
    try sheet.cell("A3").setValue(.string("Item"))
    try sheet.cell("B3").setValue(.string("Price"))
    try sheet.cell("C3").setValue(.string("Tax"))
    try sheet.cell("D3").setValue(.string("Total"))

    let items = ["Book", "Pen", "Notebook"]
    let prices = [12.99, 2.50, 5.99]

    for (index, item) in items.enumerated() {
        let row = index + 3

        try sheet.cell(column: 0, row: row).setValue(.string(item))
        try sheet.cell(column: 1, row: row).setValue(.number(prices[index]))

        // Tax calculation using ABSOLUTE reference to B1
        try sheet.cell(column: 2, row: row).setFormula("=B\(row + 1)*$B$1")

        // Total
        try sheet.cell(column: 3, row: row).setFormula("=B\(row + 1)+C\(row + 1)")
    }

    try workbook.save(to: URL(fileURLWithPath: "output_absolute.xlsx"))
    print("Saved! Tax calculations with absolute reference ($B$1)")
}

/// Example 6: Cross-Sheet References
func example6_CrossSheetReferences() throws {
    print("\n=== Example 6: Cross-Sheet References ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))

    // Assume template has at least 2 sheets
    guard workbook.sheetCount >= 2 else {
        print("Template needs at least 2 sheets")
        return
    }

    let sheet1 = try workbook.sheet(at: 0)
    let sheet2 = try workbook.sheet(at: 1)

    // Put data in Sheet2
    try sheet2.cell("A1").setValue(.number(100))
    try sheet2.cell("A2").setValue(.number(200))
    try sheet2.cell("A3").setValue(.number(300))

    // Reference Sheet2 data from Sheet1
    try sheet1.cell("A1").setValue(.string("Data from Sheet2:"))
    try sheet1.cell("B1").setFormula("=Sheet2!A1")
    try sheet1.cell("B2").setFormula("=Sheet2!A2")
    try sheet1.cell("B3").setFormula("=Sheet2!A3")

    // Sum data from Sheet2
    try sheet1.cell("A5").setValue(.string("Total:"))
    try sheet1.cell("B5").setFormula("=SUM(Sheet2!A1:A3)")

    try workbook.save(to: URL(fileURLWithPath: "output_crosssheet.xlsx"))
    print("Saved! Formulas referencing data across sheets")
}

/// Example 7: Common Text Functions
func example7_TextFunctions() throws {
    print("\n=== Example 7: Text Functions ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Sample data
    try sheet.cell("A1").setValue(.string("John"))
    try sheet.cell("B1").setValue(.string("Doe"))

    // CONCATENATE
    try sheet.cell("C1").setValue(.string("Full Name:"))
    try sheet.cell("D1").setFormula("=CONCATENATE(A1,\" \",B1)")

    // Alternative: & operator
    try sheet.cell("D2").setFormula("=A1&\" \"&B1")

    // UPPER, LOWER, PROPER
    try sheet.cell("A4").setValue(.string("hello world"))
    try sheet.cell("B4").setFormula("=UPPER(A4)")  // HELLO WORLD
    try sheet.cell("C4").setFormula("=LOWER(A4)")  // hello world
    try sheet.cell("D4").setFormula("=PROPER(A4)") // Hello World

    // LEN (length)
    try sheet.cell("A6").setValue(.string("Test String"))
    try sheet.cell("B6").setFormula("=LEN(A6)")

    // LEFT, RIGHT, MID
    try sheet.cell("A8").setValue(.string("ApachePOI"))
    try sheet.cell("B8").setFormula("=LEFT(A8,6)")   // Apache
    try sheet.cell("C8").setFormula("=RIGHT(A8,3)")  // POI
    try sheet.cell("D8").setFormula("=MID(A8,7,3)")  // POI

    try workbook.save(to: URL(fileURLWithPath: "output_text.xlsx"))
    print("Saved! Text manipulation formulas")
}

/// Example 8: Date and Time Formulas
func example8_DateTimeFunctions() throws {
    print("\n=== Example 8: Date and Time Functions ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // TODAY and NOW
    try sheet.cell("A1").setValue(.string("Today:"))
    try sheet.cell("B1").setFormula("=TODAY()")

    try sheet.cell("A2").setValue(.string("Now:"))
    try sheet.cell("B2").setFormula("=NOW()")

    // Date arithmetic
    try sheet.cell("A4").setValue(.string("Start Date:"))
    try sheet.cell("B4").setValue(.date(Date()))

    try sheet.cell("A5").setValue(.string("End Date (30 days later):"))
    try sheet.cell("B5").setFormula("=B4+30")

    try sheet.cell("A6").setValue(.string("Days between:"))
    try sheet.cell("B6").setFormula("=B5-B4")

    // YEAR, MONTH, DAY
    try sheet.cell("A8").setValue(.string("Year:"))
    try sheet.cell("B8").setFormula("=YEAR(B4)")

    try sheet.cell("A9").setValue(.string("Month:"))
    try sheet.cell("B9").setFormula("=MONTH(B4)")

    try sheet.cell("A10").setValue(.string("Day:"))
    try sheet.cell("B10").setFormula("=DAY(B4)")

    // WEEKDAY
    try sheet.cell("A11").setValue(.string("Weekday:"))
    try sheet.cell("B11").setFormula("=WEEKDAY(B4)")

    try workbook.save(to: URL(fileURLWithPath: "output_datetime.xlsx"))
    print("Saved! Date and time formulas")
}

/// Example 9: Statistical Functions
func example9_StatisticalFunctions() throws {
    print("\n=== Example 9: Statistical Functions ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "template.xlsx"))
    let sheet = try workbook.sheet(at: 0)

    // Sample data
    try sheet.cell("A1").setValue(.string("Values:"))
    let values = [10.5, 20.3, 15.8, 22.1, 18.9, 25.4, 12.7]

    for (index, value) in values.enumerated() {
        try sheet.cell(column: 1, row: index).setValue(.number(value))
    }

    // Statistics
    try sheet.cell("D1").setValue(.string("Count:"))
    try sheet.cell("E1").setFormula("=COUNT(B1:B7)")

    try sheet.cell("D2").setValue(.string("Sum:"))
    try sheet.cell("E2").setFormula("=SUM(B1:B7)")

    try sheet.cell("D3").setValue(.string("Average:"))
    try sheet.cell("E3").setFormula("=AVERAGE(B1:B7)")

    try sheet.cell("D4").setValue(.string("Min:"))
    try sheet.cell("E4").setFormula("=MIN(B1:B7)")

    try sheet.cell("D5").setValue(.string("Max:"))
    try sheet.cell("E5").setFormula("=MAX(B1:B7)")

    try sheet.cell("D6").setValue(.string("Median:"))
    try sheet.cell("E6").setFormula("=MEDIAN(B1:B7)")

    try sheet.cell("D7").setValue(.string("Std Dev:"))
    try sheet.cell("E7").setFormula("=STDEV(B1:B7)")

    try workbook.save(to: URL(fileURLWithPath: "output_statistics.xlsx"))
    print("Saved! Statistical analysis formulas")
}

/// Example 10: Marbar Report Formulas (Real-World)
func example10_MarbarReportFormulas() throws {
    print("\n=== Example 10: Marbar Report Formulas ===")

    let workbook = try ExcelWorkbook(fileURL: URL(fileURLWithPath: "marbar_template.xlsm"))
    let sheet = try workbook.sheet(named: "GENERALES")

    // Example: Calculate total hours
    // Assume hours are entered in cells B10:B20

    // Total hours formula
    try sheet.cell("B25").setFormula("=SUM(B10:B20)")

    // Average hours per day
    try sheet.cell("B26").setFormula("=AVERAGE(B10:B20)")

    // Status indicator based on hours
    try sheet.cell("C25").setFormula("=IF(B25>100,\"Over Budget\",\"On Track\")")

    // Percentage calculation
    try sheet.cell("D25").setFormula("=B25/1000*100")

    // Complex nested formula for reporting
    try sheet.cell("E25").setFormula(
        "=IF(AND(B25>0,B25<200),\"Normal\",IF(B25>=200,\"High\",\"Low\"))"
    )

    try workbook.save(to: URL(fileURLWithPath: "marbar_with_formulas.xlsm"))
    print("Saved! Marbar report with automated calculations")
}

// Run all examples
func runAllFormulaExamples() {
    print("=== ApachePOISwift Formula Examples ===\n")

    do {
        try example1_BasicArithmetic()
        try example2_SUMAndAVERAGE()
        try example3_IFStatements()
        try example4_VLOOKUP()
        try example5_AbsoluteReferences()
        // try example6_CrossSheetReferences()  // Requires multi-sheet template
        try example7_TextFunctions()
        try example8_DateTimeFunctions()
        try example9_StatisticalFunctions()
        // try example10_MarbarReportFormulas()  // Requires marbar template

        print("\n✅ All formula examples completed successfully!")
        print("\nNote: Open the generated .xlsx files in Excel to see the formulas in action.")
        print("Excel will automatically calculate the results when you open the files.")
    } catch {
        print("❌ Error: \(error)")
    }
}

// Uncomment to run:
// runAllFormulaExamples()
