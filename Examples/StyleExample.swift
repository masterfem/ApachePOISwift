//
// StyleExample.swift
// ApachePOISwift Examples
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation
import ApachePOISwift

/// Example demonstrating style reading and manipulation
func styleExample() {
    print("=== ApachePOISwift Style Example ===\n")

    // Example file path (adjust as needed)
    guard let templatePath = Bundle.main.path(forResource: "template", ofType: "xlsm") else {
        print("Error: Could not find template file")
        return
    }

    let templateURL = URL(fileURLWithPath: templatePath)

    do {
        // MARK: - Example 1: Read Cell Styles

        print("Example 1: Reading Cell Styles")
        print("--------------------------------")

        let workbook = try ExcelWorkbook(fileURL: templateURL)
        let sheet = try workbook.sheet(at: 0)

        // Find a styled cell
        for row in 0..<10 {
            for col in 0..<10 {
                if let cell = try? sheet.cell(column: col, row: row),
                   let styleIndex = cell.styleIndex {
                    print("\nCell \(cell.reference):")
                    print("  Style Index: \(styleIndex)")

                    // Access font
                    if let font = cell.font {
                        print("  Font: \(font.name ?? "unknown") \(font.size ?? 0)pt")
                        if font.bold { print("    Bold: Yes") }
                        if font.italic { print("    Italic: Yes") }
                        if let color = font.color { print("    Color: \(color)") }
                    }

                    // Access fill (background)
                    if let fill = cell.fill {
                        print("  Fill Pattern: \(fill.patternType.rawValue)")
                        if let fgColor = fill.foregroundColor {
                            print("    Foreground Color: \(fgColor)")
                        }
                    }

                    // Access border
                    if let border = cell.border {
                        var borderDesc: [String] = []
                        if let left = border.left {
                            borderDesc.append("left=\(left.style.rawValue)")
                        }
                        if let right = border.right {
                            borderDesc.append("right=\(right.style.rawValue)")
                        }
                        if let top = border.top {
                            borderDesc.append("top=\(top.style.rawValue)")
                        }
                        if let bottom = border.bottom {
                            borderDesc.append("bottom=\(bottom.style.rawValue)")
                        }
                        if !borderDesc.isEmpty {
                            print("  Border: \(borderDesc.joined(separator: ", "))")
                        }
                    }

                    // Access number format
                    if let numberFormat = cell.numberFormat {
                        print("  Number Format: \(numberFormat.getFormatCode())")
                    }

                    break  // Found one example, move on
                }
            }
        }

        // MARK: - Example 2: Copy Style from One Cell to Another

        print("\n\nExample 2: Copying Cell Style")
        print("------------------------------")

        let sourceCell = try sheet.cell("B5")
        let targetCell = try sheet.cell("C10")

        if let sourceStyleIndex = sourceCell.styleIndex {
            print("Copying style from \(sourceCell.reference) (style \(sourceStyleIndex)) to \(targetCell.reference)")

            // Copy style
            targetCell.setStyleIndex(sourceStyleIndex)

            // Set a value while preserving style
            targetCell.setValue(.string("Styled Text"))

            print("Style copied successfully!")
        } else {
            print("Source cell has no style")
        }

        // MARK: - Example 3: Find All Cells with Bold Font

        print("\n\nExample 3: Finding Bold Cells")
        print("-----------------------------")

        var boldCells: [(reference: String, value: String)] = []

        for row in 0..<20 {
            for col in 0..<10 {
                if let cell = try? sheet.cell(column: col, row: row),
                   let font = cell.font,
                   font.bold,
                   !cell.isEmpty {
                    let valueStr: String
                    switch cell.value {
                    case .string(let str): valueStr = str
                    case .number(let num): valueStr = String(num)
                    default: valueStr = "\(cell.value)"
                    }
                    boldCells.append((cell.reference, valueStr))
                }
            }
        }

        print("Found \(boldCells.count) bold cells:")
        for (ref, val) in boldCells.prefix(5) {
            print("  \(ref): \(val)")
        }
        if boldCells.count > 5 {
            print("  ... and \(boldCells.count - 5) more")
        }

        // MARK: - Example 4: Find All Cells with Background Color

        print("\n\nExample 4: Finding Colored Cells")
        print("--------------------------------")

        var coloredCells: [(reference: String, color: String)] = []

        for row in 0..<20 {
            for col in 0..<10 {
                if let cell = try? sheet.cell(column: col, row: row),
                   let fill = cell.fill,
                   fill.patternType == .solid,
                   let fgColor = fill.foregroundColor {
                    coloredCells.append((cell.reference, fgColor))
                }
            }
        }

        print("Found \(coloredCells.count) cells with background color:")
        for (ref, color) in coloredCells.prefix(5) {
            print("  \(ref): \(color)")
        }
        if coloredCells.count > 5 {
            print("  ... and \(coloredCells.count - 5) more")
        }

        // MARK: - Example 5: Save with Modified Styles

        print("\n\nExample 5: Saving with Style Preservation")
        print("-----------------------------------------")

        let outputURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("styled_output.xlsm")

        try workbook.save(to: outputURL)
        print("Saved workbook to: \(outputURL.path)")

        // Reload and verify
        let reloadedWorkbook = try ExcelWorkbook(fileURL: outputURL)
        let reloadedSheet = try reloadedWorkbook.sheet(at: 0)
        let reloadedCell = try reloadedSheet.cell("C10")

        if let reloadedFont = reloadedCell.font {
            print("Reloaded cell C10 font: \(reloadedFont.name ?? "unknown") \(reloadedFont.size ?? 0)pt")
            print("Styles preserved successfully! ✅")
        }

        // MARK: - Example 6: Inspect Workbook Styles

        print("\n\nExample 6: Workbook Style Statistics")
        print("------------------------------------")

        if let styles = workbook.stylesData {
            print("Total Fonts: \(styles.fonts.count)")
            print("Total Fills: \(styles.fills.count)")
            print("Total Borders: \(styles.borders.count)")
            print("Total Cell Styles: \(styles.cellStyles.count)")
            print("Custom Number Formats: \(styles.numberFormats.count)")

            // Show some example fonts
            print("\nSample Fonts:")
            for (index, font) in styles.fonts.prefix(5).enumerated() {
                let name = font.name ?? "unknown"
                let size = font.size ?? 0
                let attrs = [
                    font.bold ? "bold" : nil,
                    font.italic ? "italic" : nil
                ].compactMap { $0 }
                let attrsStr = attrs.isEmpty ? "" : " (\(attrs.joined(separator: ", ")))"
                print("  \(index): \(name) \(size)pt\(attrsStr)")
            }

            // Show some example fills
            print("\nSample Fills:")
            for (index, fill) in styles.fills.prefix(5).enumerated() {
                var desc = "  \(index): \(fill.patternType.rawValue)"
                if let fg = fill.foregroundColor {
                    desc += " fg=\(fg)"
                }
                print(desc)
            }
        }

        print("\n=== Style Example Complete ===")

    } catch {
        print("Error: \(error)")
    }
}

// Uncomment to run this example
// styleExample()
