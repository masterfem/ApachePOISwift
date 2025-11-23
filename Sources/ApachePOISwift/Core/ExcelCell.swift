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

    /// Reference to parent sheet for tracking modifications
    private weak var sheet: ExcelSheet?

    /// Whether this cell has been modified
    internal var isModified: Bool = false

    init(reference: String, cellData: CellData, workbook: ExcelWorkbook?, sheet: ExcelSheet? = nil) {
        self.reference = reference
        self.cellData = cellData
        self.workbook = workbook
        self.sheet = sheet
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

    // MARK: - Write Support (Phase 2)

    /// Set the cell value
    /// - Parameter newValue: The new value to set
    public func setValue(_ newValue: CellValue) {
        switch newValue {
        case .string(let text):
            setStringValue(text)

        case .number(let num):
            setNumberValue(num)

        case .boolean(let bool):
            setBooleanValue(bool)

        case .formula(let formula):
            setFormulaValue(formula)

        case .empty:
            clearValue()

        case .date(let date):
            // Excel stores dates as numbers (days since 1900-01-01)
            let excelDate = dateToExcelNumber(date)
            setNumberValue(excelDate)
        }

        isModified = true
        sheet?.markAsModified()
    }

    // MARK: - Style Support

    /// Get the style index for this cell
    public var styleIndex: Int? {
        return cellData.styleIndex
    }

    /// Get the cell style (if available)
    public var style: CellStyle? {
        guard let styleIndex = cellData.styleIndex,
              let stylesData = workbook?.stylesData,
              styleIndex >= 0 && styleIndex < stylesData.cellStyles.count else {
            return nil
        }
        return stylesData.cellStyles[styleIndex]
    }

    /// Get the font for this cell
    public var font: Font? {
        guard let style = style,
              let fontId = style.fontId,
              let stylesData = workbook?.stylesData,
              fontId >= 0 && fontId < stylesData.fonts.count else {
            return nil
        }
        return stylesData.fonts[fontId]
    }

    /// Get the fill for this cell
    public var fill: Fill? {
        guard let style = style,
              let fillId = style.fillId,
              let stylesData = workbook?.stylesData,
              fillId >= 0 && fillId < stylesData.fills.count else {
            return nil
        }
        return stylesData.fills[fillId]
    }

    /// Get the border for this cell
    public var border: Border? {
        guard let style = style,
              let borderId = style.borderId,
              let stylesData = workbook?.stylesData,
              borderId >= 0 && borderId < stylesData.borders.count else {
            return nil
        }
        return stylesData.borders[borderId]
    }

    /// Get the number format for this cell
    public var numberFormat: NumberFormat? {
        guard let style = style,
              let numFmtId = style.numberFormatId else {
            return nil
        }

        // Check custom formats first
        if let stylesData = workbook?.stylesData {
            if let customFormat = stylesData.numberFormats.first(where: { $0.formatId == numFmtId }) {
                return customFormat
            }
        }

        // Return built-in format
        return NumberFormat(formatId: numFmtId)
    }

    /// Set the style index for this cell
    public func setStyleIndex(_ index: Int?) {
        cellData = CellData(
            reference: reference,
            type: cellData.type,
            value: cellData.value,
            formula: cellData.formula,
            styleIndex: index
        )
        isModified = true
        sheet?.markAsModified()
    }

    // MARK: - Style Creation (Phase 4B)

    /// Apply a font to this cell
    /// - Parameter font: Font to apply
    public func applyFont(_ font: Font) {
        guard let workbook = workbook else { return }

        let fontId = workbook.findOrCreateFont(font)

        // Get current style components
        let currentStyle = style
        let fillId = currentStyle?.fillId
        let borderId = currentStyle?.borderId
        let numberFormatId = currentStyle?.numberFormatId

        // Create new cell style with new font
        let styleIndex = workbook.findOrCreateCellStyle(
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            numberFormatId: numberFormatId
        )

        setStyleIndex(styleIndex)
    }

    /// Apply a fill (background) to this cell
    /// - Parameter fill: Fill to apply
    public func applyFill(_ fill: Fill) {
        guard let workbook = workbook else { return }

        let fillId = workbook.findOrCreateFill(fill)

        // Get current style components
        let currentStyle = style
        let fontId = currentStyle?.fontId
        let borderId = currentStyle?.borderId
        let numberFormatId = currentStyle?.numberFormatId

        // Create new cell style with new fill
        let styleIndex = workbook.findOrCreateCellStyle(
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            numberFormatId: numberFormatId
        )

        setStyleIndex(styleIndex)
    }

    /// Apply a border to this cell
    /// - Parameter border: Border to apply
    public func applyBorder(_ border: Border) {
        guard let workbook = workbook else { return }

        let borderId = workbook.findOrCreateBorder(border)

        // Get current style components
        let currentStyle = style
        let fontId = currentStyle?.fontId
        let fillId = currentStyle?.fillId
        let numberFormatId = currentStyle?.numberFormatId

        // Create new cell style with new border
        let styleIndex = workbook.findOrCreateCellStyle(
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            numberFormatId: numberFormatId
        )

        setStyleIndex(styleIndex)
    }

    /// Apply a number format to this cell
    /// - Parameter numberFormat: Number format to apply
    public func applyNumberFormat(_ numberFormat: NumberFormat) {
        guard let workbook = workbook else { return }

        // Get current style components
        let currentStyle = style
        let fontId = currentStyle?.fontId
        let fillId = currentStyle?.fillId
        let borderId = currentStyle?.borderId

        // Create new cell style with new number format
        let styleIndex = workbook.findOrCreateCellStyle(
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            numberFormatId: numberFormat.formatId
        )

        setStyleIndex(styleIndex)
    }

    /// Apply complete style to this cell
    /// - Parameters:
    ///   - font: Font (optional)
    ///   - fill: Fill (optional)
    ///   - border: Border (optional)
    ///   - numberFormat: Number format (optional)
    ///   - horizontalAlignment: Horizontal alignment (optional)
    ///   - verticalAlignment: Vertical alignment (optional)
    ///   - wrapText: Wrap text flag
    public func applyStyle(
        font: Font? = nil,
        fill: Fill? = nil,
        border: Border? = nil,
        numberFormat: NumberFormat? = nil,
        horizontalAlignment: HorizontalAlignment? = nil,
        verticalAlignment: VerticalAlignment? = nil,
        wrapText: Bool = false
    ) {
        guard let workbook = workbook else { return }

        let fontId = font.map { workbook.findOrCreateFont($0) }
        let fillId = fill.map { workbook.findOrCreateFill($0) }
        let borderId = border.map { workbook.findOrCreateBorder($0) }
        let numberFormatId = numberFormat?.formatId

        let styleIndex = workbook.findOrCreateCellStyle(
            fontId: fontId,
            fillId: fillId,
            borderId: borderId,
            numberFormatId: numberFormatId,
            horizontalAlignment: horizontalAlignment,
            verticalAlignment: verticalAlignment,
            wrapText: wrapText
        )

        setStyleIndex(styleIndex)
    }

    // MARK: - Style Helpers

    /// Make this cell's text bold
    public func makeBold() {
        let currentFont = font ?? Font(name: "Calibri", size: 11)
        var newFont = currentFont
        newFont.bold = true
        applyFont(newFont)
    }

    /// Make this cell's text italic
    public func makeItalic() {
        let currentFont = font ?? Font(name: "Calibri", size: 11)
        var newFont = currentFont
        newFont.italic = true
        applyFont(newFont)
    }

    /// Set background color for this cell
    /// - Parameter color: Color in ARGB hex format (e.g., "FFFF0000" for red)
    public func setBackgroundColor(_ color: String) {
        applyFill(Fill(patternType: .solid, foregroundColor: color))
    }

    /// Set border around this cell
    /// - Parameters:
    ///   - style: Border style
    ///   - color: Border color in ARGB hex format (optional)
    public func setBorder(style: BorderStyle, color: String? = nil) {
        let edge = BorderEdge(style: style, color: color)
        applyBorder(Border(left: edge, right: edge, top: edge, bottom: edge))
    }

    // MARK: - Value Modification

    /// Set a string value
    private func setStringValue(_ text: String) {
        // For Phase 2, we'll use inline strings (simpler than managing shared strings)
        // Phase 3 can optimize to use shared strings
        cellData = CellData(
            reference: reference,
            type: .inlineString,
            value: text,
            formula: nil,
            styleIndex: cellData.styleIndex  // Preserve existing style
        )
    }

    /// Set a numeric value
    private func setNumberValue(_ number: Double) {
        cellData = CellData(
            reference: reference,
            type: .number,
            value: String(number),
            formula: nil,
            styleIndex: cellData.styleIndex  // Preserve existing style
        )
    }

    /// Set a boolean value
    private func setBooleanValue(_ bool: Bool) {
        cellData = CellData(
            reference: reference,
            type: .boolean,
            value: bool ? "1" : "0",
            formula: nil,
            styleIndex: cellData.styleIndex  // Preserve existing style
        )
    }

    /// Set a formula
    private func setFormulaValue(_ formula: String) {
        cellData = CellData(
            reference: reference,
            type: nil,  // Formulas don't have a type attribute
            value: nil,  // Value will be calculated by Excel
            formula: formula,
            styleIndex: cellData.styleIndex  // Preserve existing style
        )
    }

    /// Clear the cell value
    private func clearValue() {
        cellData = CellData(
            reference: reference,
            type: nil,
            value: nil,
            formula: nil,
            styleIndex: cellData.styleIndex  // Preserve existing style
        )
    }

    /// Convert Swift Date to Excel number (days since 1900-01-01)
    private func dateToExcelNumber(_ date: Date) -> Double {
        // Excel epoch: January 1, 1900 (but Excel incorrectly treats 1900 as a leap year)
        let excelEpoch = Date(timeIntervalSince1970: -2209161600) // 1900-01-01 00:00:00 UTC
        let daysSinceEpoch = date.timeIntervalSince(excelEpoch) / 86400.0
        return daysSinceEpoch + 1  // Excel is 1-indexed
    }

    /// Get the internal cell data (for XML writing)
    internal func getCellData() -> CellData {
        return cellData
    }
}

extension ExcelCell: CustomStringConvertible {
    public var description: String {
        return "\(reference): \(value)"
    }
}
