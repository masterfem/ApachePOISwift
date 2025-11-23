//
// CellStyle.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a complete cell style (combination of font, fill, border, number format, and alignment)
public struct CellStyle {
    /// Index in the cellXfs array
    public var index: Int

    /// Font index (references fonts array)
    public var fontId: Int?

    /// Fill index (references fills array)
    public var fillId: Int?

    /// Border index (references borders array)
    public var borderId: Int?

    /// Number format index
    public var numberFormatId: Int?

    /// Horizontal alignment
    public var horizontalAlignment: HorizontalAlignment?

    /// Vertical alignment
    public var verticalAlignment: VerticalAlignment?

    /// Text wrap flag
    public var wrapText: Bool

    /// Text rotation (0-180 degrees, or 255 for vertical text)
    public var textRotation: Int?

    /// Indent level
    public var indent: Int?

    /// Apply font flag
    public var applyFont: Bool

    /// Apply fill flag
    public var applyFill: Bool

    /// Apply border flag
    public var applyBorder: Bool

    /// Apply number format flag
    public var applyNumberFormat: Bool

    /// Apply alignment flag
    public var applyAlignment: Bool

    public init(
        index: Int,
        fontId: Int? = nil,
        fillId: Int? = nil,
        borderId: Int? = nil,
        numberFormatId: Int? = nil,
        horizontalAlignment: HorizontalAlignment? = nil,
        verticalAlignment: VerticalAlignment? = nil,
        wrapText: Bool = false,
        textRotation: Int? = nil,
        indent: Int? = nil,
        applyFont: Bool = false,
        applyFill: Bool = false,
        applyBorder: Bool = false,
        applyNumberFormat: Bool = false,
        applyAlignment: Bool = false
    ) {
        self.index = index
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.numberFormatId = numberFormatId
        self.horizontalAlignment = horizontalAlignment
        self.verticalAlignment = verticalAlignment
        self.wrapText = wrapText
        self.textRotation = textRotation
        self.indent = indent
        self.applyFont = applyFont
        self.applyFill = applyFill
        self.applyBorder = applyBorder
        self.applyNumberFormat = applyNumberFormat
        self.applyAlignment = applyAlignment
    }
}

/// Horizontal alignment options
public enum HorizontalAlignment: String {
    case general = "general"
    case left = "left"
    case center = "center"
    case right = "right"
    case fill = "fill"
    case justify = "justify"
    case centerContinuous = "centerContinuous"
    case distributed = "distributed"
}

/// Vertical alignment options
public enum VerticalAlignment: String {
    case top = "top"
    case center = "center"
    case bottom = "bottom"
    case justify = "justify"
    case distributed = "distributed"
}

extension CellStyle: Equatable {}
