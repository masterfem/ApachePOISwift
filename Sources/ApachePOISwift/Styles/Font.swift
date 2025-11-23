//
// Font.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a font in Excel
public struct Font {
    /// Font name (e.g., "Calibri", "Arial")
    public var name: String?

    /// Font size in points
    public var size: Double?

    /// Bold flag
    public var bold: Bool

    /// Italic flag
    public var italic: Bool

    /// Underline style
    public var underline: UnderlineStyle?

    /// Strike-through flag
    public var strikethrough: Bool

    /// Font color (ARGB hex format, e.g., "FFFF0000" for red)
    public var color: String?

    /// Font family (1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative)
    public var family: Int?

    /// Character set
    public var charset: Int?

    /// Public initializer for creating fonts
    public init(
        name: String? = nil,
        size: Double? = nil,
        bold: Bool = false,
        italic: Bool = false,
        underline: UnderlineStyle? = nil,
        strikethrough: Bool = false,
        color: String? = nil,
        family: Int? = nil,
        charset: Int? = nil
    ) {
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.color = color
        self.family = family
        self.charset = charset
    }
}

/// Font underline styles
public enum UnderlineStyle: String {
    case single = "single"
    case double = "double"
    case singleAccounting = "singleAccounting"
    case doubleAccounting = "doubleAccounting"
    case none = "none"
}

extension Font: Equatable {}
