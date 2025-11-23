//
// Border.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents borders around a cell
public struct Border {
    /// Left border
    public var left: BorderEdge?

    /// Right border
    public var right: BorderEdge?

    /// Top border
    public var top: BorderEdge?

    /// Bottom border
    public var bottom: BorderEdge?

    /// Diagonal border
    public var diagonal: BorderEdge?

    /// Diagonal up flag
    public var diagonalUp: Bool

    /// Diagonal down flag
    public var diagonalDown: Bool

    public init(
        left: BorderEdge? = nil,
        right: BorderEdge? = nil,
        top: BorderEdge? = nil,
        bottom: BorderEdge? = nil,
        diagonal: BorderEdge? = nil,
        diagonalUp: Bool = false,
        diagonalDown: Bool = false
    ) {
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown
    }
}

/// Represents a single border edge (left, right, top, bottom, or diagonal)
public struct BorderEdge {
    /// Border style
    public var style: BorderStyle

    /// Border color (ARGB hex format, e.g., "FF000000" for black)
    public var color: String?

    public init(style: BorderStyle, color: String? = nil) {
        self.style = style
        self.color = color
    }
}

/// Border line styles
public enum BorderStyle: String {
    case none = "none"
    case thin = "thin"
    case medium = "medium"
    case thick = "thick"
    case double = "double"
    case dotted = "dotted"
    case dashed = "dashed"
    case dashDot = "dashDot"
    case dashDotDot = "dashDotDot"
    case hair = "hair"
    case mediumDashed = "mediumDashed"
    case mediumDashDot = "mediumDashDot"
    case mediumDashDotDot = "mediumDashDotDot"
    case slantDashDot = "slantDashDot"
}

extension Border: Equatable {}
extension BorderEdge: Equatable {}
