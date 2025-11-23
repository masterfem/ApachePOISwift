//
// Fill.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a fill (background) pattern in Excel
public struct Fill {
    /// Pattern type
    public var patternType: PatternType

    /// Foreground color (ARGB hex format, e.g., "FFFF0000" for red)
    public var foregroundColor: String?

    /// Background color (ARGB hex format)
    public var backgroundColor: String?

    public init(
        patternType: PatternType = .none,
        foregroundColor: String? = nil,
        backgroundColor: String? = nil
    ) {
        self.patternType = patternType
        self.foregroundColor = foregroundColor
        self.backgroundColor = backgroundColor
    }
}

/// Fill pattern types
public enum PatternType: String {
    case none = "none"
    case solid = "solid"
    case gray125 = "gray125"
    case gray0625 = "gray0625"
    case darkGray = "darkGray"
    case mediumGray = "mediumGray"
    case lightGray = "lightGray"
    case darkHorizontal = "darkHorizontal"
    case darkVertical = "darkVertical"
    case darkDown = "darkDown"
    case darkUp = "darkUp"
    case darkGrid = "darkGrid"
    case darkTrellis = "darkTrellis"
    case lightHorizontal = "lightHorizontal"
    case lightVertical = "lightVertical"
    case lightDown = "lightDown"
    case lightUp = "lightUp"
    case lightGrid = "lightGrid"
    case lightTrellis = "lightTrellis"
}

extension Fill: Equatable {}
