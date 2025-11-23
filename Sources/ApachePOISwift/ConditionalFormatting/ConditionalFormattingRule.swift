//
// ConditionalFormattingRule.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Represents a single conditional formatting rule in Excel
public struct ConditionalFormattingRule {

    /// The type of conditional formatting rule
    public enum RuleType: String {
        case cellIs = "cellIs"
        case expression = "expression"
        case colorScale = "colorScale"
        case dataBar = "dataBar"
        case iconSet = "iconSet"
        case top10 = "top10"
        case uniqueValues = "uniqueValues"
        case duplicateValues = "duplicateValues"
        case containsText = "containsText"
        case notContainsText = "notContainsText"
        case beginsWith = "beginsWith"
        case endsWith = "endsWith"
        case containsBlanks = "containsBlanks"
        case notContainsBlanks = "notContainsBlanks"
        case containsErrors = "containsErrors"
        case notContainsErrors = "notContainsErrors"
        case timePeriod = "timePeriod"
        case aboveAverage = "aboveAverage"
    }

    /// Comparison operator for cellIs rules
    public enum Operator: String {
        case lessThan = "lessThan"
        case lessThanOrEqual = "lessThanOrEqual"
        case equal = "equal"
        case notEqual = "notEqual"
        case greaterThanOrEqual = "greaterThanOrEqual"
        case greaterThan = "greaterThan"
        case between = "between"
        case notBetween = "notBetween"
        case containsText = "containsText"
        case notContains = "notContains"
        case beginsWith = "beginsWith"
        case endsWith = "endsWith"
    }

    /// The type of this rule
    public let type: RuleType

    /// Priority of the rule (1 = highest priority)
    public let priority: Int

    /// Whether to stop processing further rules if this one matches
    public let stopIfTrue: Bool

    /// Comparison operator (for cellIs rules)
    public let `operator`: Operator?

    /// Formula expressions for the rule
    public let formulas: [String]

    /// Differential formatting to apply when rule matches
    public let formatting: ConditionalFormatting?

    /// Text value (for text-based rules)
    public let text: String?

    /// Time period (for timePeriod rules)
    public let timePeriod: String?

    /// Rank value (for top10 rules)
    public let rank: Int?

    /// Whether to apply to bottom values (for top10 rules)
    public let bottom: Bool?

    /// Whether to apply as percentage (for top10 rules)
    public let percent: Bool?

    /// Whether above average (for aboveAverage rules)
    public let aboveAverage: Bool?

    /// Whether equal to average (for aboveAverage rules)
    public let equalAverage: Bool?

    /// Standard deviation multiplier (for aboveAverage rules)
    public let stdDev: Int?

    public init(
        type: RuleType,
        priority: Int,
        stopIfTrue: Bool = false,
        operator: Operator? = nil,
        formulas: [String] = [],
        formatting: ConditionalFormatting? = nil,
        text: String? = nil,
        timePeriod: String? = nil,
        rank: Int? = nil,
        bottom: Bool? = nil,
        percent: Bool? = nil,
        aboveAverage: Bool? = nil,
        equalAverage: Bool? = nil,
        stdDev: Int? = nil
    ) {
        self.type = type
        self.priority = priority
        self.stopIfTrue = stopIfTrue
        self.operator = `operator`
        self.formulas = formulas
        self.formatting = formatting
        self.text = text
        self.timePeriod = timePeriod
        self.rank = rank
        self.bottom = bottom
        self.percent = percent
        self.aboveAverage = aboveAverage
        self.equalAverage = equalAverage
        self.stdDev = stdDev
    }
}

/// Represents the differential formatting applied by a CF rule
public struct ConditionalFormatting {
    /// Font properties
    public let font: Font?

    /// Fill properties
    public let fill: Fill?

    /// Border properties
    public let border: Border?

    /// Number format
    public let numberFormat: NumberFormat?

    public init(
        font: Font? = nil,
        fill: Fill? = nil,
        border: Border? = nil,
        numberFormat: NumberFormat? = nil
    ) {
        self.font = font
        self.fill = fill
        self.border = border
        self.numberFormat = numberFormat
    }
}

/// Represents a conditional formatting area (collection of rules for a range)
public struct ConditionalFormattingArea {
    /// The cell range this CF applies to (e.g., "A1:C10")
    public let range: String

    /// The rules that apply to this range
    public let rules: [ConditionalFormattingRule]

    public init(range: String, rules: [ConditionalFormattingRule]) {
        self.range = range
        self.rules = rules
    }
}
