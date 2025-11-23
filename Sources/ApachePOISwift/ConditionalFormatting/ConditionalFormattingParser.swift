//
// ConditionalFormattingParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Parses conditional formatting from worksheet XML
class ConditionalFormattingParser: NSObject, XMLParserDelegate {

    private var conditionalFormattingAreas: [ConditionalFormattingArea] = []

    // Current parsing state
    private var currentRange: String?
    private var currentRules: [ConditionalFormattingRule] = []
    private var currentRuleType: ConditionalFormattingRule.RuleType?
    private var currentPriority: Int?
    private var currentOperator: ConditionalFormattingRule.Operator?
    private var currentFormulas: [String] = []
    private var currentText: String?
    private var currentStopIfTrue: Bool = false
    private var currentTimePeriod: String?
    private var currentRank: Int?
    private var currentBottom: Bool?
    private var currentPercent: Bool?
    private var currentAboveAverage: Bool?
    private var currentEqualAverage: Bool?
    private var currentStdDev: Int?
    private var currentFormulaText = ""
    private var isParsingFormula = false

    func parse(_ xmlString: String) -> [ConditionalFormattingArea] {
        conditionalFormattingAreas = []

        guard let data = xmlString.data(using: .utf8) else {
            return []
        }

        let parser = XMLParser(data: data)
        parser.delegate = self
        parser.parse()

        return conditionalFormattingAreas
    }

    // MARK: - XMLParserDelegate

    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String: String]) {

        switch elementName {
        case "conditionalFormatting":
            // Start of a CF area
            currentRange = attributeDict["sqref"]
            currentRules = []

        case "cfRule":
            // Start of a CF rule
            currentRuleType = nil
            currentPriority = nil
            currentOperator = nil
            currentFormulas = []
            currentText = nil
            currentStopIfTrue = false
            currentTimePeriod = nil
            currentRank = nil
            currentBottom = nil
            currentPercent = nil
            currentAboveAverage = nil
            currentEqualAverage = nil
            currentStdDev = nil

            // Parse rule type
            if let typeStr = attributeDict["type"] {
                currentRuleType = ConditionalFormattingRule.RuleType(rawValue: typeStr)
            }

            // Parse priority
            if let priorityStr = attributeDict["priority"], let priority = Int(priorityStr) {
                currentPriority = priority
            }

            // Parse operator
            if let operatorStr = attributeDict["operator"] {
                currentOperator = ConditionalFormattingRule.Operator(rawValue: operatorStr)
            }

            // Parse text
            currentText = attributeDict["text"]

            // Parse stopIfTrue
            if let stopIfTrueStr = attributeDict["stopIfTrue"], stopIfTrueStr == "1" {
                currentStopIfTrue = true
            }

            // Parse timePeriod
            currentTimePeriod = attributeDict["timePeriod"]

            // Parse rank (for top10)
            if let rankStr = attributeDict["rank"], let rank = Int(rankStr) {
                currentRank = rank
            }

            // Parse bottom
            if let bottomStr = attributeDict["bottom"], bottomStr == "1" {
                currentBottom = true
            }

            // Parse percent
            if let percentStr = attributeDict["percent"], percentStr == "1" {
                currentPercent = true
            }

            // Parse aboveAverage
            if let aboveAvgStr = attributeDict["aboveAverage"], aboveAvgStr == "0" {
                currentAboveAverage = false
            } else {
                currentAboveAverage = true  // Default is true
            }

            // Parse equalAverage
            if let equalAvgStr = attributeDict["equalAverage"], equalAvgStr == "1" {
                currentEqualAverage = true
            }

            // Parse stdDev
            if let stdDevStr = attributeDict["stdDev"], let stdDev = Int(stdDevStr) {
                currentStdDev = stdDev
            }

        case "formula":
            // Start parsing formula text
            isParsingFormula = true
            currentFormulaText = ""

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if isParsingFormula {
            currentFormulaText += string
        }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {

        switch elementName {
        case "formula":
            // End of formula
            isParsingFormula = false
            if !currentFormulaText.isEmpty {
                currentFormulas.append(currentFormulaText.trimmingCharacters(in: .whitespacesAndNewlines))
            }

        case "cfRule":
            // End of CF rule - create the rule
            guard let type = currentRuleType, let priority = currentPriority else {
                return
            }

            let rule = ConditionalFormattingRule(
                type: type,
                priority: priority,
                stopIfTrue: currentStopIfTrue,
                operator: currentOperator,
                formulas: currentFormulas,
                formatting: nil,  // Not parsing differential formatting for now
                text: currentText,
                timePeriod: currentTimePeriod,
                rank: currentRank,
                bottom: currentBottom,
                percent: currentPercent,
                aboveAverage: currentAboveAverage,
                equalAverage: currentEqualAverage,
                stdDev: currentStdDev
            )

            currentRules.append(rule)

        case "conditionalFormatting":
            // End of CF area
            if let range = currentRange, !currentRules.isEmpty {
                let area = ConditionalFormattingArea(range: range, rules: currentRules)
                conditionalFormattingAreas.append(area)
            }

        default:
            break
        }
    }
}
