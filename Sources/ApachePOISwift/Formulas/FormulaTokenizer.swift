//
// FormulaTokenizer.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Tokenizes Excel formulas into a stream of tokens
class FormulaTokenizer {
    private let formula: String
    private var current: String.Index
    private var position: Int = 0

    init(formula: String) {
        // Remove leading = if present
        self.formula = formula.hasPrefix("=") ? String(formula.dropFirst()) : formula
        self.current = self.formula.startIndex
    }

    /// Tokenize the entire formula
    func tokenize() throws -> [FormulaToken] {
        var tokens: [FormulaToken] = []

        while !isAtEnd() {
            skipWhitespace()
            if isAtEnd() { break }

            let token = try scanToken()
            tokens.append(token)
        }

        tokens.append(FormulaToken(type: .eof, lexeme: "", position: position))
        return tokens
    }

    // MARK: - Token Scanning

    private func scanToken() throws -> FormulaToken {
        let start = position
        let char = advance()

        switch char {
        case "+": return makeToken(.plus, start)
        case "-": return makeToken(.minus, start)
        case "*": return makeToken(.multiply, start)
        case "/": return makeToken(.divide, start)
        case "^": return makeToken(.power, start)
        case "&": return makeToken(.concat, start)
        case "(": return makeToken(.leftParen, start)
        case ")": return makeToken(.rightParen, start)
        case ",": return makeToken(.comma, start)
        case ":": return makeToken(.colon, start)

        case "=":
            return makeToken(.equal, start)

        case "<":
            if match("=") {
                return makeToken(.lessOrEqual, start)
            } else if match(">") {
                return makeToken(.notEqual, start)
            } else {
                return makeToken(.lessThan, start)
            }

        case ">":
            if match("=") {
                return makeToken(.greaterOrEqual, start)
            } else {
                return makeToken(.greaterThan, start)
            }

        case "\"":
            return try scanString(start)

        case "0"..."9", ".":
            return try scanNumber(start, firstChar: char)

        case "A"..."Z", "a"..."z", "$":
            return try scanIdentifierOrReference(start, firstChar: char)

        default:
            throw ExcelError.parsingError("Unexpected character '\(char)' at position \(start)")
        }
    }

    // MARK: - String Scanning

    private func scanString(_ start: Int) throws -> FormulaToken {
        var value = ""

        while !isAtEnd() && peek() != "\"" {
            if peek() == "\\" && peekNext() == "\"" {
                // Escaped quote
                _ = advance()
                value.append(advance())
            } else {
                value.append(advance())
            }
        }

        if isAtEnd() {
            throw ExcelError.parsingError("Unterminated string at position \(start)")
        }

        // Consume closing "
        _ = advance()

        let lexeme = "\"" + value + "\""
        return FormulaToken(type: .string(value), lexeme: lexeme, position: start)
    }

    // MARK: - Number Scanning

    private func scanNumber(_ start: Int, firstChar: Character) throws -> FormulaToken {
        var numStr = String(firstChar)

        // Integer part
        while !isAtEnd() && peek().isNumber {
            numStr.append(advance())
        }

        // Decimal part
        if !isAtEnd() && peek() == "." && peekNext()?.isNumber == true {
            numStr.append(advance())  // .
            while !isAtEnd() && peek().isNumber {
                numStr.append(advance())
            }
        }

        // Scientific notation
        if !isAtEnd() && (peek() == "e" || peek() == "E") {
            let saved = current
            numStr.append(advance())  // e or E

            if !isAtEnd() && (peek() == "+" || peek() == "-") {
                numStr.append(advance())
            }

            if !isAtEnd() && peek().isNumber {
                while !isAtEnd() && peek().isNumber {
                    numStr.append(advance())
                }
            } else {
                // Not scientific notation, backtrack
                current = saved
                numStr.removeLast()
            }
        }

        guard let value = Double(numStr) else {
            throw ExcelError.parsingError("Invalid number '\(numStr)' at position \(start)")
        }

        return FormulaToken(type: .number(value), lexeme: numStr, position: start)
    }

    // MARK: - Identifier/Reference Scanning

    private func scanIdentifierOrReference(_ start: Int, firstChar: Character) throws -> FormulaToken {
        var text = String(firstChar)

        // Scan alphanumeric, $, !, .
        while !isAtEnd() {
            let char = peek()
            if char.isLetter || char.isNumber || char == "$" || char == "!" || char == "." || char == "_" {
                text.append(advance())
            } else {
                break
            }
        }

        // Check for range (A1:B10)
        if !isAtEnd() && peek() == ":" {
            _ = advance()  // consume :

            // Scan second reference
            var secondRef = ""
            while !isAtEnd() {
                let char = peek()
                if char.isLetter || char.isNumber || char == "$" || char == "!" {
                    secondRef.append(advance())
                } else {
                    break
                }
            }

            if !secondRef.isEmpty {
                let rangeText = text + ":" + secondRef
                return FormulaToken(type: .range(rangeText), lexeme: rangeText, position: start)
            }
        }

        // Check if it's a cell reference (A1, $A$1, Sheet1!A1, etc.)
        if isCellReference(text) {
            return FormulaToken(type: .cellReference(text), lexeme: text, position: start)
        }

        // Check for boolean literals
        let upper = text.uppercased()
        if upper == "TRUE" {
            return FormulaToken(type: .boolean(true), lexeme: text, position: start)
        }
        if upper == "FALSE" {
            return FormulaToken(type: .boolean(false), lexeme: text, position: start)
        }

        // Check if it's followed by ( - then it's a function
        skipWhitespace()
        if !isAtEnd() && peek() == "(" {
            return FormulaToken(type: .function(text.uppercased()), lexeme: text, position: start)
        }

        // Otherwise it's a named range or error
        throw ExcelError.parsingError("Unknown identifier '\(text)' at position \(start)")
    }

    // MARK: - Helper Methods

    private func isCellReference(_ text: String) -> Bool {
        // Simple check: contains letters and numbers
        // A1, $A$1, Sheet1!A1, etc.
        let hasLetter = text.contains(where: { $0.isLetter })
        let hasNumber = text.contains(where: { $0.isNumber })
        return hasLetter && (hasNumber || text.contains("!"))
    }

    private func makeToken(_ type: TokenType, _ start: Int) -> FormulaToken {
        let lexeme = formula[formula.index(formula.startIndex, offsetBy: start)..<current]
        return FormulaToken(type: type, lexeme: String(lexeme), position: start)
    }

    private func advance() -> Character {
        defer {
            current = formula.index(after: current)
            position += 1
        }
        return formula[current]
    }

    private func match(_ expected: Character) -> Bool {
        if isAtEnd() || peek() != expected {
            return false
        }
        _ = advance()
        return true
    }

    private func peek() -> Character {
        return isAtEnd() ? "\0" : formula[current]
    }

    private func peekNext() -> Character? {
        let next = formula.index(after: current)
        return next < formula.endIndex ? formula[next] : nil
    }

    private func skipWhitespace() {
        while !isAtEnd() && peek().isWhitespace {
            _ = advance()
        }
    }

    private func isAtEnd() -> Bool {
        return current >= formula.endIndex
    }
}
