//
// FormulaParser.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Parses Excel formulas using recursive descent with operator precedence
class FormulaParser {
    private var tokens: [FormulaToken] = []
    private var current: Int = 0

    /// Parse a formula string into an AST
    func parse(_ formula: String) throws -> FormulaAST {
        let tokenizer = FormulaTokenizer(formula: formula)
        tokens = try tokenizer.tokenize()
        current = 0

        let ast = try expression()

        if !isAtEnd() {
            throw ExcelError.parsingError("Unexpected tokens after expression at position \(peek().position)")
        }

        return ast
    }

    // MARK: - Expression Parsing (Precedence Climbing)

    /// expression → comparison
    private func expression() throws -> FormulaAST {
        return try comparison()
    }

    /// comparison → concatenation ( ("=" | "<>" | "<" | "<=" | ">" | ">=") concatenation )*
    private func comparison() throws -> FormulaAST {
        var expr = try concatenation()

        while match(.equal, .notEqual, .lessThan, .lessOrEqual, .greaterThan, .greaterOrEqual) {
            let op = previous()
            let right = try concatenation()

            expr = switch op.type {
            case .equal: .equal(expr, right)
            case .notEqual: .notEqual(expr, right)
            case .lessThan: .lessThan(expr, right)
            case .lessOrEqual: .lessOrEqual(expr, right)
            case .greaterThan: .greaterThan(expr, right)
            case .greaterOrEqual: .greaterOrEqual(expr, right)
            default: expr  // Should never happen
            }
        }

        return expr
    }

    /// concatenation → addition ( "&" addition )*
    private func concatenation() throws -> FormulaAST {
        var expr = try addition()

        while match(.concat) {
            let right = try addition()
            expr = .concat(expr, right)
        }

        return expr
    }

    /// addition → multiplication ( ("+" | "-") multiplication )*
    private func addition() throws -> FormulaAST {
        var expr = try multiplication()

        while match(.plus, .minus) {
            let op = previous()
            let right = try multiplication()

            expr = switch op.type {
            case .plus: .add(expr, right)
            case .minus: .subtract(expr, right)
            default: expr  // Should never happen
            }
        }

        return expr
    }

    /// multiplication → power ( ("*" | "/") power )*
    private func multiplication() throws -> FormulaAST {
        var expr = try power()

        while match(.multiply, .divide) {
            let op = previous()
            let right = try power()

            expr = switch op.type {
            case .multiply: .multiply(expr, right)
            case .divide: .divide(expr, right)
            default: expr  // Should never happen
            }
        }

        return expr
    }

    /// power → unary ( "^" unary )*
    private func power() throws -> FormulaAST {
        var expr = try unary()

        while match(.power) {
            let right = try unary()
            expr = .power(expr, right)
        }

        return expr
    }

    /// unary → ("+" | "-") unary | primary
    private func unary() throws -> FormulaAST {
        if match(.minus) {
            let expr = try unary()
            return .negate(expr)
        }

        if match(.plus) {
            let expr = try unary()
            return .positive(expr)
        }

        return try primary()
    }

    /// primary → NUMBER | STRING | BOOLEAN | CELL_REF | RANGE | FUNCTION | "(" expression ")"
    private func primary() throws -> FormulaAST {
        // Number
        if case .number(let value) = peek().type {
            advance()
            return .number(value)
        }

        // String
        if case .string(let value) = peek().type {
            advance()
            return .string(value)
        }

        // Boolean
        if case .boolean(let value) = peek().type {
            advance()
            return .boolean(value)
        }

        // Cell reference
        if case .cellReference(let ref) = peek().type {
            advance()
            return .cellReference(ref)
        }

        // Range
        if case .range(let range) = peek().type {
            advance()
            return .range(range)
        }

        // Function call
        if case .function(let name) = peek().type {
            return try parseFunction(name)
        }

        // Grouped expression
        if match(.leftParen) {
            let expr = try expression()

            if !match(.rightParen) {
                throw ExcelError.parsingError("Expected ')' after expression at position \(peek().position)")
            }

            return expr
        }

        throw ExcelError.parsingError("Expected expression at position \(peek().position)")
    }

    // MARK: - Function Parsing

    private func parseFunction(_ name: String) throws -> FormulaAST {
        advance()  // Consume function name

        if !match(.leftParen) {
            throw ExcelError.parsingError("Expected '(' after function name at position \(peek().position)")
        }

        var arguments: [FormulaAST] = []

        // Parse arguments
        if !check(.rightParen) {
            repeat {
                let arg = try expression()
                arguments.append(arg)
            } while match(.comma)
        }

        if !match(.rightParen) {
            throw ExcelError.parsingError("Expected ')' after function arguments at position \(peek().position)")
        }

        return .function(name: name, arguments: arguments)
    }

    // MARK: - Token Navigation

    private func match(_ types: TokenType...) -> Bool {
        for type in types {
            if check(type) {
                advance()
                return true
            }
        }
        return false
    }

    private func check(_ type: TokenType) -> Bool {
        if isAtEnd() { return false }
        return peek().type == type
    }

    private func advance() {
        if !isAtEnd() {
            current += 1
        }
    }

    private func peek() -> FormulaToken {
        return tokens[current]
    }

    private func previous() -> FormulaToken {
        return tokens[current - 1]
    }

    private func isAtEnd() -> Bool {
        return peek().type == .eof
    }
}
