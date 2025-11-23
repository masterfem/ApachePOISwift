//
// FormulaToken.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Token types in Excel formulas
enum TokenType: Equatable {
    // Literals
    case number(Double)
    case string(String)
    case boolean(Bool)
    case cellReference(String)  // A1, $A$1, Sheet1!A1
    case range(String)          // A1:B10

    // Operators
    case plus
    case minus
    case multiply
    case divide
    case power
    case concat              // &
    case equal
    case notEqual
    case lessThan
    case lessOrEqual
    case greaterThan
    case greaterOrEqual

    // Grouping
    case leftParen
    case rightParen
    case comma
    case colon

    // Functions
    case function(String)    // SUM, AVERAGE, etc.

    // Control
    case eof
}

/// Token with position information for error reporting
struct FormulaToken: Equatable {
    let type: TokenType
    let lexeme: String       // Original text
    let position: Int        // Position in formula

    init(type: TokenType, lexeme: String, position: Int) {
        self.type = type
        self.lexeme = lexeme
        self.position = position
    }
}
