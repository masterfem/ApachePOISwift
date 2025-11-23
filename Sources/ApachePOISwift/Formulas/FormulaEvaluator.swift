//
// FormulaEvaluator.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation

/// Evaluates Excel formula AST to produce values
public class FormulaEvaluator {
    private let workbook: ExcelWorkbook
    private let functionLibrary: ExcelFunctionLibrary
    private var evaluationCache: [String: ExcelValue] = [:]

    public init(workbook: ExcelWorkbook) {
        self.workbook = workbook
        self.functionLibrary = ExcelFunctionLibrary()
    }

    /// Evaluate a formula string
    public func evaluate(_ formula: String, in sheet: ExcelSheet) throws -> ExcelValue {
        let parser = FormulaParser()
        let ast = try parser.parse(formula)
        return try evaluate(ast, in: sheet)
    }

    /// Evaluate a formula AST
    public func evaluate(_ ast: FormulaAST, in sheet: ExcelSheet) throws -> ExcelValue {
        switch ast {
        // Literals
        case .number(let value):
            return .number(value)

        case .string(let value):
            return .string(value)

        case .boolean(let value):
            return .boolean(value)

        case .cellReference(let ref):
            return try evaluateCellReference(ref, currentSheet: sheet)

        case .range(let range):
            return try evaluateRange(range, currentSheet: sheet)

        // Arithmetic operations
        case .add(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                guard let aNum = a.toNumber(), let bNum = b.toNumber() else {
                    return .error(.value)
                }
                return .number(aNum + bNum)
            }

        case .subtract(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                guard let aNum = a.toNumber(), let bNum = b.toNumber() else {
                    return .error(.value)
                }
                return .number(aNum - bNum)
            }

        case .multiply(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                guard let aNum = a.toNumber(), let bNum = b.toNumber() else {
                    return .error(.value)
                }
                return .number(aNum * bNum)
            }

        case .divide(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                guard let aNum = a.toNumber(), let bNum = b.toNumber() else {
                    return .error(.value)
                }
                guard bNum != 0 else {
                    return .error(.divideByZero)
                }
                return .number(aNum / bNum)
            }

        case .power(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                guard let aNum = a.toNumber(), let bNum = b.toNumber() else {
                    return .error(.value)
                }
                return .number(pow(aNum, bNum))
            }

        case .concat(let left, let right):
            return try evaluateBinaryOp(left, right, sheet) { a, b in
                return .string(a.toString() + b.toString())
            }

        // Comparison operations
        case .equal(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 == .orderedSame }

        case .notEqual(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 != .orderedSame }

        case .lessThan(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 == .orderedAscending }

        case .lessOrEqual(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 == .orderedAscending || $0 == .orderedSame }

        case .greaterThan(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 == .orderedDescending }

        case .greaterOrEqual(let left, let right):
            return try evaluateComparison(left, right, sheet) { $0 == .orderedDescending || $0 == .orderedSame }

        // Unary operations
        case .negate(let expr):
            let value = try evaluate(expr, in: sheet)
            guard let num = value.toNumber() else {
                return .error(.value)
            }
            return .number(-num)

        case .positive(let expr):
            return try evaluate(expr, in: sheet)

        // Function call
        case .function(let name, let arguments):
            return try evaluateFunction(name, arguments, sheet)
        }
    }

    // MARK: - Helper Methods

    private func evaluateBinaryOp(
        _ left: FormulaAST,
        _ right: FormulaAST,
        _ sheet: ExcelSheet,
        _ op: (ExcelValue, ExcelValue) -> ExcelValue
    ) throws -> ExcelValue {
        let leftValue = try evaluate(left, in: sheet)
        let rightValue = try evaluate(right, in: sheet)

        // Propagate errors
        if leftValue.isError { return leftValue }
        if rightValue.isError { return rightValue }

        return op(leftValue, rightValue)
    }

    private func evaluateComparison(
        _ left: FormulaAST,
        _ right: FormulaAST,
        _ sheet: ExcelSheet,
        _ compare: (ComparisonResult) -> Bool
    ) throws -> ExcelValue {
        let leftValue = try evaluate(left, in: sheet)
        let rightValue = try evaluate(right, in: sheet)

        // Propagate errors
        if leftValue.isError { return leftValue }
        if rightValue.isError { return rightValue }

        guard let result = leftValue.compare(rightValue) else {
            return .error(.value)
        }

        return .boolean(compare(result))
    }

    private func evaluateCellReference(_ ref: String, currentSheet: ExcelSheet) throws -> ExcelValue {
        // Check cache first
        if let cached = evaluationCache[ref] {
            return cached
        }

        // Parse sheet name if present (Sheet1!A1)
        let (sheetName, cellRef) = parseSheetReference(ref)

        // Get the sheet
        let sheet: ExcelSheet
        if let name = sheetName {
            sheet = try workbook.sheet(named: name)
        } else {
            sheet = currentSheet
        }

        // Get the cell
        let cell = try sheet.cell(cellRef)

        // Get the value
        let value: ExcelValue
        switch cell.value {
        case .string(let str):
            value = .string(str)
        case .number(let num):
            value = .number(num)
        case .boolean(let bool):
            value = .boolean(bool)
        case .date(let date):
            // Convert date to Excel serial number
            let excelEpoch = Date(timeIntervalSince1970: -2209161600)  // December 30, 1899
            let days = date.timeIntervalSince(excelEpoch) / 86400
            value = .number(days)
        case .formula(let formula):
            // Recursively evaluate formula
            value = try evaluate(formula, in: sheet)
        case .empty:
            value = .empty
        }

        // Cache the result
        evaluationCache[ref] = value

        return value
    }

    private func evaluateRange(_ range: String, currentSheet: ExcelSheet) throws -> ExcelValue {
        // For now, return the range as an error (full array support would be complex)
        // A complete implementation would expand the range into an array of values
        // and functions like SUM would accept arrays

        // Parse range: A1:B10
        let parts = range.split(separator: ":")
        guard parts.count == 2 else {
            return .error(.ref)
        }

        // For MVP, if used in a context that expects a single value, return first cell
        let firstCell = String(parts[0])
        return try evaluateCellReference(firstCell, currentSheet: currentSheet)
    }

    private func evaluateFunction(_ name: String, _ arguments: [FormulaAST], _ sheet: ExcelSheet) throws -> ExcelValue {
        guard let function = functionLibrary.get(name) else {
            return .error(.name)
        }

        // Evaluate all arguments
        var argValues: [ExcelValue] = []
        for arg in arguments {
            // Special handling for ranges in aggregate functions
            if case .range(let range) = arg {
                let values = try expandRange(range, currentSheet: sheet)
                argValues.append(contentsOf: values)
            } else {
                let value = try evaluate(arg, in: sheet)
                argValues.append(value)
            }
        }

        // Call the function
        return try function(argValues)
    }

    private func expandRange(_ range: String, currentSheet: ExcelSheet) throws -> [ExcelValue] {
        // Parse range: A1:B10
        let parts = range.split(separator: ":")
        guard parts.count == 2 else {
            throw ExcelError.parsingError("Invalid range: \(range)")
        }

        let start = try CellReference(String(parts[0]))
        let end = try CellReference(String(parts[1]))

        var values: [ExcelValue] = []

        for row in start.row...end.row {
            for col in start.column...end.column {
                let cell = try currentSheet.cell(column: col, row: row)
                let value: ExcelValue

                switch cell.value {
                case .string(let str):
                    value = .string(str)
                case .number(let num):
                    value = .number(num)
                case .boolean(let bool):
                    value = .boolean(bool)
                case .date(let date):
                    let excelEpoch = Date(timeIntervalSince1970: -2209161600)
                    let days = date.timeIntervalSince(excelEpoch) / 86400
                    value = .number(days)
                case .formula(let formula):
                    value = try evaluate(formula, in: currentSheet)
                case .empty:
                    value = .empty
                }

                values.append(value)
            }
        }

        return values
    }

    private func parseSheetReference(_ ref: String) -> (sheet: String?, cell: String) {
        if let bangIndex = ref.firstIndex(of: "!") {
            let sheetName = String(ref[..<bangIndex])
            let cellRef = String(ref[ref.index(after: bangIndex)...])
            return (sheetName, cellRef)
        }
        return (nil, ref)
    }

    /// Clear evaluation cache (call when workbook data changes)
    public func clearCache() {
        evaluationCache.removeAll()
    }
}
