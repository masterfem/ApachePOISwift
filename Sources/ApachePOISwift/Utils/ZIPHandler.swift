//
// ZIPHandler.swift
// ApachePOISwift
//
// Copyright 2024 ApachePOISwift Contributors
// Licensed under the Apache License, Version 2.0
//

import Foundation
import ZIPFoundation

/// Handles ZIP archive operations for Excel files (.xlsx/.xlsm)
public class ZIPHandler {

    /// Extract a ZIP archive (Excel file) to a temporary directory
    /// - Parameter url: URL of the .xlsx or .xlsm file
    /// - Returns: URL of the temporary directory containing extracted files
    /// - Throws: ExcelError if the file cannot be extracted
    public static func extractArchive(at url: URL) throws -> URL {
        // Verify file exists
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw ExcelError.fileNotFound(url)
        }

        // Create temporary directory
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString, isDirectory: true)

        try FileManager.default.createDirectory(
            at: tempDir,
            withIntermediateDirectories: true,
            attributes: nil
        )

        // Extract archive
        do {
            try FileManager.default.unzipItem(at: url, to: tempDir)
        } catch {
            // Clean up temp directory on error
            try? FileManager.default.removeItem(at: tempDir)
            throw ExcelError.invalidZIPArchive(url)
        }

        return tempDir
    }

    /// List all files in a ZIP archive
    /// - Parameter url: URL of the .xlsx or .xlsm file
    /// - Returns: Array of file paths within the archive
    /// - Throws: ExcelError if the archive cannot be read
    public static func listContents(at url: URL) throws -> [String] {
        guard let archive = Archive(url: url, accessMode: .read) else {
            throw ExcelError.invalidZIPArchive(url)
        }

        return archive.map { $0.path }
    }

    /// Create a ZIP archive from a directory
    /// - Parameters:
    ///   - directory: URL of the directory to compress
    ///   - destination: URL where the .xlsx or .xlsm file should be created
    /// - Throws: ExcelError if the archive cannot be created
    public static func createArchive(from directory: URL, to destination: URL) throws {
        // Remove existing file if present
        if FileManager.default.fileExists(atPath: destination.path) {
            try FileManager.default.removeItem(at: destination)
        }

        guard let archive = Archive(url: destination, accessMode: .create) else {
            throw ExcelError.fileWriteError("Cannot create archive at \(destination.path)")
        }

        // Get all files in directory recursively
        let fileManager = FileManager.default
        guard let enumerator = fileManager.enumerator(
            at: directory,
            includingPropertiesForKeys: [.isRegularFileKey],
            options: [.skipsHiddenFiles]
        ) else {
            throw ExcelError.fileWriteError("Cannot enumerate directory \(directory.path)")
        }

        // Add each file to archive
        for case let fileURL as URL in enumerator {
            // Skip directories (archive.addEntry handles them automatically)
            guard let isRegularFile = try? fileURL.resourceValues(forKeys: [.isRegularFileKey]).isRegularFile,
                  isRegularFile else {
                continue
            }

            // Calculate relative path
            guard let relativePath = fileURL.path.replacingOccurrences(
                of: directory.path + "/",
                with: ""
            ).removingPercentEncoding else {
                continue
            }

            // Add file to archive
            do {
                try archive.addEntry(
                    with: relativePath,
                    relativeTo: directory,
                    compressionMethod: .deflate
                )
            } catch {
                throw ExcelError.fileWriteError("Cannot add file \(relativePath) to archive: \(error.localizedDescription)")
            }
        }
    }

    /// Clean up a temporary directory
    /// - Parameter url: URL of the temporary directory to remove
    public static func cleanupTempDirectory(at url: URL) {
        try? FileManager.default.removeItem(at: url)
    }
}
