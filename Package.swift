// swift-tools-version: 5.9
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PackageDescription

let package = Package(
    name: "ApachePOISwift",
    platforms: [
        .iOS(.v15),
        .macOS(.v12),
        .watchOS(.v8)
    ],
    products: [
        // Products define the executables and libraries a package produces, making them visible to other packages.
        .library(
            name: "ApachePOISwift",
            targets: ["ApachePOISwift"]),
    ],
    dependencies: [
        // ZIPFoundation for handling .xlsx/.xlsm ZIP archives
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", .upToNextMajor(from: "0.9.19"))
    ],
    targets: [
        // Targets are the basic building blocks of a package, defining a module or a test suite.
        // Targets can depend on other targets in this package and products from dependencies.
        .target(
            name: "ApachePOISwift",
            dependencies: ["ZIPFoundation"]),
        .testTarget(
            name: "ApachePOISwiftTests",
            dependencies: ["ApachePOISwift"],
            resources: [.copy("TestResources")]),
    ]
)
