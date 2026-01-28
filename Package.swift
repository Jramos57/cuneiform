// swift-tools-version: 6.0
// Pure Swift library for reading and writing Office Open XML SpreadsheetML (.xlsx) files

import PackageDescription

let package = Package(
    name: "Cuneiform",
    defaultLocalization: "en",
    platforms: [
        .macOS(.v13),
        .iOS(.v16),
        .tvOS(.v16),
        .watchOS(.v9),
        .visionOS(.v1)
    ],
    products: [
        .library(
            name: "Cuneiform",
            targets: ["Cuneiform"]
        ),
    ],
    dependencies: [
        .package(url: "https://github.com/swiftlang/swift-testing.git", from: "0.12.0"),
    ],
    targets: [
        .target(
            name: "Cuneiform"
        ),
        .testTarget(
            name: "CuneiformTests",
            dependencies: [
                "Cuneiform",
                .product(name: "Testing", package: "swift-testing"),
            ]
        ),
    ]
)
