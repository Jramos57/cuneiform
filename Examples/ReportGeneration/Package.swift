// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "ReportGenerationExample",
    platforms: [.macOS(.v13)],
    dependencies: [
        .package(path: "../..")
    ],
    targets: [
        .executableTarget(
            name: "ReportGenerationExample",
            dependencies: [
                .product(name: "Cuneiform", package: "cuneiform")
            ],
            path: "Sources"
        )
    ]
)
