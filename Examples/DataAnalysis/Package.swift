// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "DataAnalysisExample",
    platforms: [.macOS(.v13)],
    dependencies: [
        .package(path: "../..")
    ],
    targets: [
        .executableTarget(
            name: "DataAnalysisExample",
            dependencies: [
                .product(name: "Cuneiform", package: "cuneiform")
            ],
            path: "Sources"
        )
    ]
)
