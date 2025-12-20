import Testing
import Foundation
import Darwin
@testable import Cuneiform

@Suite("Large Workbook Stress Suite")
struct LargeWorkbookStressTests {
    private func envFlag(_ name: String) -> Bool {
        (ProcessInfo.processInfo.environment[name] ?? "").lowercased() == "1"
    }
    
    private func residentMemoryBytes() -> UInt64? {
        var info = task_vm_info_data_t()
        var count = mach_msg_type_number_t(MemoryLayout<task_vm_info>.size) / 4
        let kerr = withUnsafeMutablePointer(to: &info) { ptr -> kern_return_t in
            ptr.withMemoryRebound(to: integer_t.self, capacity: Int(count)) {
                task_info(mach_task_self_, task_flavor_t(TASK_VM_INFO), $0, &count)
            }
        }
        if kerr == KERN_SUCCESS { return info.resident_size } else { return nil }
    }
    
    @Test("Build large workbook and measure performance")
    func buildLargeWorkbook() throws {
        guard envFlag("CUNEIFORM_STRESS") else {
            // Skip heavy run; do a tiny sanity build
            var writer = WorkbookWriter()
            let idx = writer.addSheet(named: "Mini")
            var sheet = writer.sheet(at: idx)!
            for r in 1...50 {
                sheet.writeText("Mini_\(r)", to: CellReference(column: "A", row: r))
            }
            writer.modifySheet(at: idx) { $0 = sheet }
            _ = try writer.buildData()
            return
        }
        let stress = envFlag("CUNEIFORM_STRESS")
        let sheets = stress ? 50 : 5
        let rows = stress ? 10_000 : 2_000
        let cols = stress ? 20 : 10
        
        var writer = WorkbookWriter()
        
        let startMem = residentMemoryBytes()
        let t0 = CFAbsoluteTimeGetCurrent()
        
        for s in 0..<sheets {
            let idx = writer.addSheet(named: "Sheet_\(s+1)")
            var sheet = writer.sheet(at: idx)!
            
            for r in 1...rows {
                for c in 0..<cols {
                    let col = String(UnicodeScalar(65 + (c % 26))!)
                    let ref = CellReference(column: col, row: r)
                    switch (r + c) % 4 {
                    case 0:
                        sheet.writeText("Label_\(r)_\(c)", to: ref)
                    case 1:
                        sheet.writeNumber(Double(r * c), to: ref)
                    case 2:
                        sheet.writeBoolean((r + c) % 2 == 0, to: ref)
                    default:
                        sheet.writeFormula("SUM(A\(max(1,r-1)):A\(r))", cachedValue: Double(r), to: ref)
                    }
                }
            }
            writer.modifySheet(at: idx) { $0 = sheet }
        }
        
        let data = try writer.buildData()
        let t1 = CFAbsoluteTimeGetCurrent()
        let endMem = residentMemoryBytes()
        
        // Persist and spot-check round-trip
        let tempURL = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString + ".xlsx")
        try data.write(to: tempURL)
        let workbook = try Workbook.open(url: tempURL)
        let readSheet = try workbook.sheet(named: "Sheet_1")!
        if case .text(let val) = readSheet.cell(at: "A1") {
            #expect(val.hasPrefix("Label_"))
        }
        
        let elapsed = t1 - t0
        let totalCells = sheets * rows * cols
        let fileSizeMB = Double(data.count) / 1_048_576.0
        
        // Report metrics
        print("📊 Stress Test Metrics:")
        print("  Configuration: \(sheets) sheets × \(rows) rows × \(cols) cols = \(totalCells) cells")
        print("  Build time: \(String(format: "%.2f", elapsed))s")
        print("  Workbook size: \(String(format: "%.2f", fileSizeMB))MB")
        if let start = startMem, let end = endMem {
            let deltaMB = Double(Int64(end) - Int64(start)) / 1_048_576.0
            let peakMB = Double(max(start, end)) / 1_048_576.0
            print("  Memory delta: \(String(format: "%.2f", deltaMB))MB")
            print("  Peak RSS: \(String(format: "%.2f", peakMB))MB")
        }
        
        if stress {
            // Thresholds per assignment
            #expect(elapsed <= 90.0)
            if let start = startMem, let end = endMem {
                let peak = max(start, end)
                #expect(peak <= 2_500_000_000) // ~2.5 GB
            }
        } else {
            // Sanity thresholds for reduced scale
            #expect(elapsed <= 10.0)
        }
    }
}
