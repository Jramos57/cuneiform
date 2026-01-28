#!/usr/bin/env swift
import Foundation
import Darwin

// Minimal mock types to compile standalone (we'll run via swift test instead)
print("""
Stress Benchmark Runner
=======================
Run via: CUNEIFORM_STRESS=1 swift test --filter LargeWorkbook

This script documents the stress matrix approach.
For actual runs, use the test suite with CUNEIFORM_STRESS=1 env var.

Matrix:
- Small:  5 sheets × 2,000 rows × 10 cols (~100K cells)
- Medium: 10 sheets × 5,000 rows × 15 cols (~750K cells)
- Large:  50 sheets × 10,000 rows × 20 cols (~10M cells)

Metrics captured:
- Build time (seconds)
- Peak RSS (MB)
- Workbook size (MB)
- Shared strings count (if applicable)
""")
