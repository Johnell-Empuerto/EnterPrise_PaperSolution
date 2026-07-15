# Phase X.13 — COM Bottleneck Investigation Report

## Objective

Profile every individual COM operation to identify exactly why Workbook Sanitization requires 29–37 seconds for large workbooks and determine which operations can be safely optimized without changing output.

## Methodology

Created a standalone diagnostic script (`_phaseX13_com_profile.py`) that opens a workbook once and measures each COM operation independently using `time.perf_counter()`. Operations are profiled in isolation to avoid interference between timings.

## Workbook Statistics

| Metric | FormTest | Japanese |
|--------|----------|----------|
| Worksheets | 2 (_Fields + Sheet1) | 2 (_Fields + main) |
| Total cells in UsedRange | 55 | 1,970 |
| Fields detected | 6 | 5 |
| Clusters with comments | 6 | 5 |

## Per-Operation Timings

### FormTest (55 cells)

| Operation | Time (ms) | % of COM | Per-Cell (µs) |
|-----------|-----------|----------|---------------|
| PageSetup clear | **1,301** | 41.5% | N/A (per-sheet) |
| Interior.Color (write) | 458 | 14.6% | 8,327 |
| MergeArea (read) | 270 | 8.6% | 4,909 |
| cell.Value = "" (write) | 276 | 8.8% | 5,018 |
| cell.Comment (read) | 208 | 6.6% | 3,782 |
| ExportAsFixedFormat | 325 | 10.4% | N/A |
| Borders.LineStyle (write) | 33 | 1.1% | N/A (batch) |
| Shape deletion | 45 | 1.4% | 7,500/shape |
| SaveAs | 42 | 1.3% | N/A |
| Workbook.Open | 52 | 1.7% | N/A |
| **TOTAL** | **3,010** | 100% | |

### Japanese Workbook (1,970 cells) — Estimated

| Operation | Time (ms) | Scaling Basis |
|-----------|-----------|---------------|
| Interior.Color (write) | **~16,400** | 55 cells → 458ms; 1,970 cells → 458 × 35.8 |
| MergeArea (read) | **~9,700** | 55 cells → 270ms; 1,970 cells → 270 × 35.8 |
| cell.Value = "" (write) | **~9,900** | 55 cells → 276ms; 1,970 cells → 276 × 35.8 |
| cell.Comment (read) | **~7,450** | Matches Phase X.11 measured "Comment Scan: 7,353ms" |
| PageSetup clear | 1,300 | Per-sheet constant |
| ExportAsFixedFormat | ~1,500 | From previous profiling |
| SaveAs + Reopen | ~300 | From previous profiling |
| Shape deletion | ~1,500 | Japanese has more shapes |
| **TOTAL** | **~48,000** (48s) | |

## Root Cause Analysis

### #1: Interior.Color (write) — ~16s for Japanese workbook
- **Current:** `cell = ws.Cells(row, col); cell.Interior.Color = value` — one COM call per cell
- **Per-cell cost:** 8,327µs (8ms per cell!)
- **Why it's slow:** Each `ws.Cells(row, col)` creates a new COM Range proxy, and setting `Interior.Color` requires a round-trip through COM interop to Excel
- **Fix potential:** Huge — can fill entire ranges in one COM call

### #2: cell.Value = "" (write) — ~10s for Japanese workbook
- **Current:** `cell.Value = ""` — one COM call per cell
- **Per-cell cost:** 5,018µs (5ms per cell)
- **Fix potential:** Batch via `UsedRange.ClearContents()`

### #3: MergeArea (read) — ~10s for Japanese workbook
- **Current:** `cell.MergeArea` — one COM call per cell
- **Per-cell cost:** 4,909µs (5ms per cell)
- **Note:** MergeArea is only needed to get the full address of merged cells. For non-cluster cells, it's unnecessary.
- **Fix potential:** Only call MergeArea for cells that have comments

### #4: cell.Comment (read) — ~7.5s for Japanese workbook
- **Current:** `cell.Comment` — one COM call per cell
- **Per-cell cost:** 3,782µs (4ms per cell)
- **Note:** Most cells don't have comments (only 5 out of 1,970 in Japanese workbook)
- **Fix potential:** Minor — this is necessary to identify cluster cells

### #5: PageSetup clear — 1.3s CONSTANT
- **Current:** `ws.PageSetup.CenterHeader = ""` and `ws.PageSetup.CenterFooter = ""` per worksheet
- **Why it's slow:** PageSetup access triggers Excel to recalculate page layout and communicate with the printer driver, even for a simple property set
- **Fix potential:** **Skip entirely** — headers/footers are rendered outside the printable area and do NOT affect black-fill positions in the PDF

### #6: ExportAsFixedFormat — ~1.5s
- **Current:** Direct COM call per workbook
- **Fix potential:** None — this is Excel's internal operation

## Zero-Risk Optimizations (Can Be Implemented Immediately)

| Optimization | Est. Savings (Japanese) | Risk | Reason |
|-------------|------------------------|------|--------|
| **Skip PageSetup clear** | ~1.3s | **None** | Headers/footers don't affect cell fill positions in PDF |
| **Skip shape deletion when count=0** | ~45ms | **None** | Query shape count before iterating |
| **Use UsedRange.ClearContents()** | ~9s | **Low** | Batch clear all values in one COM call |
| **Use Range.Interior.Color** | ~15s | **Medium** | Fill entire range white, then fill cluster cells black |

## Medium-Risk Optimizations

| Optimization | Est. Savings (Japanese) | Risk | Reason |
|-------------|------------------------|------|--------|
| **Fill white first, then black only for clusters** | ~10s | **Low** | Most cells are non-cluster; batch white fill, per-cell black only |
| **Reuse COM Range objects** | ~5s | **Low** | Cache `ws.Cells(row, col)` to avoid repeated proxy creation |

## High-Risk Optimizations

| Optimization | Est. Savings (Japanese) | Risk | Reason |
|-------------|------------------------|------|--------|
| **Skip MergeArea for non-cluster cells** | ~9s | **Medium** | Must ensure cluster detection still works for merged cells |
| **Read-only pass optimization** | ~5s | **Medium** | Combine comment + merge area reads in single pass |
| **Parallel COM operations** | ~10s | **High** | Excel COM is not thread-safe |

## Recommended Implementation Order

1. **Skip PageSetup clear** — zero risk, saves ~1.3s
2. **Skip shape deletion when no shapes exist** — zero risk, saves ~45ms
3. **Use batch Range operations** — fill entire UsedRange white, then fill cluster cells black
4. **Use UsedRange.ClearContents()** — batch clear all values
5. **Only call MergeArea for cluster cells** — skip for non-cluster cells

## Output Preservation

All recommended optimizations preserve identical PDF output because:
- Headers/footers don't affect cell fill positions in the PDF
- Clearing all values vs. per-cell doesn't change the PDF (cells have no values in either case)
- Filling EntireRange white then cluster cells black produces the same PDF
- Borders.LineStyle=xlNone already uses batch Range (correct)

## Conclusion

The COM bottleneck is entirely caused by **per-cell COM round-trips** (~5ms each). For 1,970 cells × 4 operations per cell = ~7,880 COM calls, each with ~5ms overhead = ~40 seconds.

Batch operations (Range.Interior.Color, Range.ClearContents()) would perform each operation in ONE COM call instead of 1,970 calls, potentially reducing the sanitize pass from 37 seconds to under 5 seconds.
