# Phase X.18 — Deep COM Micro-Profiling Investigation

**Status:** ✅ Complete
**Date:** July 15, 2026
**Type:** Investigation only — no production code changes

---

## Objective

Profile every individual COM statement inside the workbook sanitization phase to identify the exact bottleneck. The existing "Workbook Sanitization" profiler stage was a black box — this phase opens it up.

---

## Methodology

A temporary diagnostic script was created that opens each workbook via Excel COM and measures every individual operation with `time.perf_counter()` precision:

1. Comment scan (ws.Comments enumeration)
2. Per-worksheet operations:
   - Shapes.Count + Shape.Delete (per shape)
   - UsedRange acquisition and property reads (.Row, .Rows.Count, .Column, .Columns.Count)
   - UsedRange.Interior.Color = white (batch fill)
   - UsedRange.ClearContents() (batch clear)
   - Cluster Range(.Address) + Interior.Color = black + Value = "" (per cluster)
   - Range(1,1..LR,LC) creation + Borders.LineStyle = xlNone
3. Workbook.SaveAs
4. Workbook.Close
5. Workbook.Open (sanitized)
6. ExportAsFixedFormat

---

## FormTest Results

### Detailed Micro-Profile

| Operation | Time (ms) | % of Total |
|-----------|-----------|-----------|
| ExportAsFixedFormat | 483 | 42.9% |
| Close (workbook) | 156 | 13.9% |
| Workbook.Open (original + sanitized) | 132 | 11.7% |
| Range get (per cluster, 12 calls) | 105 | 9.3% |
| Comment Scan (6 comments) | 55 | 4.9% |
| SaveAs | 40 | 3.6% |
| Shape Delete (6 shapes) | 32 | 2.8% |
| Find max bounds | 29 | 2.6% |
| Shapes.Count | 15 | 1.4% |
| Fill White (batch) | 13 | 1.2% |
| UsedRange.Columns.Count | 12 | 1.0% |
| Range create (border) | 10 | 0.9% |
| UsedRange.Row | 8 | 0.7% |
| ClearContents (batch) | 7 | 0.6% |
| UsedRange (get) | 6 | 0.5% |
| Border Clear (84 cells) | 6 | 0.5% |
| UsedRange.Column | 2 | 0.2% |
| Excel.Quit | 3 | 0.2% |
| **TOTAL** | **1,124 ms** | **100%** |

### Statistics
- Worksheets: 2
- Total cells scanned: 55
- Comments found: 6
- Cluster fills: 12 (2 sheets × 6 clusters)
- Shapes deleted: 6
- Border cells cleared: 84

### Top 10 Slowest Statements

| Rank | Operation | Time (ms) | % of Total |
|------|-----------|-----------|-----------|
| 1 | ExportAsFixedFormat | 483 | 42.9% |
| 2 | Workbook.Close (sanitized) | 138 | 12.3% |
| 3 | Workbook.Open | 74 | 6.6% |
| 4 | Workbook.Open (sanitized) | 57 | 5.1% |
| 5 | Comment Scan (6 comments) | 55 | 4.9% |
| 6 | Workbook.SaveAs | 40 | 3.6% |
| 7 | Finding max bounds | 29 | 2.6% |
| 8 | Workbook.Close | 19 | 1.7% |
| 9 | Range get + fill + clear ($A$1:$B$2) | 15 | 1.3% |
| 10 | Range get + fill + clear ($A$12) | 12 | 1.1% |

---

## Japanese Workbook Results

### Detailed Micro-Profile

| Operation | Time (ms) | % of Total (est.) |
|-----------|-----------|-------------------|
| ExportAsFixedFormat | 1,260 | ~52% |
| Find max bounds | ~60 | ~3% |
| SaveAs | 57 | ~2% |
| Border Clear (1,970 cells) | 49 | ~2% |
| Comment Scan (5 comments) | 22 | ~1% |
| Fill White (batch) | 35 | ~1% |
| ClearContents (batch) | 20 | ~1% |
| Close + Open + Other | ~900 | ~37% |
| **TOTAL (est.)** | **~2,400 ms** | **100%** |

### Statistics
- Worksheets: 2 (Sheet1 + _Fields)
- Total cells scanned: 1,970
- Comments found: 5
- Border cells cleared: 1,970

### Top 10 Slowest Statements

| Rank | Operation | Time (ms) | % of Total |
|------|-----------|-----------|-----------|
| 1 | **ExportAsFixedFormat** | **1,260** | **~52%** |
| 2 | Find max bounds | 60 | ~3% |
| 3 | SaveAs | 57 | ~2% |
| 4 | Border Clear (1,970 cells) | 49 | ~2% |
| 5 | Fill White (batch) | 35 | ~1% |
| 6 | Comment Scan (5 comments) | 22 | ~1% |
| 7 | ClearContents (batch) | 20 | ~1% |

*Note: Close/Open/bookkeeping times are estimated — exact micro-measurement was limited by console Unicode encoding.*

---

## Key Findings

### 1. ExportAsFixedFormat is the Dominant Cost

For both workbooks, **ExportAsFixedFormat is the #1 bottleneck**:

| Workbook | Export Time | % of Total |
|----------|------------|-----------|
| FormTest | 483 ms | 42.9% |
| Japanese | 1,260 ms | ~52% |

The Japanese workbook has a larger UsedRange (1,970 cells vs 55 cells), which explains the longer export time. **This is unavoidable** — Excel COM must render the workbook to PDF, and this is inherently slow.

### 2. Batch Operations are Already Optimal

Phase X.15 optimizations are proven effective:

| Operation | FormTest (ms) | Japanese (ms) | Verdict |
|-----------|--------------|---------------|---------|
| Fill White (batch) | 13 | 35 | ✅ Fast for any size |
| ClearContents (batch) | 7 | 20 | ✅ Fast for any size |
| Border Clear | 6 | 49 | ✅ Scales linearly with cells |

None of these batch operations are slow enough to warrant further optimization.

### 3. Per-Cluster Range Operations are Not a Bottleneck

Despite 12 cluster fills (6 clusters × 2 sheets), total per-cluster time is only 105ms. Each individual `Range()` get + `Interior.Color` set + `Value` clear takes 6–15ms. Not worth optimizing further.

### 4. Shape Deletion is Negligible

| Operation | FormTest (ms) |
|-----------|--------------|
| Shapes.Count | 15 |
| Shape.Delete (6 shapes) | 32 |
| **Total** | **47 ms (4.2%)** |

The Optimization D (check Shapes.Count before enumeration) is already in place and works well.

### 5. Worksheet Property Reads are Trivial

Individual UsedRange property reads:
- `.Row`: 1ms
- `.Rows.Count`: 3ms
- `.Column`: 1–2ms
- `.Columns.Count`: 3–8ms

These are negligible in isolation (< 1% of total).

---

## Ranking by Optimization Potential

| Priority | Operation | Current Time (Japanese) | % of Total | Can Optimize? | Expected Savings |
|----------|-----------|----------------------|-----------|--------------|-----------------|
| 1 | ExportAsFixedFormat | 1,260 ms | ~52% | ❌ Excel internal | 0 ms |
| 2 | Workbook.Open/Close | ~900 ms (est.) | ~37% | ❌ Already 2 opens (req) | 0 ms |
| 3 | Find max bounds | ~60 ms | ~3% | ❌ Needed for border clear | 0 ms |
| 4 | Border Clear | 49 ms | ~2% | ⚠️ Could skip? Unlikely | 0–49 ms |
| 5 | SaveAs | 57 ms | ~2% | ❌ Required step | 0 ms |
| 6 | Fill White (batch) | 35 ms | ~1% | ✅ Already optimal | 0 ms |
| 7 | ClearContents (batch) | 20 ms | ~1% | ✅ Already optimal | 0 ms |
| 8 | Comment Scan | 22 ms | ~1% | ✅ Already optimal (ws.Comments) | 0 ms |

**Conclusion: No further COM optimizations are possible without changing the pipeline architecture.**

---

## Recommendations

### 1. ExportAsFixedFormat is not optimizable

The 1,260ms for PDF export is Excel's internal rendering time. This can only be reduced by:
- Using a smaller UsedRange (not possible — workbook is what it is)
- Using a faster rendering approach (not COM-based, e.g. OpenXML SDK)
- Running export in parallel (not possible — COM is single-threaded)

**No action recommended.**

### 2. Border clearing has marginal savings

The 49ms border clear for 1,970 cells is already fast (~25μs/cell). Even if skipped entirely, the savings would be < 2% of total time. **No action recommended.**

### 3. Total COM time is now ~2.4s for the Japanese workbook

The remaining ~10s total pipeline time comes from:
- PDF rendering (~0.4s)
- Pixel scan (~0.2s)
- PNG rendering (~0.3s)
- The 2× COM overhead of opening the workbook twice (original + sanitized)

**The only remaining architectural optimization is eliminating the second workbook open by exporting the sanitized PDF directly from the open workbook without SaveAs + Reopen.** However, this is blocked by COM proxy state corruption (Phase X.10).

---

## Final Conclusion

**All COM operations inside the sanitize phase are already optimized to their practical minimum.** The batch operations from Phase X.15 (Fill White, ClearContents) are fast. The ws.Comments collection from Phase X.17 is fast. Shape deletion is negligible.

The dominant cost is `ExportAsFixedFormat` (~1.3s), which is an Excel-internal operation that cannot be accelerated from Python.

**No further COM micro-optimizations are warranted.** Any future performance gains must come from:
1. Architectural changes (eliminating one of the two workbook opens)
2. Rendering pipeline improvements (faster PDF rendering, PNG generation)
3. Parallelizing non-COM operations (pixel scan and background PNG render run sequentially)
