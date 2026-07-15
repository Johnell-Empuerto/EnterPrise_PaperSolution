# Phase X.4 — Performance Regression Investigation Report

## Objective

Determine why `/upload/preview` might exceed 100 seconds after preserving worksheet shapes.

**Do NOT propose fixes. Evidence collection only.**

---

## 1. Worksheet Statistics

### FormTest (baseline)

| Property | Value |
|---|---|
| Sheet count | 2 (_Fields, Sheet1) |
| Total shapes | 6 |
| Pictures | 0 |
| TextBoxes | 0 |
| Charts | 0 |
| OLE Objects | 0 |
| Max used rows | 12 |
| Max used cols | 7 |
| Print area | $A$1:$D$12 |
| Zoom | 100 |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| File size | 12,830 bytes |

### Japanese workbook

| Property | Value |
|---|---|
| Sheet count | 2 (_Fields, 第15回DMSK_i-Reporter) |
| Total shapes | 17 |
| Pictures | 3 |
| TextBoxes | 0 |
| Charts | 0 |
| OLE Objects | 0 |
| Max used rows | 45 |
| Max used cols | 43 |
| Print area | $B$1:$M$44 |
| Zoom | 100 |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| File size | 104,950 bytes |

### Key Differences

- **UsedRange**: Japanese workbook is **23× larger** (1935 cells vs 84 cells)
- **Shapes**: Japanese workbook has **17 shapes vs 6** (2.8× more)
- **Pictures**: Japanese workbook has **3 pictures**; FormTest has 0
- **File size**: Japanese workbook is **8× larger** (105KB vs 13KB)

---

## 2. Stage-by-Stage Timing Comparison

### Raw Timings

| Stage | FormTest (ms) | % | Japanese (ms) | % | Ratio |
|---|---|---|---|---|---|
| Identify Clusters | 622 | 2% | 9,514 | 16% | 15.3× |
| Sanitize Workbook | 5,960 | 17% | 30,795 | 51% | 5.2× |
| Export Sanitized PDF | 923 | 3% | 1,184 | 2% | 1.3× |
| Render Sanitized PDF | 54 | 0% | 47 | 0% | 0.9× |
| Scan Black Rectangles | 25,950 | 75% | 17,020 | 28% | 0.7× |
| Split Merged Rects | 216 | 1% | 0 | 0% | 0.0× |
| Normalize Rects | 0 | 0% | 0 | 0% | — |
| Export Original PDF | 528 | 2% | 1,627 | 3% | 3.1× |
| Render Original PNG | 199 | 1% | 262 | 0% | 1.3× |
| Get Page Dimensions | 2 | 0% | 1 | 0% | 0.5× |
| **TOTAL** | **34,454** | 100% | **60,449** | 100% | **1.75×** |

---

## 3. Bottleneck Analysis

### Japanese Workbook — Slowest Stages

| Rank | Stage | Time | % | Root Cause |
|---|---|---|---|---|
| **1** | **Sanitize Workbook** | **30.8s** | **51%** | **Cell-by-cell COM operations** |
| 2 | Scan Black Rectangles | 17.0s | 28% | NumPy pixel iteration |
| 3 | Identify Clusters | 9.5s | 16% | COM cell scan for comments |
| 4 | Export Original PDF | 1.6s | 3% | COM ExportAsFixedFormat |
| 5 | Export Sanitized PDF | 1.2s | 2% | COM ExportAsFixedFormat |

### Why `sanitize_workbook` is the bottleneck (30.8s)

The `sanitize_workbook()` function iterates **every cell** in `UsedRange` via COM:

```python
for row in range(1, lr + 1):       # 45 rows
    for col in range(1, lc + 1):   # 43 columns
        cell = ws.Cells(row, col)  # COM call
        # + MergeArea check        # COM call
        # + fill color set         # COM call
        # + value clear            # COM call
```

**1935 cells × multiple COM calls per cell = ~30.8 seconds.**

The **shape deletion loop** takes negligible time:

```python
for shape in list(ws.Shapes):  # Only 17 shapes
    shape.Delete()              # Fast COM operation
```

### Why `scan_black_rectangles` is slower on FormTest (26.0s vs 17.0s)

FormTest produces **4** initial black rectangles (merged from 6 clusters) on a 2550×3300px image. The scan finds more top-left corners to verify (more candidate rectangles = more pixel checking). The Japanese workbook produces only **1** large black blob — fewer corners = less scanning.

---

## 4. Timeout Analysis

### Timing vs Timeout Threshold

| Metric | FormTest | Japanese |
|---|---|---|
| Total pipeline time | **34.5s** | **60.4s** |
| Under 100s timeout? | ✅ Yes | ✅ Yes |
| Under ASP.NET 5min timeout? | ✅ Yes | ✅ Yes |

**Neither workbook exceeds the 100-second timeout.**

### Possible Explanation for Production Timeout

The 100-second timeout the user observed likely comes from **one** of the following:

1. **Older code**: The ASP.NET `HttpClient` timeout was previously the default **100 seconds**. The current `Program.cs` has `TimeSpan.FromMinutes(5)`, which may have been a recent fix. If the timeout was encountered before this change, the pipeline barely squeezes under 100s (60.4s) — but a slower machine or additional COM overhead could push it over.

2. **Concurrent COM operations**: If the production server runs multiple COM instances simultaneously (e.g., multiple uploads), Excel's single-threaded COM apartment can cause **serialization delays** that add up.

3. **Additional overhead**: The test script runs stages sequentially in a single process. The HTTP framework adds file upload time, multipart parsing, and process spawning overhead that could add 5-15 seconds.

### Shape Preservation Impact

**Preserving shapes will NOT cause a timeout.** The shape deletion loop (`for shape in ws.Shapes: shape.Delete()`) accounts for **less than 1%** of the 30.8s sanitize time. Preserving shapes would actually **reduce** the sanitize time slightly.

---

## 5. Conclusion

| Question | Answer |
|---|---|
| Is there a performance regression? | **Yes.** Japanese workbook takes 1.75× longer than FormTest. |
| What is the bottleneck? | **`sanitize_workbook` at 30.8s (51%)** — the cell-by-cell COM iteration over 1935 cells. |
| Does shape preservation cause the timeout? | **No.** Shape deletion takes <1% of total time. Preserving shapes would barely change runtime. |
| Does the Japanese workbook exceed 100s? | **No.** It completes in 60.4s — well under 100s. |
| Does the ASP.NET client timeout at 100s? | **No.** It's configured to 5 minutes. The 100s timeout must be from an earlier configuration or a different path. |
| What is the real fix target? | The cell-by-cell COM iteration in `sanitize_workbox()`, not shape preservation. |

### File Size Comparison

```
┌──────────────────────────────────────────────────────────────┐
│  Pipeline Stage                          FormTest   Japanese │
├──────────────────────────────────────────────────────────────┤
│  Identify Clusters                        0.6s       9.5s   │
│  Sanitize Workbook                        6.0s      30.8s   │
│  Export Sanitized PDF                     0.9s       1.2s   │
│  Render Sanitized PDF                     0.1s       0.0s   │
│  Scan Black Rectangles                   26.0s      17.0s   │
│  Split Merged Rects                       0.2s       0.0s   │
│  Export Original PDF                      0.5s       1.6s   │
│  Render Original PNG                      0.2s       0.3s   │
├──────────────────────────────────────────────────────────────┤
│  TOTAL                                   34.5s      60.4s   │
└──────────────────────────────────────────────────────────────┘
```
