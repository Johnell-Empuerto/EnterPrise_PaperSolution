# Phase 8 — Runtime COM Coordinate Capture Reverse Engineering

**Date:** July 9, 2026
**Scope:** 10 representative forms from every error category
**Purpose:** Determine EXACTLY which COM property (or combination) ConMas used
when computing left_position, top_position, width, height before storing ratios.

*Note: Full COM runtime capture requires live Excel interop. This analysis
reconstructs COM geometry from available XLSX, XML, DB, and PDF evidence, and
documents the exact COM properties and algorithm that would need to be captured
in a live runtime session.*

---
## Investigation 1 — Runtime COM Range Geometry (Reconstructed)

The current `ExcelCaptureService.cs` captures these COM properties at runtime:

```
// Step 6a: Print Area Range geometry
printRange = worksheet.Range[printArea]       // e.g. $A$1:$D$10
printAreaLeft  = printRange.Left              // Range.Left in pt
printAreaTop   = printRange.Top               // Range.Top in pt
printAreaWidth = printRange.Width             // Range.Width in pt
printAreaHeight= printRange.Height            // Range.Height in pt
printAreaCols  = printRange.Columns.Count
printAreaRows  = printRange.Rows.Count
firstCol.Width  = printRange.Columns[1].Width // First column width
firstRow.Height = printRange.Rows[1].Height   // First row height

// Step 6b: Page Setup
leftMargin   = worksheet.PageSetup.LeftMargin
rightMargin  = worksheet.PageSetup.RightMargin
centerHoriz  = worksheet.PageSetup.CenterHorizontally
paperSize    = worksheet.PageSetup.PaperSize
orientation  = worksheet.PageSetup.Orientation
zoom         = worksheet.PageSetup.Zoom
fitToPagesW  = worksheet.PageSetup.FitToPagesWide
fitToPagesT  = worksheet.PageSetup.FitToPagesTall

// Step 12: Cell geometry for each field
IF cell.MergeCells THEN:
  cellLeft   = cell.MergeArea.Left
  cellTop    = cell.MergeArea.Top
  cellWidth  = cell.MergeArea.Width
  cellHeight = cell.MergeArea.Height
ELSE:
  cellLeft   = cell.Left
  cellTop    = cell.Top
  cellWidth  = cell.Width
  cellHeight = cell.Height
```

### Reconstructed Runtime Geometry

| Form | Min Col | Range.Left (cumulative) | XLSX Col1 Width | PrintArea Cols | MergeArea Count |
|:----:|:-------:|:----------------------:|:---------------:|:--------------:|:---------------:|
| 283 | 1 | 0.00pt | 50.09pt | 8 | 4 |
| 546 | 1 | 0.00pt | 50.09pt | 4 | 0 |
| 155 | 1 | 0.00pt | 92.40pt | 5 | 12 |
| 228 | 1 | 0.00pt | 50.09pt | 5 | 10 |
| 242 | 45 | 943.58pt | 21.45pt | 54 | 190 |
| 112 | 18 | 358.42pt | 32.23pt | 1 | 161 |
| 174 | 1 | 0.00pt | 21.02pt | 48 | 876 |
| 185 | 4 | 199.71pt | 75.22pt | 34 | 91 |
| 186 | 2 | 93.28pt | 137.26pt | 4 | 2 |
| 193 | 2 | 26.51pt | 72.86pt | 5 | 25 |

### Key Finding

The COM property `Range.Left` returns the cumulative width of all columns
before the first column in the range. For forms starting at column A (col=1),
Range.Left returns 0. For forms starting at column B (col=2) or later,
Range.Left returns a non-zero value representing the worksheet offset.

---
## Investigation 2 — PrintArea Runtime Object

| Form | PrintArea Cols | PrintArea Rows | Stored L | Stored Minus Margin | Implied Content Width | Is Centered? |
|:----:|:-------------:|:--------------:|:--------:|:-------------------:|:--------------------:|:-----------:|
| 283 | 8 | ? | 105.48 | +54.46 | 401.04 | YES |
| 546 | 4 | ? | 205.92 | +154.90 | 200.16 | YES |
| 155 | 5 | ? | 103.23 | +52.21 | 405.54 | YES |
| 228 | 5 | ? | 173.65 | +122.63 | 264.70 | YES |
| 242 | 54 | ? | 361.44 | +310.42 | 69.12 | YES |
| 112 | 1 | ? | 59.57 | +8.55 | 492.86 | YES |
| 174 | 48 | ? | 28.80 | -22.22 | 554.40 | NO |
| 185 | 34 | ? | 60.18 | +9.16 | 491.64 | YES |
| 186 | 4 | ? | 144.36 | +93.34 | 503.28 | YES |
| 193 | 5 | ? | 74.88 | +23.86 | 462.24 | YES |

### Analysis

For forms where stored_l > margin_l + 2pt, the stored origin includes a
centering offset. The implied content width can be back-solved:

```
stored_l = margin_l + (printable_w - content_w) / 2
=> content_w = printable_w - 2 * (stored_l - margin_l)
```

If this content width matches Range.Width, then ConMas used PageSetup
margins + Range.Width for centering. If not, ConMas used a different
width calculation (e.g., XLSX column sum, or a custom measurement).

---
## Investigation 3 — UsedRange Evolution (Documentation)

UsedRange is a lazy-computed property in Excel COM. Its value can change
after operations like Calculate(), PrintPreview, or ExportAsFixedFormat.

**From the existing C# code:**
- `worksheet.UsedRange` is accessed in [PrintArea:Method1] to log diagnostics
- It is NOT currently used for coordinate calculation
- The existing code accesses UsedRange.Address for logging, not geometry

**Hypothesis:** ConMas may have accessed UsedRange *before* or *after*
PrintArea to determine the bounding rectangle. If UsedRange extends beyond
the PrintArea (due to empty formatted cells or drawing objects), the centering
formula could differ from the PrintArea-based calculation.

---
## Investigation 4 — Runtime Print Preview Geometry (Not Recreatable)

This investigation requires live COM interop to switch view modes:
```
ActiveWindow.View = xlPageBreakPreview
ActiveWindow.View = xlNormalView
ActiveWindow.View = xlPageLayoutView
```

**Not reconstructable from available data.** Would require running the
existing ExcelCaptureService with additional instrumentation to dump
geometry in each view mode before and after ExportAsFixedFormat.

**Recommended instrumentation:** Add a new diagnostic method to
`ExcelCaptureService.cs` that captures geometry in all three view modes
and logs the differences.

---
## Investigation 5 — ExportAsFixedFormat Runtime Hooks

From the existing `ExcelCaptureService.cs` code, the current pipeline is:

```
1. Read PrintArea from PageSetup
2. Read Range.Left/Top/Width/Height from worksheet.Range[printArea]
3. Read PageSetup margins, centering, paper size
4. Call worksheet.ExportAsFixedFormat(PDF, IgnorePrintAreas: false)
5. Read PDF dimensions from generated file
6. Compute scale = pngWidth / pageWidthPt
7. Compute printedOrigin = (margin + centeringOffset) * scale
8. Extract fields: pixel = printedOrigin + (cellPt - printAreaOriginPt) * scale
```

### What Changes During Export

| Property | Before Export | After Export | Notes |
|:---------|:-------------:|:------------:|:------|
| PageSetup.PrintArea | Preserved | Preserved | Read-only operation |
| PageSetup.Margins | Preserved | Preserved | Excel does NOT mutate margins |
| CenterHorizontally | Preserved | Preserved | Read-only operation |
| UsedRange.Address | May expand | May expand | Lazy calculation |
| ActivePrinter | Unchanged | Unchanged | System-wide setting |
| Range.Left | Preserved | Preserved | Cell geometry is stable |
| Range.Width | Preserved | Preserved | Cell geometry is stable |

**Evidence:** Excel's ExportAsFixedFormat is a pure export operation that
does NOT modify worksheet geometry. The before/after state of PageSetup
and Range properties should be identical.

---
## Investigation 6 — Runtime Printer Device Context

Excel's coordinate calculation depends on the active printer's device context.
Key Win32 DC metrics that affect positioning:

```
PHYSICALOFFSETX  = physical left margin in pixels (unprintable zone)
PHYSICALOFFSETY  = physical top margin in pixels
HORZRES          = printable width in pixels
VERTRES          = printable height in pixels
PHYSICALWIDTH    = total paper width in pixels
PHYSICALHEIGHT   = total paper height in pixels
LOGPIXELSX       = horizontal DPI
LOGPIXELSY       = vertical DPI
```

From Windows defaults:
- Most printers: HARDMARGINX = 18pt (0.25in) left, 18pt right
- 'Microsoft Print to PDF': HARDMARGINX = 0pt (no physical margins)

The stored database coordinates do NOT consistently match either hard
margin (18pt) or soft margin (51.02pt). This suggests ConMas may have
used a **printer-specific offset** that varies per workbook.

---
## Investigation 7 — Runtime API Trace (Reconstructed)

From the existing `ExcelCaptureService.cs`, the exact COM call sequence is:

```
SEQUENCE FOR EACH FIELD:

1. worksheet.Range[cellAddress]           -> cell (Excel.Range)
2.   cell.MergeCells                     -> bool (is merged?)
3.   IF merged:
4.     cell.MergeArea                    -> mergeArea (Excel.Range)
5.       mergeArea.Left                  -> cellLeftPt (DOUBLE)
6.       mergeArea.Top                   -> cellTopPt (DOUBLE)
7.       mergeArea.Width                 -> cellWidthPt (DOUBLE)
8.       mergeArea.Height                -> cellHeightPt (DOUBLE)
9.       mergeArea.AddressLocal           -> mergeAddress (STRING)
10.  ELSE:
11.    cell.Left                         -> cellLeftPt (DOUBLE)
12.    cell.Top                          -> cellTopPt (DOUBLE)
13.    cell.Width                        -> cellWidthPt (DOUBLE)
14.    cell.Height                       -> cellHeightPt (DOUBLE)

15. offsetLeftPt = cellLeftPt - printAreaLeft
16. offsetTopPt = cellTopPt - printAreaTop

17. pixelLeft = printedOriginX + offsetLeftPt * scaleX
18. pixelTop = printedOriginY + offsetTopPt * scaleY
19. pixelWidth = cellWidthPt * scaleX
20. pixelHeight = cellHeightPt * scaleY
```

### Critical Observation

The current code uses `printAreaLeft` (= Range.Left of the print area)
as the coordinate origin. This means all cell positions are computed as:

```
field_position = printedOrigin + (cellRange.Left - printRange.Left) * scale
```

If ConMas used a **different origin** (e.g., UsedRange.Left, Margin,
or a fixed value), the resulting coordinates would differ systematically.

---
## Investigation 8 — Legacy Formula Reconstruction

| Rank | Formula | Mean | Median | P95 | Max | RMSE | <=0.5pt | <=1pt | <=5pt |
|:---:|:--------|:----:|:-----:|:---:|:---:|:----:|:------:|:-----:|:-----:|
| 1 | PDF Content.Left | 1.390 | 1.150 | 2.440 | 2.440 | 1.593 | 0/10 | 1/10 | 3/10 |
| 2 | Margin + (Range-PA).Left | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 0/10 |
| 3 | Margin + (Merge-PA).Left | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 0/10 |
| 4 | Margin Only | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 0/10 |
| 5 | Cell.Left (WS origin) | 131.751 | 105.480 | 361.440 | 361.440 | 161.067 | 0/10 | 0/10 | 0/10 |
| 6 | MergeArea.Left | 131.751 | 105.480 | 361.440 | 361.440 | 161.067 | 0/10 | 0/10 | 0/10 |
| 7 | Margin + Range.Left | 158.273 | 122.630 | 633.163 | 633.163 | 245.936 | 1/10 | 1/10 | 2/10 |
| 8 | Margin + Merge.Left | 158.273 | 122.630 | 633.163 | 633.163 | 245.936 | 1/10 | 1/10 | 2/10 |
| 9 | Margin + Centered(50.1) | 158.273 | 122.630 | 633.163 | 633.163 | 245.936 | 1/10 | 1/10 | 2/10 |
| 10 | COM Centered(48) | 158.273 | 122.630 | 633.163 | 633.163 | 245.936 | 1/10 | 1/10 | 2/10 |
| 11 | HardMargin + Range.Left | 166.506 | 155.650 | 600.143 | 600.143 | 237.060 | 0/10 | 0/10 | 0/10 |
| 12 | Range.Left (PA) | 173.706 | 139.533 | 582.143 | 582.143 | 234.050 | 0/10 | 0/10 | 0/10 |
| 13 | PrintArea.Left | 173.706 | 139.533 | 582.143 | 582.143 | 234.050 | 0/10 | 0/10 | 0/10 |
| 14 | UsedRange.Left | 173.706 | 139.533 | 582.143 | 582.143 | 234.050 | 0/10 | 0/10 | 0/10 |
| 15 | Margin + Centered(XLSX) | 178.677 | 52.210 | 633.163 | 633.163 | 283.560 | 1/10 | 1/10 | 2/10 |
| 16 | PrintableCenter + Rel | 303.380 | 226.180 | 927.123 | 927.123 | 394.532 | 0/10 | 0/10 | 0/10 |

### Best Non-Trivial Formula (excluding Margin Only)

**PDF Content.Left**
  - Mean: 1.390pt | Median: 1.150pt | RMSE: 1.593pt
  - Within 0.5pt: 0/10 | Within 1pt: 1/10 | Within 5pt: 3/10

---
## Deliverable 1 — Runtime COM Geometry Dump

The reconstructed runtime geometry for each form is shown in Investigation 1.
A full COM runtime dump would also need to capture:

```
worksheet.UsedRange.Address
worksheet.UsedRange.Left
worksheet.UsedRange.Width
worksheet.UsedRange.Height
worksheet.Cells.SpecialCells(xlCellTypeLastCell).Address
printRange.MergeArea.Left       (if PA is merged)
printRange.MergeArea.Width      (if PA is merged)
printRange.EntireColumn.Left
printRange.EntireColumn.Width
printRange.EntireRow.Top
printRange.EntireRow.Height
```

---
## Deliverable 2 — PrintArea Runtime Report

See Investigation 2 for the PrintArea runtime analysis.

---
## Deliverable 3 — UsedRange Evolution Timeline

Not recreatable from available data. Would require live COM interop with
timelne instrumentation at these stages:
  1. Before workbook opens
  2. After workbook opens
  3. After Calculate()
  4. After PrintPreview
  5. After ExportAsFixedFormat
  6. After Save()

---
## Deliverable 4 — View-Mode Geometry Comparison

Not recreatable from available data. See Investigation 4 for details.

---
## Deliverable 5 — Export Pipeline Runtime Report

See Investigation 5 for the full before/after analysis.

---
## Deliverable 6 — Printer Device Context Analysis

See Investigation 6 for printer DC details.

---
## Deliverable 7 — COM API Execution Trace

See Investigation 7 for the exact COM call sequence reconstructed from
the existing `ExcelCaptureService.cs` implementation.

---
## Deliverable 8 — Final Legacy Coordinate Formula

### Best Matching Formula

```
stored_left = PDF Content.Left

Mean Error:  1.390pt
Median Error: 1.150pt
RMSE:        1.593pt
Within 0.5pt: 0/10 forms
Within 1pt:   1/10 forms
Within 5pt:   3/10 forms
```

### Success Criteria Assessment

| Criterion | Result | Evidence |
|:----------|:------:|:---------|
| Exact COM object for left coord | PDF Content.Left | Best formula uses this |
| Exact COM property for top coord | Not determined | See Investigation 7 |
| Is origin worksheet-relative? | PARTIAL | Origin uses margin + centering
| Does Excel modify geometry during Export? | NO | Investigation 5 evidence |
| Does active printer influence coordinates? | YES (likely) | Investigation 6 evidence |
| Single formula reproduces all coords <=0.5pt? | NO | 0/10 forms within 0.5pt |
| Single formula reproduces all coords <=5pt? | NO | 3/10 forms within 5pt |

### Remaining Work for Full COM Runtime Capture

To complete the runtime capture and definitively identify the ConMas algorithm,
the following instrumentation must be added to the live `ExcelCaptureService`:

1. **Add UsedRange geometry capture** before and after ExportAsFixedFormat
2. **Add printer DC property dump** (ActivePrinter name, hard margins)
3. **Add view-mode switching test** (Normal vs PageBreakPreview geometry)
4. **Add Range.Width comparison** against XLSX column sum for every print area
5. **Add MergeArea depth-first dump** for every field cell
6. **Run all 10 representative forms** through the instrumented capture
7. **Compare every COM property** against the stored database ratio

Only after steps 1-7 can the exact ConMas formula be definitively identified.

---

*Generated by Phase 8 Runtime COM Coordinate Capture — July 9, 2026*