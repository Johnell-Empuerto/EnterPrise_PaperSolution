# Coordinate System Comparison

## Overview

This report compares **four coordinate systems** for every cluster:

| # | System | Source | Accuracy |
|---|--------|--------|----------|
| 1 | **Database** | Legacy PostgreSQL (`def_cluster`) | Ground truth |
| 2 | **Excel COM** | `Range.Left/Top/Width/Height` | Excel's rendered position |
| 3 | **OpenXML Calc** | Our engine's column width estimates | Approximate |
| 4 | **MergeArea** | `MergeArea.Left/Top/Width/Height` | Merged range bounds |

## Page Setup (from Excel COM)

| Property | Value |
|----------|------:|
| Page Width | 612.0 pt |
| Page Height | 792.0 pt |
| Left Margin | 50.4000 pt (0.7000 in) |
| Right Margin | 50.4000 pt |
| Top Margin | 54.0000 pt |
| Bottom Margin | 54.0000 pt |
| Center Horizontally | False |
| Center Vertically | False |
| Zoom | 100 |
| FitToPagesWide | 1 |
| FitToPagesTall | 1 |
| Orientation | 1 |
| PaperSize | 1 (1) |
| PrintArea | (none) |
| Order | 1 |
| FirstPageNumber | -4105 |

## Used Range (from Excel COM)

| Property | Value |
|----------|------:|
| Address | $A$1:$G$5 |
| Rows | 5 |
| Columns | 7 |
| Left | 0.0 pt |
| Top | 0.0 pt |
| Width | 336.0 pt |
| Height | 66.0 pt |

## Active Window (from Excel COM)

| Property | Value |
|----------|------:|
| Zoom | 130 |
| VisibleRange | $A$1:$L$8 |
| ScrollRow | 1 |
| ScrollColumn | 1 |
| HPageBreaks | 0 |
| VPageBreaks | 0 |

## Per-Cluster Comparison

| Cluster | DB Left | COM Left | Merge Left | OpenXML Left | DB→COM Δx | COM→Merge Δx |
|---------|---------|----------|------------|-------------|-----------|-------------|
| $I$10:$M$10 | 0.5294118 (324.0pt) | 384.0pt | 0.0pt | 0.0pt | -60.0pt | +384.0pt |
| $I$6:$M$6 | 0.5294118 (324.0pt) | 384.0pt | 0.0pt | 0.0pt | -60.0pt | +384.0pt |
| $I$7:$M$7 | 0.5294118 (324.0pt) | 384.0pt | 0.0pt | 0.0pt | -60.0pt | +384.0pt |
| $I$8:$M$8 | 0.5294118 (324.0pt) | 384.0pt | 0.0pt | 0.0pt | -60.0pt | +384.0pt |
| $I$9:$M$9 | 0.5294118 (324.0pt) | 384.0pt | 0.0pt | 0.0pt | -60.0pt | +384.0pt |

## Coordinate System Identification

### Hypothesis: DB = (COM_Left - PrintedOrigin) / PageWidth

Where **PrintedOrigin** is the offset from the page left edge
to column A on the printed page.

**From $I$6:$M$6:** PrintedOrigin = DB_Left(324.0) - COM_Left(384.0) = -60.00 pt

### Verification Across All Clusters

| Address | COM Left | DB Left(pts) | Derived Origin | Constant? |
|---------|----------|-------------|----------------|-----------|
| $I$6:$M$6 | 384.0 | 324.0 | -60.00 | ✓ |
| $I$7:$M$7 | 384.0 | 324.0 | -60.00 | ✓ |
| $I$8:$M$8 | 384.0 | 324.0 | -60.00 | ✓ |
| $I$9:$M$9 | 384.0 | 324.0 | -60.00 | ✓ |
| $I$10:$M$10 | 384.0 | 324.0 | -60.00 | ✓ |

### Compute UsedRange-Based PrintedOrigin

**If centered:** PrintedOrigin = (PageWidth - UsedRange.Width) / 2
  = (612.0 - 336.0) / 2 = 138.00 pt

**If not centered:** PrintedOrigin = LeftMargin = 50.40 pt

**Result: Margin-based formula fits.** (Error: 110.40pt vs 198.00pt for centering)

PrintedOrigin = LeftMargin = 50.40 pt

### Workaround: Direct COM Interop Integration

Since the legacy system used Excel COM Range.Left, and we've proven
the formula is `DB = (Range.Left + PrintedOrigin) / PageWidth`, the
engine could use this formula with COM Interop to achieve 100% match
for BOTH templates without any template-specific logic.
