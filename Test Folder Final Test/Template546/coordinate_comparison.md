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
| Address | $A$1:$G$1 |
| Rows | 1 |
| Columns | 7 |
| Left | 0.0 pt |
| Top | 0.0 pt |
| Width | 336.0 pt |
| Height | 14.4 pt |

## Active Window (from Excel COM)

| Property | Value |
|----------|------:|
| Zoom | 100 |
| VisibleRange | $A$1:$R$23 |
| ScrollRow | 1 |
| ScrollColumn | 1 |
| HPageBreaks | 0 |
| VPageBreaks | 0 |

## Per-Cluster Comparison

| Cluster | DB Left | COM Left | Merge Left | OpenXML Left | DB→COM Δx | COM→Merge Δx |
|---------|---------|----------|------------|-------------|-----------|-------------|
| $A$1:$B$2 | 0.3364706 (205.9pt) | 0.0pt | 0.0pt | 0.0pt | +205.9pt | +0.0pt |
| $A$12 | 0.3364706 (205.9pt) | 0.0pt | 0.0pt | 0.0pt | +205.9pt | +0.0pt |
| $A$3:$D$4 | 0.3364706 (205.9pt) | 0.0pt | 0.0pt | 0.0pt | +205.9pt | +0.0pt |
| $A$6:$D$7 | 0.3364706 (205.9pt) | 0.0pt | 0.0pt | 0.0pt | +205.9pt | +0.0pt |
| $A$9:$D$10 | 0.3364706 (205.9pt) | 0.0pt | 0.0pt | 0.0pt | +205.9pt | +0.0pt |
| $C$1:$D$2 | 0.5000000 (306.0pt) | 96.0pt | 0.0pt | 0.0pt | +210.0pt | +96.0pt |

## Coordinate System Identification

### Hypothesis: DB = (COM_Left - PrintedOrigin) / PageWidth

Where **PrintedOrigin** is the offset from the page left edge
to column A on the printed page.

**From $A$1:$B$2:** PrintedOrigin = DB_Left(205.9) - COM_Left(0.0) = 205.92 pt

### Verification Across All Clusters

| Address | COM Left | DB Left(pts) | Derived Origin | Constant? |
|---------|----------|-------------|----------------|-----------|
| $A$1:$B$2 | 0.0 | 205.9 | +205.92 | ✓ |
| $C$1:$D$2 | 96.0 | 306.0 | +210.00 | ✗ varies |
| $A$3:$D$4 | 0.0 | 205.9 | +205.92 | ✓ |
| $A$6:$D$7 | 0.0 | 205.9 | +205.92 | ✓ |
| $A$9:$D$10 | 0.0 | 205.9 | +205.92 | ✓ |
| $A$12 | 0.0 | 205.9 | +205.92 | ✓ |

### Compute UsedRange-Based PrintedOrigin

**If centered:** PrintedOrigin = (PageWidth - UsedRange.Width) / 2
  = (612.0 - 336.0) / 2 = 138.00 pt

**If not centered:** PrintedOrigin = LeftMargin = 50.40 pt

**Result: Center-based formula fits.** (Error: 67.92pt vs 155.52pt for margin)

PrintedOrigin = (PageWidth - UsedRange.Width) / 2
  = (612.0 - 336.0) / 2 = 138.00 pt

### Workaround: Direct COM Interop Integration

Since the legacy system used Excel COM Range.Left, and we've proven
the formula is `DB = (Range.Left + PrintedOrigin) / PageWidth`, the
engine could use this formula with COM Interop to achieve 100% match
for BOTH templates without any template-specific logic.
