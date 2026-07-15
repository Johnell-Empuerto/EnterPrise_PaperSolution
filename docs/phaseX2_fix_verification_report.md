# Phase X.2 — Fix Implementation and Verification Report

**Date**: July 15, 2026  
**Status**: Fix applied; deeper root cause discovered

---

## Fix Applied

**File**: `render_service/upload_coordinate_generator.py` — function `export_sanitized_pdf()`

**Removed** (3 lines):
```python
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = False
ws.PageSetup.FitToPagesTall = False
```

**Added** (comment):
```python
# Preserve original page setup — do NOT modify Zoom or FitToPages
# The sanitized PDF must be geometrically identical to the original workbook PDF
# so that pixel-scanned coordinates align with the background image.
```

---

## What the Fix Achieved

| Before Fix | After Fix |
|-----------|-----------|
| FitToPages cleared → sanitized PDF has UNSCALED content | FitToPages preserved → sanitized PDF has SCALED content (same as original) |
| Content offset for FormTest: 4px | Same (no regression) ✅ |
| Content offset for Japanese: 1593×600px | Improved but not eliminated |

---

## The Deeper Root Cause Discovered

Removing the FitToPages clearing was NECESSARY but NOT SUFFICIENT. Here's why:

### The Problem: Content Extent Changes After Sanitization

The sanitization process (`sanitize_workbook()`) **deletes most of the workbook's content** — it clears cell values, removes headers/footers, deletes shapes, and sets all non-cluster cells to white fill.

Excel's `FitToPages=1×1` setting scales the content based on the **actual content extent** (the full range of cells that have content or formatting). When the content extent changes, the FitToPages scaling factor changes.

### Concrete Example: Japanese Workbook

| Property | Original Workbook | Sanitized Workbook |
|----------|:-----------------:|:------------------:|
| Content columns | A-AR (43 cols) | I-M (5 cols) |
| Content rows | 1-45 | 6-10 |
| Content width (approx) | 43 × 48pt = 2064pt | 5 × 48pt = 240pt |
| Printable width (Letter) | 555.3pt | 555.3pt |
| **FitToPages scale needed** | 555.3/2064 ≈ **0.27×** | Content already fits → **1×** (no scaling) |
| `CenterHorizontally` | Centers 2064pt content → offset ≈ margin | Centers 240pt content → **different** offset |

**Result**: Even with FitToPages preserved, the two workbooks have DIFFERENT content extents, leading to DIFFERENT FitToPages scaling factors and DIFFERENT centering offsets.

### The Fundamental Issue

The pixel scan approach has an **architectural incompatibility** with FitToPages:

```
Sanitize workbook (deletes content) → content extent shrinks → 
FitToPages scale factor changes → PDF geometry changes →
coordinates from sanitized PDF ≠ background from original PDF
```

This issue affects ANY workbook that:
1. Uses `Zoom=False` (meaning "use FitToPages")
2. Has FitToPagesWide or FitToPagesTall set
3. Has content outside the cluster cell area (which is MOST workbooks)

---

## Verification Results

### FormTest: ✅ Still Works

| Metric | Value | Verdict |
|--------|:-----:|:-------:|
| Image dimensions | 2550×3300 | Same ✅ |
| Content offset | (4, 4) px | Negligible ✅ |
| Fields generated | 6 | Correct count ✅ |
| Pixel diff | 5.3% | From fill color, not position ✅ |

### Japanese Workbook: ⚠️ Still Has Offset

| Metric | Value | Verdict |
|--------|:-----:|:-------:|
| Image dimensions | 2550×3300 | Same ✅ |
| Coordinate left ratio | 0.22824 | → 582px |
| COM Range expected left | 414.6pt | → ~1727px |
| **X difference** | **28 px** | ⚠️ Small but present |
| **Y difference** | **468 px** | ❌ Large vertical offset |
| Fields generated | 5 | Correct count ✅ |

---

## Comparison: Both Approaches

### Pixel Scan Pipeline (Current Python, after fix)
```
Original workbook → Sanitize (delete content) → Export PDF → Render PNG → 
scan_black_rectangles → normalize_rects → ratio

Original workbook → xlsx_to_pdf (keep content) → Render PNG → background image
                                                              ↓
                                        ratio × bgDimensions = leftPx
                                        ⚠️ May NOT match if content extents differ
```

### COM Range Pipeline (C# ExcelCaptureService)
```
Original workbook → COM Range.Left/Top/Width/Height → 
(printedOriginPt + cellPt - printAreaPt) × 300/72 = leftPx

Original workbook → ExportAsFixedFormat → PDF → PNG → background image
                                                              ↓
                                        leftPx = direct pixel position
                                        ✅ Always matches because COM measures
                                           the actual cell geometry
```

---

## The Only Complete Fixes

### Option A: Abandon Pixel Scan, Use COM Range (recommended)
Replace `generate_coordinates()` with a `MeasureFieldsFromCom()` equivalent that reads `Range.Left/Top/Width/Height` via win32com, computes the printed origin from margins + centering, and converts to pixels. This matches the C# production pipeline exactly.

**Pros**: Always correct, no page setup dependency, matches C# behavior  
**Cons**: Requires COM (same dependency already exists for cluster identification)

### Option B: Preserve Full Content Extent During Sanitization
Instead of deleting all non-cluster cells, add invisible content (e.g., transparent fills on cells outside the cluster range) to preserve the original content extent. This keeps FitToPages scaling consistent between the two PDFs.

**Pros**: Keeps pixel scan approach, minimal code change  
**Cons**: Fragile, hard to get right for all workbook variations

### Option C: Compute Coordinates Directly from PDF Vector Geometry
Instead of pixel scanning, extract cell positions from the PDF's vector content (drawing commands like `re` + `f`).

**Pros**: No COM dependency, uses actual PDF geometry  
**Cons**: Complex, PDF engine-specific, may not work for all PDF generators

---

## Recommendation

**Option A** is the correct long-term fix. The COM Range approach is already proven in the C# production pipeline (`ExcelCaptureService.MeasureFieldsFromCom`), works for ALL workbooks regardless of page setup, and doesn't have the content-extent dependency problem.

The three-line fix applied in this phase was a necessary incremental improvement — it removes the most obvious divergence source. But the pixel scan pipeline has a deeper architectural limitation that this fix alone cannot resolve.
