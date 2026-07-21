# Field Coordinate Architecture Analysis

**Date:** 2026-07-21
**Scope:** Excel → PNG background + interactive KeyboardText field overlay pipeline
**Bug:** Newly added fields rendered in wrong position / joined with neighboring fields (p1f7 $C$12:$D$12 overlaps with p1f6 $A$12)

---

## 1. Current Architecture

The upload preview pipeline (`render_service/upload_coordinate_generator.py`) uses **two coordinate sources**:

### Phase A — COM Session (while Excel is open)

| Step | Action | COM available? |
|------|--------|---------------|
| 1 | Open workbook | ✅ Yes |
| 2 | Export original PDF (background PNG) | ✅ Yes |
| 3a | Read comments → `cluster_meta[]` | ✅ Yes |
| 3a.5 | Read column widths → `column_widths_by_sheet` | ✅ Yes |
| 3b | Sanitize workbook (fill black/white, clear values) | ✅ Yes |
| 4 | Delete metadata sheets, export sanitized PDF | ✅ Yes |
| 5 | **Quit Excel** → COM proxies become stale | ❌ No |

### Phase B — Non-COM (after Excel.Quit)

| Step | Action | Source |
|------|--------|--------|
| 6 | Render sanitized PDF to image (300 DPI) | PyMuPDF (`fitz`) |
| 7 | **Pixel scan** (`scan_black_rectangles`) | Image processing |
| 8 | **Split merged rects** (`split_merged_rects`) | Grid/column-width proportions |
| 9 | Normalize pixels to ratios (`normalize_rects`) | Ratio = pixel / pageDimension |
| 10 | Merge metadata + ratios → **final coordinates** | ← **THIS IS THE BUG** |
| 11 | Render original PDF → background PNG | PyMuPDF (`fitz`) |

### Critical Finding

**Phase A reads COM Range geometry but NEVER uses it for coordinates.** The COM values (`range.Left`, `range.Top`, `range.Width`, `range.Height`, `printArea.Left`, etc.) are logged via FIELD DEBUG diagnostics but are **not stored or used in the final coordinate calculation**.

Phase B produces the final `left_ratio`/`top_ratio`/`right_ratio`/`bottom_ratio` **entirely from the pixel scan of the sanitized PDF**.

---

## 2. Root Cause

### 2a. Pixel Scan Merges Adjacent Fields

The `scan_black_rectangles()` function detects black rectangles on a white background. In the sanitized workbook:

- Cluster cells are filled **black**
- All other cells are filled **white**

When two black cluster cells are **vertically or horizontally adjacent with zero gap** (no white pixels between them), the scan merges them into **one rectangle** instead of two.

For the reported bug:

| Row 12 cells | Fill color | Column |
|-------------|-----------|--------|
| A12 | **Black** (cluster) | A (48pt) |
| B12 | White (no cluster) | B |
| C12:D12 | **Black** (merge) | C-D (96pt) |

When the pixel scan processes row 12, it finds A12's black rectangle and expands right until white. Column B is white (no cluster), so the scan should stop at A12's right edge. C12:D12 is detected as a separate rectangle starting at column C. **If column B has any width at all, the scan should produce 2 separate rectangles.**

**But the FIELD DEBUG log proves the rectangles WERE merged**, because the final ratios match the `split_merged_rects` grid calculation:

```
A12 + C12:D12 merged rect:  857.0 → 1063.0 = 206px
Grid split (4 columns):     
  A12 (col 1):      0% →  25% = 857.0 →  908.5  ← matches p1f6
  B (no cluster):   25% →  50% = gap (unassigned)
  C12:D12 (col 3-4): 50% → 100% = 960.0 → 1063.0  ← matches p1f7
```

The merged rect produces **51.5px per column** at a ~0.2575x scale factor, while the same columns in rows 1-2 render at **~1.0425x** (417px per A:B merge). **Same columns, different effective scale per row group.** This is the core symptom of an unreliable pixel-scan coordinate system.

### 2b. Column Width Map Fix is Insufficient

For the reported workbook, all four columns (A-D) are **equal width** (48pt each via COM Range.Width). Because they're equal, both the grid proportion and the column-width proportion produce **identical split ratios**:

| Method | A12 | C12:D12 |
|--------|-----|---------|
| Grid (4 equal cols) | 1/4 = 0.25 | 2/4 = 0.50 |
| Column widths (48×4) | 48/192 = 0.25 | 96/192 = 0.50 |

The `column_width_map` fix helps when columns have non-uniform widths (e.g., the Japanese form with cols 9=106.5pt, 10=106.5pt, 11=27.2pt, 12=74.7pt, 13=106.5pt), but offers **no improvement for equal-width columns**.

### 2c. Inconsistent Column Positions Across Rows

The same column C appears at **different horizontal positions** depending on the row — this is geometrically impossible in a well-formed Excel worksheet:

| Field | Range | left_ratio | Pixel left | Scale factor |
|-------|-------|-----------|-----------|-------------|
| p1f2 | $C$1:$D$2 | **0.4996078** | 1274.0 | ~1.0425x |
| p1f7 | $C$12:$D$12 | **0.3764706** | 960.0 | ~0.2575x |

Column C's worksheet-absolute position should be **identical regardless of row**. The discrepancy (1274px vs 960px) means the pixel-scan coordinate system is **not physically consistent**.

### 2d. Why the Scale Differs Per Row Group

The sanitized workbook export uses **`IgnorePrintAreas=False`**. When the original workbook has cell values that extend beyond the print area's used range, and `ClearContents()` removes those values, the sanitized workbook's **effective page layout can change** (different page breaks, different content boundaries). Excel may apply **different scaling or centering** to the sanitized workbook's print area, causing the same columns to render at different physical sizes on different rows.

---

## 3. The Fix: COM-Derived Coordinate Pipeline

### Architecture Change

```
BEFORE (BROKEN):
  Excel COM → comment scan → cluster_meta (cellAddr only, no geometry)
  Sanitize → PDF → pixel scan → split → normalize → FINAL RATIOS ❌

AFTER (FIXED):
  Excel COM → comment scan → cluster_meta + 
                ↓
  Excel COM Range (range.Left, Top, Width, Height)
  Excel COM PrintArea (printArea.Left, Top, Width, Height)
                ↓
  Compute: relativePt = range.Pt - printArea.Pt
  Convert: pixel = relativePt * (DPI / 72)
  Normalize: ratio = pixel / pageDimension
                ↓
                 → FINAL RATIOS ✅
```

### Detailed Algorithm

For each field during the COM phase (before Excel.Quit):

```
1. Resolve the range:
   if range.MergeCells == true:
       fieldRange = range.MergeArea
   else:
       fieldRange = range

2. Read COM geometry (in points):
   fieldLeftPt  = fieldRange.Left
   fieldTopPt   = fieldRange.Top
   fieldWidthPt = fieldRange.Width
   fieldHeightPt= fieldRange.Height

3. Read PrintArea geometry (in points):
   printAreaLeftPt  = printAreaRange.Left
   printAreaTopPt   = printAreaRange.Top
   printAreaWidthPt = printAreaRange.Width
   printAreaHeightPt= printAreaRange.Height

4. Compute print-area-relative positions:
   relativeLeftPt  = fieldLeftPt - printAreaLeftPt
   relativeTopPt   = fieldTopPt - printAreaTopPt
   relativeWidthPt = fieldWidthPt
   relativeHeightPt= fieldHeightPt

5. Convert to PNG pixels at render DPI:
   pixelsPerPoint = DPI / 72  (= 300/72 = 4.1667)
   pixelLeft   = relativeLeftPt   * pixelsPerPoint
   pixelTop    = relativeTopPt    * pixelsPerPoint
   pixelWidth  = relativeWidthPt  * pixelsPerPoint
   pixelHeight = relativeHeightPt * pixelsPerPoint

6. Convert to normalized ratios:
   left_ratio   = pixelLeft   / pngWidth
   top_ratio    = pixelTop    / pngHeight
   right_ratio  = (pixelLeft + pixelWidth)  / pngWidth
   bottom_ratio = (pixelTop  + pixelHeight) / pngHeight
```

### Key Invariant

After this fix, the same column C will have the **same left_ratio** regardless of row:

| Field | Range | Column C left | Column D right |
|-------|-------|--------------|---------------|
| p1f2  | $C$1:$D$2 | **same X** | **same X** |
| p1f7  | $C$12:$D$12 | **same X** | **same X** |

Only `top_ratio` and `bottom_ratio` should differ between row ranges.

### What NOT to Change

- PrintArea capture logic
- PDF export (both original and sanitized)
- PNG rendering (background layer)
- Excel styling behavior
- PaperLessConfig persistence
- Field configuration persistence
- Font/alignment/numeric restriction persistence

### What to Remove

- `scan_black_rectangles()` — no longer needed for coordinate production (keep for diagnostics only)
- `split_merged_rects()` — no longer needed
- `normalize_rects()` — replaced by COM-derived computation
- Column width reading (`column_widths_by_sheet`) — no longer needed for split
- Sanitization pipeline (MakeCluster) — still needed for PDF export, but coordinates come from COM

---

## 4. Implementation Plan

### Step 1 — Store COM geometry during COM phase

In `generate_coordinates_and_preview()`, while iterating comments, compute and store:

```python
field_geometry = {
    "cellAddr": addr,
    "sheetName": sheet_name,
    "name": ...,
    "type": ...,
    # COM-derived geometry
    "range_left_pt": float(merge_area.Left),
    "range_top_pt": float(merge_area.Top),
    "range_width_pt": float(merge_area.Width),
    "range_height_pt": float(merge_area.Height),
    "print_area_left_pt": float(pa_range.Left),
    "print_area_top_pt": float(pa_range.Top),
    "print_area_width_pt": float(pa_range.Width),
    "print_area_height_pt": float(pa_range.Height),
}
```

This replaces `cluster_meta` as the source of field metadata for coordinates.

### Step 2 — Compute ratios from COM geometry in non-COM phase

In the non-COM phase (after Excel.Quit), for each sheet:

```python
ppp = DPI / 72.0  # pixels per point

for meta in sheet_fields:
    rl_pt = meta["range_left_pt"] - meta["print_area_left_pt"]
    rt_pt = meta["range_top_pt"] - meta["print_area_top_pt"]
    rw_pt = meta["range_width_pt"]
    rh_pt = meta["range_height_pt"]
    
    pixel_left   = rl_pt * ppp
    pixel_top    = rt_pt * ppp
    pixel_right  = (rl_pt + rw_pt) * ppp
    pixel_bottom = (rt_pt + rh_pt) * ppp
    
    fields.append({
        "name": meta["name"],
        "type": meta["type"],
        "cellAddr": meta["cellAddr"],
        "left_ratio": pixel_left / page_w,
        "top_ratio": pixel_top / page_h,
        "right_ratio": pixel_right / page_w,
        "bottom_ratio": pixel_bottom / page_h,
    })
```

### Step 3 — Keep pixel scan for diagnostics only

Rename `scan_black_rectangles` usage to be behind a `DIAGNOSTIC_PIXEL_SCAN` flag. When enabled, it computes pixel-scan coordinates alongside COM-derived coordinates and logs both for comparison, but **never uses pixel-scan coordinates as the final output**.

### Step 4 — Backward compatibility for `generate_coordinates()`

The standalone `generate_coordinates()` function (used by legacy `/upload/coordinates` endpoint) doesn't have access to COM geometry. It will continue using the pixel-scan pipeline until it's also migrated.

---

## 5. Verification

After implementing the fix:

1. **Frontend build** — 0 errors
2. **Backend build** — 0 errors  
3. **Upload original workbook** — verify p1f6 ($A$12) and p1f7 ($C$12:$D$12) positions
4. **Export** — verify exported workbook opens correctly
5. **Upload exported workbook** — verify positions match
6. **Add new field**, merge new range, export, upload — verify all fields
7. **Verify invariant**: `p1f2.left_ratio === p1f7.left_ratio` (same column C)

### Expected Result

After fix, the FIELD GEOMETRY SOURCE DEBUG log should show:

```
Field: p1f7
Cell: $C$12:$D$12
MergeCells: True
MergeArea: $C$12:$D$12

COM:
  Left: 96.00
  Top: 158.40
  Width: 96.00
  Height: 14.40

PrintArea:
  Left: 0.00
  Top: 0.00
  Width: 192.00
  Height: 172.80

COM-derived:
  relativeLeftPt: 96.00
  relativeTopPt: 158.40
  relativeWidthPt: 96.00
  relativeHeightPt: 14.40

DPI: 300
pixelsPerPoint: 4.1667

COM-derived pixels:
  left: 400.00
  top: 660.00
  width: 400.00
  height: 60.00

COM-derived ratios:
  left: 0.15686    ← matches p1f2's C column position
  top: 0.20000
  right: 0.31373   ← matches p1f2's D column position
  bottom: 0.21818
```

**p1f2** ($C$1:$D$2) must have identical `left` and `right` ratios as **p1f7** ($C$12:$D$12).

---

## 6. File Changes Summary

| File | Change |
|------|--------|
| `render_service/upload_coordinate_generator.py` | Add COM geometry storage during comment scan; add COM-derived ratio computation in non-COM phase; keep pixel scan behind feature flag for diagnostics |
| `render_service/app.py` | No changes needed (already calls `generate_coordinates_and_preview`) |
| `paperless/app/page.tsx` | FIELD OVERLAY DEBUG already added — no changes needed |
| `render_service/background_renderer.py` | No changes needed |
| `render_service/pdf_converter.py` | No changes needed |
| Other files | No changes needed |
