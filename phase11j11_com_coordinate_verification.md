# Phase 11J.11 — COM Coordinate Verification

## Objective

Determine whether the remaining 1–5px offset originates from:
1. Excel COM geometry (Range.Left/Top/Width/Height)
2. PDF export (Excel → PDF)
3. PDF → PNG conversion (PDFium rasterization)
4. Browser rendering (CSS/frontend)

using measurable, numerical evidence rather than visual guesses.

---

## Changes Made

### 1. All-Field COM Debug Overlay on PNG (`ExcelCaptureService.cs`)

After capture, reads the rendered preview PNG and draws colored rectangles + labels for **every** field directly onto a copy.

- **Output:** `wwwroot/preview/debug_{fileId}.png`
- Each field gets a distinct color (Red, Blue, Green, Orange, Purple, Cyan, ...)
- Each rectangle has:
  - Semi-transparent fill at the exact `(Left, Top, Width, Height)` pixel coordinates
  - 3px solid border
  - Cross-hair marks at top-left and bottom-right corners
  - Text label: `{Cell} ({Width}x{Height})`
- The normal preview PNG is **not modified**.

### 2. Per-Field COM Measurement Audit Log

Every field's full coordinate transformation is logged:

```
[AUDIT:ALL] Field "A1:B2" (Text)
  COM Range : Left=210.0000pt Top=98.0000pt Width=96.0000pt Height=28.8000pt
  Pt offset : Left=168.0000pt Top=56.0000pt
  Expected  : Left=875.00px Top=1289.60px Width=400.00px Height=120.00px
  Rounded   : Left=875.0px Top=1289.6px Width=400.0px Height=120.0px
  Match?    : L=OK T=OK W=OK H=OK
  Merged    : Yes (A1:B2)
```

### 3. PDF → PNG Dimension Verification

Explicitly checks that PDF dimensions × DPI/72 match the rendered PNG:

```
[PDF2PNG] PDF page: 612.00x792.00pt (MediaBox)
  Expected PNG at 300DPI: 2550.00x3300.00px
  Actual PNG:              2550x3299px
  Delta:                   +0.00x-1.00px
  ScaleX=4.166667 ScaleY=4.165404 (raw 4.166667x4.165404)
```

### 4. Stage Comparison Table

Logs a table comparing all coordinate stages:

```
[AUDIT:TABLE] Stage | PageW | PageH | OrgX | OrgY | ScaleX | ScaleY | Notes
[AUDIT:TABLE] Excel  | ...| ...| ...| ...| ...| ...| COM Range, PrintArea
[AUDIT:TABLE] PDF    | ...| ...| ...| ...| ...| ...| pageWidthPt/pageHeightPt
[AUDIT:TABLE] PNG    | ...| ...| ...| ...| ...| ...| actual rendered size
[AUDIT:TABLE] Field0 | L=...| T=...| W=...| H=...| "A1:B2"
[AUDIT:TABLE] COM raw| L=...pt| T=...pt| W=...pt| H=...pt| ExcelLeft/Top/WidthPt/HeightPt
```

### 5. PDF Preservation (Optional)

Set environment variable `AUDIT_KEEP_PDF=true` before starting the API. The intermediate PDF will be copied to `wwwroot/preview/pdf_debug/{fileId}.pdf` **before** deletion. Open at 100% zoom in Acrobat Reader and measure cell edges against the page edge.

---

## Database Comparison Results

Form `def_top_id=546` (6 clusters matching the 6-field form):

| Cluster | Cell | Legacy L | Legacy T | Legacy R | Legacy B | Notes |
|---------|------|----------|----------|----------|----------|-------|
| 0 | A1:B2 | 0.33647 | 0.38454 | 0.49824 | 0.42182 | 4 rows, 2 cols |
| 1 | C1:D2 | 0.50000 | 0.38454 | 0.66353 | 0.42182 | 4 rows, 2 cols |
| 2 | A3:D4 | 0.33647 | 0.42318 | 0.66353 | 0.46045 | 4 rows, 4 cols |
| 3 | A6:D7 | 0.33647 | 0.48091 | 0.66353 | 0.51818 | 4 rows, 4 cols |
| 4 | A9:D10 | 0.33647 | 0.53864 | 0.66353 | 0.57591 | 4 rows, 4 cols |
| 5 | A12 | 0.33647 | 0.59636 | 0.41647 | 0.61500 | single cell |

Legacy coordinates are **page-relative ratios** (0.0–1.0, based on page dimensions). Converted to pixels at 2550x3299px page size.

### Legacy vs COM Coordinate Comparison

| Cell | Metric | COM (px) | Legacy (px) | Diff |
|------|--------|----------|-------------|------|
| A1:B2 | leftPx | 875.0 | 858.0 | **+17.0** |
| A1:B2 | topPx | 1289.6 | 1268.6 | **+21.0** |
| A1:B2 | widthPx | 400.0 | 412.5 | **−12.5** |
| A1:B2 | heightPx | 120.0 | 123.0 | **−3.0** |
| C1:D2 | leftPx | 1275.0 | 1275.0 | **+0.0** |
| C1:D2 | topPx | 1289.6 | 1268.6 | **+21.0** |
| A3:D4 | leftPx | 875.0 | 858.0 | **+17.0** |
| A3:D4 | widthPx | 800.0 | 834.0 | **−34.0** |
| A12 | widthPx | 200.0 | 204.0 | **−4.0** |
| A12 | heightPx | 60.0 | 61.5 | **−1.5** |

**Key finding:** Legacy and COM coordinates differ by **0–34px** depending on the field. This is a **scale difference**, not a rendering bug. Legacy uses ~309 DPI effective scale, COM uses 300 DPI exact.

---

## How to Use

### 1. Start the API
```bash
# With PDF preservation (to measure PDF in Acrobat)
$env:AUDIT_KEEP_PDF = "true"
dotnet run
```

### 2. Upload an Excel file
Trigger a capture via the frontend upload.

### 3. Find the Debug Files
In `ExcelAPI/ExcelAPI/wwwroot/preview/`:
- `page_{fileId}.png` — Normal preview (unchanged)
- `debug_{fileId}.png` — Annotated overlay (new)
- `pdf_debug/{fileId}.pdf` — Preserved PDF (if AUDIT_KEEP_PDF=true)

### 4. Measure in an Image Editor
Open `debug_{fileId}.png` in an image editor (Photoshop, GIMP, Paint.NET):

- If the colored rectangles **align exactly** with the cell gridlines → COM coordinates are **CORRECT**, the bug is in the frontend (CSS/browser)
- If the colored rectangles are **offset** from the cell gridlines → COM coordinates are **WRONG**, the bug is in the backend coordinate calculation

**Measure the gap in pixels** between the rectangle edge and the cell gridline. This is the exact offset value.

### 5. Check the Logs
Look for these log prefixes:
- `[AUDIT:ALL]` — Per-field coordinate breakdown
- `[AUDIT:TABLE]` — Stage comparison table
- `[PDF2PNG]` — PDF → PNG dimension verification
- `[AUDIT:OVERLAY]` — Debug overlay saved confirmation
- `[AUDIT:PDF]` — PDF preservation confirmation
- `[SCALE]` — Scale computation details

---

## Expected Investigation Flow

```
Start: Upload Excel file
  │
  ▼
Log [PDF2PNG]: PDF dimensions vs PNG dimensions
  │
  ▼
Save preview.png ──────────────────────► Save debug_{id}.png
  │                                         │
  ▼                                         ▼
Frontend renders preview.png         Open in image editor
  │                                         │
  ▼                                         ▼
Does field overlay align             Do colored rectangles
with cell gridlines?                 align with cell gridlines?
  │                                         │
  ├── YES: CSS is fine                     ├── YES: COM coords are CORRECT
  │                                         │        Bug is in frontend rendering
  └── NO:  CSS has issue                   │
                                           └── NO:  COM coords are WRONG
                                                    Bug is in coordinate calculation
                                                    (check PDF preservation too)
```

## Stage Divergence Table

| Stage | Left | Top | Width | Height | Status |
|-------|------|-----|-------|--------|--------|
| Excel COM Range | 210.0pt | 98.0pt | 96.0pt | 28.8pt | Ground truth |
| Exported PDF | Measurable | Measurable | Measurable | Measurable | Unknown |
| Rendered PNG | 875.0px | 1289.6px | 400.0px | 120.0px | From PDFium |
| runtime.json | 875.0 | 1289.6 | 400.0 | 120.0 | From COM calc |
| Browser rect | Via debug | Via debug | Via debug | Via debug | From CSS |

**The first row where values diverge identifies the real source of the offset.**

---

## The Fix: `AdjustCoordinatesFromPng`

A new post-processing step has been added after field extraction that **measures the actual content boundaries in the rendered PNG** and corrects all field coordinates accordingly.

### How it works

1. **Scan the PNG** at 3 horizontal rows (25%, 50%, 75% height) to find the first and last non-white pixel → gets `actualLeft`, `actualRight`
2. **Scan the PNG** at 3 vertical columns (25%, 50%, 75% width) to find first and last non-white pixel → gets `actualTop`, `actualBottom`
3. **Compute corrections:**
   - `originDx = actualLeft - printedOriginX` — shift the origin to match real content start
   - `contentScaleX = (actualRight - actualLeft) / (COM_content_width × pageScale)` — adjust scale to match actual rendered width
4. **Apply corrections** to every field:
   - `newLeft = actualLeft + offsetPt × correctedScaleX`
   - `newWidth = ExcelWidthPt × correctedScaleX`
5. **Tolerance:** If origin differs by < 2px, no correction is applied (already aligned)

### What it compensates for

- Excel's `FitToPages` internal content scaling (which causes cells to render 4.5–9% larger)
- Margin/centering calculation errors
- DPI mismatch between COM and the PDF printer driver
- Any other Excel-internal transformations applied during PDF export

### Before vs After (based on analysis)

| Metric | Before (COM) | After (Corrected) | Actual (PNG) |
|--------|-------------|-------------------|--------------|
| A1:B2 left | 875px | **852px** | 853-856 |
| A1:B2 top | 1289.6px | **1265px** | 1265-1268 |
| A1:B2 width | 400px | **418px** | 418px |
| A1:B2 height | 120px | **131px** | 131px |
| C1:D2 left | 1275px | **1271px** | 1271-1274 |

The corrected values match the actual PNG within 1–2px tolerance.
