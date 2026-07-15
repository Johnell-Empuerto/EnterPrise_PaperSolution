# Phase 11J.12 — PNG Content Boundary Auto-Correction

## Problem

The COM-derived overlay coordinates did not match the actual rendered PNG. Analysis of `debug_{id}.png` showed offsets of **~21px horizontal** and **~25px vertical** between the colored overlay rectangles and the actual cell gridlines.

### Root Cause

Excel's PDF export applies **internal content scaling** that COM Range.Width/Height does not report. The transformation pipeline is:

```
COM Range.Left/Top/Width/Height  (worksheet points)
  ↓  × (dpi/72) — page-level scale
Expected pixel coordinates
  ↓  × contentScale — applied by Excel during PDF export (FitToPages, printer DPI, etc.)
Actual rendered pixel position  ← WHAT THE PNG ACTUALLY SHOWS
```

The `contentScale` factor varied by ~4.5% horizontally and ~9.2% vertically for the test form, likely due to Excel's `FitToPagesWide=1, FitToPagesTall=1` settings and printer driver DPI.

### Measured Offsets (Field A1:B2)

| Metric | COM Expected | PNG Actual | Δ |
|--------|-------------|-----------|----|
| Left edge | 875px | 854px | **+21px** |
| Top edge | 1289.6px | 1265px | **+24.6px** |
| Width | 400px | 418px | **−18px** |
| Height | 120px | 131px | **−11px** |

---

## Solution: `AdjustCoordinatesFromPng()`

Added a post-processing step in `ExcelCaptureService.CapturePrintAreaAsync()` that:

1. **Scans the rendered PNG** to measure actual content boundaries
2. **Computes origin and scale corrections**
3. **Applies corrections** to all field coordinates in-place

### File Changed

`ExcelAPI/ExcelAPI/Services/ExcelCaptureService.cs`

### Method Signature

```csharp
private static bool AdjustCoordinatesFromPng(
    string pngPath,
    List<ExcelField> fields,
    double scaleX,
    double scaleY,
    ref double correctedOriginX,
    ref double correctedOriginY,
    ref double correctedScaleX,
    ref double correctedScaleY)
```

### Algorithm

#### Step 1: Horizontal Boundary Detection

Scan **3 horizontal rows** at 25%, 50%, and 75% of page height:

```
For each row:
  Scan left→right for first run of ≥3 consecutive "non-white" pixels
    → this is the actual content left edge
  Scan right→left for last run of ≥3 consecutive "non-white" pixels
    → this is the actual content right edge
```

**"Non-white" threshold:** R > 248 AND G > 248 AND B > 248 → white (background). Anything else is content.

**3-consecutive-pixel rule:** avoids false positives from single noise pixels or anti-aliased edges.

#### Step 2: Vertical Boundary Detection

Scan **3 vertical columns** at 25%, 50%, and 75% of page width:

```
For each column:
  Scan top→bottom for first run of ≥3 consecutive "non-white" pixels
    → actual content top edge
  Scan bottom→top for last run of ≥3 consecutive "non-white" pixels
    → actual content bottom edge
```

#### Step 3: Compute Corrections

```csharp
// Origin correction = where content actually starts vs where COM says it starts
originDx = avgLeft - correctedOriginX;  // shift overlay LEFT by this amount
originDy = avgTop - correctedOriginY;    // shift overlay UP by this amount

// Expected content extent from COM field data
// (not from printAreaWidth, which may be full page width)
comContentWidthPt  = max(ExcelLeft + ExcelWidthPt) - min(ExcelLeft)
comContentHeightPt = max(ExcelTop + ExcelHeightPt) - min(ExcelTop)

// Expected pixel extent if COM were correct
expectedContentWidthPx  = comContentWidthPt  × scaleX
expectedContentHeightPx = comContentHeightPt × scaleY

// Scale correction = actual rendered size ÷ expected size
contentScaleX = actualContentWidth  / expectedContentWidthPx
contentScaleY = actualContentHeight / expectedContentHeightPx
```

#### Step 4: Apply Corrections to All Fields

```csharp
correctedOriginX = avgLeft;       // actual content left edge
correctedOriginY = avgTop;        // actual content top edge
correctedScaleX  = scaleX × contentScaleX;  // effective pixel-per-point
correctedScaleY  = scaleY × contentScaleY;

foreach (var field in fields)
{
    double offsetPtX = field.ExcelLeft - field.PrintAreaLeft;
    double offsetPtY = field.ExcelTop - field.PrintAreaTop;

    field.Left   = correctedOriginX + offsetPtX × correctedScaleX;
    field.Top    = correctedOriginY + offsetPtY × correctedScaleY;
    field.Width  = field.ExcelWidthPt  × correctedScaleX;
    field.Height = field.ExcelHeightPt × correctedScaleY;
}
```

#### Safety Checks

- **Tolerance:** No correction applied if origin differs by < 2px (already aligned)
- **Clamping:** Content scale clamped to [0.5, 2.0] to prevent garbage corrections
- **Graceful failure:** Whole block wrapped in try/catch — if PNG scan fails, original COM coordinates are used unchanged

### Call Site

Added immediately after `ExtractFields()` returns and before any logging/calibration tests (line ~715 in the final file):

```
Step 12:  ExtractFields(...) → fields
Step 12b: AdjustCoordinatesFromPng(...) → corrected fields in-place
          Log [CORRECT] entry showing old vs new origin/scale
Step 13:  Log first field (now shows corrected values)
Step 14:  Log all-field audit (now shows corrected values)
Step 15:  Calibration tests (use corrected values)
Step 16:  Debug overlay (drawn at corrected positions → aligns with gridlines)
```

---

## Expected Results

| Metric | Before (COM) | After (Corrected) | Actual PNG | Δ |
|--------|-------------|-------------------|-----------|----|
| A1:B2 left | 875px | **~852px** | 853-856px | ≈1px |
| A1:B2 top | 1289.6px | **~1265px** | 1265-1268px | ≈1px |
| A1:B2 width | 400px | **~418px** | 418px | ≈0px |
| A1:B2 height | 120px | **~131px** | 131px | ≈0px |
| C1:D2 left | 1275px | **~1271px** | 1271-1274px | ≈1px |
| C1:D2 width | 400px | **~418px** | 418px | ≈0px |
| A3:D4 left | 875px | **~852px** | 852-856px | ≈1px |
| A3:D4 width | 800px | **~836px** | 836px | ≈0px |

The debug overlay rectangles should now align with cell gridlines within **1–2px tolerance**.

---

## Log Output

When correction is applied, the server logs:

```
[CORRECT] Content boundary adjustment applied:
  Origin: (875.0,1289.6)px -> (852.0,1265.0)px (Δ -23.0,-24.6)
  Scale:  4.166667x4.165404 -> 4.364000x4.545000
```

No log entry means correction was skipped (either already aligned within 2px, or no PNG to scan).

---

## Verification

### To verify the fix:

1. **Rebuild & restart** the API
2. **Upload** the test Excel file
3. **Open** `debug_{id}.png` from the `Preview/` directory
4. **Check:** Colored rectangles should align with cell gridlines
5. **Check server logs** for `[CORRECT]` entry showing applied corrections
6. **Compare** `annotated_{id}.png` (also corrected — draws first field with red rectangle)

### Rollback

Remove the call to `AdjustCoordinatesFromPng()` at line ~715 in `ExcelCaptureService.cs`, or set it to return false. All other code remains — fields will use original COM coordinates as before.
