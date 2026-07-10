# Phase 11J.5 — Coordinate System Unification Investigation

**Status:** Investigation Only (No Code Changes)

---

## 1. Two Coordinate Systems Exist Side-by-Side

### Coordinate System A: Excel COM (powers the PNG background)

**Entry point:** `ExcelCaptureService.CapturePrintAreaAsync()`  
**File:** `ExcelAPI/ExcelAPI/Services/ExcelCaptureService.cs`

| Step | Code Location | Value |
|------|--------------|-------|
| Page margins (points) | Lines ~660-665: `leftMarginPt = worksheet.PageSetup.LeftMargin` etc. | COM reads actual PageSetup values |
| Paper size (points) | Line ~690: `GetPaperSizePoints(paperSize, orientation)` | Letter = 612x792, A4 = 595x842, etc. |
| Printable area (points) | Lines ~696-697: `printableWidthPt = pageWidthPt - leftMarginPt - rightMarginPt` | Page minus margins |
| **Printed origin X (points)** | Lines ~700-704: `printedOriginXPts = leftMarginPt` then `+= (printableWidthPt - printAreaWidth) / 2.0` if centered | **Accounts for margins AND centering** |
| **Printed origin Y (points)** | Lines ~700-708: same pattern with topMargin and vertical centering | **Accounts for margins AND centering** |
| Scale X (px/pt) | Line ~929: `scaleX = pngWidth > 0 && pageWidthPt > 0 ? pngWidth / pageWidthPt : dpi / 72.0` | **Actual rendered PNG width / page width** |
| Scale Y (px/pt) | Line ~931: same for Y | **Actual rendered PNG height / page height** |
| **Final origin (pixels)** | Lines ~936-937: `actualPrintedOriginX = printedOriginXPts * scaleX` | Origin in pixels using actual scale |
| **Field pixel Left** | Line ~1095: `pixelLeft = printedOriginX + (cellLeftPt - printAreaLeft) * scaleX` | COM cell position minus print area offset, scaled |
| **Field pixel Top** | Line ~1098: same for Y | Same pattern |
| Cell geometry source | Line ~1065: `cellLeftPt = mergeArea.Left` or `cell.Left` | **Actual Excel COM Range.Left/Top/Width/Height** |

**Summary:** Uses Excel COM to read real cell dimensions, real page setup, actual PNG output dimensions. Everything is physically measured.

---

### Coordinate System B: OpenXML + CoordinateEngine (powers the Runtime overlay fields)

**Entry point:** `FormRuntimeBuilder.Build()`  
**File:** `ExcelAPI/ExcelAPI/Runtime/FormRuntimeBuilder.cs`

| Step | Code Location | Value |
|------|--------------|-------|
| Page origin X (points) | `FormRuntimeBuilder.Build()` parameter default | **`originXPt = 0` (default)** |
| Page origin Y (points) | Same parameter default | **`originYPt = 0` (default)** |
| Scale (px/pt) | `CoordinateEngine.PointsToPixels = 300.0 / 72.0` (hardcoded) | **4.166666...7 always** |
| Column cumulative left | `GeometryBuilder.ComputeGeometry()` → `CharWidthToPoints(w, maxDigitWidth)` | `(charWidth * 7.33 + 5.0) * 72.0 / 96.0` |
| Row cumulative top | `GeometryBuilder.ComputeGeometry()` → sum of row heights | `sheet.DefaultRowHeight = 15.0` or explicit |
| **Field pixel Left** | `CoordinateEngine.GetCellPixelBounds()` → `PtToPx(originXPt + colCumLeftPt)` | **Origin is 0, no margins, no centering** |
| **Field pixel Top** | Same → `PtToPx(originYPt + rowCumTopPt)` | **Origin is 0, no margins, no centering** |
| Cell geometry source | `OpenXmlParser.ParseColumns()` + `ParseRowsAndCells()` | **OpenXML column widths converted via ECMA-376 formula** |

**Summary:** Uses OpenXML column widths converted via formula, with **zero origin offset**. No margins, no centering, no print area offset.

---

## 2. The Caller — Where It All Breaks

**File:** `ExcelAPI/ExcelAPI/Controllers/FormController.cs` — `GetRuntime()` method (line ~163)

```csharp
var runtimeForm = _runtimeBuilder.Build(workbook);
//                     ^^^^^^^^^^
//     originXPt = 0 (default), originYPt = 0 (default)
//     No margin, centering, or print area info is passed!
```

The `Build()` method is called with **default origin parameters** = (0, 0). The `FormController.GetRuntime()` endpoint has no access to the COM page setup data (margins, centering, paper size) that `ExcelCaptureService` computed during upload. These values are never persisted to the runtime metadata.

### Compare: what the PNG background uses vs what the Runtime overlay uses

| Property | COM (PNG background) | OpenXML (Runtime overlay) | **Match?** |
|----------|---------------------|--------------------------|------------|
| Origin X (points) | `leftMargin + centering offset` | `0` | **❌** |
| Origin Y (points) | `topMargin + centering offset` | `0` | **❌** |
| Scale (px/pt) | `pngWidth / pageWidthPt` (≈4.1667) | `300/72 = 4.166666...7` | **❌** (tiny diff ±0.05%) |
| Cell width (points) | COM `Range.Width` (real Excel measurement) | `CharWidthToPoints(col.width, 7.33)` | **❌** (formula vs real) |
| Cell height (points) | COM `Range.Height` (real Excel measurement) | OOXML row height or default 15pt | **❌** (formula vs real) |

---

## 3. Exact Divergence Point

**The divergence begins at `FormRuntimeBuilder.Build()` line ~57:**

```csharp
public RuntimeForm Build(
    RenderWorkbook workbook,
    int dpi = 300,
    double originXPt = 0,    // ← NO MARGIN OFFSET
    double originYPt = 0)    // ← NO MARGIN OFFSET
```

Called from `FormController.GetRuntime()` line ~170:
```csharp
var runtimeForm = _runtimeBuilder.Build(workbook);
// Default origin = (0,0) — no page setup information available!
```

**Root cause:** The `FormController.GetRuntime()` method has no access to the page setup data (margins, centering, paper size) that was read during upload. These values are computed by `ExcelCaptureService.CapturePrintAreaAsync()` using Excel COM, but they are never stored in a location accessible to the `GetRuntime()` endpoint, which uses the OpenXML parser independently.

---

## 4. Magnitude of the Vertical Offset

For a standard Letter workbook with default margins (0.75" left/right, 0.75" top/bottom = **54pt each**) and no centering:

| Axis | Missing margin offset (points) | Missing margin offset (pixels at 300 DPI) |
|------|-------------------------------|------------------------------------------|
| X    | 54pt                          | ~225px                                   |
| Y    | 54pt                          | ~225px                                   |

If centering is enabled, the offset is even larger. For a typical 8-column letter-size page (8.5" × 11" = 612 × 792pt):

| Property | Value |
|----------|-------|
| Page width | 612pt |
| Left + Right margins (default) | 54 + 54 = 108pt |
| Printable width | 504pt |
| Print area content width | ~400pt (varies) |
| Horizontal centering offset | (504-400)/2 = 52pt additional |
| **Total missing X offset** | **54 + 52 = 106pt ≈ 442px** |

This means **every field is shifted ~400-450px up and left** from where it should be. The fields appear in the top-left corner of the page instead of at their correct printed positions.

---

## 5. Hidden _Fields Worksheet — Root Cause

**File:** `OpenXmlParser.cs` line ~84

```csharp
foreach (var sheetEntry in doc.WorkbookPart.Workbook.Descendants<Sheet>())
{
    var wsPart = doc.WorkbookPart.GetPartById(sheetEntry.Id!) as WorksheetPart;
    if (wsPart == null) continue;

    var renderSheet = ParseSheet(wsPart, sheetEntry, sheetIndex, ...);
    workbook.Sheets.Add(renderSheet);
    sheetIndex++;
}
```

**The parser iterates ALL `Sheet` elements in workbook.xml without checking the `state` attribute.**

The OpenXML `Sheet` element has a `State` property (from `DocumentFormat.OpenXml.Spreadsheet.SheetStateValues`):

| State Value | Meaning |
|-------------|---------|
| `visible` (default) | Sheet is visible in the Excel UI |
| `hidden` | Sheet is hidden but can be unhidden via Excel UI |
| `veryHidden` | Sheet is hidden and can only be unhidden via VBA/macro |

The `_Fields` hidden worksheet is created by the legacy VSTO Add-in to store form field metadata. It has `state="hidden"` in the workbook XML. The parser ignores this attribute entirely.

**Exact location of the missing filter:** `OpenXmlParser.cs:84` — the `foreach` loop that iterates `sheetEntry` from `Workbook.Descendants<Sheet>()`.

**Compare with COM path:** `ExcelCaptureService.cs:130-148` correctly skips non-visible worksheets:
```csharp
bool isVisible = ws.Visible == Excel.XlSheetVisibility.xlSheetVisible;
if (isVisible) { selectedSheet = ws; break; }
```

---

## 6. Summary of All Divergence Points

| # | Issue | File | Line(s) | Impact |
|---|-------|------|---------|--------|
| 1 | **Runtime origin = (0,0)** — no margin/centering offset passed to `Build()` | `FormRuntimeBuilder.cs` | 57-59 | Runtime fields shifted ~400-450px up-left from correct position |
| 2 | **COM scale is actual PNG ratio** — `pngWidth/pageWidthPt` | `ExcelCaptureService.cs` | 929-931 | ~0.05% scale difference vs hardcoded 300/72 |
| 3 | **Runtime scale is hardcoded** — `300/72` always | `CoordinateEngine.cs` | 15 | Never matches actual rendered PNG scale |
| 4 | **Cell width: COM + actual vs OpenXML + formula** | Both | Multiple | Per-cell width differences from font-rendering approximations |
| 5 | **Hidden sheet filter missing** in OpenXML parser | `OpenXmlParser.cs` | 84 | Hidden `_Fields` sheet leaks into Runtime JSON as a second sheet |
| 6 | **Page setup data is ephemeral** — computed during upload but never persisted for `GetRuntime()` | `FormController.cs` | ~170 | `FormRuntimeBuilder.Build()` has no access to margins/centering |

---

## 7. Recommended Fixes (Not Yet Implemented)

### Fix 1: Pass origin offset to FormRuntimeBuilder.Build()
**Files:** `FormController.cs` → `FormRuntimeBuilder.cs`
- Store `printedOriginXPts`, `printedOriginYPts` from the Excel COM capture in a metadata file or the template record
- Pass these values as `originXPt`, `originYPt` to `FormRuntimeBuilder.Build()`

### Fix 2: Store page setup data with the template
**File:** `FormController.cs` (`FromExcel()` method)
- Persist `CaptureResult.PageSetup` (margins, centering, paper size, printedOriginXPt/Y) as JSON metadata alongside the XLSX in `Forms/`
- `GetRuntime()` reads this metadata and passes it to `Build()`

### Fix 3: Filter hidden sheets in OpenXmlParser.Parse()
**File:** `OpenXmlParser.cs` line ~84
- Add a check: `if (sheetEntry.State?.Value == SheetStateValues.Hidden || sheetEntry.State?.Value == SheetStateValues.VeryHidden) continue;`
- This prevents the `_Fields` sheet and any other hidden worksheets from appearing in the runtime output

### Fix 4: Use actual PNG scale instead of hardcoded 300/72
**File:** `CoordinateEngine.cs` and/or `FormRuntimeBuilder.Build()`
- The scale should be configurable (it already is via the `dpi` parameter), but the actual rendered PNG scale from step 11 of `ExcelCaptureService` should be used if available

---

## 8. Conclusion

The overlay fields are vertically offset because the `FormRuntimeBuilder.Build()` method defaults to origin `(0,0)` — no page margins, no centering, no print area offset. The PNG background is rendered with margins and centering applied by Excel's print engine. The Runtime overlay needs the same origin offset applied to align with the PNG.

The hidden `_Fields` worksheet leaks into the Runtime because `OpenXmlParser.Parse()` does not check the `Sheet.State` attribute. The COM path correctly filters it out.
