# Phase 11J.7 — Replace Runtime Coordinate Engine with Excel COM Geometry

**Date:** July 2026
**Status:** Architecture Investigation (No Code Changes)

---

## 1. Current Architecture: Two Coordinate Pipelines

### Pipeline A: Excel COM (powers the PNG background)

```
Excel Workbook
    ↓
ExcelCaptureService.CapturePrintAreaAsync()     [ExcelCaptureService.cs:55]
    ├── Launch Excel COM (hidden)                [line 105]
    ├── Open workbook                             [line 124]
    ├── Select first VISIBLE worksheet            [line 130-148]
    ├── Read Print Area from PageSetup            [line 155+]
    ├── Read Range.Left, Top, Width, Height       [line 260-277]
    ├── Read PageSetup margins, centering         [line 300-370]
    ├── Compute printed origin (margin+center)    [line 376-400]
    ├── Export worksheet → PDF via COM            [line 410-420]
    ├── Convert PDF → PNG via PDFium              [line 440-460]
    ├── Read actual PNG dimensions                [line 470-480]
    ├── Compute actual scale (pngW / pageW)       [line 486-495]
    ├── Extract commented cells via SpecialCells  [ExtractFields:1070-1150]
    │   ├── Range.Left/Top (COM real measurement)
    │   ├── MergeArea.Left/Top/Width/Height
    │   └── Convert: pixel = printedOrigin + (cellPt - printAreaPt) * scale
    └── Return CaptureResult                      [line 530-560]
         ├── Fields[] (pixel coords from COM)
         ├── PageSetup (margins, centering, scale)
         └── ImageUrl
```

### Pipeline B: OpenXML + CoordinateEngine (powers the Runtime overlay)

```
Excel Workbook
    ↓
OpenXmlParser.Parse()                            [OpenXmlParser.cs:55]
    ├── Open XLSX via SpreadsheetDocument         [line 62]
    ├── Load shared strings                       [line 68]
    ├── Load stylesheet                           [line 71]
    ├── Iterate Sheet entries (now filters hidden) [line 84-97]
    ├── ParseSheet()                              [line 100-160]
    │   ├── ParseColumns: charWidth→pointWidth    [line 165-195]
    │   │   via CharWidthToPoints() formula       [ECMA-376: (w*7.33+5)*72/96]
    │   ├── ParseRows: row height or default 15pt [line 200-240]
    │   └── ParseMerges: column ranges only       [line 245-280]
    └── Return RenderWorkbook (point-based)
             ↓
    GeometryBuilder.ComputeGeometry()             [GeometryBuilder.cs:30]
        ├── Sum column widths via formula         [line 45-60]
        ├── Sum row heights                       [line 65-80]
        └── Compute merge bounds from cumulative  [line 85-100]
             ↓
    CoordinateEngine.GetCellPixelBounds()         [CoordinateEngine.cs:60]
        ├── PtToPx(pt) = pt * dpi/72              [hardcoded 300/72]
        └── Returns SKRect (left, top, right, bottom)
             ↓
    FormRuntimeBuilder.Build()                    [FormRuntimeBuilder.cs:55]
        ├── FieldDetector.DetectFields()          [FieldDetector.cs:35]
        │   └── CoordinateEngine pixel bounds     [line 80]
        └── Returns RuntimeForm (JSON to frontend)
```

---

## 2. Detailed Stage-by-Stage Comparison

| Stage | Excel COM Pipeline | OpenXML Pipeline | **Match?** |
|-------|-------------------|------------------|------------|
| **Cell Left (points)** | `Range.Left` — actual Excel measurement | `CharWidthToPoints(col.width)` — `(w*7.33+5)*72/96` | **❌** Formula vs real |
| **Cell Top (points)** | `Range.Top` — actual Excel measurement | `Row.Height` from XML or default 15pt | **❌** Formula vs real |
| **Cell Width (points)** | `Range.Width` — actual Excel measurement | `CharWidthToPoints(col.width)` — formula approximation | **❌** Formula vs real |
| **Cell Height (points)** | `Range.Height` — actual Excel measurement | `Row.Height` from XML | **≈** (same data source, but COM may have adjusted) |
| **Merge Size (points)** | `MergeArea.Width/.Height` — actual COM | `lastCol+1 cumLeft - firstCol cumLeft` | **❌** COM real vs formula cumulative |
| **Page Margins (points)** | `PageSetup.LeftMargin etc.` — actual COM | **Not read** (passed as originXPt/YPt from metadata) | **≈** (origin is now applied, but still OpenXML cell positions) |
| **Print Area Origin** | `Range.Left/Top of printArea` — actual COM | **Not used** (cell positions are worksheet-relative) | **❌** COM uses print-area-relative; OpenXML uses worksheet-relative + origin |
| **Centering Offset** | `(printable - content) / 2` — computed from COM values | **Not computed** (originXPt from metadata includes it) | **≈** (origin now includes centering) |
| **Scale (px/pt)** | `pngWidth / pageWidthPt` — actual rendered ratio | `dpi / 72 = 300/72 = 4.166666...7` — hardcoded | **≈** (actual vs theoretical, ≤0.05% difference) |
| **Final Pixel Left** | `printedOriginX + (cellLeftPt - printAreaLeft) * scaleX` | `(originXPt + colCumLeftPt) * dpi/72` | **❌** Different cell position sources |
| **Final Pixel Top** | `printedOriginY + (cellTopPt - printAreaTop) * scaleY` | `(originYPt + rowCumTopPt) * dpi/72` | **❌** Different cell position sources |
| **Final Pixel Width** | `cellWidthPt * scaleX` | `colWidthPt * dpi/72` | **❌** Different widths |
| **Final Pixel Height** | `cellHeightPt * scaleY` | `rowHeightPt * dpi/72` | **≈** (height data from same source) |
| **Sheet Filter** | Filters `xlSheetVisible` only | Now filters `Sheet.State == Hidden/VeryHidden` | **✅** Now consistent |
| **Field Detection** | Only cells with comments (`SpecialCells`) | Any cell with borders (`IsEditable()`) | **❌** Totally different logic |

**Key takeaway:** The biggest divergence is in **cell position** (left/top) and **cell size** (width/height). COM reads the actual Excel-measured values. OpenXML reconstructs from column widths using the ECMA-376 character-width formula. These can differ by several points depending on font, column width rounding, and Excel's internal adjustments.

---

## 3. Class Dependency Diagram

```
Current Architecture:

ExcelController.cs           FormController.cs
    │                              │
    ▼                              ├──OpenXmlParser ── GeometryBuilder ── CoordinateEngine
IExcelCaptureService               │       │                  │                 │
    │                              │       ▼                  ▼                 ▼
    ▼                              │   RenderWorkbook   RenderSheet        SKRect
ExcelCaptureService                │       │
    │                              │       ▼
    │                              ├──FormRuntimeBuilder ── FieldDetector ── FieldTypeResolver
    │                              │         │                  │
    │                              │         ▼                  ▼
    │                              │    RuntimeForm        RuntimeField
    │                              │
    ▼                              ├──PersistRuntimeMetadata()  →  {templateId}.meta.json
CaptureResult                      └──LoadRuntimeMetadata()    ←  {templateId}.meta.json
    │
    ├── Fields[] (COM pixel coords)
    ├── PageSetup (margins, scale)
    └── ImageUrl
```

---

## 4. Every Place OpenXML Recalculates What COM Already Knows

| What COM Measures | What OpenXML Recalculates | File:Line | Error Source |
|-------------------|---------------------------|-----------|-------------|
| `Range.Left` | `CharWidthToPoints(col.width, 7.33)` | `GeometryBuilder.cs:45-60` | ECMA-376 formula vs real Excel layout engine |
| `Range.Top` | Sum of previous row heights | `GeometryBuilder.cs:65-80` | Row height rounding differences |
| `Range.Width` | `CharWidthToPoints(col.width)` | `GeometryBuilder.cs:55` | Font-dependent: Calibri 11pt→7.33, but other fonts differ |
| `Range.Height` | `Row.Height ?? 15.0` | `OpenXmlParser.cs:220` | COM may auto-expand height for wrapped text |
| `MergeArea.Width` | `(lastCol+1 cumLeft - firstCol cumLeft)` | `GeometryBuilder.cs:85-100` | Cumulative position errors compound |
| `MergeArea.Height` | `(lastRow+1 cumTop - firstRow cumTop)` | `GeometryBuilder.cs:85-100` | Same compounding issue |
| `SpecialCells(xlCellTypeComments)` | `IsEditable()` border detection | `FieldDetector.cs:90-115` | Totally different field detection strategy |

---

## 5. Why the OpenXML Pipeline Cannot Match COM Exactly

The ECMA-376 column width formula (used by `CharWidthToPoints`):
```
pixelWidth = charWidth * maxDigitWidth + padding
pointWidth = pixelWidth * 72 / 96
```
where `maxDigitWidth ≈ 7.33` (Calibri 11pt) and `padding = 5 pixels`.

This is an **approximation** of how Excel calculates displayed column width. Excel's actual width calculation depends on:
1. The specific font face (not just font size — Calibri vs Arial vs Aptos have different max digit widths)
2. The font's actual max digit width at the current font size
3. The DPI setting of the display/printer
4. Excel's internal rounding rules for column width display

**Excel COM** bypasses all of this by calling `Range.Width` which returns the exact width that Excel has already calculated internally.

Similarly for row height: `Range.Height` returns the actual height after Excel has applied auto-fit, text wrapping, and font metrics. The OpenXML row height is either stored explicitly (which may not match due to Excel's auto-adjustments) or defaults to 15pt.

---

## 6. Proposed Simplified Architecture

### Target Architecture: COM-Only Coordinate Pipeline

```
Upload
  │
  ▼
ExcelCaptureService.CapturePrintAreaAsync()
  ├── Launch Excel COM
  ├── Open workbook
  ├── Select visible worksheet
  ├── Export → PDF → PNG
  ├── Read ALL cell geometry via COM:
  │   ├── Fields (commented cells with pixel coords)
  │   ├── All merged ranges
  │   └── Page setup (margins, centering, scale)
  │
  └── Persist to disk:
      ├── Preview/page_{id}.png
      ├── Forms/{id}.xlsx
      └── Forms/{id}.runtime.json   ★ NEW — persistent runtime metadata
                                          │
                                          ▼
GET /api/form/runtime/{id}
  │
  ▼
Read Forms/{id}.runtime.json
  │
  ▼
Return RuntimeForm (no OpenXML parsing, no geometry recalculation)
```

### Files in Forms/ for Each Template

```
Forms/
  {templateId}.xlsx           ← original workbook (moved from uploads)
  {templateId}.runtime.json   ← ★ NEW: complete runtime metadata
  page_{templateId}.png       ← in Preview/
```

### Runtime Metadata Schema (proposed `{id}.runtime.json`)

```json
{
  "version": "1.0",
  "capturedAt": "2026-07-10T12:00:00Z",
  "workbookName": "Form_ABC",
  "dpi": 300,
  "scaleX": 4.166667,
  "scaleY": 4.166667,
  "pageWidthPx": 2550,
  "pageHeightPx": 3300,
  "originXPt": 54.0,
  "originYPt": 54.0,
  "sheets": [
    {
      "name": "Sheet1",
      "index": 0,
      "pageWidthPx": 2550,
      "pageHeightPx": 3300,
      "fields": [
        {
          "id": "field_B5",
          "cellReference": "B5",
          "leftPx": 225.0,
          "topPx": 225.0,
          "widthPx": 420.0,
          "heightPx": 20.0,
          "dataType": "text",
          "readOnly": false,
          "required": false,
          "fontSize": 11,
          "bold": false,
          "isMerged": false
        }
      ],
      "merges": [
        {
          "reference": "A1:C3",
          "firstCol": 1,
          "firstRow": 1,
          "lastCol": 3,
          "lastRow": 3,
          "leftPx": 0.0,
          "topPx": 0.0,
          "widthPx": 600.0,
          "heightPx": 45.0
        }
      ],
      "images": [],
      "drawings": []
    }
  ]
}
```

### Target Pipeline

```
GET /api/form/runtime/{id}
  ↓
Read Forms/{id}.runtime.json
  ↓
Deserialize to RuntimeForm
  ↓
Return JSON
  ↓
Frontend renders FormPage + OverlayFields
```

**No OpenXmlParser call. No GeometryBuilder. No CoordinateEngine. No FieldDetector.**

The runtime JSON is generated ONCE during upload (when Excel COM is running and has all the real measurements) and served from disk on every subsequent request.

---

## 7. Classes That Become Obsolete

With the COM-only architecture, the following classes would no longer be needed for Runtime:

| Class | File | Status | Notes |
|-------|------|--------|-------|
| `OpenXmlParser` | `OpenXmlParser.cs` | **Obsolete for Runtime** | Still useful for other purposes (e.g., cell style resolution, drawing parsing) |
| `GeometryBuilder` | `GeometryBuilder.cs` | **Obsolete for Runtime** | Only used to compute geometry for CoordinateEngine |
| `CoordinateEngine` | `CoordinateEngine.cs` | **Obsolete** | COM provides pixel coords directly — no conversion needed |
| `FieldDetector` | `FieldDetector.cs` | **Obsolete** | COM already knows which cells have comments/types |
| `FieldTypeResolver` | `FieldTypeResolver.cs` | **Obsolete** | Type comes from cell comment, not heuristics |
| `FormRuntimeBuilder` | `FormRuntimeBuilder.cs` | **Replaced** | New simpler builder that reads persisted metadata |
| `RenderWorkbook` | (rendering models) | **Obsolete for Runtime** | Internal rendering model not needed |
| `RenderSheet` | (rendering models) | **Obsolete for Runtime** | |
| `RenderCell` | (rendering models) | **Obsolete for Runtime** | |
| `RenderMerge` | (rendering models) | **Obsolete for Runtime** | |
| `RenderColumn` | (rendering models) | **Obsolete for Runtime** | |

**Classes that remain:**

| Class | File | Reason |
|-------|------|--------|
| `ExcelCaptureService` | `ExcelCaptureService.cs` | Still generates PNG + captures COM geometry |
| `ExcelController` | `ExcelController.cs` | Still handles legacy upload (could be deprecated) |
| `FormController` | `FormController.cs` | Still hosts from-excel and runtime endpoints |
| `StyleResolver` | `StyleResolver.cs` | May still be needed for styling (font, color, border info) |
| `DrawingParser` | `DrawingParser.cs` | Still needed for image/shape extraction |
| `ImageResolver` | `ImageResolver.cs` | Still needed for embedded image data |
| `ThemeResolver` | `ThemeResolver.cs` | Still needed for color resolution |
| `ColorResolver` | `ColorResolver.cs` | Still needed for color resolution |
| `CellGeometryEngine` | `CellGeometryEngine.cs` | Only if legacy rendering is still used |

---

## 8. Migration Plan

### Phase 1: Persist Complete Runtime Metadata During Upload

**Files to modify:** `FormController.cs`, `ExcelCaptureService.cs`

- After `CapturePrintAreaAsync()` returns, iterate all visible worksheets via COM
- For each worksheet, read ALL cells (not just commented ones) using `UsedRange`
- Capture `Range.Left/Top/Width/Height` for every cell and merge area
- Serialize everything to `Forms/{templateId}.runtime.json`

**Note:** This is the most impactful change. Currently `ExcelCaptureService` only reads cells with comments. To generate complete runtime metadata, it needs to iterate every cell in the print area (or used range) and capture geometry.

**Risk:** Reading every cell via COM is slow for large worksheets.

### Phase 2: Create Metadata-Based Builder

**New file:** `RuntimeMetadataService.cs`

- Reads `{id}.runtime.json` and deserializes to `RuntimeForm`
- No OpenXML parsing involved
- Falls back to legacy OpenXML path if no `.runtime.json` exists

### Phase 3: Redirect GetRuntime() to Use Metadata

**File to modify:** `FormController.cs` — `GetRuntime()`

- Try reading `{id}.runtime.json` first
- If found, return immediately (no OpenXML, no geometry)
- If not found, fall back to current OpenXML path

### Phase 4: Parallel Run & Validation

- Run both pipelines (COM metadata and OpenXML) and compare output
- Log differences between COM and OpenXML field positions
- Verify overlay alignment with PNG using COM metadata vs OpenXML

### Phase 5: Remove Obsolete Code

- After COM metadata is validated, remove `GeometryBuilder`, `CoordinateEngine`, `FieldDetector` dependencies from Runtime pipeline
- Keep the classes if used elsewhere (e.g., for rendering or debugging)

---

## 9. Risks

| Risk | Impact | Mitigation |
|------|--------|-----------|
| **Slower upload** — iterating all cells via COM | Upload time increases by 2-10 seconds | Only read cells in the print area, not the entire worksheet. Add progress logging. |
| **COM dependency** — requires Excel installed | Server must have Excel | Already required — no change. But persists even for read-only runtime. |
| **Large metadata files** — thousands of cells | File size could be MB+ | Use compact JSON format. Only store cells that are relevant (commented cells + merged cells). |
| **Stale metadata** — runtime.json out of sync with XLSX | Metadata doesn't match workbook | Metadata is generated fresh on every upload. The XLSX is the source of truth. |
| **OpenXmlParser still needed for drawings** — COM doesn't expose drawing data | Drawing coordinates lost | Keep DrawingParser for image/shape extraction only. |
| **OpenXmlParser still needed for style resolution** — COM gives font/border info but OpenXML is more complete | Style information incomplete | Can combine: COM for geometry, OpenXML for style resolution. Two-phase approach. |

---

## 10. Recommended Implementation Plan

### Step 1 (Highest Priority): Persist COM Field Rectangles

Modify `ExcelCaptureService.CapturePrintAreaAsync()` to:
- After extracting commented cells, also iterate all cells in the print area
- For each cell, capture `Range.Left/Top/Width/Height` and whether it's merged
- For merged cells, capture `MergeArea.Left/Top/Width/Height`
- Save this as `Forms/{templateId}.runtime.json`

### Step 2: Create COM-Based Runtime Builder

New class `ComRuntimeBuilder`:
- Reads `Forms/{templateId}.runtime.json`
- Converts the COM pixel coordinates directly into `RuntimeField` list
- Returns `RuntimeForm` immediately — no OpenXML, no geometry computation

### Step 3: Update GetRuntime()

- Try `ComRuntimeBuilder` first
- Fall back to existing `FormRuntimeBuilder` (OpenXML) if metadata not found

### Step 4: Validate & Remove Legacy Code

- Compare outputs of both builders for the same workbook
- Verify overlay alignment in the browser
- Once validated, remove `GeometryBuilder`, `CoordinateEngine`, `FieldDetector`, and OpenXML parsing from the Runtime pipeline

---

## 11. Success Criteria

After migration:

- `GET /api/form/runtime/{id}` returns fields positioned using **exact COM cell measurements**, not OpenXML formula approximations
- Background PNG and HTML overlay share the **same coordinate system** (both from Excel COM)
- No `CharWidthToPoints()` calls for Runtime overlay positioning
- No `GeometryBuilder.ComputeGeometry()` for Runtime
- No `CoordinateEngine` for Runtime
- `FieldDetector` is replaced by COM's direct cell reading
- Upload is 2-5 seconds slower (COM cell iteration) but runtime is faster (no OpenXML parsing)
- Legacy templates (pre-migration) still work via fallback to OpenXML path
