# Phase 11I — Form Runtime Engine

**Date:** July 2026  
**Milestone:** Runtime Architecture  
**Build:** 0 errors, 21 warnings  
**Status:** ✅ Complete

---

## Objective

Build the Form Runtime Engine that converts a parsed Excel workbook into a runtime form definition consumable by the Next.js frontend. The Rendering Engine remains unchanged — this phase builds the runtime metadata layer that sits above rendering.

---

## Architecture

```
Excel Workbook (.xlsx)
        │
        ▼
OpenXmlParser (Rendering)
        │
        ▼
RenderWorkbook (Rendering)
        │
        ▼
FormRuntimeBuilder (NEW)
        │
        ├── GeometryBuilder.ComputeGeometry()
        ├── CoordinateEngine.GetCellOrMergePixelBounds()
        ├── FieldDetector.DetectFields()
        │   ├── FieldTypeResolver.ResolveType()
        │   └── CoordinateEngine pixel bounds
        └── RuntimeSerializer
        │
        ▼
RuntimeForm JSON
        │
        ▼
GET /api/runtime/{templateId}
        │
        ▼
Next.js Frontend
```

---

## Files Created (7 new in `Runtime/`)

| File | Purpose |
|:-----|:--------|
| `RuntimeField.cs` | Editable field model — Id, CellReference, Row, Column, LeftPx/TopPx/WidthPx/HeightPx, MergeRange, IsMerged, DataType, ReadOnly, Required, Alignment, Font, FontSize, Bold, FontColor, BackgroundColor, Border, Placeholder, DefaultValue, MaxLength, TabIndex, ValidationPattern, ValidationMessage |
| `RuntimeSheet.cs` | Sheet model — Name, Index, Fields, Images, Shapes, PrintArea, PageWidthPx, PageHeightPx |
| `RuntimeForm.cs` | Complete runtime form — WorkbookName, Title, Sheets, PageWidth, PageHeight, Scale, Dpi, Version |
| `FieldTypeResolver.cs` | Cell data type detection — checkbox (☐☑✓✗), date (parse + serial numbers), number, calculated, text. Also ReadOnly detection, Required detection, Border style resolution |
| `FieldDetector.cs` | Editable field detection — empty bordered cells, merged cells with borders, alignment heuristics. Uses CoordinateEngine for all pixel coordinates |
| `FormRuntimeBuilder.cs` | Orchestrates runtime construction — ensures geometry computed, detects fields per sheet, converts drawing objects, builds RuntimeForm |
| `RuntimeSerializer.cs` | JSON serialization — camelCase naming, null skipping, indented/minified options |

## Files Modified (2)

| File | Change |
|:-----|:-------|
| `Controllers/FormController.cs` | Added `GET /api/runtime/{templateId}` endpoint. Parses XLSX via `OpenXmlParser`, builds `RuntimeForm` via `FormRuntimeBuilder`, returns JSON. `FindTemplateFile` searches Forms and Uploads directories with exact + prefix matching. |
| `Program.cs` | Added `using ExcelAPI.Runtime;` and registered `FieldTypeResolver`, `FieldDetector`, `FormRuntimeBuilder`, `RuntimeSerializer` as singletons. |

---

## RuntimeField Properties

| Property | Type | Description |
|:---------|:----:|:------------|
| `Id` | string | Unique field identifier (e.g., "field_A1_0") |
| `CellReference` | string | Excel cell reference (e.g., "A1") |
| `Row` | uint | Row index (1-based) |
| `Column` | uint | Column index (1-based) |
| `LeftPx` | double | Left edge in pixels (page-relative) |
| `TopPx` | double | Top edge in pixels |
| `WidthPx` | double | Field width in pixels |
| `HeightPx` | double | Field height in pixels |
| `MergeRange` | string? | Merge range reference (e.g., "A1:C3") |
| `IsMerged` | bool | Part of a merged range |
| `DataType` | string | "text", "number", "date", "checkbox", "signature", "dropdown", "calculated" |
| `ReadOnly` | bool | Whether the field is read-only |
| `Required` | bool | Whether the field is required |
| `Alignment` | string? | "left", "center", "right", "general" |
| `Font` | string? | Font name |
| `FontSize` | double | Font size in points |
| `Bold` | bool | Bold font |
| `FontColor` | string? | Font color in #AARRGGBB |
| `BackgroundColor` | string? | Background fill color in #AARRGGBB |
| `Border` | string? | "none", "thin", "medium", "thick", "double" |
| `Placeholder` | string? | Placeholder text |
| `DefaultValue` | string? | Default/cell value |
| `MaxLength` | int | Maximum input length |
| `TabIndex` | int | Tab order for keyboard navigation |
| `ValidationPattern` | string? | Regex validation pattern |
| `ValidationMessage` | string? | Validation error message |

---

## Field Detection Rules

| Rule | Condition | Example |
|:-----|:----------|:--------|
| Empty bordered cell | No value + any border | `A1: empty, thin border` → editable text field |
| Filled bordered cell | Has value + any border | `B2: "Name", medium border` → editable field with default |
| Merged bordered cell | Merge index + any border | `C3:E5 merged, thick border` → editable merged field |
| Alignment signal | center/right alignment | `D4: value, center-aligned` → editable field |
| Checkbox pattern | ☐☑✓✗ characters | `E1: "☐"` → checkbox field |
| Date pattern | Date parse + serial numbers | `F2: "2026-01-15"` → date field |
| Number pattern | Numeric parse | `G3: "123.45"` → number field |

### Read-Only Detection

- Cells with **no border** → read-only label
- Cells with **text value + no visible border** → read-only label
- All other editable cells → editable (ReadOnly = false)

### Required Detection

- Empty cells with borders → required
- Cells with placeholder containing "required", "*", or "mandatory" → required

---

## Coordinate Rules

**Never recompute coordinates.**

Every `RuntimeField` pixel coordinate comes from `CoordinateEngine.GetCellOrMergePixelBounds()`:

```
bounds = _coords.GetCellOrMergePixelBounds(sheet, cell, originXPt, originYPt)
field.LeftPx   = bounds.Left
field.TopPx    = bounds.Top
field.WidthPx  = bounds.Width
field.HeightPx = bounds.Height
```

Image coordinates use `CoordinateEngine.PtToPx(DrawingParser.EmuToPt(emu))`.

No coordinate recomputation occurs anywhere in the Runtime namespace.

---

## API Endpoint

### `GET /api/runtime/{templateId}`

**Response:**

```json
{
  "success": true,
  "message": "Runtime form built: 1 sheet(s), 42 field(s)",
  "data": {
    "workbookName": "Invoice",
    "title": "Invoice",
    "sheets": [
      {
        "name": "Invoice",
        "index": 0,
        "fields": [
          {
            "id": "field_B3_0",
            "cellReference": "B3",
            "row": 3,
            "column": 2,
            "leftPx": 583.33,
            "topPx": 125.0,
            "widthPx": 250.0,
            "heightPx": 20.83,
            "mergeRange": null,
            "isMerged": false,
            "dataType": "text",
            "readOnly": false,
            "required": true,
            "alignment": "left",
            "font": "Calibri",
            "fontSize": 11,
            "bold": false,
            "backgroundColor": "#FFFFFFFF",
            "border": "thin",
            "placeholder": "B3",
            "defaultValue": null,
            "maxLength": 0,
            "tabIndex": 0
          }
        ],
        "images": [],
        "shapes": [],
        "printArea": null,
        "pageWidthPx": 2550,
        "pageHeightPx": 3300
      }
    ],
    "pageWidth": 2550,
    "pageHeight": 3300,
    "scale": 1.0,
    "dpi": 300,
    "version": "1.0"
  }
}
```

---

## Dependency Rules

| Dependency | Usage |
|:-----------|:------|
| `CoordinateEngine` | All pixel coordinate calculations in FieldDetector, FormRuntimeBuilder |
| `GeometryBuilder` | Geometry computation in FormRuntimeBuilder.Build() |
| `OpenXmlParser` | XLSX parsing in FormController.GetRuntime() |
| `RenderingContext` | Not directly used — origin coordinates passed as parameters |

No rendering code is called from the Runtime namespace. No rendering engine files were modified.

---

## Known Limitations

| Issue | Status |
|:------|:-------|
| Named range detection | Not implemented — named ranges in workbook.xml are not parsed by the current model |
| Alignment heuristic | `HorizontalAlignment = "center"/"right"` may falsely identify non-editable cells |
| FindTemplateFile prefix match | May match wrong file if multiple GUIDs share a prefix |
| Validation patterns | `ValidationPattern`/`ValidationMessage` are declared but not auto-detected |
| Dropdown fields | Not auto-detected — data validation lists would require additional OpenXml parsing |

---

## Files Summary

### New Files

```
ExcelAPI/ExcelAPI/Runtime/
├── RuntimeField.cs              # Editable field model
├── RuntimeSheet.cs              # Sheet model
├── RuntimeForm.cs               # Complete runtime form model
├── FieldTypeResolver.cs         # Data type detection
├── FieldDetector.cs             # Editable field detection
├── FormRuntimeBuilder.cs        # Runtime construction orchestrator
└── RuntimeSerializer.cs         # JSON serialization
```

### Modified Files

```
ExcelAPI/ExcelAPI/Controllers/
└── FormController.cs             # GET /api/runtime/{templateId}

ExcelAPI/
└── Program.cs                    # using + DI registrations
```

No rendering engine files were modified.

---

## Roadmap After Phase 11I

```
✅ Phase 11A — Rendering Core
✅ Phase 11B — Rendering Pipeline
✅ Phase 11C — Text Engine
✅ Phase 11D — Coordinate & Print Layout
✅ Phase 11E — Style Resolution & Theme Engine
✅ Phase 11F — Image & Shape Engine
✅ Phase 11G — Production Export Engine
✅ Phase 11H — Regression & Pixel-Diff Validation
✅ Phase 11I — Form Runtime Engine
────────────────────────────────────
⬜ Phase 11J — Field Overlay Engine (Yellow Editable Layer)
⬜ Phase 11K — Designer Enhancements
⬜ Phase 11L — Production Release
```
