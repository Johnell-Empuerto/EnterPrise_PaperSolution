# Phase 4.4 — Native WorkbookDefinition Editing & Excel Round-Trip

**Date**: 2026-07-19  
**Status**: ✅ Complete  
**Build**: 0 errors, 0 new warnings  

---

## Summary

Implemented the round-trip editing pipeline: WorkbookDefinition → WorkbookValueWriter → Edited XLSX. The `WorkbookValueWriter` uses DocumentFormat.OpenXml to surgically modify only cell values in the original workbook, preserving all formatting, styles, layout, and structure.

---

## Files Created

| File | Description |
|------|-------------|
| `Application/WorkbookValueWriter.cs` | Core cell value writer. Opens original XLSX, locates cells by CellReference, writes values as inline/shared strings or numbers, preserves StyleIndex, saves as copy. |
| `Application/WorkbookDiffValidator.cs` | Diff validator comparing original vs edited XLSX. Reports: style changes, page setup changes, merge changes, formula changes, sheet count changes, editable value changes. |

## Files Modified

| File | Change |
|------|--------|
| `Models/WorkbookDefinition/FieldDefinition.cs` | Added `Value` property (string?) for user-entered field values. |
| `Application/FormSaveService.cs` | Added `WorkbookValueWriter` field/constructor param. Added `SaveEditedValuesAsync()` method. |
| `Controllers/FormController.cs` | Added `POST /api/forms/save-edited` endpoint accepting `WorkbookDefinition` with edited values. |
| `Application/ServiceRegistration.cs` | Registered `WorkbookValueWriter` and `WorkbookDiffValidator`. |

---

## Architecture

```
Original Excel (.xlsx)
        │
        ├──→ COM Capture → WorkbookDefinition (with SourcePath)
        │
        ▼
Frontend Editor
        │
        ▼
Edited WorkbookDefinition (fields have .Value set)
        │
        ▼
POST /api/forms/save-edited
        │
        ▼
FormSaveService.SaveEditedValuesAsync()
        │
        ▼
WorkbookValueWriter.WriteValues(WbDef, sourcePath, outputPath)
        │
        ├── 1. Copy original XLSX (never modify source)
        ├── 2. Open copy with SpreadsheetDocument.Open(read-write)
        ├── 3. For each sheet → find WorksheetPart
        ├── 4. For each field → find/create Cell by CellReference
        ├── 5. Write value (inline string / number)
        ├── 6. Preserve StyleIndex (never touch formatting)
        └── 7. Save and return edited .xlsx
        │
        ▼
Edited Excel (.xlsx) — IDENTICAL to original except cell values
```

---

## Key Implementation Details

### WorkbookValueWriter
- Uses `DocumentFormat.OpenXml` for all operations (no COM dependency)
- Copies the original XLSX before modifying (source is never touched)
- Preserves `StyleIndex` on every cell — styles, fonts, borders, fills remain unchanged
- Preserves formulas — only writes into editable input cells (non-formula)
- Uses SharedStringTable for string values (adds new entries as needed)
- Handles numeric fields (Number, Calculated) as `CellValues.Number`
- Creates new cells with column-positioned insertion if they don't exist in original

### WorkbookDiffValidator
- Compares: sheet count, sheet names, visibility, stylesheet XML, merged cells, page setup (PrintOptions, PageSetup), cell style indices, formulas, cell values
- Reports: structural changes (expected: 0) and editable value changes (expected: >0)
- Uses `OuterXml` comparison for stylesheet and page setup elements

### API Endpoint
**POST** `/api/forms/save-edited`
- Accepts: `WorkbookDefinition` JSON with fields having `Value` set and `SourcePath` pointing to the original XLSX
- Returns: File download (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
- Error: 400 if no definition, 500 on save failure

---

## Known Gaps (Documented for Phase 4.5)

| Gap | Description | Priority |
|-----|-------------|----------|
| Image replacement | Signature images and photo placeholders are not written back | Medium |
| Auto-validation after save | `WorkbookDiffValidator.Compare()` is not called automatically after save | Low |
| New cell styling | Cells created during write (not in original) get default style | Low |
| DiffValidator named ranges | Named range comparison not yet implemented | Low |
| DiffValidator hidden rows/cols | Hidden row/column comparison not yet implemented | Low |
| Fallback for missing SourcePath | Currently throws FileNotFound if SourcePath is missing | Low |

---

## Readiness for Phase 4.5

**Ready** ✅

The codebase is ready for Phase 4.5 which should focus on:

1. **Image replacement** — Write signature/photo images back into the workbook's drawing layer
2. **Auto-validation pipeline** — Wire `WorkbookDiffValidator` into `SaveEditedValuesAsync`
3. **Inline strings** — Consider using inline strings (`CellValues.InlineString`) instead of SharedStringTable to avoid SST modification
4. **Full diff validation** — Add named range, hidden row/column, drawing layer checks to `WorkbookDiffValidator`
