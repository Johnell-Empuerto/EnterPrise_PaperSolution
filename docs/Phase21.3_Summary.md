# Phase 21.3 — Export Pipeline Deep Diagnostic Logging

**Date:** July 20, 2026
**Status:** Complete ✅
**Build:** 0 errors

---

## Objective

Instrument the entire export pipeline with extremely detailed logs to identify exactly where field data disappears between frontend submission and workbook save.

The hard evidence that motivated this phase:
```
Fields received (total): 0
Fields with non-empty values: 0
Cells actually written: 0
EXPORT COMPLETE: 0 fields written
```

## Implementation Summary

Three files were enhanced with 7 diagnostic stages:

---

### Stage 1 — Export Payload Received (`FormController.cs`)

At the start of `SaveEdited()`, prints structured per-field dump:

```
=========================================================
EXPORT PAYLOAD RECEIVED
=========================================================
Workbook: FormTest
Session: abc123
Pages (sheets): 1
Total Fields: 18
Fields With Values: 6
Empty Fields: 12
```

Every field logged with: Id, Page, Cell, Type, Value, Required, ReadOnly, Placeholder, Default Value, Max Length, Data Validation.

Logs CRITICAL warning if ALL fields have empty values (model binding mismatch detection).

---

### Stage 2 — Workbook Definition Diagnostic (`FormSaveService.cs`)

Before calling `WorkbookValueWriter.WriteValues()`, logs the complete `WorkbookDefinition` with per-field dump.

Enables comparison: Stage 1 output vs Stage 2 output. If fields disappear between Controller → Service, the bug is in the service call.

---

### Stage 3 — Workbook Value Writer Input (`WorkbookValueWriter.cs`)

At the start of `WriteValues()`, logs source path, output path, sheets, field counts per sheet.

Enables comparison: Stage 2 output vs Stage 3 output. If fields disappear between Service → Writer, the bug is in the call chain.

---

### Stage 4 — Cell Write Operations

Every cell write is logged with `[DIAG]` tags:
- Sheet and cell address
- Old value (if cell existed)
- New value being written
- Whether row/cell was NEW or EXISTS
- Formula vs Number vs SharedString type
- StyleIndex preserved

---

### Stage 5 — Workbook Verification (Post-write)

After the workbook is saved and closed, it is re-opened in read-only mode. Every written cell is read back and compared against the expected value:

```
  Cell Sheet1!C7:
    Expected Value: 'John Smith'
    Actual Value:   'John Smith'
    Value Result:   PASS
```

---

### Stage 6 — Style Verification

For every written cell, the actual cell style properties are read from the workbook and logged:

```
    Style Check — Cell C7:
      Style properties: not sent by frontend in export payload.
      Preserved from original workbook (template):
      Font Name:     'Arial'
      Font Size:     '14'
      Bold:          'True'
      Italic:        'False'
      Underline:     'False'
      Font Color:    '#000000'
      Fill Color:    '#FFFF00'
      H-Align:       'Center'
      V-Align:       'Center'
      Wrap Text:     'False'
    Workbook styles preserved successfully from original template.
```

Per the user's instruction, no expected style values are invented — styles come from the workbook template and are logged for diagnostics only.

---

### Stage 7 — Export Summary

Final summary with PASS/FAIL for verification:

```
=========================================================
EXPORT SUMMARY
=========================================================
Fields Received:        18
Fields With Values:     6
Fields Written:         6
Cells Written:          6
Cells Skipped (empty):  12
Cells Skipped (no sheet): 0

Workbook Verification:
  Cells Pass:           6
  Cells Fail:           0
  Style Checks:         6

Overall Result:         PASS
=========================================================
```

---

## Compilation Fixes

| Error | Fix |
|-------|-----|
| `FieldDefinition` has no `Options` property | Replaced with `DataValidation?.Formula1` |
| `stylePass++` undeclared variable | Renamed to `styleChecked++` |
| Duplicate `"Validation"` log label | Changed to `"DV Formula"` |

---

## Files Modified

| File | Changes |
|------|---------|
| `Controllers/FormController.cs` | Stage 1: EXPORT PAYLOAD RECEIVED diagnostics |
| `Application/FormSaveService.cs` | Stage 2: WORKBOOK DEFINITION diagnostics before WriteValues |
| `Application/WorkbookValueWriter.cs` | Stages 3-7: WRITER INPUT, verification + style check, EXPORT SUMMARY |

---

## Remaining Gaps

1. **Stage 4 format**: Uses `[DIAG]` tags instead of the user's requested `WRITE CELL` / `SKIPPED` structured blocks. Empty fields are skipped silently via `continue;` — no SKIPPED entry is produced.

2. **Dead code**: `BuildDesignerModel()` (~150 lines) in `FormController.cs` is unreachable since `UploadExcel` was updated to use `_designerModelReader.Read()` in Phase 21.1.

---

## Next Steps

With Phase 21.3's diagnostics in place, running a real export and analyzing the logs will reveal exactly where field data is lost. The likely candidates are:
- **Frontend model binding mismatch**: `runtimeFormToWorkbookDefinition` maps fields with property names that don't match `FieldDefinition`
- **Payload serialization**: Frontend sends `f.id` but backend expects `f.Name`
- **Empty value filter**: Backend silently skips empty values — if all fields are "empty" due to binding mismatch, nothing gets written
