# Phase 21.4 — Export Payload Deep Trace

**Date:** July 20, 2026
**Status:** Complete ✅
**Build:** 0 errors

---

## Objective

Produce enough logging to identify the exact location where field collection becomes empty during the export pipeline. No business logic was changed — only diagnostics were added.

---

## Implementation Summary

### Files Modified

| File | Stages Added |
|------|-------------|
| `Controllers/FormController.cs` | Stages 1-4, 6, 10-11 |
| `Application/FormSaveService.cs` | Stage 7 |
| `Application/WorkbookValueWriter.cs` | Stages 3/5/7/9 already exist from Phase 21.3 |

---

### Stage 1 — Raw HTTP Request

At the very start of `SaveEdited()`, before any processing:

```
=========================================================
RAW EXPORT REQUEST
=========================================================
Content-Type: application/json
Content-Length: 15284 bytes
Request Path: /api/form/save-edited
Method: POST
Session cookie: present
=========================================================
```

### Stage 2 — Raw JSON Body

Documents that `[FromBody]` consumes the request stream before the action method executes, so raw JSON capture requires an `IAsyncActionFilter` registered before model binding. Logs a clear informational message explaining how to enable this.

### Stage 3 — Model Binding Result

After `[FromBody]` deserialization completes, logs the complete model binding state:

```
=========================================================
MODEL BINDING RESULT
=========================================================
wbDef != null: True
SessionId: abc123
SourceFileName: (null/empty)
SourcePath: (null/empty)
Info: present
Sheets Count: 1
  Sheet[0]: Name='Sheet1', Fields=18
=========================================================
```

### Stage 4 — Dump Every Property

For the first 5 fields of each sheet, logs 17 properties per field:

```
    Field[0]:
      Id:             txtName
      Name:           txtName
      Cell.Address:   C7
      Cell.RowIndex:  7
      Type:           Text
      Value:          John
      Required:       True
      Locked:         False
      Formula:        (null)
      Placeholder:    Enter name
      DefaultValue:   (null)
      MaxLength:      50
      TabIndex:       0
      Visible:        True
      DataValidation: (null)
      Style.Font:     Arial
      Style.FontSize: 14
```

### Stage 5 — WorkbookDefinition Builder

Noted as frontend JavaScript (`runtimeFormToWorkbookDefinition`). Logs explain this runs client-side and would require browser instrumentation.

### Stage 6 — Calling Save Service

Before calling `FormSaveService.SaveEditedValuesAsync()`:

```
=========================================================
CALLING SAVE SERVICE
=========================================================
WorkbookDefinition
  Sheets: 1
  Fields (total): 18
  Fields (with values): 6
  Fields (empty): 12
=========================================================
```

### Stage 7 — Save Service Input

At the start of `SaveEditedValuesAsync()` in FormSaveService:

```
=========================================================
SAVE SERVICE INPUT
=========================================================
Workbook: FormTest
SessionId: abc123
SourcePath: /forms/temp.xlsx
Sheets: 1
Fields (total): 18
Fields (with values): 6
Fields (empty): 12
=========================================================
```

Includes its own FIELD LOSS DETECTED warning if 0 fields enter the service.

### Stage 8 — Calling WorkbookValueWriter

Logged before `_valueWriter.WriteValues()` (within Stage 7 section).

### Stage 9 — WorkbookValueWriter Input

Already exists from Phase 21.3 — logs source, output, sheets, field counts, and per-field details.

### Stage 10 — Field Loss Detector

After export completes, compares field counts across all pipeline stages:

```
=========================================================
FIELD LOSS DETECTOR
=========================================================
Field count tracking:
  Stage 1-4 (Controller after binding):  18
  Stage 6 (Fields with values):           6
  Stage 7-8 (FormSaveService entry):     (see FormSaveService logs)
  Stage 9 (WorkbookValueWriter entry):   (see WorkbookValueWriter logs)
  Cells Written (result):                6
```

Three loss scenarios automatically diagnosed:
- **Controller has 0 fields**: JSON property name mismatch
- **All fields empty**: values[f.id] mapping mismatch
- **0 written despite values**: Sheet name or cell address resolution failure

### Stage 11 — Export Pipeline Summary

Final pipeline trace with auto-diagnosis:

```
=========================================================
EXPORT PIPELINE SUMMARY
=========================================================
Pipeline stage                         | Fields
---------------------------------------------------------
Frontend JSON (sent by browser)        | Unknown (see Stage 2)
ASP.NET Model Binding                  | 18
Fields with non-empty values           | 6
SaveService.SaveEditedValuesAsync()     | see service logs
WorkbookValueWriter.WriteValues()       | see writer logs
Cells Actually Written                 | 6
Workbook Verification                  | see writer logs

Pipeline trace: Normal (all counts consistent)
=========================================================
```

---

## Compilation Fixes

| Error | Fix |
|-------|-----|
| `RowIndex ?? -1` (uint with negative int) | Changed to `Cell != null ? Cell.RowIndex : "(null)"` |
| `formsDir` used before declaration | Moved `Directory.CreateDirectory` before the log section |
| Extra `))` closing parenthesis | Removed duplicate parenthesis |

---

## Remaining Gaps (from code review)

1. **Stage 2 — Raw JSON capture** would require an `IAsyncActionFilter` to run before `[FromBody]`. Addressed with informational logs explaining the workaround.
2. **Stage 1 — User identity** not logged (`HttpContext.User?.Identity?.Name`).
3. **Stage 8** could be a more clearly separated section header in FormSaveService.
4. **Field dump limit** of 5 per sheet could hide partial field loss.

---

## Next Steps

With Phase 21.4's comprehensive diagnostics in place, run a real export and analyze the logs. The pipeline summary will pinpoint exactly where field data disappears. The most likely root causes are:
- **Stage 1-3 gap**: Frontend sends fields but model binding produces 0 (JSON property name mismatch)
- **Stage 3-6 gap**: Fields exist but all have empty values (values[f.id] mapping mismatch)
- **Stage 7-9 gap**: Fields with values but 0 written to workbook (sheet name/cell resolution failure)
