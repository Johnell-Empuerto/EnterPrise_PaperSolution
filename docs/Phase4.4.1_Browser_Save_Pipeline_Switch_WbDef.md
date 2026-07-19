# Phase 4.4.1 — Browser Save Pipeline Switch to WorkbookValueWriter

## Status: **Complete** ✅

---

## Objective

Switch the browser's "Export Excel" save pipeline from the legacy COM-based `WorkbookGenerator` to the native `WorkbookValueWriter` pipeline that surgically edits only cell values in the original workbook.

---

## Pipeline Audit

### Legacy Path (NOW FALLBACK ONLY)

```
Frontend "Export Excel" button
  ↓
runtimeFormToFormDefinition() ← LOSES original workbook structure
  ↓
POST /api/form/output-excel   ← LEGACY endpoint
  ↓
FormSaveService.OutputExcelAsync()
  ↓
WorkbookGenerator.Generate(FormDefinition) ← COM rebuilds workbook from scratch
  ↓
RESULT: Extra worksheet, bold headers, lost alignment, broken styles
```

### Canonical Path (NOW PRIMARY)

```
Frontend "Export Excel" button
  ↓
runtimeFormToWorkbookDefinition() ← PRESERVES original structure
  ↓
POST /api/form/save-edited        ← NEW endpoint (Phase 4.4)
  ↓
FormSaveService.SaveEditedValuesAsync()
  ↓
WorkbookValueWriter.WriteValues() ← OpenXml: edits only cell values
  ↓
WorkbookDiffValidator.Compare()   ← Auto-validation
  ↓
RESULT: Only edited values differ — everything else preserved
```

---

## Frontend Changes

### Files Modified

- **`paperless/app/page.tsx`**

### What Changed

1. **Added `sourceFileName` state** (line ~51) — stores the template filename from upload response

2. **In `handleUpload`** (upload-preview flow) — extracts `result.templateId` and stores as `sourceFileName`

3. **In `handleUploadExcel`** (upload-excel flow) — extracts `result.data.templateId` and stores as `sourceFileName`

4. **In `handleReset`** — resets `sourceFileName` along with other state

5. **Added `runtimeFormToWorkbookDefinition()`** — converts `RuntimeForm` + `runtime.values` + `sourceFileName` into a minimal WbDef JSON payload:
   ```typescript
   {
     info: { title: "..." },
     sourceFileName: "abc123.xlsx",
     sheets: [{
       name: "Sheet1",
       index: 0,
       fields: [{
         cell: { address: "B5", rowIndex: 5 },
         name: "FullName",
         type: "KeyboardText",
         value: "John Smith"
       }]
     }]
   }
   ```

6. **Added `handleSaveEdited()`** — sends WbDef to `POST /api/form/save-edited`, handles binary download, logs `X-Validation-Results` header

7. **Updated `handleExportExcel()`** — now prefers `handleSaveEdited()` when `sourceFileName` is known; falls back to legacy `POST /api/form/output-excel` otherwise

8. **Calls `runtime.markAllClean()`** after successful save — resets the dirty state

### Upload Flow Limitations

| Upload Method | Source Saved? | `sourceFileName` Set? | Uses Save-Edited? |
|---|---|---|---|
| Upload Preview (first button) | ❌ File deleted | ❌ No | ❌ Falls back to legacy |
| Upload Excel (second button) | ✅ Persisted to Forms/ | ✅ Yes | ✅ Yes |

---

## Backend Changes

### Files Modified

- **`ExcelAPI/ExcelAPI/Application/FormSaveService.cs`**
- **`ExcelAPI/ExcelAPI/Controllers/FormController.cs`**

### FormSaveResult (updated)

Added new fields to support the save-edited pipeline:
```csharp
public WorkbookDiffResult? ValidationResult { get; set; }
public int CellsWritten { get; set; }
public string SourcePath { get; set; } = "";
```

### SaveEditedValuesAsync (rewritten)

The method now:
1. **Logs comprehensive pipeline info** — "========== SAVE PIPELINE ==========" with workbook title, source file, sheets count, and every field value being written
2. **Logs "WorkbookGenerator invoked: FALSE"** — explicit confirmation that the legacy path is NOT used
3. **Logs "WorkbookValueWriter invoked: TRUE"** — confirmation of the canonical path
4. **Requires SourceFileName** — no dangerous fallback that picks random XLSX files
5. **Logs warning for zero values** — "No field values to write — workbook copied unchanged."
6. **Auto-runs WorkbookDiffValidator** after writing values
7. **Logs validation results** — sheet count changed, styles changed, merges changed, formulas changed, editable values changed

### FormSaveService Constructor

Added `WorkbookDiffValidator diffValidator` parameter — registered in `Application/ServiceRegistration.cs` since Phase 4.4.

### SaveEdited Controller (updated)

The endpoint now:
1. **Logs pipeline banner** — "========== SAVE PIPELINE ==========" with controller info
2. **Logs "WorkbookGenerator invoked: FALSE"** at controller level
3. **Logs "WorkbookValueWriter invoked: TRUE"** at controller level
4. **Adds `X-Validation-Results` response header** — base64-encoded JSON with validation results:
   ```json
   {
     "passed": true,
     "cellsWritten": 12,
     "sheetCountChanged": 0,
     "styleChanges": 0,
     "mergeChanges": 0,
     "pageSetupChanges": 0,
     "formulaChanges": 0,
     "namedRangeChanges": 0,
     "editableValueChanges": 12,
     "totalDiffs": 0
   }
   ```
5. Uses `result.CellsWritten` instead of hardcoded `0`

---

## DI Registration

Both services confirmed registered in `Application/ServiceRegistration.cs`:
```csharp
services.AddScoped<WorkbookValueWriter>();      // Phase 4.4
services.AddScoped<WorkbookDiffValidator>();     // Phase 4.4
```

---

## Build Verification

| Check | Result |
|---|---|
| Backend build | ✅ 0 errors, 0 new warnings |
| Frontend TypeScript | ✅ Compiles (no type errors) |
| DI registration | ✅ Both services registered |

---

## Console Log Output (Expected)

```
========== SAVE PIPELINE ==========
Controller: POST /api/form/save-edited
Workbook: FormTest - Copy.xlsx
SourceFileName: abc123.xlsx
Sheets: 1, Fields with values: 12
WorkbookGenerator invoked: FALSE
WorkbookValueWriter invoked: TRUE
  B5 -> John Smith
  C8 -> 2026-07-19
  F12 -> Checked
  No field values to write — workbook copied unchanged.
WorkbookValueWriter complete: 12 cells written → edited_xxx.xlsx
========== VALIDATION RESULTS ==========
Sheet Count Changed: 0
Styles Changed: 0
Merged Cells Changed: 0
Alignment/Print Changed: 0
Formula Changed: 0
Named Range Changed: 0
Editable Values Changed: 12
Validation Passed: True
========== END VALIDATION ==========
========== SAVE PIPELINE END ==========
```

---

## Architecture (After Phase 4.4.1)

```
Excel Workbook (uploaded)
      │
      ▼
COM Capture (Server-side)
      │
      ▼
WorkbookDefinition
      │
      ▼
Browser Editing
      │
      ▼
WorkbookDefinition (edited values + sourceFileName)
      │
      ▼
POST /api/form/save-edited  ← PRIMARY PATH
      │
      ├── WorkbookValueWriter.WriteValues()
      │       ↓
      │   Opens original XLSX via OpenXml
      │   Writes ONLY cell values
      │   Preserves EVERYTHING else (styles, layout, formulas, merges, print settings)
      │
      ├── WorkbookDiffValidator.Compare()
      │       ↓
      │   Validates: sheet count, styles, merges, page setup, formulas, named ranges
      │
      └── Returns edited.xlsx + X-Validation-Results header
```

---

## Risk Assessment

| Risk | Mitigation |
|---|---|
| Source file not found | Clear error message returned; frontend falls back to legacy path |
| Empty field values | Warning logged; workbook returned unchanged |
| Validation failure | Warning logged with details; file still returned |
| Legacy path still exposed | Preserved as fallback for upload-preview flow; can be removed in Phase 4.5 |

---

## Phase 4.5 Recommendation

Consider removing the legacy `POST /api/form/output-excel` endpoint entirely once:
- Frontend always has `sourceFileName` available (all uploads go through upload-excel)
- The legacy `WorkbookGenerator` is no longer needed for any workflow
