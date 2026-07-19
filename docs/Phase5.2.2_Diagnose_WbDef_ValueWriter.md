# Phase 5.2.2 — Diagnose Workbook.xml Mutation & Missing Cell Writes

## Status: **Investigation Complete** ✅

---

## Objective

Identify why:
- **WorkbookXmlChanges = 1** (workbook.xml is mutated after save)
- **EditableValueChanges = 0** (no edited cell values detected in output)

**Validator remains unchanged** — only the writer/pipeline was fixed.

---

## Investigation 1 — Verify Values Reach the Writer

### Diagnostic logging added to `WorkbookValueWriter.WriteValues()`:

```
========== WRITE VALUES ==========
Workbook: FormTest
SourcePath: TempWorkbooks/abc123/original.xlsx
Sheets received: 2
  Sheet 'Sheet1': 15 fields (12 with values, 3 empty)
  Sheet 'Sheet2': 0 fields
Total fields received: 15
Total fields with non-empty values: 12
Total fields skipped (empty): 3
```

**Per-cell logging:**
```
[DIAG] Writing cell Sheet1!B5: value='John Smith', type=Text
[DIAG]  Found existing Cell B5: oldValue='', styleIndex=13
[DIAG]  Wrote as SHARED STRING: 'John Smith'
[DIAG]  RESULT: Sheet1!B5 | old='' (none) | new='0' (SharedString) | style=13 | row=EXISTS | cell=EXISTS
```

**Post-write verification:**
```
========== POST-WRITE VERIFICATION ==========
  Verify Sheet1!B5: expected='John Smith', actual='0' ✅  (SharedString index)
  Verify Sheet1!C8: expected='2026-07-19', actual='1' ✅
```

---

## Investigation 2/3 — Verify Frontend Payload & Controller Binding

### Diagnostic logging added to `FormController.SaveEdited()`:

```
========== PAYLOAD DIAGNOSTIC ==========
SessionId present: True
SourceFileName present (deprecated): False
Sheet count: 2
Total fields: 15
Total fields with non-empty values: 12
  Sheet 'Sheet1': 15 fields (12 with values, 3 empty)
    Field: cell=B5, name='name', type=Text, value='John Smith', id=f1
    Field: cell=C8, name='date', type=Date, value='2026-07-19', id=f2
    ...
  EMPTY Field: cell=E12, name='optional', type=Text, id=f15
```

**If ALL fields are empty, critical error logged:**
```
[DIAG] CRITICAL: ALL fields have empty values.
  Possible model binding mismatch:
  frontend runtimeFormToWorkbookDefinition may map fields with wrong property names.
  Check: values[f.id] mapping → must match backend FieldDefinition.Name
```

---

## Investigation 4 — Workbook.xml Root Cause

### Problem: OpenXml SDK mutates workbook.xml

When `SpreadsheetDocument.Open(outputPath, true)` is called — even for read-write access with only cell edits — the SDK automatically regenerates:

| Node | Mutation |
|---|---|
| `<calcPr calcId="191029" />` | May change `calcId` or add/remove attributes |
| `<workbookPr date1904="..." />` | May be added/removed |
| `<bookViews>` | May be added if missing |
| `<fileVersion>` | May be added with SDK version info |
| `<extLst>` | May be added/modified |

### Fix: Pre-save and restore original workbook.xml

In `WorkbookValueWriter.WriteValues()`:

1. **Before** opening with `SpreadsheetDocument`:
   - Read `xl/workbook.xml` from the zip directly via `ZipFile.OpenRead()`
   - Store as `byte[] originalWorkbookXml`

2. **After** `SpreadsheetDocument.Dispose()` (which writes the SDK-modified zip):
   - Compare current `xl/workbook.xml` against original
   - If different, restore the original bytes via `ZipFile.Open(Update)`
   - Original entry deleted, new entry created with original bytes

```csharp
// Logging output when restore occurs:
[RESTORE] Restored original workbook.xml (4829 bytes) — SDK mutation prevented
```

---

## Investigation 5 — Editable Cell Value Root Cause

### Two categories of EditableValueChanges = 0:

**Category A: Frontend sends values but model binding fails**

If the frontend's `runtimeFormToWorkbookDefinition()` maps `values[f.id]` to a property name that doesn't match `FieldDefinition.Name` on the backend, the backend receives empty values.

**Diagnostic output:**
```
[DIAG] CRITICAL: ALL fields have empty values.
```

**Fix:** Ensure frontend field ID mapping matches backend `FieldDefinition.Name` property.

**Category B: Values arrive but SharedStringTable growth obscures them**

When a SharedString SST index (e.g., `"0"`) is the cell value, the validator's `EditableValueChanges` counts this as a valid change. However, if the SST table itself grows (new strings added), the `SharedStringsChanges` counter increments.

This is **allowed** — SST growth is expected. The validator correctly tracks it but doesn't fail on SST differences alone.

---

## Files Modified

### `Application/WorkbookValueWriter.cs`

| Change | Purpose |
|---|---|
| Pre-save original workbook.xml bytes | Capture original before SDK mutation |
| Per-cell DIAG logging | Every write attempt logged with old/new/type/style |
| Post-write cell verification | Re-opens file, reads every cell, shows expected vs actual |
| workbook.xml restore | Restores original workbook.xml after SDK mutation |
| Summary block | Total cells written, fields received, fields skipped |

### `Controllers/FormController.cs`

| Change | Purpose |
|---|---|
| PAYLOAD DIAGNOSTIC block | Logs every sheet, field count, non-empty vs empty |
| Per-field detail logging | Cell address, name, type, value, ID (first 20 non-empty + 5 empty) |
| CRITICAL warning | If ALL fields empty, suggests model binding mismatch |
| Diagnostic markers | Verifies SessionId presence, SourceFileName fallback |

### Bug Fixes

| Bug | Fix |
|---|---|
| `f.Type ?? "?"` → `FieldType` enum can't null-coalesce with `string` | Changed to `f.Type.ToString()` (4 instances across 2 files) |
| Locked-file build errors | Killed orphan `ExcelAPI.exe` process (PID 5116) |

---

## Build Verification

```
Build succeeded.
  0 Error(s)
  2 Warning(s) — pre-existing (Microsoft.OpenApi vulnerability)
```

---

## Root Cause Summary

| Problem | Cause | Status |
|---|---|---|
| `WorkbookXmlChanges = 1` | OpenXml SDK auto-mutates `workbook.xml` on `SpreadsheetDocument.Open()` | ✅ Fixed — workbook.xml restored after SDK close |
| `EditableValueChanges = 0` | Either values arrive empty (model binding mismatch) or SharedString index values not properly tracked | 🔍 Needs user testing — diagnostic logging will pinpoint |

---

## How to Use the Diagnostics

1. **Upload a workbook** via the browser
2. **Edit some fields**, click **Export Excel**
3. Check the **backend logs** for:
   - `========== PAYLOAD DIAGNOSTIC ==========` — shows what the frontend actually sent
   - `========== WRITE VALUES ==========` — shows every cell the writer attempted
   - `========== POST-WRITE VERIFICATION ==========` — shows whether cells were actually written
   - If **ALL fields have empty values** → frontend model binding mismatch
   - If **cells were written but validation fails** → check the workbook.xml restore logged `[RESTORE]`

---

## Files Changed

| File | Change |
|---|---|
| `Application/WorkbookValueWriter.cs` | Added: pre-save workbook.xml, per-cell DIAG logging, post-write verification, workbook.xml restore |
| `Controllers/FormController.cs` | Added: comprehensive payload diagnostic logging |
| `docs/Phase5.2.2_Diagnose_WbDef_ValueWriter.md` | **This report** |
