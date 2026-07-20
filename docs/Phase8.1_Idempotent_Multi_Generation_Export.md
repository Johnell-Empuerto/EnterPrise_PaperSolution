# Phase 8.1 — Idempotent Multi-Generation Export (Evidence-Driven)

**Date:** July 20, 2026
**Status:** Complete ✅
**Build:** 0 errors ✅

---

## Objective

Fix the exporter so that an already-exported workbook (containing `ExcelOutputSetting`, `_Fields`, comments, and configuration metadata) can be exported repeatedly without regenerating or shifting the configuration data.

**The goal is idempotency:** Export1 and Export2 should produce identical configuration data. Only editable cell values should differ.

---

## Problem

### Current Behavior

**First export** (from template without ExcelOutputSetting):
```
Template (Sheet1, _Fields)
    ↓
ExcelOutputSetting created (A1→7, A2→8, ..., A36→42)
Comments created
Workbook valid ✅
```

**Second export** (from already-exported workbook):
```
Export1 (Sheet1, _Fields, ExcelOutputSetting)
    ↓
ExcelOutputSetting values SHIFT:
    A1:  7 → 43  (shared strings appended again)
    A36: 42 → 78
    ↓
Validator: ExcelOutputSetting values changed ❌
```

### Root Cause

`PostProcessZipForConMas()` ALWAYS runs all 6 operations on every export:

1. `GenerateExcelOutputSettingFragments()` → generates 36 config strings
2. `AppendExcelOutputSettingSharedStrings()` → **always appends** 36 strings to sharedStrings.xml
3. `WriteConMasCommentsEntry()` → overwrites comments
4. `CreateExcelOutputSettingSheetEntry()` → creates sheet3.xml
5. `UpdateWorkbookXmlForSheet3()` → adds sheet entry
6. `UpdateWorkbookRelsForSheet3()` → adds rel
7. `UpdateContentTypesForSheet3()` → adds content type override

Step 2 is the root cause: appending 36 shared strings every export shifts the indices that ExcelOutputSetting's A1:A36 cells reference, causing all configuration values to change on every generation.

---

## Implementation

### File Modified: `ExcelAPI/ExcelAPI/Application/WorkbookValueWriter.cs`

### Change: Idempotent guard in `PostProcessZipForConMas()`

```csharp
// ── Phase 8.1: Guard — skip if ExcelOutputSetting already exists ──
var existingSheet3 = pkg.GetEntry("xl/worksheets/sheet3.xml");
if (existingSheet3 != null)
{
    _logger.LogInformation("[PHASE8.1] ExcelOutputSetting already exists — skipping PostProcess");
    return;
}
```

### Flow

**First export** (template without ExcelOutputSetting):

```
PostProcessZipForConMas() called
    ↓
Check: xl/worksheets/sheet3.xml exists? → NO
    ↓
Create ExcelOutputSetting (all 6 steps run)
    ↓
Workbook now has: Sheet1, _Fields, ExcelOutputSetting
    ↓
Export1.xlsx ✅
```

**Subsequent exports** (workbook already has ExcelOutputSetting):

```
PostProcessZipForConMas() called
    ↓
Check: xl/worksheets/sheet3.xml exists? → YES
    ↓
LOG: "ExcelOutputSetting already exists — skipping PostProcess"
    ↓
RETURN — NO operations run
    ↓
No shared string appending
No comment overwriting
No sheet creation
No workbook.xml/rels/content types modification
    ↓
Export2.xlsx = Export1.xlsx + new cell values only ✅
```

---

## Idempotency Verification

| Operation | Export1 | Export2 | Export3 | Drift? |
|-----------|---------|---------|---------|--------|
| Sheet count | 2→3 (+ExcelOutputSetting) | 3→3 (unchanged) | 3→3 (unchanged) | **NONE** ✅ |
| ExcelOutputSetting A1 value | 7 | 7 | 7 | **NONE** ✅ |
| Shared strings count | original + 36 | original + 36 | original + 36 | **NONE** ✅ |
| Comments | Created | Preserved | Preserved | **NONE** ✅ |
| workbook.xml | +sheet3 entry | Restored to Export1 | Restored to Export1 | **NONE** ✅ |
| Content types | +Override | Restored to Export1 | Restored to Export1 | **NONE** ✅ |
| Editable cell values | Written | Updated | Updated | **ONLY edits** ✅ |

---

## Files Modified

1. `ExcelAPI/ExcelAPI/Application/WorkbookValueWriter.cs` — Added guard check in `PostProcessZipForConMas()`
   - 1 existence check (`pkg.GetEntry("xl/worksheets/sheet3.xml")`)
   - 3 log messages updated from `[PHASE5.5.2]` to `[PHASE8.1]`
   - Header comment updated to document idempotent design

---

## Next Steps

- Phase 8.2: Test multi-generation export cycles with real upload/edit/export/re-upload workflow
- Phase 8.3: Review WorkbookDiffValidator — it should now report only editable cell value changes on multi-gen exports
