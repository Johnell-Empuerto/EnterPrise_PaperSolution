# Phase 5.2.5 — Forensic Investigation: Missing Configuration Sheet Regression

## Status: **Investigation Complete** ✅ (No code changes)

---

## Task 1 — Documentary Evidence

### Key Documents Found

| Document | Relevant Findings |
|---|---|
| **`docs/Phase31_MetadataReverseEngineeringReport.md`** | ConMas stores field metadata in two ways: (1) Hidden `_Fields` worksheet (7 columns: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes), (2) Cell comments as fallback |
| **`docs/PhaseX6_Legacy_OutputExcel_Complete_Reverse_Engineering.md`** | The legacy ConMas Output Excel has 11 structural differences vs our generated workbook. **Most critical: `ExcelOutputSetting` worksheet** — a 36-row XML configuration sheet |
| **`docs/PhaseX10_Implementation_Report.md`** | Documents `CreateExcelOutputSetting()` — creates the 36-cell XML config sheet during export |
| **`docs/OutputExcel_ForensicComparison.md`** | ConMas creates: `_Fields` sheet (hidden, 7 columns), `ExcelOutputSetting` sheet, cell comments with ConMas field types |
| **`docs/phase11j14_legacy_reverse_engineering.md`** | Hidden `_Fields` sheet is the primary metadata source. Field names, types, and parameters stored there |
| **`docs/phase11j5_coordinate_system_unification_investigation.md`** | **Critical regression documented here** — OpenXmlParser was modified to SKIP hidden/VeryHidden sheets. The `_Fields` sheet (state="hidden") was being included in Runtime JSON as a second sheet |
| **`ConMas_Cluster_Discovery_Reference.md`** | ConMas's `MakeCluster()` reads cell comments per-cell via `HasComment()` — NOT the `_Fields` sheet. The `_Fields` sheet is only read by `GetDefinition()` (XML export path) |

---

## Task 2 — Configuration Sheet Comparison

### What ConMas Designer Stores

| Artifact | Type | Location | Created By |
|---|---|---|---|
| **`_Fields` sheet** | Hidden worksheet | `xl/worksheets/sheet*.xml` | `WorkbookGenerator.PopulateFieldsWorksheet()` |
| **`ExcelOutputSetting` sheet** | Visible worksheet | `xl/worksheets/sheet*.xml` | `WorkbookGenerator.CreateExcelOutputSetting()` |
| **Cell comments** | Comments XML | `xl/comments*.xml` | `WorkbookGenerator.WriteCellComments()` |
| **Defined names** | workbook.xml | `xl/workbook.xml` | `WorkbookGenerator.SetDefinedNames()` |

### Two Distinct Pipelines

#### Legacy Pipeline (BROKEN — `POST /api/form/output-excel` deprecated, returns 410)

```
FormDefinition/WorkbookDefinition
    │
    ▼
WorkbookGenerator.Generate()          ← COM-based, creates ALL metadata
    ├── EnsureSheets()                 ← Creates 3 sheets: _Fields, Sheet1, ExcelOutputSetting
    ├── ApplySheetLayout()             ← Applies page setup, column widths, merged cells
    ├── WriteCellValues()              ← Writes cell values from form data
    ├── WriteCellComments()            ← Creates cell comments with ConMas field types
    ├── PopulateFieldsWorksheet()      ← Fills _Fields sheet with 7-column metadata
    ├── CreateExcelOutputSetting()     ← Creates 36-row XML configuration sheet
    └── SetDefinedNames()              ← Adds Print_Area defined name
    │
    ▼
Output workbook: HAS _Fields + ExcelOutputSetting + comments + defined names ✅
```

#### Current Pipeline (ACTIVE — `POST /api/form/save-edited`)

```
WorkbookDefinition
    │
    ▼
WorkbookValueWriter.WriteValues()      ← OpenXml-based, only edits cell values
    ├── File.Copy(source, output)      ← Copies original workbook (has _Fields from ConMas)
    ├── SpreadsheetDocument.Open()     ← SDK modifies XML internally
    ├── Write cell values              ← Only writes to VISIBLE content sheets
    ├── SpreadsheetDocument.Dispose()  ← SDK writes all parts (potentially mutating _Fields)
    └── ZIP restore                    ← Restores everything EXCEPT worksheets + shared strings
    │
    ▼
Output workbook: HAS _Fields (if original had it) ❌ UPDATED? NO
                  MISSING ExcelOutputSetting (never created) ❌
                  HAS comments (if original had them) — NOT UPDATED ❌
```

---

## Task 3 — Pipeline Trace

### Where the Configuration Sheet Disappears

```
Original XLSX (from ConMas Designer)
  Has: _Fields sheet ✅, ExcelOutputSetting sheet ✅, Comments ✅
    │
    ▼
SessionWorkbookStore (TempWorkbooks/{sessionId}/original.xlsx)
  File.Copy preserves everything ✅
    │
    ▼
WorkbookValueWriter.WriteValues()
  1. File.Copy(source, output) → preserves everything ✅
  2. Pre-save ZIP entries → all entries saved ✅
  3. SpreadsheetDocument.Open(output, true)
     └── SDK opens in read-write mode
     └── SDK internally parses all parts
     └── _Fields sheet XML is parsed (not modified unless we change values)
     └── ExcelOutputSetting sheet XML is parsed (not modified)
  4. Write cell values to VISIBLE content sheets only
     └── _Fields sheet: NOT modified (no values written here)
     └── ExcelOutputSetting: NOT modified
  5. Save SharedStringTable → new strings appended
  6. wsPart.Worksheet.Save() → only content sheets saved
  7. SpreadsheetDocument.Dispose() → SDK writes ALL parts back to ZIP
     └── _Fields sheet XML: SDK serializes it (potentially with different whitespace/ordering)
     └── ExcelOutputSetting: SDK serializes it (potentially with changes)
  8. ZIP restore loop
     └── xl/workbook.xml → RESTORED ✅ (sheet definitions, visibility preserved)
     └── xl/styles.xml → RESTORED ✅ (styles preserved)
     └── xl/worksheets/sheet*.xml → SKIPPED ❌ (includes _Fields + ExcelOutputSetting!)
     └── xl/sharedStrings.xml → SKIPPED (intentional)
  == END ==
    │
    ▼
Output XLSX
  _Fields sheet: SDK's version (potentially mutated by SDK close) ⚠️
  ExcelOutputSetting: SDK's version (potentially mutated) ⚠️
  Comments: RESTORED from original ✅
  Defined names: RESTORED from original workbook.xml ✅
    │
    ▼
Re-upload → WorkbookReaderService.Read()
  └── Reads _Fields sheet from COM
  └── If _Fields missing or empty → fallback to cell comments
  └── Cell comments are preserved → fallback SHOULD work
```

### Critical Findings in the Trace

1. **`Post /api/form/save-edited` does NOT call `WorkbookGenerator` at all.** The `WorkbookGenerator` which creates `_Fields`, `ExcelOutputSetting`, and comments is BYPASSED entirely.

2. **`WorkbookValueWriter` only edits cell values on VISIBLE content sheets.** It never writes to `_Fields` sheet or `ExcelOutputSetting` sheet.

3. **The ZIP restore SKIPS all `xl/worksheets/sheet*` entries**, which means the SDK's version of the `_Fields` sheet and `ExcelOutputSetting` sheet are kept in the output. If the SDK mutated them during open/close (different whitespace, reordered attributes, recalculated shared string indices), the output has the SDK-modified versions.

4. **The `_Fields` sheet in the output is NOT updated with new field values** — it has the ORIGINAL data from when ConMas Designer created the workbook. But since cell values ARE written to the content sheets, the field addresses in `_Fields` should still be correct.

5. **Cell comments are preserved** (they're in `xl/comments*.xml` which is restored from the ZIP). So the re-upload fallback to comments should work.

---

## Task 4 — Regression Timeline

### When It Worked

| Phase | Endpoint | Generator | _Fields Created? | ExcelOutputSetting? | Re-upload Works? |
|---|---|---|---|---|---|
| Pre-Phase 4.3 | `POST /api/form/output-excel` | `WorkbookGenerator` (COM) | ✅ Yes (PopulateFieldsWorksheet) | ✅ Yes (CreateExcelOutputSetting) | ✅ |
| Phase 4.3 | `POST /api/form/output-excel` | `Designer/Generation/WorkbookGenerator` (WbDef→COM) | ✅ Yes | ✅ Yes | ✅ |

### When It Broke

| Phase | Endpoint | Generator | _Fields Created? | ExcelOutputSetting? | Re-upload Works? |
|---|---|---|---|---|---|
| **Phase 4.4** | **`POST /api/form/save-edited`** | **`WorkbookValueWriter`** | **❌ No (only copies original)** | **❌ No (only copies original)** | **❌ Depends on original** |
| Phase 4.4.1 | `POST /api/form/save-edited` | `WorkbookValueWriter` | ❌ No | ❌ No | ❌ |
| Phase 5.2 | `POST /api/form/save-edited` | `WorkbookValueWriter` | ❌ No | ❌ No | ❌ |

### The Exact Regression Point

**Phase 4.4 — "Native WorkbookDefinition Editing & Excel Round-Trip"**

The `POST /api/form/save-edited` endpoint was introduced in Phase 4.4, using `WorkbookValueWriter` instead of `WorkbookGenerator`. `WorkbookValueWriter`:
- Does NOT call `PopulateFieldsWorksheet()` — so `_Fields` sheet is never updated
- Does NOT call `CreateExcelOutputSetting()` — so `ExcelOutputSetting` sheet is never created
- Does NOT call `WriteCellComments()` — so comments are never updated

### But Wait — The Original File Should Have These

**This is the subtlety:** If the ORIGINAL uploaded workbook is from ConMas Designer, it ALREADY has:
- `_Fields` sheet (populated with original field metadata)
- `ExcelOutputSetting` sheet (with original ConMas XML)
- Cell comments (with original field types)

`WorkbookValueWriter` copies the original file via `File.Copy`, so these artifacts should be PRESERVED in the output.

**Hypothesis: The issue is not that these artifacts are missing, but that:**

1. **The `_Fields` sheet data is STALE** — it contains the ORIGINAL field definitions, not updated with any new fields the user might have configured in the browser
2. **The SDK silently mutates the `_Fields` sheet XML** during `SpreadsheetDocument.Open/Dispose` — different whitespace, attribute ordering, or shared string index shifts
3. **The `ExcelOutputSetting` sheet might be removed or corrupted** by the SDK if the original file doesn't have one (e.g., if the uploaded workbook is NOT from ConMas Designer)

### The `WorkbookXmlChanges = 1` Connection

The validator reports `WorkbookXmlChanges = 1` for the non-sheet content of `workbook.xml`. This difference is likely:

1. **`<calcPr calcId="..."/>`** — The SDK changes the `calcId` attribute, and the ZIP restore of `xl/workbook.xml` might NOT be working (the workbook.xml entry is stored and restored correctly, but the Phase 5.2.4 SHA256 diagnostic would confirm this)

2. OR: **`<workbookPr date1904="..."/>`** — The SDK changes workbook properties

3. OR: **`<bookViews>`** — The SDK adds/modifies the bookViews node

The Phase 5.2.4 diagnostic logging (element-by-element XPath comparison) will definitively show which specific node differs.

---

## Root Cause

### Primary Root Cause

**The browser save pipeline switched from `WorkbookGenerator` (which actively creates `_Fields`, `ExcelOutputSetting`, and comments) to `WorkbookValueWriter` (which only edits cell values in a copy of the original workbook).**

- `WorkbookGenerator.PopulateFieldsWorksheet()` → **NOT called** in the new pipeline
- `WorkbookGenerator.CreateExcelOutputSetting()` → **NOT called** in the new pipeline
- `WorkbookGenerator.WriteCellComments()` → **NOT called** in the new pipeline

### Secondary Issue

**The ZIP restore in `WorkbookValueWriter` skips all `xl/worksheets/sheet*` entries**, including the `_Fields` sheet. If the SDK modifies the `_Fields` sheet XML during open/save, the output will contain the SDK-modified version, not the original. The `CalculateDefaultColumnWidths` base function in the OpenXml SDK is known to adjust column widths on open, which could affect the `_Fields` sheet's layout.

### Contributing Factor

**The legacy `POST /api/form/output-excel` endpoint was deprecated and returns 410 Gone.** The browser cannot fall back to it even if the `WorkbookValueWriter` output is missing metadata.

---

## Recommended Fix

### Option A: Restore `_Fields` & `ExcelOutputSetting` Generation in the Current Pipeline

Add metadata generation to `FormSaveService.SaveEditedValuesAsync()`:

```csharp
// After WorkbookValueWriter.WriteValues():
// 1. Update _Fields sheet with current field definitions
// 2. Create/update ExcelOutputSetting sheet
// 3. Update cell comments
```

This can be done by calling the relevant methods from `WorkbookGenerator` on the `WorkbookValueWriter` output file.

### Option B: Chain `WorkbookValueWriter` → `WorkbookGenerator` 

```
WorkbookValueWriter edits values in original workbook
    │
    ▼
WorkbookGenerator opens the edited workbook via COM
    ├── PopulateFieldsWorksheet() — updates _Fields with current field definitions
    ├── CreateExcelOutputSetting() — creates/config output sheet
    └── WriteCellComments() — updates comments with current field data
```

### Option C: Add `_Fields` Update to `WorkbookValueWriter`

Add a new method to `WorkbookValueWriter` (or a new service) that:
1. Opens the edited workbook after `WorkbookValueWriter` finishes
2. Locates the `_Fields` sheet by name
3. Populates the 7-column metadata table from `WorkbookDefinition`
4. Saves and restores metadata

### Recommendation

**Option C** is the most architecturally aligned — it keeps everything in the clean OpenXml path without needing COM interop. The `WorkbookValueWriter` already has the `SpreadsheetDocument` open; adding `_Fields` sheet population before `Dispose()` is straightforward.

---

## Files Involved

| File | Role | Status |
|---|---|---|
| `ExcelAPI/ExcelAPI/Designer/Generation/WorkbookGenerator.cs` | Creates _Fields, ExcelOutputSetting, comments via COM | ✅ Active but NOT CALLED from browser pipeline |
| `ExcelAPI/ExcelAPI/Application/WorkbookValueWriter.cs` | Edits cell values only | ❌ Does NOT create designer metadata |
| `ExcelAPI/ExcelAPI/Application/FormSaveService.cs` | Orchestrates save pipeline | ⚠️ SaveEditedValuesAsync calls WorkbookValueWriter, not WorkbookGenerator |
| `ExcelAPI/ExcelAPI/Controllers/FormController.cs` | Controller endpoints | ⚠️ SaveEdited endpoint uses WorkbookValueWriter path |
| `ExcelAPI/ExcelAPI/Application/WorkbookDiffValidator.cs` | Validates workbook fidelity | ✅ Correctly detects structural differences |
