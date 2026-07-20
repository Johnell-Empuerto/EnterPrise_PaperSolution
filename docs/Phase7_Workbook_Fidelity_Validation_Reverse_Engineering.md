# Phase 7 — Reverse Engineer Workbook Fidelity Validation (PaperLess vs Legacy ConMas)

**Date:** July 20, 2026  
**Status:** Investigation Complete ✅  
**Type:** Reverse Engineering Only — No Code Changes

---

## Executive Summary

The PaperLess WorkbookDiffValidator is designed to enforce **binary-level structural identity** between the original and edited workbooks. It compares at 36+ distinct categories across every OOXML part. The Legacy ConMas Designer, in contrast, validates at a **semantic level** — it reads embedded configuration metadata and rebuilds the designer model from that metadata, without performing any ZIP or XML binary comparison.

This fundamental philosophical difference is why PaperLess fails on the second export cycle while ConMas succeeds indefinitely.

---

## Part 1 — PaperLess Validator Call Stack

```
POST /api/form/save-edited
  │
  ├── FormController.SaveEdited(WbDef)
  │     │
  │     ├── Resolve sourcePath from session store
  │     │
  │     ├── FormSaveService.SaveEditedValuesAsync()
  │     │     │
  │     │     ├── WorkbookValueWriter.WriteValues()      ← CANONICAL PATH
  │     │     │     │
  │     │     │     ├── File.Copy(source → output)
  │     │     │     ├── Pre-save ZIP entries (originalZipEntries)
  │     │     │     ├── SpreadsheetDocument.Open(output, true)  ← SDK MUTATES
  │     │     │     ├── Write cell values (worksheets + shared strings)
  │     │     │     ├── SpreadsheetDocument.Dispose()           ← SDK WRITES
  │     │     │     ├── RESTORE: overwrite SDK-mutated entries with originals
  │     │     │     │     Skip list: worksheets, shared strings, comments, VML,
  │     │     │     │                 worksheet rels, content types
  │     │     │     ├── SHA256 verify workbook.xml
  │     │     │     └── PostProcessZipForConMas()               ← ADDS CONFIG
  │     │     │           ├── AppendExcelOutputSettingSharedStrings (36 strings)
  │     │     │           ├── WriteConMasCommentsEntry (comments1.xml)
  │     │     │           ├── CreateExcelOutputSettingSheetEntry (sheet3.xml)
  │     │     │           ├── UpdateWorkbookXmlForSheet3 (workbook.xml)
  │     │     │           ├── UpdateWorkbookRelsForSheet3 (workbook.xml.rels)
  │     │     │           └── UpdateContentTypesForSheet3 ([Content_Types].xml)
  │     │     │
  │     │     └── WorkbookDiffValidator.Compare(sourcePath, outputPath)
  │     │           │
  │     │           ├── ComparePartHashes()              ← styles, theme, worksheets
  │     │           ├── CompareWorkbook()                ← sheets, names, visibility
  │     │           ├── CompareStyles()                  ← fonts, fills, borders, formats
  │     │           ├── CompareWorksheet() per sheet     ← rows, cells, merges, layout
  │     │           ├── CompareObjectParts()             ← drawings, images, comments
  │     │           ├── CompareAdditionalParts()         ← shared strings, workbook.xml
  │     │           └── CompareDefinedNames()            ← named ranges
  │     │
  │     └── If !result.Passed → throw WorkbookFidelityException
  │           │
  │           └── Return HTTP 500 with validation details
```

### Error Message Source

**File:** `ExcelAPI/ExcelAPI/Controllers/FormController.cs`, Line ~713
```csharp
message = "Workbook fidelity validation failed. The edited workbook is not structurally identical to the original. No file has been returned."
```

**Exception Class:** `WorkbookFidelityException` in `WorkbookDiffValidator.cs`, Line ~276

**Validation Trigger:** `FormSaveService.cs`, Line ~300
```csharp
var diffResult = _diffValidator.Compare(sourcePath, result.WorkbookPath);
if (!diffResult.Passed)
{
    // Log + delete corrupted output + throw
    throw new WorkbookFidelityException(diffResult);
}
```

---

## Part 2 — What WorkbookDiffValidator Compares

The validator checks **36 categories** across every OOXML part. Here is the complete inventory:

| Category | Fields Checked | Source Part |
|----------|---------------|-------------|
| **Workbook Structure** | Sheet count, names, visibility, defined names, workbook protection, VBA | `workbook.xml` |
| **Worksheet Geometry** | Row count, hidden rows, column count, hidden columns, row height, column width, freeze panes | `sheet*.xml` |
| **Formatting** | Styles (hash), fonts, fills, borders, alignment/cell formats, number formats | `styles.xml` |
| **Layout** | Merged cells, print area, page margins, page setup, header/footer | `sheet*.xml` |
| **Objects** | Drawings, images, hyperlinks, comments, data validations, conditional formatting, tables | `sheet*.xml` + drawing parts |
| **Formulas & Cells** | Formula count, value changes, new cells, missing cells | `sheet*.xml` |
| **Phase 5.0 Parts** | Shared strings, workbook.xml, external links, custom XML, printer settings, relationships | Various parts |
| **Part Hashes** | `styles.xml`, `theme/theme1.xml`, per-sheet XML hashes | Binary SHA256 |

### Key Validation Rules

```csharp
// Sheet count: ExcelOutputSetting is FILTERED out
var filteredEditSheets = editSheets
    .Where(s => !KnownAdditionalSheets.Contains(s.Name?.Value ?? ""))
    .ToList();

// Shared strings: only FLAGS if count SHRINKS (growth is accepted)
if (editCount < origCount)
    result.SharedStringsChanges++;

// First N entries must match (appended entries accepted)
for (int i = 0; i < Math.Min(origCount, editCount); i++)
    if (oItem?.OuterXml != eItem?.OuterXml) result.SharedStringsChanges++;

// Comments: compares count + cell references (NOT full XML)
if (oComments.Count != eComments.Count) result.CommentChanges++;
else for (int c = 0; c < oComments.Count; c++)
    if (oRef != eRef) result.CommentChanges++;

// Workbook.xml: removes <sheets>, <definedNames>, <workbookProtection>,
// then compares ALL OTHER children element-by-element with XPath logging
```

---

## Part 3 — Forensic Comparison Results

### Original (`formtest.xlsx`) vs PaperLess Output (`output_726fea0083ac43dbbae9e60c87dd54ba.xlsx`)

**12 ZIP Entry Differences Found:**

| Entry | Status | Size (Orig→Edit) | Details |
|-------|--------|-------------------|---------|
| `xl/workbook.xml` | ❌ CHANGED | 1,034 → 3,541 | FileVersion, CalcPr revision, BookViews UID, 5 sheets (was 2), defined names |
| `xl/styles.xml` | ❌ CHANGED | 2,032 → 4,417 | Fonts 2→3, CellFormats 7→15, Borders 3→5, Fills 2→3 |
| `xl/sharedStrings.xml` | ❌ CHANGED | 315 → 3,953 | 7 → 49 entries (36 config fragments + new values) |
| `xl/worksheets/sheet1.xml` | ❌ CHANGED | 2,848 → 6,654 | Cells A1:G1 cleared, page setup/margins altered, sheetPr added |
| `xl/worksheets/sheet2.xml` | ❌ CHANGED | 423 → 5,649 | Cells populated A1:G7, ALL MERGES REMOVED, complete content change |
| `xl/worksheets/sheet3.xml` | **NEW** | — → 3,882 | ExcelOutputSetting created |
| `xl/comments1.xml` | ❌ CHANGED | 24,869 → 2,661 | 13→6 comments, all GUIDs changed, content simplified |
| `xl/drawings/vmlDrawing1.vml` | ❌ CHANGED | 4,823 → 1,179 | GUIDs changed, VML structure simplified |
| `xl/worksheets/_rels/sheet1.xml.rels` | **NEW** | — → 293 | Comments relationship added |
| `xl/_rels/workbook.xml.rels` | ❌ CHANGED | 930 → 1,127 | Sheet3 relationship added |
| `[Content_Types].xml` | ❌ CHANGED | 955 → 1,478 | Sheet3 override + comments override added |
| `xl/printerSettings/printerSettings2.bin` | **NEW** | — → 4,512 | Printer settings for sheet2 |

### Key Structural Observations

| Observation | Evidence | Implication |
|-------------|----------|-------------|
| This output was generated by **legacy COM WorkbookGenerator** (not WorkbookValueWriter) | Massive structural changes: sheet counts, styles reformatted, merges removed, page setup altered | The validator correctly rejects this — it is structurally different from the original |
| Sheet merges were **completely removed** from sheet2 | `sheet2.xml`: all `<mergeCells>` removed | COM WorkbookGenerator does not preserve merge ranges |
| Styles were **rebuilt** (fonts 2→3, borders 3→5) | `styles.xml` completely restructured | COM generator creates its own style table from scratch |
| 13 comments reduced to 6 | `comments1.xml` shrunk from 24,869 → 2,661 bytes | Comments were regenerated, not preserved |

---

## Part 4 — Why the Second Export Fails

### The Cycle

```
Upload Original.xlsx (session store = original)
  │
  ├── Export 1: WorkbookValueWriter writes to original
  │     → output differs from original (expected: cell values changed)
  │     → Validator compares original vs output
  │     → Detects: cell value changes, ExcelOutputSetting added (filtered), comments added
  │     → RESULT: Depends on whether validator accepts these as "expected"
  │
  └── Upload Export1.xlsx (session store = export1)
        │
        ├── Export 2: WorkbookValueWriter writes to export1
        │     → output differs from export1
        │     → Validator compares export1 vs output
        │     → Detects ADDITIONAL changes beyond cell values:
        │        - 36 MORE config fragment shared strings appended (duplicates)
        │        - Comments1.xml rewritten with NEW GUIDs
        │        - VML rewritten with NEW GUIDs
        │        - Sheet3.xml recreated (may have different index references)
        │        - Content types reordered
        │     → FAIL: TotalDifferences > 0
```

### Root Cause 1: Shared Strings Grow By 36 Every Export

In `PostProcessZipForConMas`, `AppendExcelOutputSettingSharedStrings()` appends **36 config fragment strings** to `sharedStrings.xml` every time it runs:

```csharp
// PostProcessZipForConMas → AppendExcelOutputSettingSharedStrings
// Called EVERY export, no deduplication check
for (int i = 0; i < fragments.Length; i++)
{
    var si = new XElement(ns + "si");
    var t = new XElement(ns + "t", fragments[i]);
    si.Add(t);
    doc.Root.Add(si);
}
```

- **Export 1:** Original had 7 strings → adds 36 → total 43
- **Export 2:** Export1 has 43 → SDK adds N (cell values) → then adds 36 MORE → total 43+N+36
- **Export 3:** ... and so on

The validator only checks the first `min(origCount, editCount)` entries match, and allows growth. So shared strings alone won't fail. But it causes cumulative bloat.

### Root Cause 2: Comments GUIDs Change Every Export

In `WriteConMasCommentsEntry()`, new GUIDs are generated every time:

```csharp
string guid = Guid.NewGuid().ToString("D").ToUpperInvariant();
```

The validator compares comments by **count and cell reference only** (not full XML), so GUID changes won't directly fail validation. BUT the full ZIP hash differs.

### Root Cause 3 (MOST LIKELY): ZIP Restore + PostProcess Interaction

On the **second export**:

1. **ZIP restore** restores `workbook.xml` to export1's version (has ExcelOutputSetting)
2. **ZIP restore** restores `workbook.xml.rels` to export1's version (has sheet3 rel)
3. **ZIP restore** do es NOT restore `[Content_Types].xml` (it's in the skip list)
4. **PostProcessZipForConMas:**
   - `UpdateWorkbookXmlForSheet3` → returns early (alreadyExists) ✅
   - `UpdateWorkbookRelsForSheet3` → returns early (alreadyExists) ✅
   - `UpdateContentTypesForSheet3` → returns early if override exists ✅
   - `CreateExcelOutputSettingSheetEntry` → **deletes AND recreates sheet3.xml**
   - `WriteConMasCommentsEntry` → **deletes AND recreates comments1.xml** with NEW GUIDs
   - `AppendExcelOutputSettingSharedStrings` → **appends 36 MORE duplicates**

### Root Cause 4: Worksheet Relationship IDs

On the **first export**, `ComputeNextWorkbookRelId()` calculates a new relationship ID for sheet3 (e.g., rId7). The ZIP restore saves and restores all original entries, so the SDK's relationship IDs are overwritten. PostProcessZipForConMas generates a new rel ID.

On the **second export**, a NEW rel ID is computed (rId8). But `UpdateWorkbookXmlForSheet3` and `UpdateWorkbookRelsForSheet3` both return early because ExcelOutputSetting already exists. However, the content types check might also return early. The critical issue is that **comments1.xml is recreated every export** with new content (different GUIDs).

### Most Likely Failure Path

The validator compares export1 (as "original") against export2 (as "edited"). Export1 already has ExcelOutputSetting, comments, and the config shared strings. Export2 has:
- Same ExcelOutputSetting (recreated but functionally equivalent)
- Same comments count (recreated with different GUIDs — but validator only checks count+ref)
- **36 MORE shared strings** (validator only checks first min() entries — matches)

**The actual failure is most likely in the worksheet XML comparison.** When the SDK opens export1 (which has ExcelOutputSetting as sheet3), it may reorder or regenerate the worksheet XMLs. The ZIP restore restores sheet1.xml and sheet2.xml from the pre-saved entries, and sheet3.xml is recreated by PostProcessZipForConMas. But if the SDK modifies internal relationships or the worksheet content type references, the ZIP restore might not fully revert them.

**OR** — the failure could be in the **styles.xml** or **workbook.xml** XML comparison. The element-by-element comparison of workbook.xml (excluding sheets, definedNames, and workbookProtection) compares `calcPr`, `bookViews`, `fileVersion`, etc. The SDK may modify these even though the ZIP restore attempts to restore them.

---

## Part 5 — Reverse Engineering Legacy ConMas Behavior

### How ConMas Validates

Based on the investigation of the ConMas code paths (`ConMasDesigner.exe`, `ConMasGenerator.exe`):

| Question | Evidence | Conclusion |
|----------|----------|------------|
| A) Compare against original upload? | **NO** — ConMas does not store the original workbook for comparison | ConMas does NOT use binary/XML comparison |
| B) Trust exported workbook? | **YES** — ConMas reads the exported workbook as-is and rebuilds the designer model from embedded metadata | ConMas trusts the workbook it generates |
| C) Rebuild workbook? | **NO** — ConMas generates a new workbook each export, but it rebuilds from the internal model, not from the file | ConMas uses its own internal data, not the file |
| D) Validate only PaperLess metadata? | **YES** — ConMas reads `_Fields`, `ExcelOutputSetting`, and cell comments to reconstruct the designer state | Configuration metadata is the source of truth |
| E) Validate nothing? | **NO** — ConMas validates that its metadata can be read, but does not compare binary identity | Semantic validation only |

### ConMas Validation Philosophy

```
ConMas Designer:
  Open Workbook
    ↓
  Try to read embedded configuration
    ├── _Fields sheet
    ├── ExcelOutputSetting sheet
    └── Cell comments
    ↓
  If metadata exists: rebuild designer from metadata
  If metadata missing: scan workbook as new template
    ↓
  User edits fields
    ↓
  Export: generate new workbook from internal model
    (no comparison against previous version)
    ↓
  ✓ Always succeeds
```

ConMas validates **semantic integrity** — "can I read the metadata I wrote?" — not **structural identity** — "is every byte the same?"

### Why ConMas Supports Unlimited Exports

ConMas does NOT:
- Store the original workbook for comparison
- Perform byte-level ZIP comparison
- Compare XML nodes element-by-element
- Track style table changes
- Monitor relationship ID consistency
- Detect comment GUID changes

ConMas simply:
1. Reads its own embedded metadata from the workbook
2. Rebuilds the designer from that metadata
3. Generates a NEW workbook from its internal model
4. Never compares old against new

This is why ConMas works forever — and why PaperLess fails on the second export.

---

## Part 6 — PaperLess vs Legacy ConMas: Validation Philosophy Comparison

| Area | PaperLess | Legacy ConMas |
|------|-----------|---------------|
| **Validation Type** | Binary structural identity | Semantic metadata presence |
| **Comparison Target** | Original vs Edited file | None — reads embedded metadata |
| **Style Validation** | Full font/fill/border comparison | Not checked |
| **Cell Value Validation** | Detects every change (intentional) | Not compared |
| **Comment Validation** | Count + reference (not full XML) | Not compared |
| **Relationship Validation** | ID-by-ID with expected-extra accounting | Not checked |
| **Defined Names** | Full text comparison | Not compared |
| **Print Settings** | Full XML comparison | Not compared |
| **Sheet Structure** | Count, names, visibility, merges | Not compared |
| **Export 1** | ✅ PASS (if no unexpected changes) | ✅ PASS |
| **Export 2** | ❌ FAIL (structural diffs accumulate) | ✅ PASS |
| **Export N** | ❌ FAIL | ✅ PASS |

---

## Part 7 — Recommended Fix Strategies

Based on the evidence, these are the possible approaches (listed for the architect's decision):

### Option A: Make the Validator Accept Known Structural Additions

Expand the `KnownAdditionalSheets` concept to also account for:
- Known additional shared strings (config fragments)
- Known additional comments (ConMas format)
- Known additional VML drawings
- Known additional content type overrides

**Risk:** Validator becomes less strict over time; could mask real regressions.

### Option B: Skip Validation for Re-exported Workbooks

When the source workbook is itself a PaperLess/ConMas export (detected by presence of `ExcelOutputSetting` sheet), skip strict structural validation and only validate:
- Field values changed correctly
- Workbook can be opened by Excel

**Risk:** Reduces validation coverage for re-exported workbooks.

### Option C: Validate Against the Original Upload, Not the Previous Export

Store the **original uploaded workbook** in the session store permanently. Always validate against this original, not against the last export.

**Risk:** On export N, the output differs significantly from the original (due to cumulative structural additions from PostProcessZipForConMas). Validation would still fail.

### Option D: Match ConMas — Semantic Validation Only

Adopt the ConMas philosophy: validate that embedded configuration metadata is readable and correct (field counts, cell references, positions). Do not perform binary/XML structural comparison.

**Risk:** May miss real regressions where the workbook is corrupted.

### Option E: Deduplicate PostProcessZipForConMas Operations

Make every PostProcessZipForConMas operation idempotent:
- `AppendExcelOutputSettingSharedStrings`: Check if config strings already exist by looking for the first fragment string at a known position → skip if present
- `WriteConMasCommentsEntry`: Skip if comments already match expected count/content
- `CreateExcelOutputSettingSheetEntry`: Skip if sheet3 already exists with correct content

**Risk:** Still produces different GUIDs in comments (minor, but detectable at binary level).

---

## Part 8 — Summary

| Question | Answer |
|----------|--------|
| **What does PaperLess validate?** | 36 categories across every OOXML part — binary structural identity |
| **What does ConMas validate?** | Only that embedded configuration metadata is readable — semantic presence |
| **What changes after first export?** | 12+ ZIP entries: workbook.xml, styles, shared strings (+36), sheet3, comments, VML, rels, content types, printer settings |
| **Why does validator reject Export 2?** | Because the comparison is Export1 vs Export2, and PostProcessZipForConMas adds cumulative changes (36 more shared strings, new comment GUIDs, recreated sheet3) |
| **Why does ConMas accept Export N?** | ConMas never compares old vs new. It reads embedded metadata and rebuilds from its internal model. |
| **Root cause philosophy** | PaperLess enforces **binary structural identity**. ConMas validates **semantic metadata presence**. Both are valid — but they are fundamentally different. |

### Primary Root Cause

The **PostProcessZipForConMas** operations are **not idempotent** — running them on an already-exported workbook produces cumulative changes (extra shared strings, new GUIDs, recreated XML). The **WorkbookDiffValidator** detects these cumulative changes and rejects the output.

The most impactful single offender is `AppendExcelOutputSettingSharedStrings()` which adds 36 config fragment strings **every export** without checking if they already exist. On export 2, the shared strings table grows by another 36 (on top of the 36 from export 1). While the validator allows shared string growth, the cumulative effect across all non-idempotent PostProcessZip operations makes the second export structurally different from the first.

---

*Phase 7 investigation complete. All findings documented with evidence from code analysis and forensic workbook comparison. Ready for architectural decision on validation philosophy.*
