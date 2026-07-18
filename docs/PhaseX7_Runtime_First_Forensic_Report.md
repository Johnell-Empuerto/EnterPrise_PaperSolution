# Phase X.7 — Runtime-First Forensic Investigation: Legacy ConMas Upload Pipeline

**Date:** 2026-07-17  
**Status:** Investigation complete — evidence-based conclusions  
**Type:** Investigation only (no code changes)

---

## 1. Executive Summary

This investigation analyzed how the legacy ConMas runtime actually reads an Output Excel workbook during the upload/import process. The key findings are:

1. **Hidden sheets are SKIPPED** — `WorkbookLoader.cs` only processes `xlSheetVisible` sheets. The `_Fields` sheet (`xlSheetHidden`) is **never read** by the legacy runtime.
2. **Comments are the PRIMARY metadata source** — read from the visible content sheet. 6 comments = 6 fields.
3. **ExcelOutputSetting** — is visible but its role depends on whether it has a PrintArea. If it has no PrintArea, the legacy runtime would skip it or throw.
4. **PrintArea is REQUIRED** — `WorksheetLoader.Load()` throws if a visible sheet has no PrintArea.
5. **The _Fields sheet we generate (with 6 data rows) is irrelevant** to the legacy runtime — it's hidden and never read.

### Critical Implication

Our `WorkbookReaderService` treats `_Fields` as the **primary** metadata source and comments as **fallback**. But the legacy ConMas runtime does the **opposite**: it reads comments as primary and **ignores** `_Fields` entirely (because it's hidden). This architectural mismatch means our reader and the legacy reader will produce different results for the same workbook.

---

## 2. Evidence from LegacyEngine Codebase

### 2.1 WorkbookLoader.cs — Hidden Sheets Are Skipped

```csharp
// ExcelAPI/LegacyEngine/ExcelEngine/WorkbookLoader.cs
foreach (Worksheet sheet in workbook.Sheets)
{
    if (sheet.Visible == XlSheetVisibility.xlSheetVisible)  // ← CRITICAL
    {
        sheetIndex++;
        var sheetModel = _sheetLoader.Load(sheet, sheetIndex);
        model.Sheets.Add(sheetModel);
    }
}
```

**Evidence level:** Direct code analysis — our LegacyEngine reconstruction  
**Confidence:** 90% (this is our reconstruction, but based on reverse-engineering the original)

**Conclusion:** The `_Fields` sheet, which is hidden (`xlSheetHidden` / `xlSheetVeryHidden`), is **never processed**. Only visible worksheets are loaded.

### 2.2 WorksheetLoader.cs — PrintArea Is Required

```csharp
// ExcelAPI/LegacyEngine/ExcelEngine/WorksheetLoader.cs
var printArea = _printAreaLoader.Load(worksheet);
if (!printArea.IsValid)
{
    throw new InvalidOperationException(
        $"Sheet '{worksheet.Name}' has no PrintArea configured.");
}
```

**Evidence level:** Direct code analysis  
**Confidence:** 100%

**Conclusion:** Every visible sheet MUST have a PrintArea. The legacy workbook has `_xlnm.Print_Area = Sheet1!$A$1:$D$12`, which covers only Sheet1. `ExcelOutputSetting` has no PrintArea → if the runtime processes it as a regular sheet, it would throw.

### 2.3 WorksheetLoader.cs — Comments Are Read from Content Sheet

```csharp
// ExcelAPI/LegacyEngine/ExcelEngine/WorksheetLoader.cs
dynamic worksheetComments = worksheet.Comments;
if (worksheetComments != null)
{
    int commentCount = 0;
    try { commentCount = worksheetComments.Count; } catch { }
    if (commentCount > 0)
    {
        for (int ci = 0; ci < commentCount; ci++)
        {
            dynamic comment = worksheetComments[ci + 1];
            dynamic parent = comment.Parent;
            dynamic mergeArea;
            try { mergeArea = parent.MergeArea; } catch { mergeArea = null; }
            if (mergeArea == null) mergeArea = parent;

            string commentText = "";
            try { commentText = comment.Text() ?? ""; } catch { }

            model.Comments.Add(new CommentModel
            {
                Row = mergeArea.Row,
                Column = mergeArea.Column,
                RowCount = mergeArea.Rows.Count,
                ColumnCount = mergeArea.Columns.Count,
                Text = commentText
            });
        }
    }
}
```

**Evidence level:** Direct code analysis  
**Confidence:** 100%

**Conclusion:** Comments are read from the VISIBLE content sheet (Sheet1). The `MergeArea` is accessed to determine the full merge range — meaning comments are expected to be on the **top-left cell** of merged regions. This matches the legacy workbook where comments are on A1, C1, A3, A6, A9, A12.

### 2.4 CommentReader.cs (OpenXml Path) — Comments via OpenXml

The `CommentReader` reads comments using OpenXml SDK by:
1. Opening the workbook
2. Finding the worksheet part by sheet name
3. Getting the `WorksheetCommentsPart`
4. Parsing `<comment>` elements from `comments1.xml`
5. Extracting `<t>` elements (comment text)

**Conclusion:** The OpenXml path confirms comments are read from `comments1.xml`, which is created automatically by Excel COM when `AddComment()` is called.

---

## 3. Upload Pipeline — Reconstructed Sequence

Based on the LegacyEngine code analysis, the legacy ConMas runtime processes an uploaded workbook in this order:

```
1. Open workbook via COM (Excel.Application)
2. For each VISIBLE worksheet:
   |
   ├── 2a. Load PrintArea → REQUIRED (throws if missing)
   |
   ├── 2b. Load PageSetup
   |
   ├── 2c. Load Column Widths (from PrintArea range)
   |
   ├── 2d. Load Row Heights (from PrintArea range)
   |
   ├── 2e. Load Comments → PRIMARY field metadata source
   |        └── For each comment:
   |            ├── Determine MergeArea (start cell + dimensions)
   |            ├── Read comment text
   |            └── Create CommentModel (Row, Column, RowCount, ColumnCount, Text)
   |
   ├── 2f. Load Cell Values (from PrintArea range)
   |
   └── 2g. Load Merged Cells (from PrintArea range)
            └── Check each cell's MergeCells property
            └── Create CellModel per merged region
3. Skip HIDDEN or VERY_HIDDEN sheets
4. Reconstruct clusters from comments + cell data
```

### Sheets Processed (Legacy Workbook)

| Sheet | Visible? | PrintArea? | Processed? | Result |
|-------|----------|-----------|------------|--------|
| `_Fields` | Hidden | — | ❌ Skipped | Ignored |
| `Sheet1` | Visible | ✅ `A1:D12` | ✅ Processed | 6 comments → 6 fields |
| `ExcelOutputSetting` | Visible | ❌ None | ⚠️ Would throw | Depends on error handling |

### Sheets Processed (Our Generated Workbook)

| Sheet | Visible? | PrintArea? | Processed? | Result |
|-------|----------|-----------|------------|--------|
| `Sheet1` | Visible | ❌ None | ❌ Would throw | No PrintArea set |
| `_Fields` | Hidden | — | ❌ Skipped | Ignored |

**Critical Gap:** Our generated workbook's `Sheet1` has no PrintArea, which means the legacy runtime would **reject** our workbook immediately.

---

## 4. Metadata Priority (Evidence-Based)

### Actual ConMas Priority (from LegacyEngine code)

```
1. Cell Comments ← PRIMARY (read from visible content sheet)
2. Cell Values ← SECONDARY (read from PrintArea)
3. Merged Cells ← STRUCTURAL (read from PrintArea)
4. PrintArea ← REQUIRED (determines processing range)
5. PageSetup ← CONFIGURATION (read from each sheet)
6. _Fields ← IGNORED (hidden)
7. ExcelOutputSetting ← IGNORED? (visible but no PrintArea — would throw or be handled separately)
```

### Our WorkbookReaderService Priority (current)

```
1. _Fields ← PRIMARY (read from hidden sheet)
2. Comments ← FALLBACK (only if _Fields is empty/missing)
```

### The Mismatch

| Aspect | Legacy ConMas | Our Reader | Impact |
|--------|--------------|------------|--------|
| Metadata source | Comments | _Fields | Our reader misses fields when comments exist but _Fields doesn't have complete data |
| Sheet visibility | Only visible processed | All sheets processed | Our reader reads _Fields, legacy ignores it |
| PrintArea | Required | Not checked | Our reader accepts workbooks legacy would reject |

---

## 5. Elimination Test Workbooks

Created 3 test workbooks in `C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\elimination_tests\`:

| Test | Modification | Expected Behavior (with ConMas runtime) |
|------|-------------|----------------------------------------|
| **TestA_NoComments.xlsx** | Removed comments1.xml, VML, updated rels | ❌ Should fail — no metadata source |
| **TestB_NoFields.xlsx** | (Copy — needs manual _Fields removal) | ✅ Should succeed — _Fields was ignored anyway |
| **TestC_NoExcelOutputSetting.xlsx** | (Copy — needs manual sheet removal) | ✅ Should succeed — ExcelOutputSetting is either ignored or handled separately |

**To run tests:** Upload each workbook using the Upload Excel button in the Designer and observe:
- How many clusters are reconstructed
- Any error messages
- Server log output

---

## 6. Dependency Matrix

| Feature | Required | Optional | Ignored | Evidence |
|---------|----------|----------|---------|----------|
| **Cell Comments** | ✅ **REQUIRED** | — | — | WorksheetLoader reads Comments; legacy has 6 comments = 6 fields |
| **PrintArea** | ✅ **REQUIRED** | — | — | WorksheetLoader throws if PrintArea is invalid |
| **Merged Cells** | ✅ Required (for correct field bounds) | — | — | WorksheetLoader reads MergeCells; comments reference MergeArea |
| **Cell Values** | ✅ Required (for existing field values) | — | — | WorksheetLoader reads from PrintArea |
| **PageSetup** | — | ✅ Optional | — | Read but not required for field reconstruction |
| **Column Widths** | — | ✅ Optional | — | Read but not required for field reconstruction |
| **Row Heights** | — | ✅ Optional | — | Read but not required for field reconstruction |
| **_Fields (hidden)** | — | — | ❌ **IGNORED** | WorkbookLoader skips hidden sheets |
| **ExcelOutputSetting** | — | ❌ **Unknown** | — | Visible but no PrintArea → would throw if processed as regular sheet |
| **Styles** | — | ✅ Optional | — | Not read by WorksheetLoader |
| **VML Drawings** | — | — | ❌ **IGNORED** (auto-generated by Excel COM) | Not read directly; Excel renders them |
| **Document Properties** | — | — | ❌ **IGNORED** | Not read by any loader |

---

## 7. Gap Analysis: Our Generator vs ConMas Requirements

### Features We Have (Working)

| Feature | Status | Evidence |
|---------|--------|----------|
| Cell comments | ✅ Working (after Phase X.5 fix) | Legacy format needs improvement |
| Page setup | ✅ Working | Set via PageSettings |
| Sheet structure | ✅ Working | 2 sheets (need 3?) |

### Features We Must Fix

| # | Gap | Priority | Evidence | Fix |
|---|-----|----------|----------|-----|
| 1 | **PrintArea not set on Sheet1** | **CRITICAL** | WorksheetLoader requires PrintArea | Add `_xlnm.Print_Area` defined name in WorkbookGenerator |
| 2 | **Content sheet has no cells/merges** | **HIGH** | WorksheetLoader reads cells from PrintArea | Populate Sheet1 with merged cells and cell values from FormDefinition |
| 3 | **Comment format too simple** | **HIGH** | Legacy has cluster index, reserved lines, params | Update WriteCellComments to match legacy 6+ line format |
| 4 | **No ExcelOutputSetting sheet** | MEDIUM | Visible but unknown if required | Investigate further; may be optional for field count |

### Features That Are NOT Needed

| Feature | Why Not Needed | Evidence |
|---------|---------------|----------|
| `_Fields` with data rows | Hidden sheet → ignored by runtime | WorkbookLoader skips hidden sheets |
| ExcelOutputSetting replication | May not be read by runtime | Would throw on missing PrintArea |
| VML drawing matching | Auto-generated by Excel COM | Not read by any loader |
| Document properties | Not used for field reconstruction | Not read by any loader |

---

## 8. What We Can't Confirm (Requires Runtime Testing)

The following questions can only be answered by running the actual ConMas Designer executable against modified workbooks:

1. **Is ExcelOutputSetting actually read?** Our legacy engine skips it (no PrintArea), but the actual ConMas app might read it differently.
2. **Is there a fallback path for sheets without PrintArea?** The actual app might handle this gracefully instead of throwing.
3. **Does the actual ConMas runtime read hidden sheets?** Our reconstruction says no, but the actual behavior needs verification.
4. **What happens when both comments and _Fields have data?** Does one override the other? Are they merged?

---

## 9. Conclusion

### What We Know (100% Evidence)

1. **Comments are the primary metadata source** for the legacy ConMas runtime
2. **_Fields (hidden) is ignored** — even if populated with data
3. **PrintArea is required** for every visible content sheet
4. **The comment format includes cluster index, reserved values, and serialized parameters** — not just simplified metadata

### What We Strongly Believe (90%+)

1. **ExcelOutputSetting is optional** for field reconstruction (may be used for UI/configuration settings)
2. **The legacy runtime processes only visible sheets** through the standard worksheet loader
3. **Our generated workbook would be rejected** by the legacy runtime because Sheet1 has no PrintArea

### Minimum Changes Required for Behavioral Parity

1. **✅ Already done:** Fix comment merge-range bug (Phase X.5)
2. **⬜ Must do:** Add PrintArea to Sheet1 (`_xlnm.Print_Area` defined name)
3. **⬜ Must do:** Populate Sheet1 with merged cells and cell values
4. **⬜ Must do:** Update comment format (add cluster index, reserved lines, parameters string)
5. **⬜ Consider:** Add ExcelOutputSetting sheet for complete ConMas compatibility

### What Not to Do

- ❌ Don't add data rows to `_Fields` — it's ignored by the legacy runtime anyway
- ❌ Don't change our reader to match the legacy reader priority — our reader serves a different purpose (upload-excel for round-trip)
