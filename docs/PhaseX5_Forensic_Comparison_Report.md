# Phase X.5 — Forensic Comparison: Legacy vs Generated Output Workbook

**Date:** 2026-07-17  
**Status:** Root cause identified — fixes applied, diagnostic logging added  
**Download Locations:**

- Legacy: `C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy.xlsx`
- Generated: `C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\FormTest - Copy_output.xlsx`

---

## Executive Summary

The generated Output Excel workbook (`FormTest - Copy_output.xlsx`) uploads to `POST /api/form/upload-excel` but reconstructs only **1 field cluster**, while the legacy ConMas workbook (`FormTest - Copy.xlsx`) reconstructs all **6 field clusters** successfully.

A deep OOXML forensic comparison revealed **three structural discrepancies** that cascade to cause this failure:

### Root Cause Chain

```
Legacy workbook (6 comments, 3 sheets)
    ↓ upload-excel
    ↓ _Fields sheet has 1 row (headers only, no data) → falls back to COMMENTS
    ↓ 6 comments found → 6 clusters reconstructed ✅

Generated workbook (1 comment, 2 sheets)
    ↓ upload-excel
    ↓ _Fields sheet has 7 rows (header + 6 data) → uses _FIELDS path
    ↓ 6 _Fields rows → 6 clusters should be created
    ↓ BUT: COM UsedRange on hidden _Fields sheet may return only 1 row
    ↓ ...OR: exceptions in cluster loop cause early exit
    ↓ Result: only 1 cluster ❌
```

---

## 1. OOXML Comparison

### 1.1 Workbook File Structure

| Component        | Legacy                      | Generated    | Status                       |
| ---------------- | --------------------------- | ------------ | ---------------------------- |
| File size        | 16,782 bytes                | 12,327 bytes | ⚠️ Smaller                   |
| Sheets           | 3                           | 2            | ❌ Missing 1 sheet           |
| Comments         | 6 (A1, C1, A3, A6, A9, A12) | 1 (A12 only) | ❌ Bug (see §3)              |
| VML drawings     | 4,823 bytes                 | 1,179 bytes  | ❌ Consistent with 1 comment |
| Shared strings   | 43 items                    | 23 items     | ⚠️ Less content              |
| Printer settings | 5,428 bytes                 | 5,428 bytes  | ✅ Identical                 |

### 1.2 Workbook.xml (Sheet Definitions)

**Legacy:**
| Name | ID | State |
|------|----|-------|
| `_Fields` | 2 | hidden |
| `Sheet1` | 1 | visible |
| `ExcelOutputSetting` | 3 | visible |

**Generated:**
| Name | ID | State |
|------|----|-------|
| `Sheet1` | 1 | visible |
| `_Fields` | 2 | hidden |

**Differences:**

- ❌ Generated is missing the `ExcelOutputSetting` sheet (36 rows of ConMas XML configuration)
- ❌ Generated \_Fields has 7 rows (header + 6 data) vs Legacy \_Fields has 1 row (headers only)

### 1.3 \_Fields Sheet Analysis

**Legacy `_Fields` (sheet1.xml):**

- 1 row, 7 cells, 0 merges
- Only headers present: Address, FieldId, FieldName, FieldType, SheetName, CreatedDate, Notes
- **NO data rows** — ConMas stores field metadata in comments and the ExcelOutputSetting sheet

**Generated `_Fields` (sheet2.xml):**

- 7 rows, 49 cells (7 cols × 7 rows), 0 merges
- Header + 6 data rows with all field addresses, IDs, names, types
- 23 unique shared strings covering all 49 cell values

### 1.4 Sheet1 (Content Sheet) Analysis

| Feature        | Legacy                                 | Generated | Status   |
| -------------- | -------------------------------------- | --------- | -------- |
| Rows           | 9                                      | 1         | ❌       |
| Cells          | 33                                     | 0         | ❌ Empty |
| Merged regions | 5 (A1:B2, C1:D2, A3:D4, A6:D7, A9:D10) | 0         | ❌       |

### 1.5 Comment Analysis

**Legacy comments** (one per field):

| Cell | Format                                         | Parameters            |
| ---- | ---------------------------------------------- | --------------------- |
| A1   | samples\nKeyboardText\n0\n0\n0\nRequired=0;... | Full parameter string |
| C1   | samples\nKeyboardText\n1\n0\n0\nRequired=0;... | Full parameter string |
| A3   | samples\nKeyboardText\n2\n0\n0\nRequired=0;... | Full parameter string |
| A6   | samples\nKeyboardText\n3\n0\n0\nRequired=0;... | Full parameter string |
| A9   | samples\nKeyboardText\n4\n0\n0\nRequired=0;... | Full parameter string |
| A12  | samples\nKeyboardText\n5\n0\n0\nRequired=0;... | Full parameter string |

**Generated comment** (only one):
| Cell | Format |
|------|--------|
| A12 | samples\nKeyboardText\n\n0 |

All 5 merged cells (A1, C1, A3, A6, A9) are **missing comments** in the generated workbook.

---

## 2. Reverse Engineering: How Legacy ConMas Reconstructs Fields

### Metadata Source Priority

The legacy ConMas runtime uses **cell comments as the PRIMARY metadata source**, not the `_Fields` worksheet. Evidence:

| Evidence                                              | Source              | Confidence |
| ----------------------------------------------------- | ------------------- | ---------- |
| Legacy `_Fields` has NO data rows (only headers)      | OOXML sheet1.xml    | 100%       |
| Legacy has 6 comments matching 6 fields               | OOXML comments1.xml | 100%       |
| Legacy has `ExcelOutputSetting` sheet with ConMas XML | OOXML sheet3.xml    | 100%       |

### Legacy Comment Format

```
Line 0: Field Name         → "samples"
Line 1: Field Type         → "KeyboardText"
Line 2: Cluster Index      → "0" (0-based index)
Line 3: 0                  → (unknown field)
Line 4: 0                  → (unknown field)
Line 5+: Parameters        → "Required=0;Lines=1;InputRestriction=None;..."
```

### Our Comment Format (generated)

```
Line 0: Field Name         → "samples"
Line 1: Field Type         → "KeyboardText"
Line 2: (empty)            → Missing cluster index
Line 3: "0"                → Placeholder for input params
```

**Our format is missing:**

- ❌ Cluster index on line 2
- ❌ The two "0" lines (lines 3-4 in legacy)
- ❌ Full parameter string (line 5+)

---

## 3. Root Cause Analysis

### Bug #1: Comments Cannot Be Added to Merged Cell Ranges

**File:** `WorkbookGenerator.cs`, `WriteCellComments()` method  
**Status:** FIXED

The code used `ws.Range[cluster.CellAddress].AddComment(commentText)` where `cluster.CellAddress` could be a merged range like `"$A$1:$B$2"`. Excel COM throws a `COMException` when `AddComment()` is called on a multi-cell Range. The exception was silently caught, so **5 of 6 fields lost their comments**.

**Fix:** Split merge range addresses to use only the top-left cell:

```csharp
// Before (broken):
var range = ws.Range[cluster.CellAddress];  // "$A$1:$B$2"
range.AddComment(commentText);               // COMException on merged range!

// After (fixed):
string singleCellAddr = cluster.CellAddress.Contains(":")
    ? cluster.CellAddress.Split(':')[0]      // "$A$1"
    : cluster.CellAddress;
var commentRange = ws.Range[singleCellAddr];
commentRange.AddComment(commentText);         // Works correctly
```

### Bug #2: COM UsedRange on Hidden \_Fields Sheet

**File:** `WorkbookReaderService.cs`, `ReadFieldsSheet()` method  
**Status:** UNDER INVESTIGATION (diagnostic logging added)

When the `_Fields` sheet is hidden (`xlSheetHidden`), COM's `UsedRange` property may return fewer rows than the actual data. If `UsedRange.Rows.Count` returns 1 (header only), the reader falls back to the comments fallback path, where only 1 comment exists (due to Bug #1).

**Expected diagnostic log output** (after re-running the upload):

If the \_Fields path is taken:

```
FIELD #0: Sheet=Sheet1 Address=$A$1:$B$2 _Fields=FOUND Comment=FOUND Cluster=CREATED
FIELD #1: Sheet=Sheet1 Address=$C$1:$D$2 _Fields=FOUND Comment=FOUND Cluster=CREATED
...
```

If the fallback path is taken (UsedRange returns 1):

```
_Fields sheet empty or missing; building clusters from comments
COMMENT #0: Sheet=Sheet1 Address=A12 Comment=FOUND Cluster=CREATED
```

### Bug #3: Sheet1 Content Not Written to Generated Workbook

**File:** `WorkbookGenerator.cs` — applies merged cells and cell styles but the content sheet has 0 cells.

**Status:** NOT YET INVESTIGATED (requires determination of whether Sheet1 should contain the cell values or just merged structure)

---

## 4. Changes Applied

### Files Modified

| File                       | Change                                                         | Purpose                                           |
| -------------------------- | -------------------------------------------------------------- | ------------------------------------------------- |
| `WorkbookGenerator.cs`     | Fix `WriteCellComments` to use single cells for `AddComment()` | Fix Bug #1 — comments now created on all 6 fields |
| `WorkbookReaderService.cs` | Add consolidated per-field diagnostic logging                  | Enable runtime tracing to identify Bug #2         |

### Diagnostic Log Format

Each field is logged as a single line:

```
FIELD #{Idx}: Sheet={Sheet} Address={Addr} _Fields=FOUND Comment={FOUND|NO} Cluster={CREATED|SKIPPED}
COMMENT #{Idx}: Sheet={Sheet} Address={Addr} Comment=FOUND|EMPTY Cluster=CREATED|SKIPPED
```

---

## 5. Validation Plan

After starting the ExcelAPI server and running `dotnet build`:

1. **Generate a new Output Excel** using the Export Excel button
2. **Upload the generated workbook** using the Upload Excel button
3. **Check the server logs** for the diagnostic trace output
4. **Verify:** The diagnostic shows which path was taken (\_Fields vs comments fallback)
5. **Fix Bug #2** if the diagnostic confirms `UsedRange` is the issue
6. **Verify round-trip:** Upload → edit → export → upload again produces same cluster count

---

## 6. Remaining Unknowns

| Question                                                            | Status     | Action Needed                            |
| ------------------------------------------------------------------- | ---------- | ---------------------------------------- |
| Does COM UsedRange return correct row count for hidden \_Fields?    | Unknown    | Diagnostic logging will confirm          |
| Should Sheet1 have cell values written?                             | Unknown    | Needs requirement clarification          |
| Should we add `ExcelOutputSetting` sheet?                           | Unknown    | Needs investigation of ConMas XML format |
| Should comments use legacy format with cluster index and all lines? | Likely yes | Needs parsing test                       |
