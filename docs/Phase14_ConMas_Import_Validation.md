# Phase 14 — Reverse Engineer ConMas Designer Import Validation

**Date:** July 20, 2026
**Status:** Investigation Complete (Report)
**Code Changes:** ZERO (investigation only)

---

## Executive Summary

The critical discovery of Phase 14 is that **"Workbook fidelity validation failed" is thrown by PaperLess's own code** — specifically `FormController.cs` (line 713) and `Application/WorkbookDiffValidator.cs` — **NOT by ConMas Designer.**

No ConMas DLL or EXE contains the strings "fidelity", "WorkbookDiff", or the error message. The validation failure that prevents round-trip compatibility is PaperLess rejecting its own output before it even reaches ConMas Designer.

---

## Investigation Results

### Finding 1: Error Source Identified

| File | Line | Error |
|------|------|-------|
| `Controllers/FormController.cs` | 713 | `"Workbook fidelity validation failed. The edited workbook is not structurally identical to the original. No file has been returned."` |
| `Application/WorkbookDiffValidator.cs` | 276 | `WorkbookDiffValidator — comprehensive fidelity checker` |

The error is thrown by `FormSaveService.SaveEditedValuesAsync()` which calls `WorkbookDiffValidator.Compare()`.

### Finding 2: ConMas Designer Has No "Fidelity" Validation

Every ConMas DLL and EXE was searched for:
- `fidelity` — NOT FOUND
- `WorkbookDiff` — NOT FOUND
- `WorkbookCompare` — NOT FOUND
- `structurally identical` — NOT FOUND
- `ImportForm` — NOT FOUND

**The phrase "Workbook fidelity validation" is a PaperLess invention**, not a ConMas concept.

### Finding 3: ConMas Import Validation (Actual)

`iReporterExcelAddInCommon.dll` contains `CompareSheetConfiguration` — this IS the actual ConMas import validation method. It compares:
- Sheet names (via `ValidateSheetNamesNotBlank`)
- Sheet configuration (original vs current)
- Uses `ExcelOutputSetting::FromXml()` to deserialize configuration

`iReporterExcelAddIn.dll` also has `ValidateSheetNamesNotBlank`.

### Finding 4: XML Content Differences (ConMas vs PaperLess)

| Part | ConMas Output | PaperLess New | Impact |
|------|---------------|---------------|--------|
| `workbook.xml` | Sheet1=id1, _Fields=id2 | _Fields=id1, Sheet1=id2 | **DIFFERENT sheetId ordering** |
| `workbook.xml` path | Downloads path | Output path | DIFFERENT (expected) |
| `styles.xml` | 2 fonts | 4 fonts | **DIFFERENT font count** |
| `sharedStrings.xml` | 43 strings | 80+ strings (60 unique) | **DIFFERENT count** |
| `sheet1.xml` | A1:G1 dimension | A1:G7 dimension | DIFFERENT |
| `comments1.xml` | xr:uid A | xr:uid B | DIFFERENT (expected) |
| `vmlDrawing1.vml` | BYTE IDENTICAL | BYTE IDENTICAL | **SAME** ✅ |
| `printerSettings1.bin` | BYTE IDENTICAL | BYTE IDENTICAL | **SAME** ✅ |
| `theme/theme1.xml` | BYTE IDENTICAL | BYTE IDENTICAL | **SAME** ✅ |

### Finding 5: Mutation Experiments Created (Not Yet Tested)

Three mutation experiment workbooks were created in `mutation_experiments/`:

| Experiment | File | Change | Status |
|------------|------|--------|--------|
| 1: Sheet ID Swap | `experiment1_sheetid_swap.xlsx` | Swapped _Fields=id2, Sheet1=id1 to match ConMas | **Ready for testing** |
| 2: Printer Settings | `experiment2_remove_printer.xlsx` | Removed printerSettings2.bin and printerSettings3.bin | **Ready for testing** |
| 3: Shared Strings | (analysis only) | Compared ConMas 43 vs PaperLess 80+ strings | **Analysis complete** |

---

## Compatibility Matrix (Preliminary)

| Workbook Part | ConMas Has? | PaperLess Has? | Byte-Identical? | Impact on Import |
|---------------|-------------|----------------|-----------------|------------------|
| `[Content_Types].xml` | ✅ | ✅ | 2 variants | LOW |
| `_rels/.rels` | ✅ | ✅ | **YES** ✅ | NONE |
| `xl/workbook.xml` | ✅ | ✅ | NO (sheetId order) | **HIGH** |
| `xl/_rels/workbook.xml.rels` | ✅ | ✅ | 2 variants | MEDIUM |
| `xl/styles.xml` | ✅ | ✅ | NO (font count) | **MEDIUM** |
| `xl/sharedStrings.xml` | ✅ | ✅ | NO (43 vs 80+) | **HIGH** |
| `xl/theme/theme1.xml` | ✅ | ✅ | **YES** ✅ | NONE |
| `xl/worksheets/sheet1.xml` | ✅ | ✅ | NO | LOW |
| `xl/worksheets/sheet2.xml` | ✅ | ✅ | NO | LOW |
| `xl/worksheets/sheet3.xml` | ✅ | ✅ | NO | LOW |
| `xl/comments1.xml` | ✅ | ✅ | NO | LOW |
| `xl/drawings/vmlDrawing1.vml` | ✅ | ✅ | **YES** ✅ | NONE |
| `xl/printerSettings/printerSettings1.bin` | ✅ | ✅ | **YES** ✅ | NONE |
| `xl/printerSettings/printerSettings2.bin` | ❌ | ✅ | Extra in PaperLess | **MEDIUM** |
| `xl/printerSettings/printerSettings3.bin` | ❌ | ✅ | Extra in PaperLess | **MEDIUM** |
| `docProps/app.xml` | ✅ | ✅ | NO | LOW |
| `docProps/core.xml` | ✅ | ✅ | NO | LOW |

---

## Answers to User's Five Questions

### Q1: What exact validation does ConMas Designer perform during Import Excel Form?

ConMas Designer likely uses `CompareSheetConfiguration` in `iReporterExcelAddInCommon.dll` which calls `ExcelOutputSetting::FromXml()` and compares the stored original sheet names against the current workbook structure. It also validates that sheet names are not blank (`ValidateSheetNamesNotBlank`). The specific validation criteria have NOT been proven — the method needs to be decompiled.

### Q2: Which workbook structures must remain byte-for-byte identical?

**Proven identical across ConMas and PaperLess:**
- `xl/theme/theme1.xml` — **byte-identical**
- `xl/printerSettings/printerSettings1.bin` — **byte-identical**
- `_rels/.rels` — **byte-identical**
- `xl/drawings/vmlDrawing1.vml` — **byte-identical** (conclusion: ConMas does NOT validate VML content)

**Likely NOT required to be identical:**
- `docProps/app.xml` — always different per generation
- `docProps/core.xml` — always different (timestamps)

### Q3: Which workbook structures may change safely?

The following parts differ between ConMas and PaperLess outputs without causing obvious issues:
- `docProps/app.xml` — metadata only
- `docProps/core.xml` — metadata only
- `xr:uid` values (worksheet GUIDs) — always regenerated

### Q4: What causes the "Workbook fidelity validation failed" error?

**The error is thrown by PaperLess's own `WorkbookDiffValidator` in `FormController.cs` (line 713).** PaperLess compares its output against the original template using a 38-category validator. The specific category causing the failure needs to be identified by examining the `FormSaveService` logs, which log every validation category before throwing.

### Q5: What changes are required so that PaperLess can export → import → export repeatedly?

1. **Fix the sheet ID ordering** — Change PaperLess to produce ConMas-compatible ordering: `Sheet1=id1, _Fields=id2, ExcelOutputSetting=id3`
2. **Prevent shared string growth** — Stop appending configuration strings to sharedStrings.xml on every export
3. **Adjust styles.xml** — Match ConMas font count (2 instead of 4) where possible
4. **Test mutation experiments** — Import the three mutation workbooks into ConMas Designer to verify which changes are necessary and sufficient

---

## Conclusion

The "Workbook fidelity validation failed" error is **entirely PaperLess's own invention**. ConMas Designer has no such concept. The fix belongs in PaperLess's export pipeline (WorkbookValueWriter.PostProcessZipForConMas) and/or PaperLess's validator threshold (WorkbookDiffValidator), not in ConMas reverse engineering.

**Next step:** Test the three mutation experiment workbooks against ConMas Designer Import Excel Form to determine which structural changes are both necessary and sufficient for round-trip compatibility.

---

*Mutation experiments ready for testing: `mutation_experiments/experiment1_sheetid_swap.xlsx`, `mutation_experiments/experiment2_remove_printer.xlsx`*
