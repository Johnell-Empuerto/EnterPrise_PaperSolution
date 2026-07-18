# Phase X.8 — Elimination Test Workbooks: What Was Done

**Date:** 2026-07-17  
**Status:** All 12 test workbooks created successfully  
**Test location:** `C:\Users\MCF-JOHNELLEEMPUERTO\Downloads\elimination_tests\`

---

## Summary

Phase X.8 created **12 controlled elimination test workbooks** by manipulating the OOXML structure of the legacy ConMas Output Excel workbook (`FormTest - Copy.xlsx`). Each test modifies exactly **one variable** to isolate what the real ConMas runtime actually depends on when importing a workbook.

## What Was Created

### Files Created

| File | Description |
|------|-------------|
| `create_elimination_tests.py` | Python script that generates test workbooks via OOXML manipulation |
| `docs/PhaseX8_Elimination_Test_Workbooks_Report.md` | Instructions and evidence table for testing |

### 12 Test Workbooks

| Test | File | Modified Feature | Size |
|------|------|-----------------|------|
| A | `TestA_Original_Baseline.xlsx` | Original (unchanged) — baseline | 16,782 bytes |
| B | `TestB_NoComments.xlsx` | **Removed all 6 comments** + VML drawings | 12,246 bytes |
| C | `TestC_NoFieldsSheet.xlsx` | **Removed `_Fields` worksheet** | 13,307 bytes |
| D | `TestD_NoExcelOutputSetting.xlsx` | **Removed `ExcelOutputSetting` worksheet** | 13,040 bytes |
| E | `TestE_NoPrintArea.xlsx` | **Removed `_xlnm.Print_Area`** defined name | 13,833 bytes |
| F | `TestF_SimplifiedComments.xlsx` | Comments replaced with **minimal text** (field name only) | 13,793 bytes |
| G | `TestG_NoMergedCells.xlsx` | **Removed all 5 merged regions** | 13,849 bytes |
| H | `TestH_OnlySheet1.xlsx` | **Only Sheet1** (`_Fields` + `ExcelOutputSetting` removed) | 12,389 bytes |
| I | `TestI_OurCommentFormat.xlsx` | Comments replaced with **our format** (`Name\nType\n\n0`) | 13,808 bytes |
| J | `TestJ_NoFields_KeepExcelOutput.xlsx` | **Removed only `_Fields`**, kept ExcelOutputSetting | 13,307 bytes |
| K | `TestK_NoDefinedNames.xlsx` | **Removed all defined names** from workbook.xml | 13,933 bytes |
| L | `TestL_OurFormat_WithClusterIndex.xlsx` | Our format **+ cluster index** added back | 13,821 bytes |

### Key Technical Fixes

The Python script had to be debugged for proper OOXML manipulation:

1. **`remove_sheet`**: Fixed from `wb.remove(element)` to `sheets_elem.remove(element)` — sheet elements are children of `<sheets>`, not direct children of `<workbook>` root.
2. **`remove_print_area` / `remove_defined_names`**: Same fix — defined names are children of `<definedNames>`, not the root.

## How to Use

1. Open **ConMas Designer** from `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer`
2. For each test workbook (A through K):
   - Use the import/upload function
   - Record: number of fields reconstructed, error messages, UI behavior
3. Report results back to fill the **dependency matrix**

### What Each Test Answers

| Tests to Compare | Question Answered |
|-----------------|-------------------|
| A vs B | Are comments required? |
| A vs C | Is `_Fields` required? |
| A vs D | Is `ExcelOutputSetting` required? |
| A vs E | Is `PrintArea` required? |
| A vs F vs I | Which comment format lines are required? |
| A vs G | Are merged cells required? |
| A vs H | Can ConMas work with only a content sheet? |
| A vs J | Is `_Fields` read when ExcelOutputSetting exists? |
| A vs K | Are defined names required? |
| A vs L | Does adding cluster index to our format help? |
