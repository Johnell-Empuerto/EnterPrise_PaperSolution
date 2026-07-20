# Phase 5.5.4 — rId Mismatch Fix & Single-Compute Refactoring

**Date:** 2026-07-20  
**Status:** Fix Applied — Build Succeeded (0 errors)  
**Next:** Generate workbook → Open in Excel → Iterate

---

## Root Cause Fixed

The ExcelOutputSetting sheet in `xl/workbook.xml` used `id=rId3`, but `xl/_rels/workbook.xml.rels` mapped `rId3 → theme/theme1.xml` (not a worksheet). The actual sheet3.xml was at `rId6`.

Excel detected: "workbook.xml says ExcelOutputSetting uses rId3, but rId3 is theme/theme1.xml" → repair mode.

## Fix Applied

### New Helper: `ComputeNextWorkbookRelId(ZipArchive pkg)`

Single source of truth for relationship ID computation. Scans ALL relationships in `xl/_rels/workbook.xml.rels` and returns `rId{max+1}`.

### Refactored: Single-Compute Pattern (per user's explicit requirement)

```
PostProcessZipForConMas
    ↓
string worksheetRelId = ComputeNextWorkbookRelId(pkg);
    ↓
UpdateWorkbookXmlForSheet3(pkg, worksheetRelId);
UpdateWorkbookRelsForSheet3(pkg, worksheetRelId);
```

Both methods now receive the **exact same** relId from one computation.

### Code Removed

- `UpdateWorkbookXmlForSheet3`: Removed all rId computation logic (previously: scan rels, fallback to sheet elements, compute rId from maxRelId+1)
- `UpdateWorkbookRelsForSheet3`: Removed the maxId scanning loop

## Build Verification

- **0 errors**, 0 new warnings

## Code Review

✅ Relationship ID computed once in `PostProcessZipForConMas`  
✅ Same `worksheetRelId` passed to both methods  
✅ Both methods use the parameter directly  
✅ No latent fallback bugs remain  
✅ Safe edge case handling (missing rels file → returns rId1)

## Next Steps

1. **Run the application** and generate a fresh workbook through the browser
2. **Open in Microsoft Excel** — check if repair dialog appears
3. **If repair persists**: Extract `xl/recoveryLog.xml` from the repaired workbook and identify the next structural issue
4. **Run forensic three-way comparison** again to identify any remaining differences between generated and ConMas output
5. **Iterate** until workbook opens without repair and structure matches ConMas
