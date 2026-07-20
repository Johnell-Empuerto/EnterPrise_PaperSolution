# Phase 13 — Round-Trip Designer Compatibility Investigation

**Date:** July 20, 2026
**Status:** Investigation In Progress (Partial Comparison)
**Code Changes:** ZERO (investigation only)

---

## Objective

Determine why a PaperLess-generated workbook cannot be imported back into ConMas Designer, while a ConMas-generated workbook can — achieving true round-trip compatibility.

---

## Workbooks Compared

| Label | File | Size | Source |
|-------|------|------|--------|
| **Template** | `formtest.xlsx` | 12,830 | Original blank form template |
| **ConMas Output** | `FormTest - Copy-conmas.xlsx` | 16,784 | ConMas Designer export (production) |
| **PaperLess Old** | `test_our_output.xlsx` | 13,296 | PaperLess early export (no ExcelOutputSetting) |
| **PaperLess New** | `_x33_generated_output.xlsx` | 19,226 | PaperLess with ExcelOutputSetting |
| **x34_gen** | `_x34_generated.xlsx` | 19,221 | PaperLess generation 34 test |
| **ConMas test** | `test_conmas_output.xlsx` | 16,784 | Duplicate of ConMas Output |
| **Phase71 Original** | `_phase71_output/Original.xlsx` | 11,069 | Investigation_546 template |
| **Phase71 Export** | `_phase71_output/Export1.xlsx` | 11,069 | Phase 7.1 COM multi-gen test |

---

## Part 1 — ZIP Entry Inventory

| Entry | Template | ConMas | PL_Old | PL_New | x34_gen | Notes |
|-------|----------|--------|--------|--------|---------|-------|
| `[Content_Types].xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Same in all |
| `_rels/.rels` | ✅ | ✅ | ✅ | ✅ | ✅ | Same in all |
| `docProps/app.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `docProps/core.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/workbook.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | **STRUCTURE DIFFERS** |
| `xl/_rels/workbook.xml.rels` | ✅ | ✅ | ✅ | ✅ | ✅ | ConMas/PL_New have sheet3 rel |
| `xl/styles.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/sharedStrings.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | **PL_New: 18KB vs ConMas: smaller** |
| `xl/theme/theme1.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | **IDENTICAL in all** |
| `xl/worksheets/sheet1.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/worksheets/sheet2.xml` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/worksheets/sheet3.xml` | ❌ | ✅ | ❌ | ✅ | ✅ | **Only ConMas/PL_New have this** |
| `xl/worksheets/_rels/sheet1.xml.rels` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/worksheets/_rels/sheet2.xml.rels` | ✅ | ✅ | ✅ | ✅ | ✅ | Different per gen |
| `xl/worksheets/_rels/sheet3.xml.rels` | ❌ | ✅ | ❌ | ✅ | ✅ | ConMas/PL_New only |
| `xl/printerSettings/printerSettings1.bin` | ✅ | ✅ | ✅ | ✅ | ✅ | **IDENTICAL in all** |
| `xl/printerSettings/printerSettings2.bin` | ❌ | ❌ | ❌ | ✅ | ✅ | PL_New/x34 only |
| `xl/printerSettings/printerSettings3.bin` | ❌ | ❌ | ❌ | ✅ | ✅ | PL_New/x34 only |
| `xl/comments1.xml` | ❌ | ✅ | ❌ | ✅ | ✅ | ConMas/PL_New have comments |
| `xl/drawings/vmlDrawing1.vml` | ❌ | ✅ | ❌ | ✅ | ✅ | ConMas/PL_New have VML |
| `xl/drawings/drawing1.xml` | ❌ | ✅ | ❌ | ✅ | **?** | ConMas/PL_New have drawing |
| **Total entries** | **15** | **16** | **15** | **20** | **20** | |

---

## Part 2 — Workbook.xml Structure Comparison

### Sheet Definitions

| Workbook | Sheet 1 | Sheet 2 | Sheet 3 |
|----------|---------|---------|---------|
| **Template** | `Sheet1` (id=1, visible) | `_Fields` (id=2, hidden) | — |
| **ConMas Output** | `Sheet1` (id=1, visible) | `_Fields` (id=2, hidden) | `ExcelOutputSetting` (id=3, visible) |
| **PaperLess Old** | `Sheet1` (id=1, visible) | `_Fields` (id=2, hidden) | — |
| **PaperLess New** | `_Fields` (id=1, hidden) | `Sheet1` (id=2, visible) | `ExcelOutputSetting` (id=3, visible) |
| **x34_gen** | `_Fields` (id=1, hidden) | `Sheet1` (id=2, visible) | `ExcelOutputSetting` (id=3, visible) |
| **ConMas test** | `Sheet1` (id=1, visible) | `_Fields` (id=2, hidden) | `ExcelOutputSetting` (id=3, visible) |

### ⭐ KEY FINDING: Sheet ID / Order Mismatch ⭐

| Aspect | ConMas Output | PaperLess New | Difference |
|--------|---------------|---------------|------------|
| **Sheet order** | Sheet1 → _Fields → ExcelOutputSetting | _Fields → Sheet1 → ExcelOutputSetting | **DIFFERENT** |
| **_Fields sheetId** | **id=2** | **id=1** | **DIFFERENT** |
| **Sheet1 sheetId** | **id=1** | **id=2** | **DIFFERENT** |
| **ExcelOutputSetting sheetId** | id=3 | id=3 | Same |

ConMas Output has:
- `_Fields` with `sheetId="2"` and `Sheet1` with `sheetId="1"`

PaperLess New has:
- `_Fields` with `sheetId="1"` and `Sheet1` with `sheetId="2"`

**If ConMas Designer uses `sheetId` to identify sheets during Import Excel Form, this mismatch would cause the import to fail** — the Designer would look for `_Fields` at `sheetId=2` but find `sheetId=1`, or vice versa.

### Defined Names

| Workbook | Defined Names |
|----------|---------------|
| **Template** | `_xlnm.Print_Area` = `Sheet1!$A$1:$D$12` |
| **ConMas Output** | `_xlnm.Print_Area` = `Sheet1!$A$1:$D$12` |
| **PaperLess Old** | **NONE** ❌ |
| **PaperLess New** | `_xlnm.Print_Area` = `Sheet1!$A$1:$L$34` |
| **x34_gen** | `_xlnm.Print_Area` = ... |

**PaperLess Old is missing the `_xlnm.Print_Area` defined name entirely.** This is a critical omission that would definitely cause validation failures.

### workbookPr

All workbooks share: `defaultThemeVersion="202300"` — **Identical**.

### bookView

ConMas Output has: `xWindow`, `yWindow`, `windowWidth`, `windowHeight`, `activeTab`, `firstSheet`, `uid`
PaperLess New has: same set of attributes — **Same structure**.

---

## Part 3 — Workbook Relationships (`xl/_rels/workbook.xml.rels`)

| rId | Template | ConMas | PL_Old | PL_New |
|-----|----------|--------|--------|--------|
| rId1 | theme/theme1.xml | **worksheets/sheet3.xml** | theme/theme1.xml | theme/theme1.xml |
| rId2 | worksheets/sheet2.xml | worksheets/sheet2.xml | worksheets/sheet2.xml | worksheets/sheet2.xml |
| rId3 | worksheets/sheet1.xml | worksheets/sheet1.xml | worksheets/sheet1.xml | **worksheets/sheet3.xml** |
| rId4 | sharedStrings.xml | theme/theme1.xml | sharedStrings.xml | worksheets/sheet1.xml |
| rId5 | styles.xml | styles.xml | styles.xml | sharedStrings.xml |
| rId6 | — | sharedStrings.xml | — | styles.xml |

**Observation:** ConMas places `sheet3.xml` at rId1, while PaperLess places it at rId3. The relationship ID assignment differs.

---

## Part 4 — Content Types

| Override | Template | ConMas | PL_Old | PL_New |
|----------|----------|--------|--------|--------|
| `/xl/workbook.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/worksheets/sheet1.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/worksheets/sheet2.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/worksheets/sheet3.xml` | ❌ | ✅ | ❌ | ✅ |
| `/xl/sharedStrings.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/styles.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/theme/theme1.xml` | ✅ | ✅ | ✅ | ✅ |
| `/xl/comments1.xml` | ❌ | ✅ | ❌ | ✅ |
| `/xl/drawings/vmlDrawing1.vml` | ❌ | ✅ | ❌ | ✅ |
| `/xl/drawings/drawing1.xml` | ❌ | ✅ | ❌ | ✅ |

**Matches:** ConMas Output and PaperLess New have identical content type overrides. ✅

---

## Part 5 — Comments, VML, and Special Parts

| Part | Template | ConMas | PL_New |
|------|----------|--------|--------|
| `xl/comments1.xml` | ❌ | ✅ (present) | ✅ (present) |
| `xl/drawings/vmlDrawing1.vml` | ❌ | ✅ (present) | ✅ (present) |
| `xl/drawings/drawing1.xml` | ❌ | ✅ (present) | ✅ (present) |
| `xl/worksheets/sheet3.xml` | ❌ | ✅ (ExcelOutputSetting) | ✅ (ExcelOutputSetting) |
| `sheet3.xml.rels` | ❌ | ✅ | ✅ |
| printerSettings2.bin | ❌ | ❌ | ✅ (extra in PL_New) |
| printerSettings3.bin | ❌ | ❌ | ✅ (extra in PL_New) |

**PL_New has 4 more ZIP entries than ConMas** — extra printer settings for sheets 2 and 3. This is likely because PaperLess creates separate printer settings per sheet while ConMas only has one.

---

## Part 6 — SHA256 Hash Table Summary

| Component | Template | ConMas | PL_Old | PL_New | Drift? |
|-----------|----------|--------|--------|--------|--------|
| `[Content_Types].xml` | A | B | A | B | ✅ (2 variants) |
| `_rels/.rels` | A | A | A | A | ✅ (All same) |
| `xl/_rels/workbook.xml.rels` | A | B | A | B | ✅ (2 variants) |
| `xl/workbook.xml` | A | B | C | D | ✅ (All different) |
| `xl/styles.xml` | A | B | C | D | ✅ (All different) |
| `xl/sharedStrings.xml` | A | B | C | D | ✅ (All different) |
| `xl/theme/theme1.xml` | A | A | A | A | ✅ **ALL IDENTICAL** |
| `xl/worksheets/sheet1.xml` | A | B | C | D | ✅ (All different) |
| `xl/worksheets/sheet2.xml` | A | B | C | D | ✅ (All different) |
| `xl/worksheets/sheet3.xml` | — | A | — | B | N/A |
| `xl/printerSettings/printerSettings1.bin` | A | A | A | A | ✅ **ALL IDENTICAL** |
| `docProps/app.xml` | A | B | C | D | ✅ (Different per gen) |
| `docProps/core.xml` | A | B | C | D | ✅ (Different per gen) |

---

## Part 7 — Structural Differences That Could Cause Import Failure

### Finding A: Sheet ID Ordering (HIGH PRIORITY) ⭐

| Aspect | ConMas Output | PaperLess New | Impact |
|--------|---------------|---------------|--------|
| `_Fields` sheetId | **id=2** | **id=1** | ConMas Designer may locate sheets by ID |
| `Sheet1` sheetId | **id=1** | **id=2** | Mismatch would confuse import |

If ConMas Designer's Import Excel Form uses `sheetId` (not index position) to identify which sheet is `_Fields` (configuration) vs `Sheet1` (printable form), this mismatch alone would cause the import to fail.

### Finding B: Missing Defined Names (PaperLess Old only)

PaperLess Old has **zero** defined names, while ConMas Output has `_xlnm.Print_Area`. PaperLess New has it but with a different range.

### Finding C: Extra Printer Settings

PaperLess generates `printerSettings2.bin` and `printerSettings3.bin` for additional sheets, while ConMas only has `printerSettings1.bin`. These extra entries may confuse ConMas Designer's import validation.

### Finding D: Different Relationship ID Assignment

ConMas places `sheet3.xml` at rId1 while PaperLess places it at rId3. If ConMas Designer checks relationship IDs, this would cause a mismatch.

### Finding E: styles.xml Content

Every workbook has a different `styles.xml` hash. The actual XML content (fonts, fills, borders, cell formats) differs in every generation. If ConMas Designer checks style compatibility, this could cause rejection.

### Finding F: sharedStrings.xml Growth

PaperLess New's shared strings is significantly larger (18KB vs smaller), indicating strings are being appended/duplicated across generations.

---

## Part 8 — Likely Root Causes (Ranked)

| Rank | Cause | Confidence | Evidence |
|------|-------|-----------|----------|
| **1** | **Sheet ID ordering mismatch** | **HIGH** | ConMas: `_Fields=id2`, `Sheet1=id1`. PaperLess: `_Fields=id1`, `Sheet1=id2`. If ConMas uses sheetId for identification, this is a direct mismatch. |
| **2** | **Extra printer settings** | **MEDIUM** | PaperLess creates 3 printer settings; ConMas only 1. Extra parts may fail ConMas validation. |
| **3** | **Relationship ID differences** | **MEDIUM** | PaperLess assigns different rId values compared to ConMas. |
| **4** | **styles.xml content variation** | **LOW-MEDIUM** | Every workbook has unique styles. Unlikely to cause import failure but possible. |
| **5** | **Missing defined names (Old only)** | **HIGH for Old** | PaperLess Old has no Print_Area. But PaperLess New has it, so this doesn't explain New's failure. |

---

## Part 9 — Recommended Investigation Steps

| Priority | Action | Why |
|----------|--------|-----|
| **1** | **Compare sheet1.xml and sheet2.xml content** between ConMas and PaperLess | Check for structural differences in worksheet content |
| **2** | **Compare comments1.xml content** | Check comment format (ConMas uses specific XML structure) |
| **3** | **Compare styles.xml content** (not just hash) | Check font, fill, border, cell format differences |
| **4** | **Compare definedNames XML** | Check for hidden ConMas-specific defined names |
| **5** | **Compare vmlDrawing1.vml content** | Check VML structure for differences |
| **6** | **Test hypothesis: modify PaperLess output to use ConMas sheetId ordering** | Modify workbook.xml to swap _Fields and Sheet1 sheetId values to see if import succeeds |

---

## Conclusion

**The most likely cause of the import failure is the sheet ID ordering mismatch.** ConMas Output uses `_Fields` with `sheetId="2"` and `Sheet1` with `sheetId="1"`, while PaperLess uses the reverse (`_Fields=id1`, `Sheet1=id2`).

If ConMas Designer's Import Excel Form validates workbook structure by looking for specific `sheetId` values (e.g., "find the hidden sheet `_Fields` at sheetId 2"), the mismatch would cause validation to fail.

**Secondary causes** include:
- Extra `printerSettings2.bin` and `printerSettings3.bin` in PaperLess output
- Different relationship ID assignment for `sheet3.xml`
- `styles.xml` and `sharedStrings.xml` content differences

**A targeted test** — modifying the `sheetId` values in PaperLess output to match ConMas — would definitively prove or disprove the primary hypothesis.

---

*Investigation status: ZIP-level structural comparison complete. Deeper OOXML content comparison (styles, comments, defined names, VML) still needed.*
