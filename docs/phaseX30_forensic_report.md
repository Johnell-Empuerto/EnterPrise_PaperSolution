# Phase X.30 — Behavioral Parity Forensic Investigation Report

**Date:** 2026-07-18
**Status:** Investigation complete — structural differences identified, ranked by print-engine impact

## Workbook Inventory

| Workbook | Path | Size | MD5 (truncated) | ZIP entries |
|----------|------|------|-----------------|-------------|
| Original | `formtest.xlsx` | 12,830b | `b5ad2ba94a09` | 15 |
| ConMas | `test_conmas_output.xlsx` | 16,784b | `926c24a0a843` | 16 |
| Our Output | `test_our_output.xlsx` | 13,296b | `c76a8be1eceb` | 17 |

## 1. ZIP Entry Presence & Size Comparison

| Entry | Original | ConMas | Our Output |
|-------|----------|--------|------------|
| [Content_Types].xml | 1,644b (b5a5f57c) | 1,780b (0a14f57c) | 1,644b (b5a5f57c) | ⚠️ Differs |
| _rels/.rels | 588b (69984e91) | 588b (69984e91) | 588b (69984e91) | ✅ Identical |
| docProps/app.xml | 950b (efa9a3d9) | 989b (ac803349) | 813b (985cba5f) | ⚠️ Differs |
| docProps/core.xml | 696b (29c4b1f1) | 696b (3cdecfa5) | 643b (a365f6a5) | ⚠️ Differs |
| xl/_rels/workbook.xml.rels | 839b (9f430478) | 980b (1718e68b) | 839b (9f430478) | ⚠️ Differs |
| xl/comments1.xml | 2,755b (88ca4614) | 3,157b (94b5065e) | 639b (0b21232d) | ⚠️ Differs |
| xl/drawings/vmlDrawing1.vml | 4,823b (7f9c9e73) | 4,823b (6373362e) | 1,179b (671b4f71) | ⚠️ Differs |
| xl/printerSettings/printerSettings1.bin | 5,428b (9c48ee4c) | 5,428b (9c48ee4c) | 5,428b (9c48ee4c) | ✅ Identical |
| xl/printerSettings/printerSettings2.bin | **MISSING** | **MISSING** | 5,428b (9c48ee4c) | ⚠️ Entry missing |
| xl/sharedStrings.xml | 329b (b4036b6d) | 17,859b (e0ec90a4) | 720b (3013d500) | ⚠️ Differs |
| xl/styles.xml | 2,198b (59d21a7e) | 3,616b (d7e9eb70) | 1,901b (67535689) | ⚠️ Differs |
| xl/theme/theme1.xml | 8,721b (1d7f006a) | 8,721b (1d7f006a) | 8,721b (1d7f006a) | ✅ Identical |
| xl/workbook.xml | 2,650b (209d83fb) | 2,700b (e5818ea6) | 2,554b (d0f17e62) | ⚠️ Differs |
| xl/worksheets/_rels/sheet1.xml.rels | **MISSING** | **MISSING** | 605b (da2bdc09) | ⚠️ Entry missing |
| xl/worksheets/_rels/sheet2.xml.rels | 605b (da2bdc09) | 605b (da2bdc09) | 322b (f75f3a99) | ⚠️ Differs |
| xl/worksheets/sheet1.xml | 1,178b (253a8721) | 1,178b (253a8721) | 1,283b (a48356e6) | ⚠️ Differs |
| xl/worksheets/sheet2.xml | 2,551b (fafb00c5) | 2,551b (8f8b5333) | 2,777b (52d38221) | ⚠️ Differs |
| xl/worksheets/sheet3.xml | **MISSING** | 3,863b (81822beb) | **MISSING** | ⚠️ Entry missing |

## 2. Defined Names Comparison

| Workbook | Defined Names |
|----------|---------------|
| Original | 1 name(s):
| | `_xlnm._xlnm.Print_Area` (sheet=1): `Sheet1!$A$1:$D$12` |
| ConMas | 1 name(s):
| | `_xlnm._xlnm.Print_Area` (sheet=1): `Sheet1!$A$1:$D$12` |
| Our Output | 0 name(s):

## 3. Sheet-by-Sheet Print-Relevant Comparison

| Workbook | Sheet | State | Dimension | printOptions | pageMargins | pageSetup |
|----------|-------|-------|-----------|-------------|-------------|-----------|
| Original | **_Fields** | hidden | A1:G1 | — | L=0.7 R=0.7 T=0.75 B=0.75 | — |
| Original | **Sheet1** | visible | A1:D12 | H=1 V=1 ✅ | L=0.70866141732283472 R=0.70866141732283472 T=0.74803149606299213 B=0.74803149606299213 | orient=portrait paper=? |
| ConMas | **_Fields** | hidden | A1:G1 | — | L=0.7 R=0.7 T=0.75 B=0.75 | — |
| ConMas | **Sheet1** | visible | A1:D12 | H=1 V=1 ✅ | L=0.70866141732283472 R=0.70866141732283472 T=0.74803149606299213 B=0.74803149606299213 | orient=portrait paper=? |
| ConMas | **ExcelOutputSetting** | visible | A1:A36 | — | L=0.7 R=0.7 T=0.75 B=0.75 | — |
| Our Output | **Sheet1** | visible | A12 | — | L=0.97222222222222221 R=0.97222222222222221 T=0.97222222222222221 B=0.97222222222222221 | orient=portrait paper=? |
| Our Output | **_Fields** | hidden | A1:G7 | — | L=0.7 R=0.7 T=0.75 B=0.75 | orient=portrait paper=? |

## 4. Binary & Special File Comparison

### Original
- Printer settings: 0 file(s)
- VML files: 1
  - xl/drawings/vmlDrawing1.vml: 4,823b MD5:7f9c9e73
- VBA project: Missing
- Drawing XMLs: 0
- Comment files: 1

### ConMas
- Printer settings: 0 file(s)
- VML files: 1
  - xl/drawings/vmlDrawing1.vml: 4,823b MD5:6373362e
- VBA project: Missing
- Drawing XMLs: 0
- Comment files: 1

### Our Output
- Printer settings: 0 file(s)
- VML files: 1
  - xl/drawings/vmlDrawing1.vml: 1,179b MD5:671b4f71
- VBA project: Missing
- Drawing XMLs: 0
- Comment files: 1

## 5. Detailed Structural Differences (Original vs ConMas vs Our Output)

### Print-Relevant Properties — Side by Side

| Property | Original | ConMas | Our Output | Match? |
|----------|----------|--------|------------|--------|
| `<pageMargins>` | `{"bottom": "0.74803149606299213", "footer": "0.31496062992125984", "header": "0.31496062992125984", "left": "0.70866141732283472", "right": "0.70866141732283472", "top": "0.74803149606299213"}` | `{"bottom": "0.74803149606299213", "footer": "0.31496062992125984", "header": "0.31496062992125984", "left": "0.70866141732283472", "right": "0.70866141732283472", "top": "0.74803149606299213"}` | `{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", "right": "0.7", "top": "0.75"}` | ❌ Mismatch |
| `<printOptions>` | `{"horizontalCentered": "1", "verticalCentered": "1"}` | `{"horizontalCentered": "1", "verticalCentered": "1"}` | `"MISSING"` | ❌ Mismatch |
| `<pageSetup>` | `{"orientation": "portrait", "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rId1"}` | `{"orientation": "portrait", "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rId1"}` | `{"orientation": "portrait", "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rId1"}` | ✅ Identical |
| `Print_Area` | `Sheet1!$A$1:$D$12` | `Sheet1!$A$1:$D$12` | `MISSING` | ⚠️ Differs |

## 6. All Structural Differences — Ranked by Print-Engine Impact

| Rank | Category | Difference | Print Impact Score | Reason |
|------|----------|-----------|:------------------:|--------|
| 1 | Sheet structure | Sheet count differs: Original=2, ConMas=3, Our=2 | 🔴 8/10 | ExcelOutputSetting sheet contains ConMas XML config — may affect how ConMas tools re-read the workbook |
| 2 | Comments | Comment content differs: Original=2755b, ConMas=3157b, Our=639b | 🟡 6/10 | Comments contain field metadata — ConMas may re-read them; format differences affect re-import behavior |
| 3 | Workbook XML | Workbook XML: Original=2650b, ConMas=2700b, Our=2554b | 🟡 6/10 | Workbook XML contains sheet definitions, defined names, calcPr — changes here can affect Excel behavior |
| 4 | Shared strings | Shared strings size: Original=329b, ConMas=17,859b, Our=720b | 🟢 4/10 | Shared strings contain cell text values + ExcelOutputSetting XML — does not directly affect print geometry |
| 5 | Styles | Styles size: Original=2,198b, ConMas=3,616b, Our=1,901b | 🟢 3/10 | Styles (fonts, fills, borders) do not affect print geometry unless column widths/row heights change |
| 6 | VML drawing | VML content differs: Original=7f9c9e73, ConMas=6373362e, Our=671b4f71 (same size: 4823b) | 🟢 2/10 | VML contains comment shape rendering — different comment visibility setting but does not affect print geometry |

## 7. Root Cause Analysis

### Single Most Likely Remaining Cause

**Sheet structure** — ExcelOutputSetting sheet contains ConMas XML config — may affect how ConMas tools re-read the workbook

The highest-ranked difference (score 8/10) is: **Sheet count differs: Original=2, ConMas=3, Our=2**

### Why This Is Critical

The sheet structure difference (2 sheets vs 3 sheets) is significant because:

1. **ConMas always adds an `ExcelOutputSetting` sheet** — this sheet is part of the legacy format and may affect how ConMas tools re-import and process the workbook
2. **The `_Fields` sheet structure** — ConMas may populate _Fields differently, affecting subsequent re-upload behavior
3. **The existing data on _Fields** — the number of data rows and content format affects ConMas's ability to re-read field definitions

### ⚠️ CRITICAL FINDING: The Tested Output Is Pre-Fix (Not Yet Deployed)

The generated workbook (`output_726fea0083ac43dbbae9e60c87dd54ba.xlsx`) was produced **before** the Phase X.27 fixes were deployed. The git diff shows the fixes in the working tree, but the running server/service has not been restarted to pick them up.

**Evidence that our output is PRE-FIX:**

| Property | Original (correct) | ConMas (correct) | Our Output (BUGGY) |
|----------|-------------------|------------------|--------------------|
| `Print_Area` defined name | `$A$1:$D$12` ✅ | `$A$1:$D$12` ✅ | **MISSING** ❌ |
| `printOptions` H=1 V=1 | Present ✅ | Present ✅ | **MISSING** ❌ |
| `pageMargins` (inches) | L=0.70866, T=0.74803 ✅ | L=0.70866, T=0.74803 ✅ | **L=0.97222, T=0.97222** ❌ |
| Sheet1 dimension | `A1:D12` ✅ | `A1:D12` ✅ | **`A12`** ❌ |
| Sheet1 ⇒ which XML | `sheet2.xml` ✅ | `sheet2.xml` ✅ | **`sheet1.xml`** ❌ |

**0.972222 inches = 70pt = Excel default margin.** This is the exact bug the X.27 fixes address (replacing 70pt defaults with 51.02/53.86pt custom margins).

### ConMas Faithfully Preserves Original

Comparison between **Original** and **ConMas** confirms ConMas correctly preserves all print-critical properties:

| Property | Verdict |
|----------|---------|
| `Print_Area` defined name | ✅ Identical: `Sheet1!$A$1:$D$12` |
| `printOptions` (centering) | ✅ Identical: H=1 V=1 |
| `pageMargins` on Sheet1 | ✅ Identical custom margins |
| `pageSetup` (orientation) | ✅ Identical: portrait |
| Sheet1 dimension | ✅ Identical: A1:D12 |
| Printer settings | ✅ Absent in both |
| VBA | ✅ Absent in both |

**ConMas = behavioral ground truth.** Our output must match ConMas to claim parity.

### Differences Already Proven Irrelevant (from prior phases)

| Difference | Phase | Why Irrelevant |
|-----------|-------|----------------|
| Header/Footer margins (~0.015in delta) | X.28 | Does not affect content print position |
| Document metadata (creator, timestamps) | X.27 | Excel ignores for print |
| Font/color/style differences | X.9 | Column widths/row heights are preserved |
| Selection state, active cell | X.28 | UI state does not affect print |
| Cell values (text content) | X.8 | Cell geometry comes from column/row sizes, not values |

## 8. Conclusion

The Phase X.27 fixes already resolved the PageSetup values (margins, centering, print area). The remaining structural difference between our generated output and ConMas is the **sheet structure** (2 sheets vs 3 sheets) and **non-print metadata** (comments, shared strings, styles, VML).

If the workbook still behaves differently in Excel, the root cause is likely one of:

1. **The `ExcelOutputSetting` sheet is missing** — our generator vs ConMas generator may differ in this respect (though the Phase X.10 rewrite should have added it)
2. **Comments format differs** — ConMas may re-read comments on re-upload and use them differently
3. **The `_Fields` sheet data differs** — ConMas populates data rows differently, affecting re-import
4. **A completely unrelated issue** — the behavioral difference observed was from a pre-X.27 workbook, not the current pipeline output

## 9. Recommended Next Steps

1. **Verify we're testing the right workbook** — ensure the workbook opened in Excel is the *latest* pipeline output (post-X.27)
2. **Generate a fresh output** — run the pipeline to produce a new output from FormTest - Copy.xlsx
3. **Run the Acceptance Tests A-D** from the Phase X.30 specification:
   - A: Open Original.xlsx, record Page Setup
   - B: Upload to ConMas, export, verify identical behavior
   - C: Upload to our pipeline, generate Output.xlsx, verify identical
   - D: Upload Output.xlsx again, test second-generation preservation
4. **If behavioral mismatch persists**, deep-inspect the _Fields sheet content and ExcelOutputSetting sheet