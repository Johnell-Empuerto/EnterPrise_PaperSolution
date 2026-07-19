# Phase X.25 — Deep Forensic Investigation: Print Position Deviation

## Files Compared

| File | Role | Size | Notes |
|------|------|------|-------|
| `formtest.xlsx` | Original (ConMas-processed) | 12,830 B | 2 sheets, 6 comments, 6 VML shapes |
| `Investigation_546\original.xlsx` | Original (raw) | 11,069 B | 2 sheets, no comments/VML |
| `Output\output_726fea...xlsx` | Generated | 13,296 B | From pipeline (before Phase X.23/X.24 fixes) |

---

## 1. Complete PageSetup Comparison

### Read from _Fields sheet (where authoritative settings live)

| Property | Original (`_Fields`) | Generated (`_Fields`) | Generated (`Sheet1`) | Status |
|----------|---------------------|----------------------|----------------------|--------|
| `CenterHorizontally` | `1` | *not present* | *not present* | To fix (X.24 merge) |
| `CenterVertically` | `1` | *not present* | *not present* | To fix (X.24 merge) |
| `LeftMargin` | `0.70866` (51.023pt) | `0.7` (50.4pt) | `0.972` (70pt) | To fix (X.24 merge) |
| `RightMargin` | `0.70866` (51.023pt) | `0.7` (50.4pt) | `0.972` (70pt) | To fix (X.24 merge) |
| `TopMargin` | `0.74803` (53.858pt) | `0.75` (54pt) | `0.972` (70pt) | To fix (X.24 merge) |
| `BottomMargin` | `0.74803` (53.858pt) | `0.75` (54pt) | `0.972` (70pt) | To fix (X.24 merge) |
| `HeaderMargin` | `0.31496` | `0.3` | `0.3` | ❌ Not captured/written |
| `FooterMargin` | `0.31496` | `0.3` | `0.3` | ❌ Not captured/written |
| `Orientation` | `portrait` | `portrait` | `portrait` | ✅ Matches |
| `PaperSize` | Letter (default) | Not written | Not written | ❌ Read but not applied |
| `Zoom` | not set (100) | not set (100) | not set (100) | ✅ Default |
| `FitToPagesWide` | not set (0) | not set (0) | not set (0) | ✅ Default |
| `FitToPagesTall` | not set (0) | not set (0) | not set (0) | ✅ Default |

**Direct OOXML evidence** — the original `_Fields` sheet has:
```xml
<printOptions horizontalCentered="1" verticalCentered="1"/>
<pageMargins left="0.70866141732283472" right="0.70866141732283472" 
             top="0.74803149606299213" bottom="0.74803149606299213" 
             header="0.31496062992125984" footer="0.31496062992125984"/>
<pageSetup orientation="portrait" r:id="rId1"/>
```

---

## 2. Worksheet XML Comparison (Sheet1 — Content Sheet)

| Element | Original `Sheet1.xml` | Generated `Sheet1.xml` | Impact |
|---------|----------------------|----------------------|--------|
| `dimension` | `ref="A1:G1"` | `ref="A12"` | Wrong range |
| `sheetView` | `workbookViewId="0"` | `tabSelected="1" workbookViewId="0"` | Minor (active tab) |
| `sheetData` | Row 1: A1-G1 header cells | Row 12: empty A12 | Content data missing |
| `pageMargins` | L=0.7 R=0.7 T=0.75 B=0.75 H=0.3 F=0.3 | L=0.972 R=0.972 T=0.972 B=0.972 H=0.3 F=0.3 | Wrong margins (old code) |
| `pageSetup` | ❌ Absent | `orientation="portrait" r:id="rId1"` | Will be overwritten by merge fix |
| `printOptions` | ❌ Absent | ❌ Absent | ❌ **Missing centering** |
| `legacyDrawing` | ❌ Absent | `r:id="rId2"` | VML for comment shapes |

---

## 3. Workbook XML Comparison

| Element | Original | Generated | Impact |
|---------|----------|-----------|--------|
| `definedNames` | ✅ Has `_xlnm.Print_Area` = `Sheet1!$A$1:$D$12` | ❌ **Completely absent** | **CRITICAL** — No Print_Area |
| `workbookView firstSheet` | `"1"` | ❌ Missing | Minor |
| `workbookView activeTab` | `"1"` | ❌ Missing | Minor |
| Sheet order | `_Fields` first, `Sheet1` second | `Sheet1` first, `_Fields` second | Cosmetic |
| `absPath` | `ConMasDesigner\SaveFont\` | `ExcelAPI\Output\` | Metadata only |
| `documentId` | Original UUID | Different UUID | Expected |

---

## 4. Styles XML Comparison

| Feature | Original | Generated | Impact |
|---------|----------|-----------|--------|
| Fonts | 1 (Aptos Narrow 11pt) | 3 (added bold Aptos, bold Tahoma 9pt) | Cosmetic |
| Borders | 2 (default + thin) | 1 (default only) | ❌ Missing thin border (X.23 fix) |
| cellXfs | 3 (default, centered+border x2) | 2 (default, bold) | ❌ Missing alignment/border styles |
| Alignment | `horizontal="center"` on xf1, xf2 | None | ❌ Cells lose center alignment |

**Font 2 in original:** Bold+Tahoma+9pt+indexed81 — added by ConMas for comment text (matches comment `<rPr>`).

---

## 5. Defined Names

| Name | Original | Generated | Impact |
|------|----------|-----------|--------|
| `_xlnm.Print_Area` | ✅ `Sheet1!$A$1:$D$12` | ❌ **Missing** | **PRIMARY ROOT CAUSE** |
| `_xlnm.Print_Titles` | ❌ Absent | ❌ Absent | Not used |

`docProps/app.xml` confirms: originals have `TitlesOfParts` = `[ _Fields, Sheet1, Sheet1!Print_Area ]`; generated only has `[ Sheet1, _Fields ]`.

---

## 6. Hidden Sheets Comparison

### _Fields Sheet Structure

| Feature | Original `_Fields.xml` | Generated `_Fields.xml` |
|---------|-----------------------|------------------------|
| `dimension` | `A1:D12` | `A1:G7` |
| `sheetView` | `tabSelected="1"` + `selection sqref="A1:D12"` | Plain `workbookViewId="0"` |
| Merge cells | 5 ranges (A1:B2, C1:D2, A3:D4, A6:D7, A9:D10) | ❌ **None** |
| `printOptions` | ✅ `horizontalCentered="1" verticalCentered="1"` | ❌ **Missing** |
| `pageMargins` | Custom (0.70866/0.74803) | Excel defaults (0.7/0.75) |
| `pageSetup` | ✅ `orientation="portrait"` | ✅ `orientation="portrait"` |
| Cell data | Empty metadata cells (style-only) | **Actual uploaded data values** (7 cols × 7 rows) |
| `legacyDrawing` | ✅ Present (formtest only) | ❌ Missing |

**The generated _Fields sheet contains uploaded data instead of the metadata structure.** This indicates the upload pipeline wrote content data to the _Fields sheet instead of preserving its metadata layout.

---

## 7. Comments and VML

| Feature | Original (formtest) | Generated | Preserved |
|---------|--------------------|-----------|-----------|
| Comments | 6 (A1, C1, A3, A6, A9, A12) | 1 (A12 only) | ❌ **17%** |
| VML shapes | 6 comment shapes | 1 shape | ❌ **17%** |

The comment text in the original contains ConMas form-definition metadata (samples, KeyboardText, etc.). Only A12's comment survived.

---

## 8. Relationships Comparison

| Relationship | Original (formtest) | Generated |
|-------------|--------------------|-----------|
| `sheet1.rels` | (none) | comments + vmlDrawing + printerSettings |
| `sheet2.rels` | comments + vmlDrawing + printerSettings | printerSettings only |
| Total printerSettings | 1 (`_Fields` sheet) | 2 (one per sheet) |

---

## 9. Document Properties (app.xml)

| Property | Original | Generated |
|----------|----------|-----------|
| Named Ranges count | 1 (`Sheet1!Print_Area`) | 0 |
| TitlesOfParts | `[_Fields, Sheet1, Sheet1!Print_Area]` | `[Sheet1, _Fields]` |

---

## 10. Hypothesis Validation

### Hypothesis A — Excel PageSetup controls everything; we missed one property
**PARTIALLY TRUE.** We fixed centering and margins, but **PaperSize** is read but never applied, and **HeaderMargin/FooterMargin** are never read or written. However, these are secondary factors. The primary issue is:

### Hypothesis B — ConMas stores layout outside PageSetup
**TRUE.** The `_xlnm.Print_Area` defined name is the critical layout mechanism. While not strictly "outside PageSetup," it's stored in `workbook.xml > definedNames` — a completely separate location from `pageSetup`/`pageMargins`. The Print_Area constrains which cells Excel considers printable, directly affecting page positioning.

### Hypothesis C — ConMas never regenerates PageSetup
**UNVERIFIED** (requires ConMas runtime test).

### Hypothesis D — Another worksheet affects print layout
**FALSE.** No evidence of cross-sheet print layout interaction.

---

## Root Cause Determination

### Primary Root Cause (Print Position)

**Missing `_xlnm.Print_Area` defined name in the generated workbook.**

| Aspect | Original | Generated |
|--------|----------|-----------|
| Print Area | `Sheet1!$A$1:$D$12` | Auto-detected by Excel |
| Print range | Constrained (A1:D12) | Entire used range |
| Page fill behavior | Centered within A1:D12 | Spread across full used range |

**How this manifests differently in print:**
- Original with Print_Area: Excel fits A1:D12 to the page → centered within that 4-col × 12-row box
- Generated without Print_Area: Excel auto-detects the used range → may be larger or positioned differently → shifts the printed content

**Confidence: HIGH** — the missing `definedNames` block is the largest structural difference between the files.

### Secondary Contributing Factors

1. **Missing `printOptions` (centering)** — The PageSetup merge (Phase X.24) should fix this, but centering is meaningless without a defined Print_Area to center within.
2. **HeaderMargin/FooterMargin not preserved** — `0.31496` vs `0.3` — small but measurable difference.
3. **Missing thin borders on cells** — Phase X.23 fixes this for A12, but the generated sample still lacks them.
4. **Wrong cell data on _Fields sheet** — The content sheet's data is written to _Fields instead, which may affect automatic used-range detection.

### Recommendation

1. **Verify Print_Area is actually written** after the Phase X.23 rename fix. The `SetDefinedNames` code exists and `CalculatePrintArea` computes the correct address, but the rename bug previously caused it to fail. Confirm with forensic logging that `Names.Add("_xlnm.Print_Area", ...)` succeeds.
2. **Add PaperSize to `ApplyPageSettings`** — it's already read but never written.
3. **Read and apply HeaderMargin/FooterMargin** — add to both `ReadPageSettings` and `ApplyPageSettings`.
4. **Print_Area is the single highest-impact fix** for print positioning. Without it, all other PageSetup adjustments operate on an undefined canvas.

---

## Complete XML Difference Classification

| Category | Differences Found | Fixed? |
|----------|-----------------|--------|
| **Print** | Missing `definedNames/Print_Area` | ❌ Not yet verified |
| **Print** | Missing `printOptions` on content sheet | Phase X.24 |
| **Print** | Missing HeaderMargin/FooterMargin | ❌ Not addressed |
| **Print** | PaperSize not applied | ❌ Not addressed |
| **Worksheet** | Wrong dimension on Sheet1 | Phase X.23 |
| **Worksheet** | No merge cells on _Fields | Phase X.23 |
| **Worksheet** | Missing `selection` in sheetViews | ❌ Not addressed |
| **Styles** | Missing thin border definition | Phase X.23 |
| **Styles** | Missing center alignment styles | Phase X.23 |
| **Metadata** | Comments/VML only 17% preserved | ❌ Not addressed |
| **Metadata** | _Fields has content data instead of structure | ❌ Not addressed |
| **Metadata** | Different documentId/absPath | Expected |
| **Metadata** | Missing firstSheet/activeTab in workbookView | ❌ Minor |
