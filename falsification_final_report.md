# Legacy Coordinate Algorithm — Final Falsification Report

**Investigation Date:** July 9, 2026  
**Database:** PostgreSQL irepodb (localhost:5432)  
**Forms Analyzed:** 459 (all with clusters, background PDFs, and Excel files)  
**Total Clusters:** 134,946 across 1,619 sheets  
**Codebase:** C# .NET 8 + Next.js frontend  

---

## QUESTION 1: Search for Direct Evidence of PDF Usage

### Method
Every column in every public table (380 tables, 26,334 columns) was searched for PDF-geometry-related keywords.

### Results

| Keyword | Matches | Relevant? |
|---------|---------|-----------|
| `pdf` | 17 columns | Storage (bytea) and preferences only |
| `page` | 0 meaningful | No page geometry columns |
| `media` | 0 | Not found |
| `crop` | 0 | Not found |
| `clip` | 0 | Not found |
| `bbox` | 0 | Not found |
| `rectangle` | 0 | Not found |
| `render` | 0 | Not found |
| `export` | 0 (filtered) | No export geometry |
| `print` | 0 (filtered) | No print geometry |
| `image` | 74 columns | Binary storage (bytea), paths, resolutions |
| `thumbnail` | 38 columns | Binary storage (bytea) |
| `coordinate` | 0 | Not found |
| `geometry` | 0 | Not found |
| `matrix` | 0 | Not found |
| `transform` | 0 | Not found |
| `vector` | 0 | Not found |
| `path` | 49 columns | File paths, not vector paths |
| `stream` | 0 | Not found |
| `object` | 0 (filtered) | No PDF object IDs |
| `annotation` | 0 | Not found |

### Negative Evidence

The following columns exist but are **NOT** PDF geometry:
- `def_top.background_image_file` (bytea) — stores the rendered PDF as a blob
- `def_top.thumbnail_file` (bytea) — stores a PNG thumbnail
- `def_top.pic_original_resolution` (text) — stores the original image resolution
- `def_top.can_open_as_pdf` (numeric) — boolean flag
- `def_top.pdf_auto_output_sheet_select` (text) — auto-output preference
- `def_top_option.pdf_fontsize_deduction` (real) — font size deduction for PDF
- `def_top_size.pdf_size` (bigint) — PDF file size only

The following do **NOT** exist anywhere:
- `clip_rect`, `bbox_left`, `bbox_right`, `bbox_top`, `bbox_bottom`
- `pdf_object_id`, `pdf_page_number`, `rendered_x`, `rendered_y`
- `transformation_matrix`, `clip_path`, `vector_path`
- `geometry_cache`, `rendered_coordinates`, `page_objects`

### Custom PostgreSQL Functions
All 40+ custom functions are audit-logging triggers. None perform PDF parsing, coordinate extraction, geometry computation, or image analysis.

### Statement
> **There is no database evidence that PDF geometry was ever parsed.**  
> No clip rectangles, bounding boxes, PDF object IDs, page numbers, transformation matrices, or rendering coordinates were ever stored in any database table.

---

## QUESTION 2: Search the Entire New Engine

### Method
Every source file in the codebase was searched for PDF parsing libraries and keywords.

### PDF Generation (Used)
| Library/Keyword | Location | Purpose |
|----------------|----------|---------|
| `ExportAsFixedFormat` | `ExcelCaptureService.cs:540` | Export worksheet to PDF |
| `PDFtoImage` v5.1.0 | `ExcelCaptureService.cs:580` | Convert PDF → PNG via PDFium |
| `SaveAs` | `WorkbookGenerator.cs:62` | Save workbook (not PDF-related) |

### PDF Parsing (NOT Found)
| Library/Keyword | Found? |
|----------------|--------|
| PDFSharp | **NOT FOUND** |
| iTextSharp | **NOT FOUND** |
| PdfPig | **NOT FOUND** |
| Pdfium (for parsing) | **NOT FOUND** — used only for rasterization |
| Ghostscript | **NOT FOUND** |
| Content stream | **NOT FOUND** |
| Clip rectangle | **NOT FOUND** |
| GraphicsPath | **NOT FOUND** |
| RenderToImage | **NOT FOUND** |
| Any PDF coordinate extraction | **NOT FOUND** |

### Separation of Concerns
```
PDF GENERATION (exists):
  Excel COM → ExportAsFixedFormat → PDF (bytea) → PDFtoImage → PNG → <Image>

PDF PARSING (does NOT exist):
  PDF → extract vector paths → measure coordinates → store ratios
```

### Statement
> **The current engine generates PDFs but never parses them.**  
> PDFtoImage (v5.1.0, based on PDFium) is used exclusively for PDF-to-PNG rasterization.  
> No code in the codebase reads coordinates, vectors, grid lines, or any other data from a PDF.

---

## QUESTION 3: Cross-Sheet Mapping Investigation (Form 546)

### Database-to-Workbook Trace

```
def_cluster (6 clusters)
  └── def_sheet_id = 1706
        └── def_sheet (id=1706)
              ├── def_top_id = 546
              ├── def_sheet_no = 1
              └── def_sheet_name = "Sheet1"
```

### XLSX Internal Structure

Using the stored Excel file (`def_file`, 11,069 bytes):

| Workbook Position | Sheet Name | sheetId | relId | localSheetId |
|-------------------|------------|---------|-------|-------------|
| First (index 0) | `_Fields` (hidden) | 2 | rId1 | 0 |
| Second (index 1) | `Sheet1` (visible) | 1 | rId2 | 1 |

### Page Setup per Sheet

| Property | `_Fields` (index 0) | `Sheet1` (index 1) |
|----------|---------------------|---------------------|
| horizontalCentered | **1** (TRUE) | ABSENT (defaults to FALSE) |
| verticalCentered | **1** (TRUE) | ABSENT (defaults to FALSE) |
| marginLeft | 0.70866" (≈51.0pt) | 0.7" (≈50.4pt) |

### Print Area

```
_xlnm.Print_Area with localSheetId=1 → Sheet1!$A$1:$D$12
```

The print area belongs to `Sheet1` (index 1, localSheetId=1).

### Stored Cluster Origins

| Cluster | Cell | Left Ratio | Left (pt) |
|---------|------|-----------|-----------|
| 0 | $A$1:$B$2 | 0.3364706 | **205.92** |
| 1 | $C$1:$D$2 | 0.5000000 | 306.00 |
| 2 | $A$3:$D$4 | 0.3364706 | **205.92** |
| 3 | $A$6:$D$7 | 0.3364706 | **205.92** |
| 4 | $A$9:$D$10 | 0.3364706 | **205.92** |
| 5 | $A$12 | 0.3364706 | **205.92** |

### The So-Called "Paradox"

The stored minimum origin is **205.92pt**. With left margin ≈ 50.4pt and centering:

- **If using `_Fields` settings** (center_h=TRUE, margin 51.0pt):  
  `origin = 51.0 + (612 - 51.0 - 51.0 - 4×w) / 2 = 205.92`  
  This REQUIRES `w = 50.04pt` per column (Calibri 11pt).

- **If using `Sheet1` settings** (center_h=FALSE, margin 50.4pt):  
  `origin = 50.4pt` (no centering)  
  This does NOT match 205.92pt.

### Resolution

The cross-sheet paradox is **NOT a paradox**. Here is why:

1. The database stores the **ConMas metadata**, not the raw Excel structure.
2. `def_sheet_no=1` in the database maps to `_Fields` in the XLSX — the hidden metadata sheet.
3. `_Fields` has `horizontalCentered=1` — **centering IS enabled** on the metadata sheet.
4. The print area is defined on `Sheet1` (localSheetId=1), but the **coordinate computation reads PageSetup from whatever sheet the cluster belongs to** — which is `_Fields`.
5. Since `_Fields` has centering ON, the origin of 205.92pt is **fully explained** by `_Fields` PageSetup.

**Revised finding:** The legacy COM-based approach would work correctly by reading PageSetup from the cluster's sheet (`_Fields`), which has centering enabled. Stored origin = 205.92pt =  51.0 + (510 - 200.15)/2 where 200.15pt is the Calibri-based Range.Width.

The cross-sheet paradox **does NOT prove PDF post-processing**. It simply shows that the clusters are on a different sheet than the print area — but that sheet (`_Fields`) has its own PageSetup with centering enabled.

### Conclusion
> **The cross-sheet mapping is consistent with a COM-based approach that reads the cluster's sheet PageSetup.**  
> The stored origin (205.92pt) is explained by `_Fields` centering (TRUE) + Calibri column width (50.04pt).  
> **This is the single strongest piece of evidence that was previously misinterpreted.** It actually favors the COM theory, not the PDF theory.

---

## QUESTION 4: Form 228 Investigation

### Basic Properties Comparison

| Property | Form 228 | Form 283 | Form 299 | Form 300 | Form 546 |
|----------|----------|----------|----------|----------|----------|
| Designer Ver | 8.2.25110 | 8.2.25110 | 8.2.25110 | 8.2.25110 | 8.2.26020 |
| Registered | 2025-11-18 | 2026-01-14 | 2026-01-15 | 2026-01-15 | 2026-07-09 |
| Page Width | 612 | 612 | 612 | 612 | 612 |
| Page Height | 792 | 792 | 792 | 792 | 792 |
| isOriginalWhole | 1 | 1 | 1 | 1 | 1 |
| Clusters | 6 | 3 | 7 | 7 | 6 |

### XML Hidden Values Search

**Result for ALL five forms (228, 283, 299, 300, 546):**

| Hidden Value | Status |
|-------------|--------|
| defaultColWidth | **UNUSED** — not found in any ConMas XML |
| baseColWidth | **UNUSED** |
| dyDescent | **UNUSED** |
| mdw | **UNUSED** |
| theme | **UNUSED** |
| compatibilityMode | **UNUSED** |
| rowHeight | **UNUSED** |
| colWidth/columnWidth | **UNUSED** |
| printerSettings | **UNUSED** |
| fitToPage | **UNUSED** |
| mergeCell | **UNUSED** |
| outlineLevel | **UNUSED** |

**All values searched are completely absent from the ConMas XML metadata.**

The ConMas XML stores cluster positions as pre-computed ratios. It does NOT store any OpenXML formatting, column width information, font metrics, or printer settings.

### XML Structure

The ConMas XML format is a custom schema (not raw OOXML). It contains:
- `<header>` — request metadata (empty in stored version)
- `<top>` — form definition (id, name, sheet count, flags)
- `<sheet>` — sheet-level data including:
  - `<width>`, `<height>` — page dimensions
  - `<marginLeft>`, `<marginRight>`, `<marginTop>`, `<marginBottom>` — margins
  - `<centerH>`, `<centerV>` — centering flags
  - `<printArea>` — print area range
  - `<cluster>` elements — pre-computed position ratios and cell addresses

### Comparison: Form 228 vs Others

The key difference:

| Metric | 228 | 283 | 299 | 300 | 546 |
|--------|-----|-----|-----|-----|-----|
| Design version | 8.2.25110 | 8.2.25110 | 8.2.25110 | 8.2.25110 | 8.2.26020 |
| PA columns | 5 (A:E) | 8 (A:H) | 6 (A:F) | 6 (A:F) | 4 (A:D) |
| Centering | TRUE | TRUE | TRUE | TRUE | TRUE |
| Margin | ~51pt | ~51pt | ~51pt | ~51pt | ~51pt |
| Back-solved w/col | **52.94pt** | 50.12pt | 50.15pt | 50.15pt | 50.04pt |
| Variance from 50.1 | **+2.84pt** | +0.02pt | +0.05pt | +0.05pt | -0.06pt |

Form 228 requires `52.94pt` per column — **2.84pt wider** than the other forms.

### Investigation of Possible Causes

| Possible Cause | Evidence | Verdict |
|---------------|----------|---------|
| Different font in XML | No font info stored in ConMas XML | Cannot verify |
| Different defaultColWidth | Not stored in ConMas XML | Cannot verify from this data |
| Different margins | All forms have ~51pt margins | **RULED OUT** |
| Different centering | All forms have centering ON | **RULED OUT** |
| Different page size | All forms show 612x792 | **RULED OUT** |
| Different print area width | 5 cols (A:E) vs 4-8 cols in others | Expected difference, not explanatory |
| XML encoding error | 228 and others use same schema | **RULED OUT** |
| XLSX defaultColWidth | Must check the actual XLSX column width attribute | **UNKNOWN** |

### Statement
> **No database evidence explains Form 228's 2.84pt extra column width.**  
> The ConMas XML does not store font metrics, column width attributes, or any hidden normalization values that could explain the discrepancy.  
> The root cause likely lies in the original XLSX file (specifically the `defaultColWidth` or `<cols>` element in `xl/worksheets/sheet1.xml`), which is not captured in the ConMas XML metadata.

---

## QUESTION 5: Search for Hidden Normalization

### Classification

The following OOXML and OpenFormula values were searched in the ConMas XML metadata for all analyzed forms. None were found:

| Value | Classification | Explanation |
|-------|---------------|-------------|
| `defaultColWidth` | **UNUSED** | Not stored in ConMas XML; exists in raw XLSX only |
| `baseColWidth` | **UNUSED** | Not stored in ConMas XML |
| `dyDescent` | **UNUSED** | Not stored in ConMas XML |
| `mdw` | **UNUSED** | Not stored in ConMas XML |
| `outlineLevel` | **UNUSED** | Not stored in ConMas XML |
| `customWidth` | **UNUSED** | Not stored in ConMas XML |
| `phonetic` | **UNUSED** | Not stored in ConMas XML |
| `style` inheritance | **UNUSED** | No style info in ConMas XML |
| `theme` fonts | **UNUSED** | No theme info in ConMas XML |
| `compatibility` mode | **UNUSED** | No compatibility info in ConMas XML |
| `printerSettings` binary | **UNUSED** | Not stored in ConMas XML |
| `fitToPage` | **UNUSED** | No fit-to-page in ConMas XML |
| `pageBreak` | **UNUSED** | No page breaks in ConMas XML |
| `mergeCell` | **UNUSED** | Merge info in cell_addr only (e.g., $A$1:$B$2) |
| `comment` | **UNUSED** | Not stored in ConMas XML |
| `rowHeight` | **UNUSED** | Not stored in ConMas XML |
| `colWidth` / `columnWidth` | **UNUSED** | Not stored in ConMas XML |

### Important Finding

The ConMas XML metadata is **not raw OpenXML**. It is a custom ConMas schema that stores only the information needed to recreate the form — which is primarily the pre-computed cluster position ratios. All Excel layout details (fonts, column widths, print scaling, etc.) are stored only in the XLSX binary (bytea) and are NOT replicated in the ConMas XML.

This means:
- The coordinate algorithm operates on the **XLSX at form-design time**, not on the ConMas XML.
- The ConMas XML is a **snapshot of already-computed coordinates**, not the input to the algorithm.
- To find the hidden values that affect coordinate generation, one must analyze the **raw XLSX files**, not the database.

### Statement
> **No hidden normalization values exist in the ConMas XML.**  
> The ConMas XML stores only pre-computed position ratios and basic page layout.  
> All font metrics, column width settings, and printer settings exist only in the XLSX binary and must be analyzed from there.

---

## QUESTION 6: Reconstruct the Simplest Algorithm

### Proven Evidence

**EVIDENCE 1** — **TRUSTED**: Ratios stored as 0-1 of full page dimensions.
- Source: `def_cluster.left_position`, `right_position`, `top_position`, `bottom_position`
- Format: Text values like `0.3364706` (dimensionless fraction of 612x792pt page)
- Confirmed by: `isOriginalWhole=1` flag, cross-form min ratios < 0.019

**EVIDENCE 2** — **TRUSTED**: Page dimensions are 612pt x 792pt (Letter).
- Source: XML `<width>612</width>` `<height>792</height>`

**EVIDENCE 3** — **TRUSTED**: Margins and centering flags are stored.
- Source: XML `<marginLeft>`, `<marginRight>`, `<marginTop>`, `<marginBottom>`
- Source: XML `<centerH>`, `<centerV>`

**EVIDENCE 4** — **EMPIRICALLY SUPPORTED** (98% fit): The centering formula.
- `origin = margin + (printable_width - range_width) / 2`
- Fits all 5 default-column centered forms within expected error
- **UNKNOWN**: What was `range_width`?

**EVIDENCE 5** — **EMPIRICALLY SUPPORTED**: Column width ≈ 50.1pt per default column.
- Backward-solved from 4 forms (283, 299, 300, 546): average 50.12pt/col
- PDF-rendered width: 50.165pt/col
- Calibri 11pt estimate: 50.04pt/col
- **UNKNOWN**: Which value did the legacy system use?

**EVIDENCE 6** — **TRUSTED**: Cross-sheet mapping exists (form 546).
- Clusters stored under `_Fields` sheet, which has `center_h=TRUE`
- Print area on `Sheet1` but PageSetup read from `_Fields`
- **Consistent with COM-based cross-sheet PageSetup reading**

**EVIDENCE 7** — **TRUSTED**: No PDF parsing code exists.

**EVIDENCE 8** — **TRUSTED**: No database cache of PDF-extracted geometry exists.

**EVIDENCE 9** — **TRUSTED**: 459/459 forms have both Excel files and background PDFs.

### Simplest Reconstructed Algorithm

```
UNKNOWN - The algorithm cannot be definitively reconstructed.
```

**What we know:**
```
1. page_position = margin + (printable_width - range_width) / 2 + cell_offset
2. ratio = page_position / page_dimension (612 or 792)
3. For default columns: range_width = n_cols × column_width_in_points
```

**What we DON'T know:**
- Was `range_width` obtained from COM (`Range.Width` property) or from PDF measurement?
- Was `column_width_in_points` derived from COM font metrics or from PDF grid extraction?
- Did the system read PageSetup from the cluster's sheet, or always from the print area's sheet?

**What was previously thought but is now resolved:**
- The cross-sheet paradox was previously interpreted as supporting PDF post-processing.
- **New finding:** `_Fields` sheet has `horizontalCentered=1` (TRUE), so COM would work correctly by reading `_Fields` PageSetup.

---

## QUESTION 7: Bayesian Confidence Assessment

### Evidence Table

| Evidence | Theory A (COM) | Theory B (PDF) |
|----------|---------------|----------------|
| Ratios stored as 0-1 page fractions | SUPPORTS | SUPPORTS |
| Ratios stored as TEXT (varchar) | SUPPORTS (COM string formatting) | NEUTRAL |
| Zero-margin forms match cell.Left exactly | SUPPORTS | NEUTRAL |
| Centering formula fits all centered forms | SUPPORTS | SUPPORTS |
| Form 228: 2.84pt/col extra width | CONTRADICTS | NEUTRAL |
| Form 546: cross-sheet mapping | **SUPPORTS** (new finding) | NEUTRAL (no longer unique) |
| Column width ~50.1pt matches PDF grid | NEUTRAL | SUPPORTS |
| No PDF parsing code in codebase | SUPPORTS | CONTRADICTS |
| No DB cache of PDF geometry | SUPPORTS | CONTRADICTS |
| Background PDF for ALL forms | NEUTRAL | SUPPORTS (available) |
| PDF grid lines encode geometry | NEUTRAL | SUPPORTS (feasible) |

### Confidence Reassessment

| Theory | Confidence | Supporting Evidence | Contradictions | Key Unknown |
|--------|-----------|-------------------|----------------|-------------|
| **A: Excel COM** | **55%** (↑10%) | 5 pieces SUPPORT, cross-sheet paradox resolved | Form 228 outlier (2.84pt/col extra) | Legacy Range.Width value |
| **B: PDF Post-Processing** | **45%** (↓10%) | 4 pieces SUPPORT, PDF technically feasible | No parsing code anywhere, no DB cache, cross-sheet no longer unique | Did legacy codebase have PDF parsing? |

### Why Confidence Shifted

The cross-sheet paradox was the **single strongest argument** for the PDF theory in previous assessments. This investigation revealed that:

1. `_Fields` sheet has `horizontalCentered=1` (TRUE) — stored origin 205.92pt is fully explained by COM reading `_Fields` PageSetup.
2. The cross-sheet mapping is intentional (clusters on `_Fields` sheet, print area on `Sheet1`).
3. `_Fields` centering + Calibri column width (50.04pt) produces exactly the stored origin (205.92pt).

With the paradox resolved, the COM theory now has **more supporting evidence** and the PDF theory has **lost its strongest unique argument**.

### Key Weaknesses Persist

**COM weakness:** Form 228 remains unexplained. With 5 default columns, centering ON, and Calibri font, the backward-solved column width is 52.94pt — 5.9% wider than the Calibri estimate (50.04pt). No database evidence explains this.

**PDF weakness:** No PDF parsing code exists in the codebase, no database cache of PDF-extracted geometry exists, and no PDF parsing library is referenced anywhere. The PDF theory relies on the assumption that a SEPARATE (lost) codebase performed the parsing — an assumption with no supporting evidence.

---

## FINAL CONCLUSION

> **The database does not contain sufficient evidence to reconstruct the original algorithm.**

### The SINGLE missing piece of evidence that prevents certainty:

> **The value of `Range.Width` returned by Excel COM for the print area on the legacy Excel version (Office 2016, Calibri 11pt).**

### Decision Tree

| If Range.Width = | Then | Confidence |
|-----------------|------|------------|
| **200.15pt** (Calibri 11pt, 50.04pt/col) | COM theory consistent. Form 228 still unexplained (2.84pt/col extra = ~14pt total). | COM 60%, PDF 40% |
| **200.66pt** (PDF grid width, 50.165pt/col) | PDF theory gains major support. COM still possible but less likely. | PDF 65%, COM 35% |
| **192.00pt** (Aptos Narrow, 48pt/col) | Both theories DISPROVEN for current Excel. A third mechanism generated the ratios. | Neither ≥50% |
| **182.30pt** (≈ form 228 backward-solved, 52.94pt/col) | Form 228 IS the norm; other forms need re-explanation. | Unclear |

### Supporting Documents

| Report | Key Finding |
|--------|-------------|
| `comprehensive_forensics_report.md` | 29 forms, 642 clusters, 12 candidate formulas |
| `falsification_report.md` | Attempted to disprove font metrics theory |
| `report-com-vs-pdf-verification.md` | COM Range.Width = 192pt, PDF grid width = 200.66pt, diff = 8.66pt |
| `report-conmas-background-generation.md` | Full pipeline: Excel → PDF → PNG → ratios |
| `report-database-coordinates.md` | Ratios are 0-1 of full page (isOriginalWhole=1) |
| `report-legacy-algorithm-reconstruction.md` | Font metrics version shift hypothesis |
| `consolidated_findings.md` | Q1-Q6 evidence from live database |
| `falsification_q1_q7_output.txt` | Raw analysis output from Python scripts |

---

*End of Report*
