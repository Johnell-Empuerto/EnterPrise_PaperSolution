# Phase 6 — Reverse Engineering Legacy Coordinate Generation

**Date:** July 9, 2026
**Scope:** 10 representative forms from every error category
**Purpose:** Determine exactly how the legacy ConMas engine generated the
database coordinates.

---
## Investigation 1 — COM Geometry Dump

| Property | Form 283 | Form 546 | Form 155 | Form 228 | Form 242 | Form 112 | Form 174 | Form 185 | Form 186 | Form 193 |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **Page W (pt)** | 612.00 | 612.00 | 612.00 | 612.00 | 792.00 | 612.00 | 612.00 | 612.00 | 792.00 | 612.00 |
| **Page H (pt)** | 792.00 | 792.00 | 792.00 | 792.00 | 612.00 | 792.00 | 792.00 | 792.00 | 612.00 | 792.00 |
| **Left Margin (pt)** | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 |
| **Right Margin (pt)** | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 |
| **Top Margin (pt)** | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 |
| **Bottom Margin (pt)** | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 | 51.02 |
| **Header Margin (pt)** | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| **Footer Margin (pt)** | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| **Center Horizontally** | No | No | No | No | No | No | No | No | No | No |
| **Center Vertically** | No | No | No | No | No | No | No | No | No | No |
| **Print Area** |  |  |  |  |  |  |  |  |  |  |
| **Print Title Rows** |  |  |  |  |  |  |  |  |  |  |
| **Print Title Cols** |  |  |  |  |  |  |  |  |  |  |
| **Zoom (%)** |  |  |  |  |  |  |  |  |  |  |
| **Fit W** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **Fit H** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **First Page #** |  |  |  |  |  |  |  |  |  |  |
| **Print Quality (DPI)** | 600 | 600 | 600 | 600 | 600 | 600 | 600 | 600 | 600 | 600 |
| **B&W** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **Draft** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **Print Order** |  |  |  |  |  |  |  |  |  |  |
| **Print Gridlines** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **Print Headings** | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| **Paper Size** |  |  |  |  |  |  |  |  |  |  |
| **Orientation** |  |  |  |  |  |  |  |  |  |  |

### Analysis

- All forms use consistent **51.02pt margins** (ConMas default)
- **0/10 forms** have centering enabled in the ConMas XML
- **0/10 forms** have FitToPages or Zoom adjustments
- No print titles, gridlines, headings, or page numbering configured

---
## Investigation 2 — Printer Geometry & Scaling

| Form | Page W | Soft Margin | Hard Margin | Soft Printable | Hard Printable | PDF L | PDF W | Scale(Soft) | Scale(Hard) |
|:----:|:-----:|:-----------:|:-----------:|:--------------:|:--------------:|:-----:|:-----:|:----------:|:----------:|
| 283 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 104.9 | 401.8 | 0.787983 | 0.697639 |
| 546 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 204.8 | 201.6 | 0.395364 | 0.350035 |
| 155 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 228 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 176.1 | 259.0 | 0.507844 | 0.449618 |
| 242 | 792 | 51.0 | 18.0 | 690.0 | 756.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 155 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 155 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 155 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 242 | 792 | 51.0 | 18.0 | 690.0 | 756.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |
| 155 | 612 | 51.0 | 18.0 | 510.0 | 576.0 | 0.0 | 0.0 | 0.000000 | 0.000000 |

### Analysis

- Standard hard margins (0.25in = 18pt) do NOT match PDF content left edge for centered forms
- PDF content width is consistently **smaller than printable width** — content doesn't fill the page

---
## Investigation 4 — Shape Bounding Rectangle

| Form | Has Drawings | Shape Count | Shape BBox Left | Shape BBox Right | Shape BBox Width |
|:----:|:-----------:|:-----------:|:---------------:|:----------------:|:----------------:|
| 283 | False | 0 |  |  |  |
| 546 | False | 0 |  |  |  |
| 155 | False | 0 |  |  |  |
| 228 | False | 0 |  |  |  |
| 242 | False | 0 |  |  |  |
| 112 | True | 1 | 15.0 | 132.65 | 117.65 |
| 174 | True | 0 |  |  |  |
| 185 | False | 0 |  |  |  |
| 186 | False | 0 |  |  |  |
| 193 | False | 0 |  |  |  |

### Analysis

- **2/10 forms** have drawing objects
  - Form 112: shape at L=15.0 T=5.2 117.7x40.5

---
## Investigation 5 — Hidden Geometry

| Form | Hidden Rows | Hidden Cols | Merged Cells | Outline Groups |
|:----:|:----------:|:-----------:|:------------:|:--------------:|
| 283 | 0 | 0 | 4 | 0 |
| 546 | 0 | 0 | 5 | 0 |
| 155 | 0 | 0 | 19 | 0 |
| 228 | 0 | 0 | 10 | 0 |
| 242 | 0 | 0 | 190 | 0 |
| 112 | 0 | 0 | 161 | 0 |
| 174 | 0 | 0 | 876 | 0 |
| 185 | 0 | 0 | 91 | 0 |
| 186 | 0 | 0 | 2 | 0 |
| 193 | 0 | 0 | 25 | 0 |

### Analysis

- **10/10 forms** have merged cells — could affect Range.Width/Left
- **0/10 forms** have hidden rows/columns

---
## Investigation 6 — Print Pipeline Comparison (4 Rectangles)

| Form | Rect A: DB Stored (L/W) | Rect B: Printable (L/W) | Rect C: PDF (L/W) | Rect D: Engine (L/W) | Best Match | Diff |
|:----:|:----------------------:|:----------------------:|:-----------------:|:-------------------:|:----------:|:----:|
| 283 | 105.5/401.0 | 51.0/510.0 | 104.9/401.8 | 51.0/400.8 | PrintArea (DB) | 0.000 |
| 546 | 205.9/200.2 | 51.0/510.0 | 204.8/201.6 | 51.0/200.4 | PrintArea (DB) | 0.000 |
| 155 | 103.2/430.7 | 51.0/510.0 | 0.0/0.0 | 51.0/250.5 | PrintArea (DB) | 0.000 |
| 228 | 173.7/262.3 | 51.0/510.0 | 176.1/259.0 | 51.0/250.5 | PrintArea (DB) | 0.000 |
| 242 | 361.4/369.7 | 51.0/690.0 | 0.0/0.0 | 51.0/2705.4 | PrintArea (DB) | 0.000 |
| 112 | 59.6/489.9 | 51.0/510.0 | 0.0/0.0 | 51.0/50.1 | PrintArea (DB) | 0.000 |
| 174 | 28.8/505.1 | 51.0/510.0 | 0.0/0.0 | 51.0/2404.8 | PrintArea (DB) | 0.000 |
| 185 | 60.2/530.7 | 51.0/510.0 | 0.0/0.0 | 51.0/1703.4 | PrintArea (DB) | 0.000 |
| 186 | 144.4/467.6 | 51.0/690.0 | 0.0/0.0 | 51.0/200.4 | PrintArea (DB) | 0.000 |
| 193 | 74.9/454.0 | 51.0/510.0 | 0.0/0.0 | 51.0/250.5 | PrintArea (DB) | 0.000 |

### Best Rectangle Summary

- **PrintArea (DB)**: 10 forms

---
## Investigation 7 — Database Coordinate Reconstruction (Back-Solve)

| Form | Stored L | Margin | Printable W | Cols | Solved Content W | Solved Margin | Solved Printable W | Is Centered? |
|:----:|:--------:|:------:|:-----------:|:----:|:---------------:|:------------:|:-----------------:|:-----------:|
| 283 | 105.48 | 51.02 | 509.96 | 8 | 401.04 | 105.48 | 509.72 | YES |
| 546 | 205.92 | 51.02 | 509.96 | 4 | 200.16 | 205.92 | 510.2 | YES |
| 155 | 103.23 | 51.02 | 509.96 | 5 | 405.53 | 103.23 | 354.93 | YES |
| 228 | 173.65 | 51.02 | 509.96 | 5 | 264.7 | 173.65 | 495.76 | YES |
| 242 | 361.44 | 51.02 | 689.96 | 54 | 69.12 | 361.44 | 3326.24 | YES |
| 112 | 59.57 | 51.02 | 509.96 | 1 | 492.86 | 59.57 | 67.2 | YES |
| 174 | 28.80 | 51.02 | 509.96 | 48 | N/A | 28.8 | N/A | NO |
| 185 | 60.18 | 51.02 | 509.96 | 34 | 491.64 | 60.18 | 1721.72 | YES |
| 186 | 144.36 | 51.02 | 689.96 | 4 | 503.28 | 144.36 | 387.08 | YES |
| 193 | 74.88 | 51.02 | 509.96 | 5 | 462.24 | 74.88 | 298.22 | YES |

### Solution Clusters

- **Back-solved margin**: avg=131.75pt, min=28.80pt, max=361.44pt
- **Back-solved content width**: avg=365.62pt, min=69.12pt, max=503.28pt

---
## Investigation 8 — Legacy Algorithm Ranking

| Rank | Candidate | Mean | Median | P95 | Max | RMSE | <=0.5pt | <=1pt |
|:---:|:----------|:----:|:-----:|:---:|:---:|:----:|:------:|:-----:|
| 1 | Stored_DB | 0.000 | 0.000 | 0.000 | 0.000 | 0.000 | 10/10 | 10/10 |
| 2 | Stored_Minus_Margin | 51.020 | 51.020 | 51.020 | 51.020 | 51.020 | 0/10 | 0/10 |
| 3 | PDF_Measured | 52.394 | 22.220 | 310.420 | 310.420 | 104.416 | 0/10 | 1/10 |
| 4 | Margin | 85.176 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 |
| 5 | XML_50.1_Centered | 85.176 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 |
| 6 | COM_48_Centered | 85.176 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 |
| 7 | XLSX_Centered | 85.176 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 |
| 8 | Hard_Margin_18pt | 113.752 | 87.480 | 343.440 | 343.440 | 146.709 | 0/10 | 0/10 |
| 9 | HardMargin_Centered | 113.752 | 87.480 | 343.440 | 343.440 | 146.709 | 0/10 | 0/10 |

### Per-Form Best Algorithm

| Form | Stored L | Best Algorithm | Best Origin | Error | 2nd Best | 2nd Error |
|:----:|:--------:|:--------------:|:-----------:|:-----:|:--------:|:---------:|
| 283 | 0.00 | Stored_DB | 105.48 | 0.000 | PDF_Measured | 0.580 |
| 546 | 0.00 | Stored_DB | 205.92 | 0.000 | PDF_Measured | 1.150 |
| 155 | 0.00 | Stored_DB | 103.23 | 0.000 | Stored_Minus_Margin | 51.020 |
| 228 | 0.00 | Stored_DB | 173.65 | 0.000 | PDF_Measured | 2.442 |
| 242 | 0.00 | Stored_DB | 361.44 | 0.000 | Stored_Minus_Margin | 51.020 |
| 112 | 0.00 | Stored_DB | 59.57 | 0.000 | Margin | 8.552 |
| 174 | 0.00 | Stored_DB | 28.80 | 0.000 | Hard_Margin_18pt | 10.800 |
| 185 | 0.00 | Stored_DB | 60.18 | 0.000 | Margin | 9.162 |
| 186 | 0.00 | Stored_DB | 144.36 | 0.000 | Stored_Minus_Margin | 51.020 |
| 193 | 0.00 | Stored_DB | 74.88 | 0.000 | Margin | 23.860 |

---
## Deliverable 1 — Complete COM Property Dump

See Investigation 1 table above for the full property dump of every
representative form. All properties were extracted from the ConMas XML data.

---
## Deliverable 2 — Printer Geometry Report

| Form | Standard HW (18pt) | Soft Margin (51.02pt) | PDF Left | Which Matches? |
|:----:|:-----------------:|:--------------------:|:--------:|:--------------:|
| 283 | 18.0pt | 51.02pt | 104.9pt | SOFT |
| 546 | 18.0pt | 51.02pt | 204.8pt | SOFT |
| 155 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 228 | 18.0pt | 51.02pt | 176.1pt | SOFT |
| 242 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 112 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 174 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 185 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 186 | 18.0pt | 51.02pt | 0.0pt | HARD |
| 193 | 18.0pt | 51.02pt | 0.0pt | HARD |

---
## Deliverable 3 — Shape Geometry Report

See Investigation 4 for full per-form shape details.

---
## Deliverable 4 — Rectangle Comparison

See Investigation 6 for the full 4-rectangle comparison table.

---
## Deliverable 5 — Legacy Algorithm Ranking

### Top 3 Algorithms

**1. Stored_DB**
   - Mean error: 0.000pt | Median: 0.000pt | RMSE: 0.000pt
   - Within 0.5pt: 10/10 forms | Within 1pt: 10/10 forms

**2. Stored_Minus_Margin**
   - Mean error: 51.020pt | Median: 51.020pt | RMSE: 51.020pt
   - Within 0.5pt: 0/10 forms | Within 1pt: 0/10 forms

**3. PDF_Measured**
   - Mean error: 52.394pt | Median: 22.220pt | RMSE: 104.416pt
   - Within 0.5pt: 0/10 forms | Within 1pt: 1/10 forms

---
## Deliverable 6 — Root Cause Matrix

| Root Cause | Forms Affected | Impact | Evidence |
|:-----------|:-------------:|:-----:|:---------|
| Already Solved | 10 forms ([283, 546, 155, 228, 242, 112, 174, 185, 186, 193]) | HIGH | Best algorithm matched this category |

### Final Assessment

Based on all 8 investigations across 10 representative forms:

1. **The best candidate equation is: Stored_DB** (mean error: 0.000pt)

2. **No single equation achieves 0.5pt accuracy for all forms** — multiple variables interact

3. **The missing variable is most likely**:
   - A **combination of printer margins + column width convention**
   - Neither margin alone nor column width alone explains all forms

4. **Recommended next step**:
   - Deploy Phases 1-2 (XLSX column widths + worksheet resolution)
   - Add font metric lookup from styles.xml
   - Test with candidate margin values: 54pt (0.75in), 36pt (0.5in), 56.7pt (2cm)
   - Re-run Phase 4 PDF validation after each adjustment

---

*Generated by Phase 6 Legacy Coordinate Reverse Engineering — July 9, 2026*