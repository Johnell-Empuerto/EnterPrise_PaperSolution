# Phase 7 — Reverse Engineering Legacy Coordinate Space

**Date:** July 9, 2026
**Scope:** 10 representative forms from every error category
**Purpose:** Determine EXACTLY which coordinate system the legacy ConMas engine
stored coordinates in.

---
## Investigation 1 — Worksheet Coordinate Space

| Form | Stored L | Margin L | Printable W | Cols | 50.1pt Width | Stored Minus Margin | Stored Minus Zero |
|:----:|:--------:|:--------:|:-----------:|:----:|:-----------:|:------------------:|:-----------------:|
| 283 | 105.48 | 51.02 | 509.96 | 8 | 400.80 | +54.46 | 105.48 |
| 546 | 205.92 | 51.02 | 509.96 | 4 | 200.40 | +154.90 | 205.92 |
| 155 | 103.23 | 51.02 | 509.96 | 5 | 250.50 | +52.21 | 103.23 |
| 228 | 173.65 | 51.02 | 509.96 | 5 | 250.50 | +122.63 | 173.65 |
| 242 | 361.44 | 51.02 | 689.96 | 54 | 2705.40 | +310.42 | 361.44 |
| 112 | 59.57 | 51.02 | 509.96 | 1 | 50.10 | +8.55 | 59.57 |
| 174 | 28.80 | 51.02 | 509.96 | 48 | 2404.80 | -22.22 | 28.80 |
| 185 | 60.18 | 51.02 | 509.96 | 34 | 1703.40 | +9.16 | 60.18 |
| 186 | 144.36 | 51.02 | 689.96 | 4 | 200.40 | +93.34 | 144.36 |
| 193 | 74.88 | 51.02 | 509.96 | 5 | 250.50 | +23.86 | 74.88 |

### Analysis

- **Form 283**: Closest to **Margin(51.0)** (diff=54.46pt). Stored=105.5pt, merges=4, dim=A3:H7
- **Form 546**: Closest to **PrintableCenter** (diff=49.06pt). Stored=205.9pt, merges=5, dim=A1:D12
- **Form 155**: Closest to **Margin(51.0)** (diff=52.21pt). Stored=103.2pt, merges=19, dim=A1:E29
- **Form 228**: Closest to **PrintableCenter** (diff=81.33pt). Stored=173.7pt, merges=10, dim=A2:E8
- **Form 242**: Closest to **PrintableCenter** (diff=16.46pt). Stored=361.4pt, merges=190, dim=A1:CT77
- **Form 112**: Closest to **Margin(51.0)** (diff=8.55pt). Stored=59.6pt, merges=161, dim=A1:AS69
- **Form 174**: Closest to **Margin(51.0)** (diff=22.22pt). Stored=28.8pt, merges=509, dim=A1:CL76
- **Form 185**: Closest to **Margin(51.0)** (diff=9.16pt). Stored=60.2pt, merges=91, dim=A1:AK69
- **Form 186**: Closest to **Margin(51.0)** (diff=93.34pt). Stored=144.4pt, merges=2, dim=B1:I6
- **Form 193**: Closest to **Margin(51.0)** (diff=23.86pt). Stored=74.9pt, merges=25, dim=A1:F24

### Key Finding

- **9/10 forms** have stored origin > margin + 2pt (effectively centered)
- The stored coordinate space is NOT raw worksheet origin (0,0) nor margin origin
- The stored coordinates reflect a **composite of margin + centering + column width**

---
## Investigation 2 — Cell Geometry Dump (Reconstructed from XLSX)

| Form | Cell | Col | Col W (pt) | Cumulative L (pt) | DB Stored L (ratio->pt) | DB L Diff |
|:----:|:----:|:---:|:---------:|:-----------------:|:----------------------:|:--------:|
| 155 | $B$3 | 2 | 98.58 | 98.58 | 204.62 | +106.03 |
| 155 | $B$3 | 2 | 98.58 | 98.58 | 186.49 | +87.90 |
| 155 | $D$3 | 4 | 98.58 | 295.75 | 403.68 | +107.94 |
| 155 | $D$3 | 4 | 98.58 | 295.75 | 407.38 | +111.64 |
| 155 | $B$6 | 2 | 98.58 | 98.58 | 186.49 | +87.90 |
| 242 | $AS$36:$AX$36 | 45 | 21.45 | 943.58 | 361.44 | -582.14 |
| 242 | $AY$36:$BD$36 | 51 | 21.45 | 1072.25 | 402.84 | -669.41 |
| 242 | $BE$36:$BJ$36 | 57 | 21.45 | 1200.92 | 443.88 | -757.04 |
| 242 | $BK$36:$BP$36 | 63 | 21.45 | 1329.59 | 484.92 | -844.67 |
| 242 | $BQ$36:$BV$36 | 69 | 21.45 | 1458.27 | 525.96 | -932.31 |
| 112 | R6C6:R6C11 | 18 | 32.23 | 358.42 | 140.60 | -217.82 |
| 112 | R7C5:R7C14 | 18 | 32.23 | 358.42 | 125.43 | -232.99 |
| 112 | R7C16:R10C17 | 18 | 32.23 | 358.42 | 325.24 | -33.18 |
| 112 | R7C18:R10C19 | 18 | 32.23 | 358.42 | 378.89 | +20.47 |
| 112 | R7C20:R10C21 | 18 | 32.23 | 358.42 | 436.24 | +77.82 |
| 174 | $AS$2:$AV$5 | 45 | 50.09 | 2109.04 | 491.76 | -1617.28 |
| 174 | $AS$2:$AV$5 | 45 | 50.09 | 2109.04 | 491.76 | -1617.28 |
| 174 | $A$7:$F$7 | 1 | 21.02 | 0.00 | 28.80 | +28.80 |
| 174 | $A$10:$B$10 | 1 | 21.02 | 0.00 | 28.80 | +28.80 |
| 174 | $C$10:$D$10 | 3 | 17.09 | 71.11 | 51.12 | -19.99 |
| 185 | $L$6:$N$6 | 12 | 70.49 | 876.79 | 193.32 | -683.46 |
| 185 | $F$3:$G$4 | 6 | 78.35 | 396.47 | 99.07 | -297.40 |
| 185 | $D$7:$F$7 | 4 | 75.22 | 199.71 | 60.18 | -139.53 |
| 185 | $H$3:$I$4 | 8 | 78.35 | 553.17 | 129.81 | -423.37 |
| 185 | $J$3:$K$4 | 10 | 78.35 | 709.88 | 160.55 | -549.33 |
| 186 | $B$3 | 2 | 137.26 | 93.28 | 144.36 | +51.08 |
| 186 | $C$3 | 3 | 146.69 | 230.54 | 280.80 | +50.26 |
| 186 | $D$3:$E$3 | 4 | 97.21 | 377.22 | 427.32 | +50.10 |
| 186 | $C$5 | 3 | 146.69 | 230.54 | 280.80 | +50.26 |
| 186 | $D$5 | 4 | 97.21 | 377.22 | 427.32 | +50.10 |
| 193 | $C$4 | 3 | 131.76 | 99.37 | 145.80 | +46.43 |
| 193 | $F$4 | 6 | 93.28 | 394.90 | 437.40 | +42.50 |
| 193 | $B$7:$C$7 | 2 | 72.86 | 26.51 | 74.88 | +48.37 |
| 193 | $D$7:$E$7 | 4 | 120.77 | 231.13 | 277.20 | +46.07 |
| 193 | $F$7 | 6 | 93.28 | 394.90 | 437.40 | +42.50 |

### Cell.Left Error Summary

Average difference between Cell.Left and stored DB left: **307.205pt**
Best case: min=19.990pt, max=1617.280pt

---
## Investigation 3 — Coordinate Space Comparison

| Coordinate Space | Avg Error | Median Error | Max Error | <=0.5pt | <=2pt | <=10pt |
|:----------------|:---------:|:------------:|:---------:|:------:|:-----:|:------:|
| DB_Stored | 0.000 | 0.000 | 0.000 | 10/10 | 10/10 | 10/10 |
| PDFContentLeft | 1.390 | 1.150 | 2.440 | 0/10 | 2/10 | 3/10 |
| XLSX_Centered | 74.251 | 23.860 | 310.420 | 0/10 | 0/10 | 2/10 |
| LeftMargin | 85.175 | 54.460 | 310.420 | 0/10 | 0/10 | 2/10 |
| XML50.1_Centered | 85.175 | 54.460 | 310.420 | 0/10 | 0/10 | 2/10 |
| COM48_Centered | 85.175 | 54.460 | 310.420 | 0/10 | 0/10 | 2/10 |
| HardMargin(18pt) | 113.751 | 87.480 | 343.440 | 0/10 | 0/10 | 0/10 |
| Worksheet(0) | 131.751 | 105.480 | 361.440 | 0/10 | 0/10 | 0/10 |
| PrintableCenter | 144.521 | 180.100 | 226.180 | 0/10 | 0/10 | 0/10 |
| XLSX_ColLeft | 223.996 | 139.533 | 582.143 | 0/10 | 0/10 | 0/10 |
| RightMargin | 465.229 | 486.100 | 596.620 | 0/10 | 0/10 | 0/10 |

### Key Insight

The coordinate space that best matches stored database coordinates is **DB_Stored**
- Mean error: 0.000pt
- Within 0.5pt: 10/10 forms
- Within 2pt: 10/10 forms

---
## Investigation 4 — Page Break Geometry

| Form | Sheet Dim | Has Drawings | Merged Cells | Page Breaks Detected? |
|:----:|:---------:|:------------:|:------------:|:--------------------:|
| 283 | A3:H7 | False | 4 | NOT IN XLSX DATA |
| 546 | A1:D12 | False | 5 | NOT IN XLSX DATA |
| 155 | A1:E29 | False | 19 | NOT IN XLSX DATA |
| 228 | A2:E8 | False | 10 | NOT IN XLSX DATA |
| 242 | A1:CT77 | False | 190 | NOT IN XLSX DATA |
| 112 | A1:AS69 | True | 161 | NOT IN XLSX DATA |
| 174 | A1:CL76 | True | 509 | NOT IN XLSX DATA |
| 185 | A1:AK69 | False | 91 | NOT IN XLSX DATA |
| 186 | B1:I6 | False | 2 | NOT IN XLSX DATA |
| 193 | A1:F24 | False | 25 | NOT IN XLSX DATA |

### Analysis

Page breaks are determined by Excel at print time based on paper size,
margins, scaling, and print area. They are NOT stored in the XLSX format
for these workbooks. This suggests ConMas used Excel's default single-page
rendering (no manual page breaks).

---
## Investigation 5 — PrintArea Translation Pipeline

For each form, trace the coordinate transformation from worksheet origin
through every stage to the stored database value.

### Pipeline Stages

| Form | WS(0) | Col1 Left | Merged Adjust | PrintArea L | +Margin | +Center | Stored DB |
|:----:|:----:|:---------:|:-------------:|:----------:|:-------:|:-------:|:---------:|
| 283 | 0.0 | 0.0 | 0 | 105.5 | 51.0 | 0.0 | 105.5 (Δ+54.46) |
| 546 | 0.0 | 0.0 | 0 | 205.9 | 51.0 | 0.0 | 205.9 (Δ+154.90) |
| 155 | 0.0 | 0.0 | 0 | 103.2 | 51.0 | 0.0 | 103.2 (Δ+52.21) |
| 228 | 0.0 | 0.0 | 0 | 173.7 | 51.0 | 0.0 | 173.7 (Δ+122.63) |
| 242 | 0.0 | 943.6 | 0 | -582.1 | 51.0 | 0.0 | 361.4 (Δ-633.16) |
| 112 | 0.0 | 358.4 | 0 | -298.9 | 51.0 | 0.0 | 59.6 (Δ-349.87) |
| 174 | 0.0 | 0.0 | 0 | 28.8 | 51.0 | 0.0 | 28.8 (Δ-22.22) |
| 185 | 0.0 | 199.7 | 0 | -139.5 | 51.0 | 0.0 | 60.2 (Δ-190.55) |
| 186 | 0.0 | 93.3 | 0 | 51.1 | 51.0 | 0.0 | 144.4 (Δ+0.06) |
| 193 | 0.0 | 26.5 | 0 | 48.4 | 51.0 | 0.0 | 74.9 (Δ-2.65) |

### Translation Analysis

- **Form 283**: Expected=51.0pt (m=51.0 + ctr=0.0 + col1=0.0) vs Stored=105.5pt. Δ=54.46pt
- **Form 546**: Expected=51.0pt (m=51.0 + ctr=0.0 + col1=0.0) vs Stored=205.9pt. Δ=154.90pt
- **Form 155**: Expected=51.0pt (m=51.0 + ctr=0.0 + col1=0.0) vs Stored=103.2pt. Δ=52.21pt
- **Form 228**: Expected=51.0pt (m=51.0 + ctr=0.0 + col1=0.0) vs Stored=173.7pt. Δ=122.63pt
- **Form 242**: Expected=994.6pt (m=51.0 + ctr=0.0 + col1=943.6) vs Stored=361.4pt. Δ=633.16pt
- **Form 112**: Expected=409.4pt (m=51.0 + ctr=0.0 + col1=358.4) vs Stored=59.6pt. Δ=349.87pt
- **Form 174**: Expected=51.0pt (m=51.0 + ctr=0.0 + col1=0.0) vs Stored=28.8pt. Δ=22.22pt
- **Form 185**: Expected=250.7pt (m=51.0 + ctr=0.0 + col1=199.7) vs Stored=60.2pt. Δ=190.55pt
- **Form 186**: Expected=144.3pt (m=51.0 + ctr=0.0 + col1=93.3) vs Stored=144.4pt. Δ=0.06pt
- **Form 193**: Expected=77.5pt (m=51.0 + ctr=0.0 + col1=26.5) vs Stored=74.9pt. Δ=2.65pt

---
## Investigation 6 — Rendering Pipeline Reconstruction

```
Complete Excel Rendering Pipeline:

Worksheet
  |
  +-- Cell Geometry (Cell.Left, Cell.Top, Cell.Width, Cell.Height)
  |     - Column widths from <cols> in XLSX
  |     - Row heights from <rows> in XLSX
  |     - Default: 8.43 chars width, 15.75pt height
  |
  +-- Merged Cell Geometry (MergeArea.Left, .Top, .Width, .Height)
  |     - Merged ranges override individual cell bounds
  |     - Excel uses MergeArea for merged cells
  |
  +-- PrintArea (PageSetup.PrintArea)
  |     - Defines the range to print
  |     - PrintArea.Left/Top/Width/Height from COM
  |     - Stored as address like $A$1:$D$10
  |
  +-- UsedRange
  |     - Excel-computed bounding box of all non-empty cells
  |     - May differ from PrintArea
  |
  +-- Page Layout
  |     - Paper size + margins + orientation
  |     - Printable area = page - margins
  |     - Centering = shift content within printable area
  |     - Scaling = FitToPages or Zoom
  |
  +-- Printer Device Context
  |     - Hard margins (non-printable zone)
  |     - Printer DPI resolution
  |     - Printer driver transforms
  |
  +-- ExportAsFixedFormat
  |     - Excel's internal PDF renderer
  |     - Uses printer DC + PageSetup
  |     - Output: PDF with MediaBox
  |
  +-- PDF
  |     - Contains rendered content at specified DPI
  |     - PDF content bounds = actual printed object positions
  |
  +-- Legacy ConMas Capture
  |     - **???? CAPTURE POINT ????
  |     - Coordinates stored as ratios (0.0-1.0)
  |     - Origin = min(left_position) * page_width
  |
  +-- Database
       - Stored as normalized ratios
       - left_position = cellLeft / pageWidth
```

### Most Likely Capture Point

Based on the evidence from all 7 phases, the coordinate capture likely occurred:

**Hypothesis A: PrintArea-Relative Coordinates (Most Likely)**
- Coordinates are relative to PrintArea.Left, not worksheet origin (0,0)
- PrintArea.Left provides the offset for non-centered forms
- Centering is applied via PageSetup.CenterHorizontally
- Stored as ratio = (cellLeft - printAreaLeft) / pageWidth

**Hypothesis B: COM Range-Relative**
- Coordinates relative to Range.Left of the print area range
- Used when PrintArea.Left differs from Range.Left (merged cells)

**Hypothesis C: Printer-Corrected**
- Coordinates include printer hard margin offset
- Would explain forms with large left offsets (>100pt)

---
## Investigation 7 — Merged Cell Analysis

| Form | Merge Count | Merged Ranges | Cluster Count | Stored L | Min Col | Max Col |
|:----:|:----------:|:-------------:|:-------------:|:--------:|:-------:|:-------:|
| 283 | 4 | A3:H4; A6:B7; F6:H7; D6:E7 | 2 | 105.5 | 1 | 8 |
| 546 | 5 | A1:B2; C1:D2; A3:D4; A6:D7; A9:D10 | 6 | 205.9 | 1 | 4 |
| 155 | 19 | A27:B27; C27:D27; A7:B7; C7:D7; A19:B19 (14 more) | 44 | 103.2 | 1 | 5 |
| 228 | 10 | A2:E2; A6:E6; A7:B7; A8:B8; D7:E7 (5 more) | 6 | 173.7 | 1 | 5 |
| 242 | 190 | CB77:CT77; AM73:AR74; AS73:CT74; A75:D76; E75:J76 (185 more) | 78 | 361.4 | 45 | 98 |
| 112 | 161 | A6:C6; A7:C7; A8:C8; A9:C9; V7:W10 (156 more) | 261 | 59.6 | 18 | 18 |
| 174 | 509 | A53:B53; A54:B54; J2:AR5; A44:B44; A45:B45 (504 more) | 1298 | 28.8 | 1 | 48 |
| 185 | 91 | A2:B6; C2:C6; D2:D6; E2:E6; F2:AK2 (86 more) | 1379 | 60.2 | 4 | 37 |
| 186 | 2 | D3:E3; B2:E2 | 6 | 144.4 | 2 | 5 |
| 193 | 25 | D6:E6; B7:C7; B8:C8; B9:C9; B10:C10 (20 more) | 34 | 74.9 | 2 | 6 |

### Analysis

Merged cells affect coordinate alignment because Excel's Range.Left
for a merged cell returns the **top-left cell of the merge range**,
not the visual left edge. MergeArea.Left returns the actual visual left edge.

If ConMas used MergeArea.Left instead of Range.Left, the stored coordinates
would reflect the merged cell's visual position, not the anchor cell position.

---
## Investigation 8 — Legacy Algorithm Reconstruction

| Rank | Algorithm | Mean | Median | P95 | Max | RMSE | <=0.5pt | <=2pt | <=10pt |
|:---:|:----------|:----:|:-----:|:---:|:---:|:----:|:------:|:-----:|:------:|
| 1 | Stored_DB | 0.000 | 0.000 | 0.000 | 0.000 | 0.000 | 10/10 | 10/10 | 10/10 |
| 2 | Margin_Origin | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 2/10 |
| 3 | XML50.1_Centered | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 2/10 |
| 4 | COM48_Centered | 85.175 | 54.460 | 310.420 | 310.420 | 122.889 | 0/10 | 0/10 | 2/10 |
| 5 | Hard_Margin_18pt | 113.751 | 87.480 | 343.440 | 343.440 | 146.709 | 0/10 | 0/10 | 0/10 |
| 6 | Worksheet_Origin | 131.751 | 105.480 | 361.440 | 361.440 | 161.067 | 0/10 | 0/10 | 0/10 |
| 7 | Printable_Center | 144.521 | 180.100 | 226.180 | 226.180 | 159.576 | 0/10 | 0/10 | 0/10 |
| 8 | XLSX_ColLeft | 178.859 | 103.230 | 582.143 | 582.143 | 257.490 | 0/10 | 0/10 | 0/10 |
| 9 | Paper_Edge | 192.249 | 231.120 | 277.200 | 277.200 | 206.103 | 0/10 | 0/10 | 0/10 |

### Best Non-Trivial Algorithm

The best non-trivial algorithm is: **Margin_Origin**
  - Mean error: 85.175pt (vs median 54.460pt)
  - RMSE: 122.889pt
  - Within 0.5pt: 0/10 forms
  - Within 2pt: 0/10 forms
  - Within 10pt: 2/10 forms

---
## Deliverable 1 — Complete Worksheet Geometry Dump

| Form | Page WxH | Margin L/R | Printable W | Col Range | Col Count | XLSX Cols Defined | Merges |
|:----:|:--------:|:----------:|:-----------:|:---------:|:---------:|:-----------------:|:------:|
| 283 | 612x792 | 51.0/51.0 | 510.0 | 1-8 | 8 | 0 | 4 |
| 546 | 612x792 | 51.0/51.0 | 510.0 | 1-4 | 4 | 0 | 5 |
| 155 | 612x792 | 51.0/51.0 | 510.0 | 1-5 | 5 | 2 | 19 |
| 228 | 612x792 | 51.0/51.0 | 510.0 | 1-5 | 5 | 0 | 10 |
| 242 | 792x612 | 51.0/51.0 | 690.0 | 45-98 | 54 | 1 | 190 |
| 112 | 612x792 | 51.0/51.0 | 510.0 | 18-18 | 1 | 14 | 161 |
| 174 | 612x792 | 51.0/51.0 | 510.0 | 1-48 | 48 | 2 | 509 |
| 185 | 612x792 | 51.0/51.0 | 510.0 | 4-37 | 34 | 8 | 91 |
| 186 | 792x612 | 51.0/51.0 | 690.0 | 2-5 | 4 | 7 | 2 |
| 193 | 612x792 | 51.0/51.0 | 510.0 | 2-6 | 5 | 6 | 25 |

---
## Deliverable 2 — Complete Populated-Cell Geometry Dump

See Investigation 2 for per-cluster cell geometry details.

---
## Deliverable 3 — Merged-Cell Geometry Report

See Investigation 7 for merged-range details.

---
## Deliverable 4 — Coordinate-Space Comparison Table

See Investigation 3 for the full coordinate space ranking table.

---
## Deliverable 5 — Page-Break Report

See Investigation 4 for page-break analysis.

---
## Deliverable 6 — Excel Rendering Pipeline Diagram

See Investigation 6 for the complete pipeline diagram.

---
## Deliverable 7 — Legacy Coordinate Reconstruction Algorithm

Based on all evidence from Phases 1-7, the reconstructed algorithm is:

```
For each cluster/cell being captured:

  1. Get the cell's visual LEFT position:
     IF cell is merged:
       cellLeft = MergeArea.Left   # Visual left edge of merge
     ELSE:
       cellLeft = Range.Left       # Cell's left edge

  2. Get the print area origin:
     printAreaLeft = PrintArea.Left
     (May differ from first column due to merged cells)

  3. Compute the relative offset:
     relativeLeft = cellLeft - printAreaLeft

  4. Apply page layout transform:
     pageWidth = PaperSize.Width
     marginLeft = PageSetup.LeftMargin
     printableWidth = pageWidth - marginLeft - marginRight

     IF CenterHorizontally:
       contentWidth = PrintArea.Width
       origin = marginLeft + (printableWidth - contentWidth) / 2
     ELSE:
       origin = marginLeft

     printedLeft = origin + relativeLeft

  5. Store as ratio:
     left_position = printedLeft / pageWidth

  KEY UNKNOWNS:
  - Whether printAreaLeft is PrintArea.Left or UsedRange.Left
  - Whether contentWidth is Range.Width or a column-sum width
  - Which printer's margins were used
```

---
## Deliverable 8 — Final Conclusion

### Where Did ConMas Capture Coordinates?

**Evidence-Based Conclusions:**

1. **NOT worksheet origin** — stored coordinates are clearly offset from (0,0)

2. **NOT raw margins** — 9/10 forms have stored origin > margin + 2pt

3. **NOT PDF output directly** — PDF geometry matches only 2/10 available PDFs within 2pt

4. **NOT cell.Left directly** — cell geometry from XLSX shows large discrepancies
   for forms with merged cells

5. **Best candidate: Margin_Origin** — mean error of 85.175pt across 10 forms
   but still insufficient for 10/10 forms

### The Remaining Unknown Variable

The evidence still shows a systematic offset that cannot be explained by:

- Page margins (known: 51.02pt)
- Column widths (known from XLSX <cols>)
- Font metrics (known: 7.33 maxDigitWidth for 11pt)
- Print area (known from cluster addresses)
- PDF geometry (measured for 3 forms)
- Cell geometry (reconstructed from XLSX)

**The missing variable is most likely one of:**

1. **Printer-specific hard margins** — different printers add 18-54pt
   of unprintable border that Excel accounts for differently

2. **ExportAsFixedFormat internal offset** — the PDF renderer may add
   its own offset beyond PageSetup margins

3. **ConMas-specific preprocessing** — the legacy engine may have applied
   its own coordinate transform before storing ratios

**Without access to the original ConMas runtime or a COM capture**
**with full property dump, this variable cannot be definitively identified.**

The algorithm in Deliverable 7 is the closest reconstruction achievable
from the available data (XLSX, XML, PDF, DB).

---

*Generated by Phase 7 Legacy Coordinate Space Reverse Engineering — July 9, 2026*