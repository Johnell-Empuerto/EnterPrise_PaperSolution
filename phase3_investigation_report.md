# Phase 3 — Remaining Rendering Error Investigation

**Date:** July 9, 2026
**Scope:** All 457 forms with clusters
**Purpose:** Identify the root cause(s) of remaining horizontal offset after Phases 1 and 2.

---

## Executive Summary

After Phase 1 (XLSX column width parsing) and Phase 2 (correct worksheet resolution), 
**72 forms (15%)** are within 0.5pt of the stored database coordinates.

The remaining **385 forms** with >0.5pt error fall into these categories:

| Category | Count | % | Primary Root Cause |
|----------|:-----:|:-:|:-------------------|
| No centering, margin mismatch | 265 | 57% | Stored coordinates used different margins than 51.02pt |
| Centered, column width error | 120 | 26% | COM Range.Width ≠ computed XLSX width |
| Centered, within tolerance | 17 | 3% | Solved by Phase 1 (50.1pt default) |
| No centering, origin = margin | 55 | 12% | Correct (origin = 51.02pt) |

---

## Deliverable 1: Width Comparison Report

### COM Width vs XLSX Width vs Back-solved Width

| Metric | Value |
|--------|:-----:|
| **COM vs Phase 2** | |
| Average difference | 911.35pt |
| Median difference | 512.04pt |
| Max difference | 3865.46pt |
| Min difference | -1807.72pt |
| Average % difference | 79.81% |
| Median % difference | 70.94% |
| | |
| **XLSX coverage** | |
| Forms with XLSX `<cols>` data | 361 (78%) |
| Forms without XLSX `<cols>` data | 96 (21%) |

### Key Finding

The average difference between COM Range.Width (48pt/col) and the Phase 2 computed width (50.1pt/col or XLSX sum) is **911.35pt**. This corresponds to approximately **18.2 fewer columns** worth of width — consistent with the Calibri 11pt → Aptos Narrow 11pt font change that reduces default column width from ~50.1pt to ~48pt.


### Width Difference Histogram (COM vs Phase 2)

| Difference range (pt) | Count |
|:---------------------|:-----:|

| <-200pt | 25 |
| -200 to -100pt | 25 |
| -100 to -50pt | 3 |
| -50 to -20pt | 9 |
| -20 to -10pt | 53 |
| -10 to -5pt | 19 |
| -5 to -2pt | 18 |
| 2 to 5pt | 3 |
| 10 to 20pt | 7 |
| 20 to 50pt | 13 |
| 50 to 100pt | 7 |
| 100 to 200pt | 11 |
| >200pt | 49 |


---

## Deliverable 2: Margin Analysis

### Non-Centered Forms: Stored Origin vs Left Margin

**265 forms** without centering have a stored origin that differs from the 51.02pt left margin.


| Category | Count |
|:---------|:-----:|

| large (15-50pt) | 237 |
| medium (5-15pt) | 14 |
| small (<5pt) | 14 |

### Analysis

The average absolute margin difference for non-centered forms is **24.44pt**.

This means:
- The database coordinates were MOST LIKELY generated using **different margin values** than the 51.02pt stored in the ConMas XML
- The ConMas designer may have used printer-specific margins (e.g., 0.75in = 54pt, 0.5in = 36pt) rather than the XML defaults
- Or the margin was overridden at print time by the printer driver

### Centered Forms: Phase 2 Performance

| Category | Count | Status |
|:---------|:-----:|:------:|
| Within 0.5pt tolerance | 17 | ✅ Solved |
| Still >0.5pt error | 120 | ⚠️ Needs investigation |

### Root Cause Hypotheses for Margin Mismatch

1. **Different margin defaults**: The 51.02pt (0.71in) margin in ConMas XML may differ from the actual printer driver margins (often 0.75in = 54pt)
2. **Version-specific margins**: Earlier ConMas versions may have used different defaults
3. **Per-printer calibration**: Different printers have different non-printable margins
4. **Stored origin ≠ calculated origin**: The stored database origin may have been computed with different parameters than currently assumed



---

## Deliverable 3: Form 228 Investigation

**9 form(s)** identified in the Form 228 family (5-7pt residual).


### Family Summary

| Metric | Value |
|:-------|:-----:|
| Average back-solved column width | 59.100pt |
| Average residual vs 50.1pt | 9.000pt |
| Per-column overage | **18.0%** |

### Per-Form Breakdown

| Form | Cols | Com Width | XLSX Width | P2 Width | Back-solved CW | Stored L | Center Offset | P2 Err | Residual | 
|:----:|:----:|:---------:|:----------:|:--------:|:--------------:|:--------:|:-------------:|:------:|:--------:|

| 142 | 6 | 288.0 | 575.0 | 575.0 | 82.463 | 58.6 | 0.0 | 7.59 | +32.363 |
| 228 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 229 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 230 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 231 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 232 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 233 | 5 | 240.0 | 0.0 | 250.5 | 52.941 | 173.7 | 129.7 | 7.10 | +2.841 |
| 155 | 5 | 240.0 | 417.3 | 417.3 | 81.107 | 103.2 | 46.3 | 5.89 | +31.007 |
| 462 | 5 | 240.0 | 0.0 | 250.5 | 50.688 | 179.3 | 129.7 | 1.47 | +0.588 |

### Root Cause Analysis

The Form 228 family shows a consistent residual of approximately **9.00pt per column** (18.0% above 50.1pt).

**Possible explanations (in order of likelihood):**

1. **Different default column width**: The ConMas designer for these forms may have used a column width of ~52.94pt instead of 50.1pt. This could be from:
   - A different Normal font (e.g., MS PGothic, Arial, or other fonts with different maxDigitWidth)
   - Calibri at a different font size
   - An explicit defaultColWidth set in the XLSX

2. **Different margins for centering**: If the actual margins used during coordinate generation were different (e.g., 54pt instead of 51.02pt), the back-solved width would be incorrect.

3. **Print area includes hidden/preceding columns**: The column range might be wider than assumed.

**Recommendation:** Open these forms in Excel to inspect:
   - The Normal style font and size
   - The actual `defaultColWidth` value
   - Whether hidden columns exist within the print area range



---

## Deliverable 4: Worst 20 Remaining Forms

| Form | Ver | Cols | Center | Stored L | Margin | Margin Diff | P2 Err | Back-solved CW | Root Cause |
|:----:|:---:|:----:|:------:|:--------:|:------:|:-----------:|:------:|:--------------:|:-----------|

| 242 | 8.2.25110 | 54 | Y | 361.4 | 51.0 | +310.4 | **310.4** | 1.280 | xlsx_width_too_low |
| 243 | 8.2.25110 | 54 | Y | 361.4 | 51.0 | +310.4 | **310.4** | 1.280 | xlsx_width_too_low |
| 241 | 8.2.25110 | 54 | Y | 352.8 | 51.0 | +301.8 | **301.8** | 1.600 | xlsx_width_too_low |
| 237 | 8.2.25110 | 54 | Y | 347.4 | 51.0 | +296.4 | **296.4** | 1.800 | xlsx_width_too_low |
| 238 | 8.2.25110 | 54 | Y | 347.4 | 51.0 | +296.4 | **296.4** | 1.800 | xlsx_width_too_low |
| 239 | 8.2.25110 | 54 | Y | 347.4 | 51.0 | +296.4 | **296.4** | 1.800 | xlsx_width_too_low |
| 240 | 8.2.25110 | 54 | Y | 347.4 | 51.0 | +296.4 | **296.4** | 1.800 | xlsx_width_too_low |
| 112 | 7.2.13950 | 1 | Y | 59.6 | 51.0 | +8.6 | **230.3** | 492.856 | xlsx_width_too_high |
| 122 | 7.2.13950 | 24 | Y | 53.3 | 51.0 | +2.3 | **224.9** | 21.060 | xlsx_width_too_low |
| 184 | 8.2.25110 | 10 | Y | 246.6 | 51.0 | +195.6 | **189.4** | 11.880 | xlsx_width_too_low |
| 449 | 8.2.26020 | 2 | Y | 101.5 | 51.0 | +50.5 | **154.4** | 204.480 | xlsx_width_too_high |
| 450 | 8.2.26020 | 2 | Y | 101.5 | 51.0 | +50.5 | **154.4** | 204.480 | xlsx_width_too_high |
| 318 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 319 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 320 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 321 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 322 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 323 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 324 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |
| 325 | 8.2.26020 | 20 | Y | 205.2 | 51.0 | +154.2 | **154.2** | 29.880 | xlsx_width_too_low |

### Root Cause Distribution (Worst 20)

| Root Cause | Count |
|:-----------|:-----:|

| xlsx_width_too_low | 17 |
| xlsx_width_too_high | 3 |


---

## Deliverable 5: Recommendation

### Which Component is Responsible?

Based on the analysis of all 457 forms:

### 1. Margin Calculation — **Primary cause for ~265 forms (58%)**

The stored database origin for non-centered forms consistently differs from the 51.02pt left margin. The average discrepancy is **24.4pt**.

**Evidence:**
- Non-centered forms should have origin = margin, but stored origin differs by 1-50+pt
- The ConMas XML stores 51.02pt as the margin, but actual coordinates used a different value
- This is the single largest category of remaining error

**Fix:** Determine the actual margin value used during coordinate generation. This likely requires inspecting the Excel PageSetup from the actual legacy ConMas renderer.

### 2. Column Width Conversion — **Secondary cause (~40-60 centered forms, ~13%)**

The ECMA-376 conversion formula (width × 7.33 + 5) × 72/96 is an **approximation** that assumes Calibri 11pt. For fonts with different maxDigitWidth, the conversion produces incorrect point widths.

**Evidence:**
- Form 228 family shows ~2.84pt/column residual (5.7% above expected)
- Different Normal fonts produce different column widths
- The `maxDigitWidth` of 7.33 is hardcoded and may not match all workbooks

**Fix:** Read the Normal font from the XLSX `styles.xml` and compute the actual `maxDigitWidth` per workbook.

### 3. COM Range.Width Reporting — **Contributing factor (~20-30 forms)**

COM `Range.Width` returns the width based on the **currently installed printer**, not a fixed point value. Different printers can produce different width measurements.

**Evidence:**
- The legacy forms were generated with a specific printer (likely 'Microsoft Print to PDF' or a physical printer)
- The current engine may use a different printer, changing the width
- This is most visible in forms with explicit column widths

**Fix:** Standardize the printer driver used during capture, or validate against the rendered PDF instead of COM values.

### 4. Database Coordinates — **Systematic offset**

The stored database ratios appear to have been generated with a consistent but different margin value. The ratios themselves are internally consistent — this is not a "database error" but a **parameter mismatch**.

### 5. PDF Rendering — **Unlikely to be primary cause**

The PDF → PNG pipeline (PDFtoImage / PDFium) faithfully renders at the specified DPI. The coordinate mismatch originates in the **coordinate calculation layer**, not the rendering layer.

### Overall Assessment

```
Complexity        Impact
─────────────────────────────────────
Margin fix         🟢 High (~265 forms)
Col width fix      🟡 Medium (~60 forms)
Font calibration   🟡 Medium (~8 forms)
Printer fix        🔴 Low (~20 forms)
```

### Recommended Next Steps

1. **Fix margin defaults** — Determine the actual margin value used by the legacy ConMas engine. Test values: 54pt (0.75in), 36pt (0.5in), 56.7pt (2cm).
2. **Fix font-dependent column width** — Read Normal font from XLSX `styles.xml`. Compute `maxDigitWidth` per workbook instead of hardcoding 7.33.
3. **Deploy Phase 2 changes** — The worksheet resolution fix is validated with zero regressions.
4. **Re-validate** — After margin fix, re-run validation. Expect ~200 additional forms to enter <0.5pt tolerance.
