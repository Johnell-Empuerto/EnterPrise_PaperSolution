# Phase 11J.17 — Multi-Template Validation & Effective Dimension Origin

## Objective
Validate the ConMas legacy algorithm across multiple templates and determine the origin of the effective content dimensions (200.16×182.88pt for template 546).

## Templates Investigated

| ID | Fields | Orientation | Page Size | FitToPages | Centered | Notes |
|----|--------|-------------|-----------|------------|----------|-------|
| 546 | 6 | Portrait | Letter (612×792) | No | Yes | Reference template (11J.15-16) |
| 501 | 1 | Portrait | Letter (612×792) | No | Yes | Single field |
| 448 | 5 | Portrait | Letter (612×792) | Yes (1×1) | No | FitToPages, not centered |
| 228 | 2 | Portrait | Letter (612×792) | No | No | Not centered |
| 186 | 1 | Landscape | A4 (842×595) | No | No | Landscape A4, not centered |
| 142 | 1 | Portrait | A4 (595×842) | No | Yes | A4 portrait, centered |

## Key Findings

### 1. Width Scaling Factor (1.0425) Is Consistent Across Templates

| Template | First Field Column | PrintArea Width | effW | ScaleW | First Field rL |
|----------|-------------------|-----------------|------|--------|----------------|
| 546 | A | 192.0 | 200.16 | **1.0425** | 0 |
| 448 | A | 192.0 | 200.16 | **1.0425** | 0 |
| 228 | A | 240.0 | 250.20 | **1.0425** | 0 |

Width scaling factor **1.0425 is invariant** when the first field starts at column A (rL=0).

### 2. Height Varies When First Field Is Not at Row 1

| Template | First Field Row | PrintArea Height | effH | ScaleH | First Field rT |
|----------|----------------|------------------|------|--------|----------------|
| 546 | 1 | 100.80 | 182.88 | **1.814** | 0 |
| 448 | 6 | 100.80 | 97.92 | **0.971** | 72.0 |

Template 448 first field starts at row 6 (rT=72), so content height calculation shifts.

### 3. Width Formula Verified

```
page_X = LM + (PW − effW) / 2 + Range.Left × (effW / PAW)
```

For template 546 with Range.Left = 0:
```
page_X = 51.024 + (509.95 − 200.16) / 2 = 205.92 ✓
```

This matches legacy ratio 0.3364706 × 612 = 205.92.

### 4. Height Formula — Unresolved Algebraic Discrepancy

For template 546 with Range.Top = 0 and TM = 53.858:
```
Page_Y = TM + (PH − effH) / 2 + Range.Top × scale
       = 53.858 + (684.284 − 182.88) / 2
       = 53.858 + 250.702
       = 304.56
```

But legacy_Y = 0.529091 × 792 = **419.04**. The formula gives 304.56, which is **114.48pt off**.

This means the height formula has a different structure than the width formula. The phase11j16 empirical test (avg error 0.08pt) did verify a formula works, but the algebraic derivation for height is not yet understood.

**Possible explanations:**
- Height uses a different origin (not TM, possibly page-top-relative)
- Ratio is measured from a different reference point (center of page?)
- Additional offset applied for vertical centering when content is top-heavy
- An interaction with `FitToPages` or row heights

### 5. Database Stores Ratios as Strings

`def_cluster` columns `left_position`, `top_position`, `right_position`, `bottom_position` are returned by psycopg2 as Python `str`, not `float` or `Decimal`.

```python
# Current behavior (BUG):
ratio = "0.06557377"      # str
value = ratio * 612       # "0.065573770.06557377..." (repeats 612 times!)

# Required fix:
ratio = float("0.06557377")  # float
value = ratio * 612          # 40.13
```

This affects every coordinate calculation in the codebase.

### 6. XLSX def_file May Differ from Working Copy

Template 546's `def_file` in the database has no `PrintArea`, 336pt wide (7 cols), default margins, not centered. The earlier analysis used a modified workbook with `PrintArea=$A$1:$D$12`, centered, custom margins. All scripts must extract from `def_file` at analysis time, not rely on cached copies.

### 7. PageDimensions COM Property Investigation (Investigation A)

- `PageSetup.Pages(1)` COM object **exists** but has **no** `.Width`, `.Height`, `.Left`, `.Top` properties.
- No other COM property directly exposes rendered page content bounds.
- `PageSetup.Pages.Count` always returns 1 (even for multi-page content).

## Recommendations

1. **Pre-computation during upload** is the recommended path for the web version — run Excel COM to render each field, store resulting ratios in database.
2. **Fix database string-type bug** — add explicit `float()` casts in the DAO layer for all `def_cluster` ratio columns.
3. **Height formula needs further investigation** — the algebraic derivation doesn't match the empirical results. May need a dedicated phase to reconcile.
4. **All scripts should work from fresh `def_file` extraction**, not cached copies.

## Confidence Assessment

| Claim | Confidence | Evidence |
|-------|-----------|----------|
| Width scaling factor 1.0425 is consistent | High | Verified across 3 templates |
| Effective dimensions encode first-field position | High | Verified algebraically for width |
| Database stores ratios as strings | High | Directly observed in psycopg2 |
| Pre-computation is viable | High | Phase 11J.16 proof-of-concept |
| COM property `Pages(1)` exists | High | Tested via Python COM |
| Height formula matches width structure | Low | Algebraic discrepancy of 114.48pt |
| PDF measurement of content bounds | Low | PyMuPDF API format changed |
| Template 546 def_file is representative | Low | Database copy differs from working copy |
