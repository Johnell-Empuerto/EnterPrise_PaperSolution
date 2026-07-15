# Phase X.11 — Performance Optimization Report

## Objective

Optimize the upload/preview pipeline while preserving 100% identical output (Phase X.10 is the golden reference).

## Critical Rule

No coordinate generation, pixel normalization, page geometry, ratio calculation, or JSON schema changes.

## Profiling Methodology

Instrumented every stage of `generate_coordinates_and_preview()` to measure actual execution time for both test workbooks.

## Profiling Results (Before Optimization)

### FormTest (total: 17,339ms)

| Stage | Time (ms) | % of Total |
|-------|-----------|-----------|
| Pixel scan (scan_black_rectangles) | **13,370** | **77.1%** |
| Sanitize (COM cell fills) | 1,909 | 11.0% |
| Reopen + export sanitized PDF | 681 | 3.9% |
| Export original PDF | 473 | 2.7% |
| Comment scan (COM) | 376 | 2.2% |
| Cluster detection | 434 | 2.5% |
| Render sanitized PDF (fitz) | 90 | 0.5% |
| Render background PNG (fitz) | 126 | 0.7% |
| Split + normalize rects | 173 | 1.0% |
| Save sanitized workbook | 61 | 0.4% |
| Excel.Quit | 2 | <0.1% |

### Japanese Workbook (partial: stopped at ~16s COM)

| Stage | Time (ms) |
|-------|-----------|
| Cluster detection | 7,100 |
| Export original PDF | 1,481 |
| Comment scan | 7,353 |

## Optimization Applied

### #1: NumPy-Vectorized Pixel Scan (13,370ms → 195ms, 67.8x speedup)

**File:** `render_service/upload_coordinate_generator.py`
**Function:** `scan_black_rectangles()`

**Before:** Pure Python nested loops (8.4M pixels) with per-pixel function calls:
```python
for y in range(1, h - 1):
    for x in range(1, w - 1):
        if _is_black(r, g, b):
            # while-loop for right expansion
            # while-loop for bottom expansion
```

**After:** NumPy-vectorized operations with pre-computed binary masks:
```python
black = np.all(img < BLACK_THRESHOLD, axis=2)   # vectorized
white = np.all(img > WHITE_THRESHOLD, axis=2)   # vectorized
not_black = ~black                               # vectorized
corners = black & left_pad & above_pad           # vectorized corner detection
corner_ys, corner_xs = np.where(corners)         # find all corners at once

# Right expansion: np.where() replaces while-loop
first_white = np.where(white[y, x + 1:])[0]

# Bottom expansion: np.where() replaces per-column while-loop
first_white_in_col = np.where(col_white)[0]
```

**Key changes:**
- Binary masks (`black`, `white`, `not_black`) replace per-pixel `_is_black()`/`_is_white()`/`_is_not_black()` calls
- Shifted mask comparison (`np.pad` + slice) replaces manual neighbor pixel checks
- `np.where()` vectorized white-pixel search replaces while-loop expansion
- `np.any()` array check replaces manual 6-pixel noise verification loops

**Identical algorithm guaranteed by:**
- Same `BLACK_THRESHOLD` (64) and `WHITE_THRESHOLD` (192) constants
- Corner condition: `black & not_black_left & not_black_above` (same as old)
- Right expansion: starts at `x+1`, finds first white pixel (same as old)
- Bottom expansion: per-column from `y+1`, finds first white pixel (same as old)
- 6-pixel noise checks use same range (6 pixels) and same threshold

**Validation:** Direct comparison on same image — all 4 rectangles match exactly.

## Validation Results

### Full Production Pipeline

| Workbook | Phase X.10 (Before) | Phase X.11 (After) | Improvement |
|----------|---------------------|--------------------|-------------|
| FormTest | 17,874ms | **4,420ms** | **4.0x faster** |
| Japanese | 50,079ms | **36,952ms** | **1.35x faster** |

### Coordinate Comparison

**FormTest — 6 fields, identical coordinates:**

| Field | CellAddr | Phase X.10 L/T | Phase X.11 L/T | Match |
|-------|----------|---------------|---------------|-------|
| samples | $A$1:$B$2 | 0.3360784 / 0.3845455 | 0.3360784 / 0.3845455 | ✅ |
| samples | $C$1:$D$2 | 0.4996078 / 0.3845455 | 0.4996078 / 0.3845455 | ✅ |
| samples | $A$3:$D$4 | 0.3360784 / 0.4225758 | 0.3360784 / 0.4225758 | ✅ |
| samples | $A$6:$D$7 | 0.3360784 / 0.480303 | 0.3360784 / 0.480303 | ✅ |
| samples | $A$9:$D$10 | 0.3360784 / 0.5381818 | 0.3360784 / 0.5381818 | ✅ |
| samples | $A$12 | 0.3360784 / 0.5957576 | 0.3360784 / 0.5957576 | ✅ |

**Japanese — 5 fields, identical coordinates:**
All 5 fields match Phase X.10 coordinates exactly.

## Remaining Bottlenecks

| Stage | FormTest (ms) | Japanese (ms) | Notes |
|-------|--------------|---------------|-------|
| Sanitize (COM) | ~1,900 | ~8,000 | Cell-by-cell COM access |
| Comment scan (COM) | ~400 | ~7,300 | Large worksheet cell scan |
| Export PDF (COM) | ~1,150 | ~3,000 | COM ExportAsFixedFormat |
| Render PNG (fitz) | ~200 | ~200 | Fast already |
| **Total** | **~4,400** | **~37,000** | COM section dominates |

## Files Modified

**Only:** `render_service/upload_coordinate_generator.py` — `scan_black_rectangles()` function.

## Files Cleaned Up

- `render_service/_phaseX11_profile.py` — temporary profiling script (deleted)
- `render_service/_phaseX11_validate_scan.py` — temporary validation script (deleted)

## Conclusion

Phase X.11 achieves its primary objective:
- **Pixel scan: 67.8x faster** (13,370ms → 195ms)
- **FormTest total: 4x faster** (17.9s → 4.4s)
- **Japanese total: 1.35x faster** (50.1s → 37.0s)
- **100% identical output** (all coordinates match Phase X.10)
- **No algorithm changes** to coordinate generation, normalization, or geometry
- **Single file changed** with minimal diff

The remaining bottleneck is COM cell-by-cell operations (sanitize, comment scan), which are inherently limited by Excel COM interop speed.
