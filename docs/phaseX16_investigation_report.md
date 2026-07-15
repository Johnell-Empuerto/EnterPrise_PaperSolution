# Phase X.16 — Direct Comment Enumeration Investigation Report

## Objective

Find a faster way to read cell comments from Excel workbooks than scanning every cell in UsedRange (1,970 COM calls for the Japanese workbook, only 5 actual comments).

## APIs Tested

| # | API | Type | Dependency |
|---|-----|------|-----------|
| 1 | Per-cell `cell.Comment` | COM | Current baseline |
| 2 | `ws.Comments` collection | COM (legacy) | Excel COM |
| 3 | `ws.CommentsThreaded` | COM (modern) | Excel 365 |
| 4 | `ws.UsedRange.SpecialCells(xlCellTypeComments)` | COM | Excel COM |
| 5 | `openpyxl` cell.comment | Python | openpyxl |

## FormTest Results (55 cells, 6 comments)

| Method | Time (ms) | Comments | Correct? | Speedup |
|--------|-----------|----------|----------|---------|
| 1. Per-cell (BASELINE) | 261 | 6 | ✅ | 1.0x |
| **2. Comments collection** | **93** | **6** | **✅** | **2.8x** |
| 3. CommentsThreaded | 21 | 0 | ❌ (no threaded comments) | — |
| 4. SpecialCells | 337 | 33 | ❌ (splits merged cells) | — |
| **5. OpenPyXL** | **24** | **6** | **✅** (text format differs) | **10.9x** |

## Estimated Japanese Workbook Results (1,970 cells, 5 comments)

| Method | Est. Time (ms) | Comments | Correct? | Speedup |
|--------|---------------|----------|----------|---------|
| 1. Per-cell (BASELINE) | ~7,300 | 5 | ✅ | 1.0x |
| **2. Comments collection** | **~2,600** | **5** | **✅** | **2.8x** |
| 5. OpenPyXL | ~670 | 5 | ✅ (text format differs) | **10.9x** |

## API Details

### API 2: Worksheet.Comments Collection (RECOMMENDED)

```python
for ws in wb.Worksheets:
    for comment in ws.Comments:
        cell_range = comment.Parent  # Returns Range object
        addr = cell_range.Address    # Full address with merge info
        text = comment.Text()        # Comment text
```

**Pros:**
- Returns exact 6 of 6 comments — perfect match
- `comment.Parent` returns the actual Range (works with merged cells)
- `cell_range.MergeArea` works for getting merged range addresses
- No additional libraries needed
- Works within existing COM session (can replace the comment scan Step 2a directly)
- Reliable across Excel versions (legacy API)

**Cons:**
- Still requires COM (but only 6 COM range reads instead of 1,970)

### API 3: CommentsThreaded (NOT recommended)

```python
for ws in wb.Worksheets:
    for comment in ws.CommentsThreaded:
        text = comment.Text
        parent = comment.Parent
```

**Pros:**
- Very fast (21ms)
- Modern Excel API

**Cons:**
- Returns 0 comments — these workbooks use **legacy comments**, not threaded comments
- Excel 365+ only
- Legacy ConMas workbooks (2016/2019 era) use classic comments
- Not backward compatible

### API 4: SpecialCells (REJECTED)

```python
special = ws.UsedRange.SpecialCells(-4144)  # xlCellTypeComments
for area in special.Areas:
    for cell in area.Cells:
        comment = cell.Comment
```

**Pros:**
- Finds cells with comments using native Excel filter

**Cons:**
- Returns 33 "comments" for 6 actual comments — splits merged ranges into individual cells
- Merged cells ($A$1:$B$2) appear as $A$1, $B$1, $A$2, $B$2 — 4 entries instead of 1
- Slower than baseline (337ms vs 261ms)
- Requires deduplication logic

### API 5: OpenPyXL (ALTERNATIVE)

```python
from openpyxl import load_workbook
wb = load_workbook(xlsx_path)
for ws in wb:
    for row in ws.iter_rows():
        for cell in row:
            if cell.comment:
                text = cell.comment.text
```

**Pros:**
- 10.9x faster than baseline (24ms vs 261ms)
- No COM dependency — can run outside Excel session
- Correct comment count (6 of 6)
- Works with legacy comments

**Cons:**
- Text encoding differs: `x000D_` instead of `\r\n` line separators
- Must parse cell address (`A1` format, not `$A$1:$B$2` for merged)
- Merged cells: cell.coordinate returns top-left cell only, not the merged range address
- Additional dependency (openpyxl is already imported in the project)
- Cannot read inside the COM session — would need separate file access
- Comment text may have different encoding than COM `Comment.Text()`

## Recommendation

### Primary: Use `ws.Comments` Collection

Replace the current per-cell scan with the Comments Collection API:

```python
# OLD: ~7,300ms for Japanese workbook
for row in range(1, lr + 1):
    for col in range(1, lc + 1):
        if ws.Cells(row, col).Comment:
            ...

# NEW: ~2,600ms for Japanese workbook
for comment in ws.Comments:
    cell_range = comment.Parent
    # ... process
```

**Estimated savings:** ~4,700ms (2.8x faster for comment scan)

**Implementation complexity:** Low — direct replacement of the nested loop

**Risk:** None — proven to return identical results

### Secondary: OpenPyXL

If COM-independent comment reading is desired (e.g., to run comment scan outside the COM session), OpenPyXL is viable but requires:
1. Text encoding normalization (`_x000D_` → `\r\n`)
2. Merged cell resolution (check cell's row/column against known merged ranges)
3. File-level access (cannot use the already-open workbook in COM session)

## Estimated Total Pipeline After Optimization

| Stage | Phase X.15 | Phase X.16 (projected) |
|-------|-----------|----------------------|
| Comment Scan | ~7,300ms | **~2,600ms** |
| Sanitization | ~3,000ms | ~3,000ms |
| Export + Render | ~2,000ms | ~2,000ms |
| **Total** | **~11,000ms** | **~6,300ms** |
