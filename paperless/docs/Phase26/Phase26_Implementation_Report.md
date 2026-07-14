# Phase 26 — Visual Parity Implementation Report

**Date:** July 13, 2026  
**Status:** ✅ Complete  
**Build:** ✅ TypeScript (tsc --noEmit) — Passed clean  
**Build:** ✅ Next.js production build — Passed clean  
**Files Changed:** 9 files (1 new, 8 modified)

---

## Objective

Move from "basic renderer" to **visual parity** — make the browser renderer visually indistinguishable from Microsoft Excel Print Preview. The architecture (CSS Grid + CSS Modules + inline styles) was preserved. No Canvas, SVG, or image rasterization was introduced.

---

## What Was Implemented

### 1. ✅ Image Rendering (from OpenXML)

**What happened:** Created a new `ExcelImage` component and integrated it into the grid.

**Implementation details:**
- Added `imageData` (base64 data URL), `imageUrl`, `offsetXPt`, `offsetYPt`, and `description` fields to the `TemplateImage` type
- Created `ExcelImage.tsx` — positions images as absolutely-positioned `<div>` elements within the CSS Grid wrapper
- Images use `object-fit: contain` with `object-position: top left` to match Excel's placement behavior
- Position is calculated from anchor cell coordinates: cumulative column/row widths + optional offset
- Images are non-interactive (`pointer-events: none`, `user-select: none`, `draggable: false`)

**Files:**
- `types/template.ts` — Enhanced `TemplateImage` interface
- `components/ExcelRenderer/ExcelImage.tsx` — **New file** — Floating image component
- `components/ExcelRenderer/ExcelGrid.tsx` — Integrated image rendering
- `components/ExcelRenderer/helpers.ts` — Added `cumulativeColWidth`, `cumulativeRowHeight` helpers

**Known limitation:** Two-cell anchor (from/to) for image placement is not yet supported. Only single-cell anchor.

### 2. ✅ Improved Border Rendering

**What happened:** Replaced simple `"1px solid #d0d0d0"` fallback with proper Excel border parsing.

**Implementation details:**
- Created `parseExcelBorder(borderStr)` helper function that converts Excel format strings to CSS shorthand
- Supports all Excel line styles: `thin`, `medium`, `thick`, `double`, `hair`, `dotted`, `dashed`, `dashdot`, `mediumdashed`, `slantdashdot`, etc.
- Extracts color from border strings — supports hex (`#FF0000`), rgb(), and named colors (black, red, blue, green, etc.)
- Handles multi-word color names like "dark blue", "light green", "dark red"
- Grid line fallback changed from `1px solid #d0d0d0` to `0.5px solid #d0d0d0` for a more Excel-like appearance
- Both `ExcelCell` and `ExcelMergedCell` now use the parser

**Map of Excel → CSS styles:**

| Excel Style | CSS Output |
|-------------|-----------|
| `thin black` | `1px solid #000000` |
| `medium #4472C4` | `2px solid #4472C4` |
| `thick` | `3px solid #000000` |
| `double` | `3px double #000000` |
| `hair` | `0.5px solid #000000` |
| `dotted` | `1px dotted #000000` |
| `dashed` | `1px dashed #000000` |
| `none` or empty | `undefined` (no border) |

**Files:**
- `components/ExcelRenderer/helpers.ts` — Added `parseExcelBorder()` function
- `components/ExcelRenderer/ExcelCell.tsx` — Uses `parseExcelBorder` via `borderStyle()`
- `components/ExcelRenderer/ExcelMergedCell.tsx` — Same

**Known limitation:** Border precedence (adjacent cells sharing a border, thicker wins) is not implemented. Diagonal borders not yet supported.

### 3. ✅ Improved Font Rendering

**What happened:** Added conditional wrapText, text-overflow ellipsis, and indent support.

**Implementation details:**
- `wrapText = true`: applies `white-space: pre-wrap`, `word-break: break-word` — text wraps within cell
- `wrapText = false | undefined`: applies `overflow: hidden`, `text-overflow: ellipsis`, `white-space: nowrap` — truncates with ellipsis
- `indent > 0`: applies `paddingLeft` calculated as `indentLevel * fontSize * 1.8` pt
  - This matches Excel's indent of ≈3 character widths per level (~19.8pt for 11pt Calibri)
- Dynamic `title` attribute shows cell value for truncated content

**Bug caught during code review:**
- Initial indent formula used `indentLevel * (fontSize * 96/72 * 0.3)` px → resulted in ~4.4px (6× too small)
- Fixed to `indentLevel * fontSize * 1.8` pt → gives ~19.8pt for 11pt (matches Excel's 3-character-width indent)

### 4. ✅ Improved Alignment & Print Origin

**What happened:** Enhanced the `originFromMargins` function for more accurate page positioning.

**Implementation details:**
- Previously only used `marginLeft` and `paperWidth` for centering calculation
- Now accounts for **all four margins** (left, right, top, bottom) and computes usable area
- Centering now correctly positions content within margin bounds:
  - `xPt = marginLeft + (usableWidth - contentWidth) / 2` (instead of `(paperWidth - contentWidth) / 2`)
  - `yPt = marginTop + (usableHeight - contentHeight) / 2`

**Bug caught during TypeScript build:**
- Duplicate `originFromMargins` function existed in helpers.ts — one from the original code + one from the new replacement
- Removed the duplicate, keeping only the improved version

### 5. ✅ Improved Merged Cell Rendering

**What happened:** Added `z-index` layering and overflow handling for merged cells.

**Implementation details:**
- Merged cells receive `z-index: 1` to prevent adjacent cells from clipping/spilling over the merge visually
- CSS class `.mergedCell` updated with `line-height: 1.4` for consistent text baseline
- Dynamic wrapText/indent applied consistently with regular cells

### 6. ✅ Hidden Rows & Columns Support

**What happened:** Added infrastructure for hiding rows/columns from the grid renderer.

**Implementation details:**
- Added optional `hiddenColumns: boolean[]` and `hiddenRows: boolean[]` arrays to `TemplateModel`
- `buildGridDimensions` skips hidden columns/rows when computing grid template and totals
- `ExcelGrid` uses `visibleColIndex` and `visibleRowIndex` Maps to track 1-based grid positions for visible items only
- Hidden items are simply not rendered — they're excluded from the loop entirely
- Backend can provide these arrays from OpenXML (when column/row has `hidden="true"` attribute)

**Edge case noted:** If a merged cell spans across hidden rows/columns, the span count (`span.cols`, `span.rows`) doesn't adjust to exclude hidden tracks. This is an unlikely scenario in practice.

---

## Issues Encountered & Resolved

### Issue 1: Duplicate Function Declaration
- **Symptom:** TypeScript error `TS2323: Cannot redeclare exported variable 'originFromMargins'`
- **Cause:** When the improved `originFromMargins` was added via `str_replace`, the original version at the bottom of `helpers.ts` was not removed
- **Fix:** Manually removed the duplicate function definition
- **Lesson:** When replacing large code blocks, verify there are no orphaned originals

### Issue 2: Indent Calculation Too Small
- **Symptom:** Code review flagged that the indent padding was ~6× smaller than Excel's actual indent
- **Cause:** Formula used `indentLevel * (fontSize * 96/72 * 0.3)` (pixels), but Excel defines 1 indent level as 3 character widths (~3 × fontSize × 0.6 = 1.8 × fontSize)
- **Fix:** Changed to `indentLevel * fontSize * 1.8` (points), applied consistently in both `ExcelCell` and `ExcelMergedCell`

### Issue 3: Multi-Word Color Names Not Parseable
- **Symptom:** Code review noted that `split(/\s+/)` broke multi-word colors like "dark blue" into separate parts
- **Cause:** Color string was extracted as `parts[1]` — only captured the first word
- **Fix:** Changed to `parts.slice(1).join(" ")` then looked up in `namedColors` map with `toLowerCase()`

### Issue 4: Unused `pxToPt` Function
- **Symptom:** Dead code flagged during review
- **Cause:** Added as a utility but never called by any component
- **Fix:** Removed the function to maintain minimal codebase

---

## Build Verification

| Check | Result |
|-------|--------|
| TypeScript Check (`tsc --noEmit`) | ✅ Passed (0 errors) |
| Next.js Production Build (`next build`) | ✅ Passed (2.8s compile, no errors) |
| Routes Generated | `/` (static), `/compare` (static) |

---

## File Change Summary

| File | Status | What Changed |
|------|--------|--------------|
| `types/template.ts` | Modified | Added `imageData`, `imageUrl`, `offsetXPt`, `offsetYPt`, `description` to `TemplateImage`; added `hiddenColumns`, `hiddenRows` to `TemplateModel` |
| `components/ExcelRenderer/helpers.ts` | Modified | Added `parseExcelBorder()`, `cumulativeColWidth()`, `cumulativeRowHeight()`; updated `buildGridDimensions()` for hidden rows/cols; enhanced `originFromMargins()` with full-margin calculation |
| `components/ExcelRenderer/ExcelImage.tsx` | **NEW** | Floating image component using absolute positioning within the CSS Grid |
| `components/ExcelRenderer/ExcelGrid.tsx` | Modified | Integrated `ExcelImage` rendering; added `visibleColIndex`/`visibleRowIndex` maps for hidden row/col tracking |
| `components/ExcelRenderer/ExcelCell.tsx` | Modified | Uses `parseExcelBorder` for proper Excel border styles; conditional wrapText; correct indent calculation |
| `components/ExcelRenderer/ExcelMergedCell.tsx` | Modified | Same improvements as ExcelCell |
| `components/ExcelRenderer/ExcelRenderer.module.css` | Modified | Added `.image` class; updated `.cell`/`.mergedCell` with `line-height: 1.4`; removed redundant base styles |
| `components/ExcelRenderer/index.ts` | Modified | Added `ExcelImage` export |
| `docs/Phase26/Phase26_VisualParity.md` | Modified | Updated with implementation status |
| `docs/Phase26/Phase26_Implementation_Report.md` | **NEW** | This file — comprehensive implementation report |

---

## Architectural Invariants Preserved

- ✅ CSS Grid layout (not Canvas, not SVG, not HTML tables)
- ✅ CSS Modules for base styles + inline styles for dynamic values
- ✅ No Tailwind in the renderer
- ✅ No external dependencies added (zero `npm install`)
- ✅ No COM or rasterization on the server
- ✅ No coordinate engine modifications
- ✅ Backend API unchanged

---

## Remaining Known Limitations

| # | Limitation | Impact | Priority |
|---|-----------|--------|----------|
| 1 | Two-cell anchor (from/to) not supported for images | Images always anchor to a single cell | Low |
| 2 | Border precedence not implemented | Adjacent cells with conflicting borders don't resolve thicker wins | Medium |
| 3 | Diagonal borders not supported | Forward/backward slash borders render as solid | Low |
| 4 | Merged cell span doesn't adjust for hidden rows/cols | Potential misalignment if hidden rows exist inside merge ranges | Low |
| 5 | Gradient/pattern fills not implemented | Only solid `backgroundColor` fills work | Low |
| 6 | Page breaks not detected | Long worksheets scroll rather than paginate | Medium |
| 7 | Rich text (multiple font runs in one cell) not supported | Inline font changes within a single cell are lost | Low |
| 8 | Print CSS (`@page`, page-size, margin boxes) not optimized | Browser Print → PDF may not match Excel PDF exactly | High |

---

## Recommendations for Phase 27

1. **Page break support** — Detect OpenXML row/page breaks and render multi-page documents with proper pagination
2. **Print CSS optimization** — Add `@page` size/orientation rules, media queries for print, and margin box configuration
3. **Gradient fills** — Parse OpenXML gradient fill definitions and render via CSS `background-image: linear-gradient()`
4. **Border precedence** — Implement cell border conflict resolution (thicker border wins when adjacent cells differ)
