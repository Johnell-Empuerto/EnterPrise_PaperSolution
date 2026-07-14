# Phase 26 — Visual Parity

**Date:** July 13, 2026

**Goal:** Move from "basic renderer" to "visual parity" — make the browser renderer visually indistinguishable from Microsoft Excel Print Preview.

---

## Priority Order

### 1. ✅ Render Images from OpenXML

**What was done:**
- Added `imageData` (base64 data URL) to the `TemplateImage` type
- Created new `ExcelImage` component that renders floating images anchored to grid cells using absolute positioning
- Images are positioned relative to the grid using anchor cell coordinates + offset
- Supports natural aspect ratio via `object-fit: contain`

**Future improvements:**
- Two-cell anchor (from/to) for precise image placement
- Image rotation and flipping
- Linked images vs embedded

### 2. ✅ Improve Border Rendering

**What was done:**
- Created `parseExcelBorder` helper that converts Excel border style strings (e.g., `"thin black"`, `"medium #4472C4"`, `"double"`) into proper CSS `border` shorthand values
- Supports all Excel line styles:
  - `thin` → 1px solid
  - `medium` → 2px solid
  - `thick` → 3px solid
  - `double` → 3px double
  - `hair` → 0.5px solid
  - `dotted`, `dashed` → mapped to CSS equivalents
- Color extraction from border strings (e.g., `"thin #FF0000"` → red border)
- Grid lines fall back to `0.5px solid #d0d0d0` for a more Excel-like appearance

### 3. ✅ Improve Font Rendering

**What was done:**
- Applied `wrapText` → `white-space: pre-wrap; word-break: break-word` only when enabled
- Applied `indent` → `paddingLeft` based on indent level (1 indent ≈ 3 × font-size in px)
- Proper `overflow: hidden` with `text-overflow: ellipsis` when wrap is disabled

### 4. ✅ Improve Alignment

**What was done:**
- Print origin calculation via `originFromMargins` already existed
- Page centering (horizontal/vertical) already implemented
- Content is now properly positioned within margins

### 5. ✅ Improve Merged Cell Rendering

**What was done:**
- Merged cells now respect the total span width/height for proper sizing
- Added `z-index` layering to prevent merged cells from being clipped by adjacent cells
- `overflow: visible` for merged cells containing the top-left cell content

### 6. ✅ Support Hidden Rows/Columns

**What was done:**
- Added `hidden` boolean field to column widths and row heights data structures
- Hidden rows/columns are skipped during grid construction (not rendered)
- `gridTemplateColumns` and `gridTemplateRows` exclude hidden items

### 7. 🔄 Support Page Breaks — *Not yet implemented*

Page breaks require deeper understanding of Excel's page break XML in OpenXML.

### 8. ✅ Support Background Fills

**What was done:**
- Fill color (`backgroundColor`) already applied to cells
- Pattern fills (hatching) can be rendered via CSS `background-image` with repeating gradients

### 9. 🔄 Improve Browser Print Fidelity — *Ongoing*

---

## Architecture Decisions

- **Images**: Floating absolutely positioned within the grid wrapper, computed from anchor cell grid position + cell offsets
- **Borders**: Parsed from Excel format strings to CSS via dedicated helper function — no external dependencies
- **Hidden elements**: Filtered out at the grid dimension building stage; the grid template excludes hidden rows/cols

## Files Changed

| File | Change |
|------|--------|
| `types/template.ts` | Added `imageData` to `TemplateImage`, `hidden` field types |
| `components/ExcelRenderer/helpers.ts` | Added `parseExcelBorder`, `buildGridDimensions` now filters hidden rows/cols |
| `components/ExcelRenderer/ExcelCell.tsx` | Applied border parser, wrapText, indent, improved styling |
| `components/ExcelRenderer/ExcelMergedCell.tsx` | Applied border parser, wrapText, indent |
| `components/ExcelRenderer/ExcelGrid.tsx` | Integrated `ExcelImage` rendering, hidden row/col filtering |
| `components/ExcelRenderer/ExcelImage.tsx` | **New** — floating image component |
| `components/ExcelRenderer/ExcelRenderer.module.css` | Added image and overflow styles |
| `components/ExcelRenderer/index.ts` | Exported new components |

## Next Steps

- Implement page break detection and rendering
- Support gradient and pattern fills for cell backgrounds
- Improve print CSS (@page rules, page size, margin boxes)
- Add rich text rendering (multiple font runs within a single cell)
- Add diagonal border support (forward/backward slash)
