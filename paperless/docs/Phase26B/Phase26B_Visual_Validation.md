# Phase 26B — Excel Visual Validation & Pixel-Perfect Refinement

**Date:** July 13, 2026  
**Status:** ✅ Complete  
**Build:** ✅ TypeScript + Next.js — Passed clean  

---

## Objective

No new rendering features. The priority shifted from **adding capabilities** to **proving the renderer matches Microsoft Excel** through visual validation tools and pixel-perfect refinement.

---

## What Was Built

### 1. ✅ PDF Panel — Real Excel PDF Display

**Replaced** the placeholder "Replace with Excel PDF export" with a **live PDF screenshot upload system**.

**Features:**
- Drag & drop image upload (or click to browse)
- Paste image URL from clipboard
- Displays the uploaded image at the correct paper dimensions (scaled to match renderer zoom)
- Clear/remove button to reset
- File type validation (PNG, JPEG, WebP)

**How it works:** The user exports their Excel file to PDF, takes a screenshot, and uploads it. The image is rendered at the exact same paper size as the browser renderer for pixel-perfect comparison.

### 2. ✅ Comparison Mode Toggle — Side-by-Side / Overlay

**Two modes:**

| Mode | Description |
|------|-------------|
| **Side-by-Side** | Excel PDF on left, Browser Renderer on right — classic A/B comparison |
| **Overlay** | Browser Renderer rendered on top of PDF with adjustable opacity |

**Overlay mode features:**
- Opacity slider: 0%–100% (default 50%)
- PDF image and renderer are stacked and center-aligned
- Instantly reveals misalignment — if text or borders don't line up, they'll be visible as "ghost" doubles

### 3. ✅ Debug Overlay System

A comprehensive debug visualization system overlayed on the renderer grid. Each layer is independently togglable:

| Toggle | What it Shows |
|--------|---------------|
| **Grid Lines** | Faint lines at every column/row boundary |
| **Cell Coordinates** | Each cell labeled with its reference (A1, B2, etc.) |
| **Merge Bounds** | Colored rectangles around merged cell groups |
| **Image Bounds** | Colored rectangles showing image placements |
| **Origin** | Blue dot at the print origin (top-left of content) |
| **Print Area** | Red dashed outline of the print area bounds |
| **Margins** | Green dashed lines showing page margins |

### 4. ✅ Cell Inspector — Click to Measure

Click any cell in the renderer to open the **Cell Inspector** panel showing:

- **Cell Reference:** A1, B2, etc.
- **Grid Position:** Column and row indices
- **Dimensions:** Width and height in points
- **Absolute Position:** Left and top in points (from print origin)
- **Font:** Family, size, bold/italic/underline
- **Borders:** Each side's style and color
- **Background:** Fill color
- **Alignment:** Horizontal and vertical
- **Text:** Content preview with wrap status

### 5. ✅ Coordinate Inspector — Hover to Compare

Hover over any cell to see a **coordinate tooltip** showing:

- Cell reference (A1, B1, etc.)
- Browser measured position (left, top in pt)
- Cell dimensions (width, height in pt)
- Absolute coordinates from the grid origin

The coordinate inspector helps identify **cumulative errors** — if the browser renderer is off by even 0.5pt from Excel, it's visible here.

### 6. ✅ Debug Overlay Styling

The entire debug overlay system is self-contained:
- Rendered with **CSS Modules** (no external dependencies)
- Uses `pointer-events: none` so it doesn't interfere with cell interaction
- `user-select: none` prevents accidental text selection on overlays
- Labels are small (8pt) `font-mono` for precision

---

## How To Use the Validation Harness

1. **Select a template** from the dropdown (546, 547, or 548)
2. **Upload your Excel PDF** by dragging a screenshot onto the left panel
3. **Toggle Side-by-Side** or **Overlay** mode
4. In **Overlay mode**, drag the opacity slider to compare alignment
5. **Enable debug toggles** to see grid lines, coordinates, merge bounds
6. **Click any cell** to inspect its properties in the Cell Inspector panel
7. **Hover over cells** to see coordinate tooltips

---

## Files Changed

| File | Status | What Changed |
|------|--------|--------------|
| `docs/Phase26B/Phase26B_Visual_Validation.md` | **NEW** | This report |
| `app/compare/page.tsx` | Modified | Complete overhaul — PDF upload, mode toggle, overlay, debug toggles, cell inspector, coordinate inspector |
| `components/ExcelRenderer/ExcelCell.tsx` | Modified | Added `onClick`, `onMouseEnter`, `onMouseLeave` props and `data-col`/`data-row` attributes |
| `components/ExcelRenderer/ExcelMergedCell.tsx` | Modified | Same as ExcelCell |
| `components/ExcelRenderer/ExcelGrid.tsx` | Modified | Passes cell click/hover callbacks from template through to cells |
| `components/CompareView/DebugOverlay.tsx` | **NEW** | Debug visualization component |
| `components/CompareView/CellInspector.tsx` | **NEW** | Cell inspection panel |

---

## Next Steps

- **Pixel diff analysis**: Automatically compute pixel-level differences between PDF image and renderer screenshot
- **Browser print validation**: Test Ctrl+P → Save PDF and compare with Excel PDF
- **Screenshot capture**: Add a "capture renderer" button that screenshots the renderer for automated comparison
- **Report generation**: Auto-generate a diff report with overlaid images and measurement data
