# Phase 28 — Overlay Coordinate Engine (PaperLess Core)

**Date:** July 13, 2026  
**Status:** ✅ Complete  
**Build:** ✅ Next.js production build — Passed clean  

---

## Objective

Transition from "Excel Renderer" to the actual **PaperLess Engine**. The renderer now serves as the foundation for the overlay system — every Excel comment becomes an overlay definition with pixel-accurate coordinates computed purely from template data (no DOM queries, no screenshots).

---

## What Was Built

### 1. ✅ OverlayModel Type System (`types/overlay.ts`)

Defined the complete overlay type hierarchy:

```typescript
type OverlayType = "textbox" | "signature" | "checkbox" | "date" | "number" 
                 | "qr" | "barcode" | "image" | "ocr" | "unknown";

interface OverlayModel {
  id: string;           // e.g., "field_A5", "img_0"
  type: OverlayType;    // Inferred from comment text
  cell: string;         // Cell reference (e.g., "A5", "C3:D4" for merges)
  leftPt: number;       // Distance from print origin (pt)
  topPt: number;
  widthPt: number;
  heightPt: number;
  rotation: number;     // Degrees (0 = no rotation)
  metadata: Record<string, unknown>;
}
```

### 2. ✅ OverlayEngine Service (`services/overlayEngine.ts`)

Pure mathematical coordinate calculator. No DOM, no screenshots — just column widths, row heights, and cumulative math.

**Three overlay sources:**

| Source | Description |
|--------|-------------|
| **Comments** | Each `TemplateComment` becomes a field overlay. Type is inferred from comment text (e.g., `"signature"` → signature, `"checkbox"` → checkbox) |
| **Merged Cells** | Un-commented merged regions get `ocr` overlays covering the full merge span |
| **Images** | Each `TemplateImage` becomes an image overlay positioned at its anchor |

**Key algorithms:**
- `inferOverlayType()` — Parses comment text to determine overlay type using keyword matching (e.g., "sign here" → signature)
- `getCellRect()` — Returns the full rectangle for a cell, accounting for merged cells (returns the entire merge span)
- `generateOverlays()` — Main entry point, produces cached `OverlayCollection`

**Export formats:**
- `exportOverlays()` — JSON export (full detail)
- `exportOverlaysBrief()` — CSV export (id, type, cell, coordinates)

### 3. ✅ OverlayRectangles Component (`components/ExcelRenderer/OverlayRectangles.tsx`)

Debug visualization that renders colored rectangles over every overlay on the rendered grid:

- **Color-coded by type**: Blue=textbox, Purple=signature, Green=checkbox, Amber=date, Red=number, Pink=QR, Orange=barcode, Violet=image, Teal=OCR
- **Badges**: Type label (TEXT, SIG, CHK, etc.) in top-left corner, cell reference in bottom-right
- **Clickable**: Click any overlay rectangle to open the Overlay Inspector
- **Pointer events**: Rectangles have `pointer-events: auto` for interaction; container is `pointer-events: none` for grid pass-through

### 4. ✅ Overlay Inspector Panel

When an overlay rectangle is clicked, a panel opens showing:
- Overlay ID and type badge
- Cell reference
- Coordinates (left, top, width, height, rotation) in points
- Raw metadata JSON (comment text, source info, etc.)

### 5. ✅ Overlay Export Buttons

In the compare page, the overlays section includes:
- **Copy JSON** — Copies full overlay collection as JSON to clipboard
- **Copy CSV** — Copies brief overlay collection as CSV to clipboard
- **Type count breakdown** — Shows overlay count by type (e.g., "textbox: 5, checkbox: 2")

---

## Architecture

```
TemplateModel (from OpenXML)
       │
       ▼
OverlayEngine.generateOverlays(template)
       │
       ├── Comments → inferOverlayType() → getCellRect() → OverlayModel[]
       ├── MergedCells → getCellRect() → OverlayModel[]
       └── Images → cumulativeColWidth/RowHeight → OverlayModel[]
       │
       ▼
OverlayCollection { overlays, byId, byCell }
       │
       ├── OverlayRectangles (debug visualization)
       └── OverlayInspector (click to inspect)
```

**Performance:** Generation is done via `useMemo`, so it only recalculates when the template changes. Target: <10ms for normal templates.

---

## File Change Summary

| File | Status | Change |
|------|--------|--------|
| `docs/Phase28/Phase28_Overlay_Engine.md` | **NEW** | This report |
| `types/overlay.ts` | **NEW** | OverlayModel, OverlayType, OverlayCollection types |
| `services/overlayEngine.ts` | **NEW** | generateOverlays(), inferOverlayType(), getCellRect(), export helpers |
| `components/ExcelRenderer/OverlayRectangles.tsx` | **NEW** | Debug visualization with color-coded rectangles |
| `components/ExcelRenderer/index.ts` | Modified | Added OverlayRectangles export |
| `app/compare/page.tsx` | Modified | Added Overlays toggle, OverlayRectangles integration, OverlayInspector, export buttons |

---

## Build Status

| Check | Result |
|-------|--------|
| Next.js Production Build | ✅ Passed |

---

## How to Use

1. Navigate to `/compare`
2. Select a template (546, 547, or 548)
3. Click **"Overlays"** in the header to enable overlay visualization
4. Each overlay appears as a colored rectangle over the grid
5. **Click an overlay rectangle** to inspect its coordinates and metadata in the Overlay Inspector panel
6. Use **Copy JSON** or **Copy CSV** to export the coordinate data for use in downstream systems (OCR, form filling)

---

## Next Steps (Phase 29+)

- **OCR Engine**: Use overlay coordinates to crop regions from page images and run OCR
- **Form Runtime**: Render interactive input fields at overlay positions
- **Digital Signatures**: Place signature pads at signature overlay positions
- **API Integration**: Expose `/api/template/{id}/overlays` on the backend
