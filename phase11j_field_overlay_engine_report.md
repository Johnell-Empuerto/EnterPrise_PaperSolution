# Phase 11J — Field Overlay Engine (Yellow Editable Layer)

**Date:** July 2026  
**Target Build:** 0 errors (Next.js `tsc --noEmit`)  
**Status:** ✅ Complete

---

## Objective

Build the interactive field overlay layer in the Next.js frontend that renders editable HTML inputs on top of the rendered PNG background — exactly like the old system's yellow overlay.

## Architecture

```
Excel
   │
Upload
   │
Backend
   ├── Render PNG (GET /preview/page_{templateId}.png)
   └── Runtime JSON (GET /api/form/runtime/{templateId})
             │
             ▼
Next.js
   │
Background PNG  ←── FormPage.tsx (<img />)
   │
Yellow Overlay  ←── OverlayRenderer.tsx → OverlayField.tsx
   │
   ├── TextField.tsx        (text / textarea)
   ├── NumberField.tsx      (number)
   ├── DateField.tsx        (date picker)
   ├── CheckboxField.tsx    (checkbox)
   ├── DropdownField.tsx    (select)
   └── SignatureField.tsx   (canvas signature)
```

## Files Created (12)

### `types/runtime.ts`
TypeScript type definitions mirroring the C# Runtime Engine models:
- `RuntimeForm` — workbook name, sheets, page dimensions, scale, DPI, version
- `RuntimeSheet` — name, fields, images, shapes, print area, page pixel dimensions
- `RuntimeField` — full field metadata: id, cellReference, coordinates, dataType, readOnly, required, alignment, font, border, validation, etc.
- `RuntimeImage`, `RuntimeShape`, `RuntimePrintArea` — supporting types
- `FieldValues`, `FieldErrors`, `DirtyState` — state management types

### `hooks/useRuntime.ts`
- Fetches `RuntimeForm` from `GET /api/form/runtime/{templateId}`
- Manages loading/error states
- Supports manual `reload()` trigger
- Handles cancellation on unmount

### `hooks/useFieldState.ts`
- Manages form field values, dirty tracking, and validation
- Validation rules: required, regex pattern, max length, number parsing, date parsing
- `validateField()` for per-field validation
- `validateAll()` for full-form validation
- `reset()` to restore defaults
- ⚠️ *Standalone hook — not yet wired into field components (available for future use)*

### `components/FormViewer/FormPage.tsx`
- Single page container: background image + overlay fields
- Loads background from `{API_BASE_URL}/preview/page_{templateId}.png`
- Applies zoom scaling with `transform: scale(zoom)`
- Shows fallback placeholder when background fails to load

### `components/FormViewer/OverlayRenderer.tsx`
- Renders all `RuntimeField` objects from a sheet as absolutely-positioned overlays
- Shows "No editable fields detected" placeholder when field list is empty

### `components/FormViewer/OverlayField.tsx`
- Dispatches to the correct field component based on `field.dataType`
- Passes `position: absolute; left: field.leftPx; top: field.topPx; width: field.widthPx; height: field.heightPx` to each component
- Handles `calculated` type as a read-only div

### `components/FormViewer/fields/TextField.tsx`
- Single-line `<input>` or multi-line `<textarea>` based on height
- Yellow highlight (#FFFDE7 → #FFF9C4 on focus)
- Focus glow effect with `box-shadow`
- Font/color/bold inherited from field definition
- `readOnly` support with reduced opacity

### `components/FormViewer/fields/NumberField.tsx`
- `<input type="number">` with yellow highlight
- Right-aligned text (respects field alignment)
- `step="any"` for decimal support

### `components/FormViewer/fields/DateField.tsx`
- `<input type="date">` with browser-native date picker
- Yellow highlight with focus effects

### `components/FormViewer/fields/CheckboxField.tsx`
- Custom checkbox div with ✓ mark when checked
- Auto-sized (max 20px) and centered within field bounds
- Click to toggle (touch-friendly)

### `components/FormViewer/fields/DropdownField.tsx`
- `<select>` dropdown with yellow highlight
- Placeholder option ("Select...")
- ⚠️ *Options are currently hardcoded (Option 1/2/3)*

### `components/FormViewer/fields/SignatureField.tsx`
- `<canvas>`-based signature capture
- Click to sign — draws a bezier curve signature line
- "Click to sign" placeholder overlay
- Gold highlight on hover

## File Modified (1)

### `app/page.tsx`
- Added **Runtime Viewer** tab alongside existing Designer tab
- Tab switcher in header with "Designer" / "Runtime Viewer" buttons
- Template ID text input + Load button
- Loading spinner, error display
- Sheet metadata bar (workbook name, sheet count, field count, DPI)
- Zoom controls (- / 100% / +)
- Focused field info bar (cell reference, data type, required/read-only status, position)
- Conditional rendering: Designer tab vs Runtime tab

## Validation

| Check | Result |
|-------|--------|
| `tsc --noEmit` | ✅ 0 errors |

## Key Design Decisions

1. **Backend coordinates as single source of truth** — Every overlay field uses `field.leftPx`, `field.topPx`, `field.widthPx`, `field.heightPx` directly from the Runtime API. Positions are never recalculated in the frontend.

2. **Field components use local `useState`** — Each field component manages its own value state. The `useFieldState` hook is available but not wired in. This simplifies the initial Phase 11J implementation. Future phases can connect shared state for form submission.

3. **Yellow highlight color scheme** — Consistent `#FFFDE7` (unfocused) / `#FFF9C4` (focused) background with `#F9A825` border, matching the classic yellow editable overlay convention.

4. **Tab-based UI** — Designer tab retains all existing cluster overlay functionality. Runtime tab provides the new interactive form view. Both coexist without conflict.

## Known Limitations

| Issue | Impact | Resolution |
|-------|--------|------------|
| `useFieldState` hook is dead code | No shared form state, no validation UI | Wire hook in Phase 11K |
| Keyboard navigation not implemented | No Tab/Arrow/Enter navigation | Add in Phase 11K |
| Background URL assumes single page | Multi-sheet forms load same image for all | Use sheet-indexed URL |
| DropdownField has hardcoded options | Can't configure options from backend | Add `options` field to RuntimeField |
| No form submit/save | Values can't be collected or submitted | Add in Phase 11K |

## Rendering Pipeline (Updated)

```
Workbook
    │
    ▼
OpenXmlParser → StyleResolver → GeometryBuilder → PrintLayoutEngine
    │
    ▼
PageRenderer → RenderingContext
    │
    ├── FillEngine
    ├── GridlineLayer
    ├── BorderEngine
    ├── TextEngine
    ├── ImageEngine
    └── ShapeEngine
    │
    ▼
ExportCoordinator → PNG / PDF
    │
    ▼
Runtime Engine (Phase 11I) → Runtime JSON
    │
    ▼
Field Overlay Engine (Phase 11J) → Next.js Interactive Form
    │
    ├── Background PNG
    └── Yellow HTML Inputs
```

## Next Steps

- **Phase 11K** — Designer Enhancements (wire `useFieldState`, keyboard navigation, form submission, multi-page support)
- **Phase 11L** — Production Release
