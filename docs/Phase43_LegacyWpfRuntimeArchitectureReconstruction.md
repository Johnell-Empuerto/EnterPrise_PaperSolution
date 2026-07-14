# Phase 43 — Legacy WPF Runtime Architecture Reconstruction in HTML

**Date:** 2026-07-13
**Status:** Complete
**Type:** Frontend architecture reconstruction (code changes)

---

## Objective

Replace the browser layout with an HTML rendering pipeline that mirrors the legacy WPF runtime, ensuring the background image and overlay controls always exist in a single, immutable coordinate space.

## Problem

The previous architecture had the background `<img>` and `<RuntimeCanvas>` as siblings inside a responsive `<div>` that was subject to Tailwind CSS layout (flex, padding, centering). The Tailwind preflight `img { max-width: 100% }` constrained the 2550px-wide PNG to viewport width (~1182px measured in Phase 42), while the overlay container remained at 2550px. This caused a **470px offset** between fields and their cells.

## Solution — PageSurface Architecture

Mirrors the legacy WPF rendering hierarchy:

```
Legacy WPF:                        Browser Equivalent:
                                    
ScrollViewer                       Viewport (overflow: auto div)
  └── Grid                          └── card (bg-white, rounded, shadow)
        └── Canvas (fixed size)           └── PageSurface (fixed w/h, position:relative)
              ├── Image (Stretch=None)          ├── BackgroundLayer (absolute, maxWidth:none)
              └── Overlay Controls              └── FieldLayer (absolute positioned)
```

## Files Modified

| File | Change | Type |
|------|--------|------|
| `paperless/components/Runtime/PageSurface.tsx` | **NEW** — Fixed-size container (WPF Canvas equivalent) | New component |
| `paperless/components/Runtime/BackgroundLayer.tsx` | **NEW** — Image layer with `maxWidth: "none"` | New component |
| `paperless/components/Runtime/RuntimeFormViewer.tsx` | Refactored to use `<PageSurface>` + `<BackgroundLayer>` + `<RuntimeCanvas>` | Modified |
| `paperless/app/page.tsx` | Form view wrapper: removed `p-6 flex justify-center`, simplified to bare `overflow-auto` | Modified |
| `paperless/components/Runtime/index.ts` | Added exports for `PageSurface` and `BackgroundLayer` | Modified |
| `paperless/app/globals.css` | Updated print selectors for `[data-page-surface]` and `[data-background-layer]` | Modified |
| `ExcelAPI/ExcelAPI/Program.cs` | Fixed typo: `eapp.Run()` → `app.Run()` | Bug fix |

## Detailed Changes

### 1. PageSurface.tsx — The WPF Canvas Equivalent

```tsx
export function PageSurface({ widthPx, heightPx, children }: PageSurfaceProps) {
  return (
    <div
      data-page-surface
      style={{
        position: "relative",   // Creates positioning context for absolute children
        width: widthPx,          // Fixed pixel width (e.g., 2550px)
        height: heightPx,        // Fixed pixel height (e.g., 3299px)
        overflow: "hidden",      // Content clips to page surface
        flexShrink: 0,           // Prevents flex parents from shrinking it
      }}
    >
      {children}
    </div>
  );
}
```

Key design decisions:
- **`position: relative`** — creates the positioning context so all child layers can use `position: absolute` relative to the PageSurface
- **Fixed `width` and `height`** — guaranteed to match the background PNG pixel dimensions
- **`overflow: hidden`** — content outside the page surface is clipped
- **`flexShrink: 0`** — prevents any flex parent from shrinking the surface
- **No responsive sizing, no percentage widths, no flex** — the surface is always at its native pixel size

### 2. BackgroundLayer.tsx — Stretch=None Equivalent

```tsx
<img
  src={src}
  data-form-background
  style={{
    display: "block",
    width: widthPx,
    height: heightPx,
    maxWidth: "none",  // CRITICAL: overrides Tailwind preflight img { max-width: 100% }
  }}
  draggable={false}
/>
```

Key design decisions:
- **`maxWidth: "none"`** — explicitly overrides Tailwind CSS preflight's `img { max-width: 100%; height: auto; }`. This was the root cause of the 470px offset identified in Phase 38/42.
- **Explicit `width` and `height`** — the image renders at its native pixel dimensions, matching WPF `Stretch=None`
- **Wrapped in `position: absolute` div** — takes it out of normal document flow, fixes it to the PageSurface coordinate space

### 3. RuntimeFormViewer.tsx — Refactored Architecture

**Before:**
```tsx
<div className="flex flex-col items-center gap-6">
  <div className="relative inline-block shadow-sm" data-form-page>
    <img ... />
    <RuntimeCanvas ... />
  </div>
</div>
```

**After:**
```tsx
<PageSurface widthPx={sheet.pageWidthPx} heightPx={sheet.pageHeightPx}>
  <BackgroundLayer ... />
  {hasSheetFields && <RuntimeCanvas ... />}
</PageSurface>
```

Changes:
- Removed the responsive `flex flex-col items-center gap-6` wrapper
- Removed `data-form-page` attribute (replaced by `data-page-surface` on PageSurface)
- Background image is now inside a dedicated layer component
- RuntimeCanvas (FieldLayer) is a sibling inside PageSurface — shares the same coordinate space

### 4. page.tsx — Simplified Viewport

**Before:**
```tsx
<div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-auto p-6 flex justify-center">
  <RuntimeFormViewer ... />
</div>
```

**After:**
```tsx
<div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-auto">
  <RuntimeFormViewer ... />
</div>
```

Changes:
- **Removed `p-6`** — no padding inside the viewport. The PageSurface sits flush against the viewport edges, matching WPF ScrollViewer behavior.
- **Removed `flex justify-center`** — the PageSurface is not centered. On wide viewports, it starts at the left edge and the overflow scrollbar handles overflow.
- **Kept `overflow: auto`** — the viewport scrolls when the PageSurface is larger than the viewport, same as WPF ScrollViewer.

### 5. globals.css — Print Styles

Updated selectors to support the new `[data-page-surface]` and `[data-background-layer]` data attributes in print mode.

### 6. Program.cs — Bug Fix

Fixed `eapp.Run()` → `app.Run()` on line 175. The variable `app` was declared via `var app = builder.Build()` but the final call mistakenly used `eapp` which doesn't exist. This caused a C# compilation error CS0103.

## Verification

### Build Status

Frontend: ✅ 0 errors
Backend: ✅ 0 errors (after fixing `eapp` typo)

### Expected Browser Behavior

After the fix, with the frontend and backend running:

1. **Image renders at native size**: `naturalWidth=2550`, `renderedWidth=2550` (was 1182px before)
2. **All layers share one coordinate space**: PageSurface, BackgroundLayer, FieldLayer all at 2550×3299px
3. **Overlay fields align with cells**: field_A1 at 875px appears exactly over cell A1 in the PNG
4. **Viewport scrolls**: on smaller screens, the overflow-auto container scrolls the PageSurface
5. **No responsive layout**: the PageSurface never scales or constrains

### How to Verify (user to run):

1. Start backend: `cd ExcelAPI/ExcelAPI && dotnet run --launch-profile http`
2. Start frontend: `cd paperless && npm run dev`
3. Open browser to the frontend URL
4. Upload a template (e.g., FormTest - Copy.xlsx)
5. Open DevTools and verify:
   ```javascript
   const img = document.querySelector('[data-form-background]');
   console.log('natural:', img.naturalWidth, 'x', img.naturalHeight);  // 2550 x 3299
   console.log('rendered:', img.width, 'x', img.height);                // 2550 x 3299
   console.log('maxWidth:', getComputedStyle(img).maxWidth);           // "none"
   
   const surface = document.querySelector('[data-page-surface]');
   const sr = surface.getBoundingClientRect();
   console.log('surface:', sr.width, 'x', sr.height);                  // 2550 x 3299
   
   const field = surface.querySelector('[style*="position: absolute"]');
   const fr = field.getBoundingClientRect();
   console.log('firstField:', fr.left, fr.top, fr.width, fr.height);
   ```

## Files Referenced

- `paperless/components/Runtime/PageSurface.tsx` — New component
- `paperless/components/Runtime/BackgroundLayer.tsx` — New component
- `paperless/components/Runtime/RuntimeFormViewer.tsx` — Refactored
- `paperless/components/Runtime/RuntimeCanvas.tsx` — Unchanged (used as FieldLayer)
- `paperless/components/Runtime/RuntimeField.tsx` — Unchanged
- `paperless/components/Runtime/index.ts` — Updated exports
- `paperless/app/page.tsx` — Form view wrapper simplified
- `paperless/app/globals.css` — Print styles updated
- `ExcelAPI/ExcelAPI/Program.cs` — Fixed `eapp` typo
