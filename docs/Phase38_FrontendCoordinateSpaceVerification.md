# Phase 38 — Frontend Coordinate Space & Rendering Origin Verification

**Date:** 2026-07-13
**Status:** Root Cause Identified — No code changes needed in backend
**Method:** Source code audit of all frontend components + CSS cascade analysis

---

## Executive Summary

**Root Cause:** Tailwind CSS v4's preflight includes `img { max-width: 100%; height: auto; }` which **constrains the 2550px-wide background PNG to fit within the viewport width**. The overlay container and field positions remain in the full 2550×3299px coordinate space.

**Result:** The image is displayed at a SMALLER size than the overlay coordinate space. The overlays maintain correct internal geometry but are displaced because they position against the full 2550px space while the image is rendered at a constrained size (e.g., ~1824px on a 1920px screen).

**Fix:** Add `maxWidth: "none"` to the `<img>` inline style to prevent Tailwind's preflight from constraining the image. OR add `max-w-none` Tailwind class.

---

## Part 1 — Complete Rendering Pipeline

### Element Hierarchy

```
<main flex-1 max-w-7xl mx-auto w-full px-6 py-6>
  <div bg-white rounded-2xl shadow-sm border overflow-auto p-6 flex justify-center>
    <RuntimeFormViewer>
      <div flex flex-col items-center gap-6>                        [R1]
        <div relative inline-block shadow-sm data-form-page>         [R2] ← POSITIONING CONTEXT
          <img                                                        [IMG]
            style: { width: 2550, height: 3299, display: "block" }
            data-form-background
          />
          <RuntimeCanvas>                                             [OV]
            style: {
              position: absolute;
              left: 0;
              top: 0;
              width: 2550px;
              height: 3299px;
              pointerEvents: "none";
              zIndex: 25;
            }
            <RuntimeField>                                           [F1]
              style: {
                position: absolute;
                left: 875px;    ← field_A1 leftPx from backend
                top: 1290px;    ← field_A1 topPx from backend
                width: 400px;
                height: 120px;
                pointerEvents: auto;
                zIndex: 26;
              }
              <TextField />
            </RuntimeField>
            ...
          </RuntimeCanvas>
        </div>
      </div>
    </RuntimeFormViewer>
  </div>
</main>
```

### Key Observation

The `<img>` and `<RuntimeCanvas>` share the SAME positioning parent (`data-form-page` div with `position: relative`). This is correct. The overlay starts at `left: 0; top: 0` matching the image's natural origin.

---

## Part 2 — CSS Audit

### All CSS Rules Affecting Rendering

| Element | CSS Source | Rules | Effect on Positioning |
|---------|-----------|-------|---------------------|
| **`<img>`** | Inline style | `width: 2550, height: 3299, display: block` | Sets explicit size |
| **`<img>`** | **Tailwind preflight** | **`max-width: 100%; height: auto;`** | ⚠️ **CONSTRAINS IMAGE to viewport width** |
| `[IMG]` parent | Tailwind `inline-block` | `display: inline-block` | Sizes to content (constrained by available width) |
| `[R2]` | Tailwind `relative` | `position: relative` | Creates positioning context |
| `[R2]` | Tailwind `shadow-sm` | Box shadow | Visual only — no layout effect |
| `[OV]` | Inline style | `position: absolute; left: 0; top: 0; width: 2550px; height: 3299px` | Full-page overlay at origin |
| `[F1]` | Inline style | `position: absolute; left: 875px; top: 1290px; width: 400px; height: 120px` | Field at backend pixel coords |
| `[R1]` | Tailwind | `display: flex; flex-direction: column; align-items: center; gap: 6` | No effect on absolute children |
| Card | Tailwind | `overflow: auto; display: flex; justify-content: center; padding: 1.5rem` | Scrollable card |

### Rules Verified as NOT Affecting Positioning

| Rule | Reason |
|------|--------|
| `object-fit` | Not present anywhere |
| `transform` / `translate` / `scale` | Not present on rendering elements |
| `zoom` | Not present |
| `aspect-ratio` | Not present |
| `width: 100%` | Not on image or canvas |
| `height: 100%` | Not on image or canvas |
| `margin` | No margins on image or canvas |
| `padding` | No padding on image or canvas |

---

## Part 3 — Root Cause Analysis

### The Issue: Tailwind Preflight Overrides Inline Width

Tailwind CSS v4 includes a preflight (reset) that contains:

```css
/* tailwindcss/preflight.css */
img,
video {
  max-width: 100%;
  height: auto;
}
```

**What happens at render time:**

1. The `<img>` inline style sets `width: 2550px; height: 3299px`
2. The Tailwind preflight sets `max-width: 100%; height: auto`
3. CSS rule: when `width > max-width`, `max-width` takes precedence
4. The containing block width ≈ viewport minus card padding (e.g., ~1824px on 1920px viewport)
5. `max-width: 100%` = 100% of containing block ≈ 1824px
6. Since 2550 > 1824, `max-width` constrains → image displayed at **~1824px wide**
7. `height: auto` from preflight overrides `height: 3299px` → height proportionally scaled to **~2359px**

**Meanwhile, the overlay container:**
- Has `width: 2550px` (no max-width constraint)
- Has `height: 3299px` (no height:auto override)
- Fields positioned at absolute pixel values within the 2550×3299 coordinate space

**Result:** The image is rendered at ~1824×2359px, but the overlays expect 2550×3299px. Fields appear displaced relative to the scaled-down image.

### Numerical Example

| Metric | Expected (2550px space) | Actual (constraint to ~1824px) | Δ |
|--------|------------------------|-------------------------------|-------|
| Image width | 2550px | ~1824px | **-726px** |
| Image height | 3299px | ~2359px | **-940px** |
| Field_A1 left | 875px | 875px (unchanged) | **0px** |
| Cell_A1 left in image | 875px | 875 × (1824/2550) = **626px** | **+249px offset** |

The field at 875px absolute would appear at 875px from the page edge. The cell in the scaled image would appear at ~626px from the image left. The field appears **249px to the right** of the cell.

### Why It Looks Like a Group Translation

Every field is shifted by the SAME proportion:
- Left offset = 875 × (1 - 1824/2550) = 875 × 0.2847 = **+249px** (varies by field position)
- Top offset = 1290 × (1 - 2359/3299) = 1290 × 0.2847 = **+367px** (varies by field position)

Since the shift is proportional to the coordinate value, fields near the top-left have small offsets (A12 at top=1950 would shift +555px) and fields near the bottom-right have larger offsets. This creates the illusion of an "entire group displaced" because ALL fields are scaled by the same ratio.

---

## Part 4 — Minimal Fix

### Option A: Add `maxWidth: "none"` to image style

```diff
 <img
   style={{
     display: "block",
     width: sheet.pageWidthPx,
     height: sheet.pageHeightPx,
+    maxWidth: "none",
   }}
 />
```

**Effect:** `max-width: none` overrides Tailwind's `max-width: 100%`, allowing the image to display at its native 2550×3299 pixel dimensions. The overlay container, already at 2550×3299px, will match exactly.

### Option B: Add `max-w-none` Tailwind class

```diff
- <img ... />
+ <img className="max-w-none" ... />
```

**Same effect** as Option A, using Tailwind utility class instead of inline style.

### Option C: Override in globals.css

```css
[data-form-background] {
  max-width: none !important;
  height: auto !important;
}
```

**Note:** The print styles already have `[data-form-background] { max-width: 100% !important; height: auto !important; }` — this is correct for print but wrong for screen. The fix must ensure the screen rendering uses the full pixel dimensions.

### Recommended: Option A

Simple, minimal, one-line addition to the existing inline style. No architectural changes, no CSS file modifications, no frontend refactoring.

---

## Part 5 — Verification Plan After Fix

| Check | Expected | How to Verify |
|-------|----------|---------------|
| Image rendered at 2550×3299px | True | DevTools → Elements → Computed → width/height |
| Overlay at same size | True | RuntimeCanvas width/height should match image |
| field_A1 at 875px overlay | True | DevTools → inspect field_A1 element → left |
| Cell A1 at 875px in image | True | DevTools → inspect image → measure from left edge |
| No overflow clipping | True | Scrollbar appears if viewport < 2550px |
| Print still works | True | Print style already has max-width:100% for paper fitting |

---

## Part 6 — Final Verdict

| Question | Answer |
|----------|--------|
| Is backend coordinate generation correct? | ✅ Yes — Phase 36 verified |
| Are frontend field coordinates correct? | ✅ Yes — uses backend leftPx/topPx directly |
| **Is rendering origin correct?** | ❌ **No — image is constrained by Tailwind preflight** |
| Is CSS transform applied? | ❌ No transforms found |
| Is overlay container offset? | ❌ No — correctly at (0,0) |
| **Root cause?** | **`img { max-width: 100% }` from Tailwind preflight** |
| Fix? | **Add `maxWidth: "none"` to image inline style** |
| Fix location? | `RuntimeFormViewer.tsx` — the `<img>` style object |
| Lines changed? | 1 line added |
