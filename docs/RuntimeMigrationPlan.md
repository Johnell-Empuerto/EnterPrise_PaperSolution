# Runtime Migration Plan: PaperLess Web Runtime v2

## Executive Summary

Our current runtime mistakenly combined Designer yellow-styling with runtime input controls, plus introduced non-standard event interception (`stopPropagation`, `pointerEvents: "none"`) that broke native browser focus. This plan restructures the runtime to match the original i-Reporter architecture while keeping React as the rendering layer.

---

## 1. Desired Architecture

```
PageSurface (fixed w/h, position:relative, overflow:hidden)
├── BackgroundLayer (<img> with maxWidth:none)
└── RuntimeLayer (position:absolute, inset:0, zIndex:25, pointerEvents:none)
    └── RuntimeField × N (position:absolute, pointerEvents:auto, zIndex:26)
        ├── TextField   → <input> or <textarea>  (permanent, transparent)
        ├── CheckboxField → <input type="checkbox"> (permanent)
        ├── DateField   → <input type="date">    (permanent)
        ├── NumberField → <input type="number">  (permanent)
        ├── DropdownField → <select>             (permanent)
        └── SignatureField → <div>               (permanent)
```

Key properties of `RuntimeLayer`:
- **`pointerEvents: "none"`** — the layer div itself passes all events through
- Individual `RuntimeField` containers have **`pointerEvents: "auto"`** — only field areas are interactive
- Absolutely positioned to overlay the background exactly

Each field is a **permanent HTML form control** — always mounted, always editable, never replaced.

---

## 2. Component Hierarchy (Current vs Proposed)

| Layer | Current (Broken) | Proposed (Fixed) |
|-------|------------------|------------------|
| Viewer | `RuntimeFormViewer` | `RuntimeFormViewer` — unchanged |
| Page container | `PageSurface` | `PageSurface` — unchanged |
| Background | `BackgroundLayer` | `BackgroundLayer` — unchanged |
| Field layer | `RuntimeCanvas` with `pointerEvents:none`, re-renders on parent change | `RuntimeFieldLayer` — **memoized**, stable reference, re-renders only on value/sheet change |
| Field wrapper | `RuntimeField` with `stopPropagation()` | `RuntimeField` — **no stopPropagation()**, **no pointerEvents override** |
| Field components | Yellow backgrounds, debug logging | **Transparent backgrounds**, no debug logging in production |
| Selection overlay | Rendered inside `RuntimeFormViewer` alongside fields | Selection overlay moves to **PaperlessDesigner** (outside the Runtime layer) |

---

## 3. Violation Analysis: Every Component That Breaks the Runtime Model

### Component: `RuntimeField.tsx`
**Violations:**
1. **`onMouseDown` with `stopPropagation()`** (line 83-93) — Interferes with React's synthetic event system and native browser focus. WPF does NOT stop event propagation on field containers.
2. **Z-index mixing** — `zIndex: 26` is the same range as selection overlays. Should use a separate layer.
3. **Re-renders on every parent change** — Not memoized. Every `PaperlessDesigner` state change cascades through `RuntimeFormViewer` → `RuntimeCanvas` → `RuntimeField` → `TextField`, risking DOM replacement.

**Fix:**
- Remove `onMouseDown` handler entirely (let events flow naturally)
- Wrap in `React.memo` so it only re-renders when its specific `overlay.id` or `value` changes
- Use stable z-index (e.g., `zIndex: 10` for fields, vs `zIndex: 30` for designer overlays)

### Component: `TextField.tsx`
**Violations:**
1. **Yellow background** (`rgba(254, 249, 195, 0.25)`) — Runtime fields should be transparent
2. **Yellow border** (`rgba(234, 179, 8, 0.6)`) — Runtime fields should have no border, or only on focus
3. **`outline: "none"`** — Removes native focus indicator. Runtime should use the browser's default focus ring (or a subtle blue outline).
4. **Controlled component with `value={...}`** — Creates a round-trip dependency on parent re-render for every keystroke. While controlled inputs can work, the current implementation breaks because the parent re-renders and resets the value before the browser processes the input.
5. **Extensive debug logging** — `console.log` on every render, every event. This should be behind a `debug` flag or removed in production.

**Fix:**
- Remove yellow background → use `background: "transparent"`
- Remove yellow border → use `border: "none"`, or `border: "1px solid transparent"` with focus style
- Remove `outline: "none"` → use `outline: "none"` only if providing custom focus style
- Add focus style: `border: "1px solid #1976d2"` on focus
- Move debug logging behind a `debug` prop or `process.env.NODE_ENV !== "production"`
- Consider `React.memo` to prevent unnecessary re-renders

### Component: `RuntimeCanvas.tsx`
**Violations:**
1. **`pointerEvents: "none"` on the canvas div** (line 48) — This is actually CORRECT for the field layer container, but combined with `RuntimeField`'s `pointerEvents: "auto"`, it creates an isolated event island that may confuse React's event delegation.
2. **Re-creates `interactiveOverlays` array on every render** — Even when no overlay data changes, the `useMemo` recalculates because `overlayCollection` is a new object each time (created in `RuntimeFormViewer`).

**Fix:**
- Ensure `RuntimeFieldLayer` (renamed from `RuntimeCanvas`) uses `React.memo` with stable comparison
- The `overlayCollection` object should be stable between renders when overlay data hasn't changed. This may require changes in `RuntimeFormViewer`.

### Component: `RuntimeFormViewer.tsx`
**Violations:**
1. **Selection overlay rendered inside the PageSurface** (lines 147-166) — The blue selection highlight shares the same coordinate space as fields. While it has `pointerEvents: "none"`, its presence in the DOM at `zIndex: 30` (above fields at `zIndex: 26`) could cause layout/focus interference.
2. **`sheetCollection` is a new object every render** (lines 81-87) — Even when overlays haven't changed, this creates new object references, causing all children to re-render.

**Fix:**
- Move selection overlay to PaperlessDesigner (parent component) — completely outside the Runtime layer
- Memoize `sheetCollection`

### Component: `PaperlessDesigner.tsx`
**Violations:**
1. **`handleMouseDown` with guarded `preventDefault`** (lines 266-284) — Even though it correctly skips `preventDefault` for form controls, the fact that this handler exists on the workspace div means ANY mouse event passes through this check. The `target.closest()` check adds unnecessary overhead and risk of false positives.
2. **`mergedForm` is recalculated on every `fieldConfigs` change** (lines 88-99) — This creates new objects that cascade through the entire component tree, causing unnecessary re-renders.

**Fix:**
- The workspace div's `onMouseDown` should be for panning only. Remove the form-control guard — if the event reaches the workspace, it genuinely missed the field (which is correct behavior).
- Memoize `mergedForm` more aggressively

---

## 4. Root Cause Analysis: Focus Loss

The investigation logging (once deployed) will confirm which hypothesis is correct, but based on code analysis:

**Most likely root cause**: The `stopPropagation()` in `RuntimeField.onMouseDown` prevents React's root-level event delegation from properly tracking focus changes on controlled inputs.

In React 18:
1. Controlled inputs use a "value tracker" that listens for native `input` events
2. When `stopPropagation()` is called on `mousedown` in a child component, React's internal focus tracking at the root level does not receive the event
3. React then fails to properly initialize the controlled input's value tracker
4. When a `keydown` event fires, React's value tracker is not active, and React treats the incoming value change as invalid
5. React reverts the DOM value to the React state value (empty string)
6. The revert causes the browser to reset the input cursor position, which triggers a focus re-evaluation
7. If the re-render also replaces the DOM node (due to key changes or component unmount), focus is permanently lost

**Second possibility**: `RuntimeField` re-renders cause the `<textarea>` DOM node to be unmounted and remounted. This would happen if:
- `overlay.id` changes (unlikely)
- The `key` prop on `RuntimeField` changes (`key={overlay.id}` is stable, so unlikely)
- A parent component unmounts the entire `RuntimeCanvas` (e.g., `showOverlay` toggles, or `hasSheetFields` changes)

The `TEXTAREA MOUNT`/`UNMOUNT` logs will confirm this.

---

## 5. Migration Plan (Step by Step)

### Phase 1: Fix Runtime Styling (Low Risk)
**Goal**: Make runtime fields look like native controls, not designer overlays.

Files to modify:
- `TextField.tsx` — Remove yellow background/border, add focus-only border
- `NumberField.tsx` — Same
- `DateField.tsx` — Same
- `CheckboxField.tsx` — Remove yellow accent/outline
- `SignatureField.tsx` — Remove yellow border/background

New default styles for ALL runtime fields:
```css
background: transparent;
border: none;            /* or 1px solid transparent */
outline: none;
padding: 1px 2px;
box-sizing: border-box;
```

Focus style (via CSS `:focus` or React state):
```css
border: 1px solid #1976d2;
/* OR */
outline: 1px solid #1976d2;
outline-offset: -1px;
```

### Phase 2: Remove Event Interception (Medium Risk)
**Goal**: Let browser events flow naturally without `stopPropagation`.

Files to modify:
- `RuntimeField.tsx` — Remove `onMouseDown` handler entirely
- `PaperlessDesigner.tsx` — Simplify `handleMouseDown` to only handle panning (no form-control guard)

This is the critical fix. After this change:
- `mousedown` on a textarea will bubble naturally up to the workspace
- The workspace's `handleMouseDown` will fire
- It should not call `preventDefault` (because the target is within the field layer)
- The browser will handle focus normally

But wait — if the mousedown bubbles up to the workspace div, `handleMouseDown` might start panning. We need to ensure the workspace handler correctly distinguishes between clicking on a form control vs clicking on empty space.

Currently the guard is:
```tsx
if (target.closest("input, textarea, select, button, [contenteditable]")) return;
e.preventDefault();
startPan(e.clientX, e.clientY);
```

This should still work without `stopPropagation` — when clicking a textarea, `target.closest("textarea")` returns the textarea, so it returns early without starting a pan.

### Phase 3: Stabilize DOM Nodes (Medium Risk)
**Goal**: Prevent unnecessary re-renders that could cause DOM replacement.

Files to modify:
- `RuntimeField.tsx` → Wrap in `React.memo`
- `RuntimeCanvas.tsx` → Wrap in `React.memo`, rename to clarify it's the field layer
- `RuntimeFormViewer.tsx` — Memoize `sheetCollection`
- `TextField.tsx` → Wrap in `React.memo`

Create a stable comparison function:
```tsx
function fieldPropsAreEqual(prev: RuntimeFieldProps, next: RuntimeFieldProps) {
  return prev.overlay.id === next.overlay.id 
    && prev.value === next.value
    && prev.overlay.leftPt === next.overlay.leftPt
    // ... etc
}
```

### Phase 4: Separate Selection from Runtime (Low Risk)
**Goal**: The selection overlay should not be inside the Runtime layer.

Files to modify:
- `RuntimeFormViewer.tsx` — Remove selection overlay rendering (lines 147-166)
- `PaperlessDesigner.tsx` — Add selection overlay rendering in the designer area

This ensures that clicking a field for input never triggers a visual selection change that could re-render the runtime layer.

### Phase 5: Debug Cleanup (Low Risk)
**Goal**: Remove or gate debug logging.

Files to modify:
- `TextField.tsx` — Remove `console.log` calls or gate behind `debug` prop
- `RuntimeField.tsx` — Same
- `RuntimeCanvas.tsx` — Same
- `useRuntimeState.ts` — Same

---

## 6. State Ownership Diagram

```
PaperlessDesigner
│
├── Runtime State (useRuntimeState)
│   ├── values: Record<string, string|boolean|null>
│   ├── setValue(id, val)
│   └── dirty: Record<string, boolean>
│
├── Designer State
│   ├── selectedFieldId: string | null
│   ├── fieldConfigs: Record<string, FieldConfig>
│   └── zoom, offset, etc.
│
├── [Designer Overlays] ← only rendered when selectedFieldId is set
│   └── Selection highlight (blue border)
│
└── Runtime Layer ← always rendered, stable
    └── RuntimeField × N (React.memo)
        └── <input>/<textarea> (controlled)
            └── value: runtime.values[id]
            └── onChange: runtime.setValue(id, newVal)
```

**Data flow for typing:**
```
User types 'A'
  → native input event fires
  → React onChange fires
  → handleChange('A')
  → runtime.setValue('p1f1', 'A')
  → runtime.values changes reference
  → TextField re-renders with value='A'
  → textarea DOM updated with value='A'
  
  (Only RuntimeLayer and the single TextField re-render)
  (PaperlessDesigner does NOT re-render)
  (No selection change, no config panel re-render)
```

---

## 7. Event Lifecycle (Desired)

```
1. mousedown on <input>/<textarea>
   → Browser gives focus to element (default action)
   → Event bubbles up through RuntimeField → RuntimeLayer → PageSurface → workspace
   → Workspace handleMouseDown checks target.closest("input, textarea, ...")
   → Returns early (no preventDefault, no pan start)

2. focus on <input>/<textarea>
   → onFocus fires (optional: apply focus style)
   → Browser shows caret

3. keydown → keypress → input → change → blur
   → Normal browser input processing
   → No React interference, no re-renders during typing (except the field itself)

4. blur
   → onBlur fires (optional: remove focus style)
   → Value is already in runtime state
```

---

## 8. Rendering Lifecycle (Desired)

```
Initial mount:
  1. PaperlessDesigner mounts
  2. RuntimeFormViewer mounts → PageSurface mounts
  3. BackgroundLayer mounts (<img> loads PNG)
  4. RuntimeFieldLayer mounts (with pointerEvents:none)
  5. RuntimeField × N mount (with pointerEvents:auto)
  6. TextField/CheckboxField/etc mount
  7. <input>/<textarea> DOM nodes created

During editing:
  User types
    → runtime.setValue called
    → runtime.values reference changes
    → ONLY TextField re-renders (React.memo)
    → <input>/<textarea> value prop updates
    → DOM node stays mounted
    → Focus preserved

Page change:
  currentPage changes
    → RuntimeFormViewer recalculates overlays
    → RuntimeFieldLayer receives new overlayCollection
    → RuntimeField × N re-render (or remount if keys differ)
    → Fields on new page mount
    → Focus may be lost (acceptable — user navigated to new page)

Config change:
  fieldConfigs[id] changes
    → PaperlessDesigner recalculates mergedForm
    → RuntimeFormViewer receives new mergedForm
    → overlayCollection recalculated
    → RuntimeField for id re-renders with new config
    → TextField re-renders with new config
    → DOM node stays mounted (if key stable)
    → Focus preserved
```

---

## 9. Risk Assessment

| Change | Risk | Mitigation |
|--------|------|------------|
| Remove `stopPropagation()` | Medium — may cause workspace to start panning when clicking fields | Ensure `handleMouseDown` guard works correctly without stopPropagation |
| Change field styling to transparent | Low — purely cosmetic | Can be done incrementally |
| Add `React.memo` to components | Low — only affects re-renders | Test with console logs to verify reduced re-renders |
| Move selection overlay out of RuntimeFormViewer | Low — only moves rendering location | Verify selection still renders correctly |
| Remove debug logging | Low — no functional impact | Can be done last |

**The highest-risk change is Phase 2 (removing `stopPropagation()`).** This should be tested carefully by:
1. First, just add a `console.log` to see if `handleMouseDown` fires when it shouldn't
2. Then, conditionally remove `stopPropagation` based on a flag
3. Finally, remove it permanently once verified

---

## 10. Implementation Order (Recommended)

1. **Phase 1** — Fix styling (transparent fields, focus borders) — Immediate visual improvement, zero risk
2. **Phase 5** — Remove debug logging — Clean up console noise
3. **Phase 3** — Stabilize DOM nodes with `React.memo` — Prevents unnecessary re-renders
4. **Phase 4** — Separate selection from runtime — Stops selection changes from affecting runtime
5. **Phase 2** — Remove event interception — The critical fix for focus loss

After each phase, deploy and test. The focus issue may resolve before Phase 2 if the root cause is re-render cascade (which Phase 3 + 4 address).

If the focus issue persists after Phase 3 + 4, then Phase 2 (removing `stopPropagation`) is the necessary fix.
