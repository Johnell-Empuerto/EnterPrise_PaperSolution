# Runtime V2 Architecture Specification

## Chief Architect's Review

**Date:** 2026-07-16
**Status:** Draft for Review
**Version:** 1.0

---

## Table of Contents

1. [Overall Architecture](#1-overall-architecture)
2. [Layer Hierarchy](#2-layer-hierarchy)
3. [Runtime Component Tree](#3-runtime-component-tree)
4. [State Architecture](#4-state-architecture)
5. [Rendering Architecture](#5-rendering-architecture)
6. [Event Architecture](#6-event-architecture)
7. [Runtime vs Designer Separation](#7-runtime-vs-designer-separation)
8. [Performance Strategy](#8-performance-strategy)
9. [Scalability Strategy](#9-scalability-strategy)
10. [Future Extensibility](#10-future-extensibility)
11. [Migration Roadmap](#11-migration-roadmap)
12. [Risks and Trade-offs](#12-risks-and-trade-offs)

---

## 1. Overall Architecture

### 1.1 Guiding Philosophy

> If the original PaperLess team built this today for the browser, how would they architect it?

The original WPF runtime had a simple, effective architecture:
- One `Window` per form
- One `Canvas` per page
- Permanent WPF controls per field
- Direct data binding to a DataTable

There was no designer overlay logic, no selection state, no editing modes, no temporary canvas-drawn rectangles. Fields were simply there, always editable, always mounted.

**Runtime V2 replicates this simplicity while taking advantage of React's strengths.**

### 1.2 High-Level Architecture

```
┌─────────────────────────────────────────────────────────┐
│                     Application Shell                     │
│  (Routing, Auth, Layout, Theme - designer or runtime)     │
├─────────────────────────────────────────────────────────┤
│                                                          │
│   ┌───────────────────────────────────────────────┐      │
│   │              Runtime Application                │      │
│   │  (FormFiller - the production runtime)          │      │
│   │                                                 │      │
│   │  ┌─────────────────────────────────────────┐    │      │
│   │  │         Runtime Store (Zustand)          │    │      │
│   │  │  - ValueStore (field values)             │    │      │
│   │  │  - ValidationStore (errors, warnings)    │    │      │
│   │  │  - DirtyStore (unsaved changes)          │    │      │
│   │  │  - SaveQueue (auto-save, debounced)      │    │      │
│   │  │  - HistoryStore (undo/redo)              │    │      │
│   │  └─────────────────────────────────────────┘    │      │
│   │        │                                         │      │
│   │        ▼                                         │      │
│   │  ┌─────────────────────────────────────────┐    │      │
│   │  │          RuntimeViewer                    │    │      │
│   │  │  - Page navigation                       │    │      │
│   │  │  - Zoom/pan (viewport)                   │    │      │
│   │  │  - Page lifecycle                         │    │      │
│   │  │                                           │    │      │
│   │  │  ┌───────────────────────────────────┐   │    │      │
│   │  │  │     RuntimePage (per page)         │   │    │      │
│   │  │  │  - BackgroundImage (PNG)           │   │    │      │
│   │  │  │  - FieldLayer (permanent controls) │   │    │      │
│   │  │  │                                     │   │    │      │
│   │  │  │  ┌─────────────────────────────┐   │   │    │      │
│   │  │  │  │  RuntimeField (per field)     │   │   │    │      │
│   │  │  │  │  - TextField                  │   │   │    │      │
│   │  │  │  │  - CheckboxField             │   │   │    │      │
│   │  │  │  │  - NumberField               │   │   │    │      │
│   │  │  │  │  - DateField                 │   │   │    │      │
│   │  │  │  │  - SelectField               │   │   │    │      │
│   │  │  │  │  - SignatureField            │   │   │    │      │
│   │  │  │  │  - CalculatedField (readonly) │   │   │    │      │
│   │  │  │  └─────────────────────────────┘   │   │    │      │
│   │  │  └───────────────────────────────────┘   │    │      │
│   │  └─────────────────────────────────────────┘    │      │
│   │                                                 │      │
│   │  ┌─────────────────────────────────────────┐    │      │
│   │  │        Runtime Services                   │    │      │
│   │  │  - SaveService (auto-save, debounce)     │    │      │
│   │  │  - ValidationService                     │    │      │
│   │  │  - FocusService (tab order, IME)         │    │      │
│   │  │  - AccessibilityService                   │    │      │
│   │  │  - AutoSaveService                       │    │      │
│   │  │  - CollaborationService (future)         │    │      │
│   │  └─────────────────────────────────────────┘    │      │
│   └───────────────────────────────────────────────┘      │
│                                                          │
│   ┌───────────────────────────────────────────────┐      │
│   │           Designer Application                  │      │
│   │  (Completely separate from Runtime V2)          │      │
│   │  - FormDesigner - template layout               │      │
│   │  - FieldConfigurator - field properties         │      │
│   │  - PreviewMode - uses RuntimeViewer read-only   │      │
│   │  - CanvasModel - ZoomableCanvas with overlays   │      │
│   └───────────────────────────────────────────────┘      │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

### 1.3 Key Architecture Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| State management | **Zustand** | Minimal boilerplate, no providers, excellent TypeScript, supports slices, middlewares (persist, undo), selector-based subscriptions, tree-shakeable, 1KB. Better than Redux (too much boilerplate), better than Context (re-render issues), better than Jotai/Recoil (atomic = overkill for form) |
| Rendering model | **HTML/CSS absolute positioning** | Matches WPF Canvas model. Native input elements for accessibility and browser autofill. No canvas rendering for fields |
| Field approach | **Permanent DOM** | Always mounted, always editable. Matches WPF behavior. Never unmount/remount fields during page navigation. Use `display: none` / `visibility: hidden` for hidden pages |
| Page navigation | **CSS visibility toggle** | All pages rendered in DOM, only current page visible. Prevents unmount/remount of field controls, preserves focus, scroll position, IME state |
| Zoom/Pan | **CSS transform** on viewport wrapper | No scaling of individual elements. Single `transform: scale(Z) translate(X, Y)` on viewport container. Fields remain at native coordinates |
| Backend data | **Pixel coordinates** (single system) | Eliminates pt/px ambiguity. Backend returns `leftPx, topPx, widthPx, heightPx`. No frontend coordinate conversion |

---

## 2. Layer Hierarchy

### 2.1 Visual Layers (per page)

```
RuntimePage (position: relative, width: Npx, height: Npx)
│
├── Layer 0: BackgroundImage
│   position: absolute; inset: 0
│   <img> with native pixel dimensions
│   z-index: 0
│   pointer-events: none (image is non-interactive)
│
├── Layer 1: FieldLayer
│   position: absolute; inset: 0
│   z-index: 1
│   Contains all field controls
│   Each field is position: absolute with left/top/width/height
│
│   ├── RuntimeField (TextField)
│   │   z-index: auto (within layer)
│   │   Native <input>/<textarea>
│   │   Standard browser focus handling
│   │
│   ├── RuntimeField (CheckboxField)
│   │   Native <input type="checkbox">
│   │   Centered within bounds
│   │
│   ├── RuntimeField (NumberField)
│   │   Native <input type="text" inputMode="decimal">
│   │   Controlled value, local validation
│   │
│   ├── RuntimeField (DateField)
│   │   Native <input type="date">
│   │
│   ├── RuntimeField (SelectField)
│   │   Native <select> or custom dropdown
│   │
│   └── RuntimeField (SignatureField)
│       <canvas> or SVG for signature capture
│
└── Layer 2: ActiveSelectionHighlight (optional)
    z-index: 2
    Only rendered when a field is focused
    Thin outline (2px blue), NOT yellow
    Does NOT affect layout or events
```

### 2.2 Layer Implementation Rules

1. **No layer overlap** - Each layer has a distinct z-index range
2. **No pointer-events: none** on field layer - Fields must receive events naturally
3. **No stopPropagation** on field layer or field wrappers - Browser events flow naturally
4. **Background layer is pointer-events: none** - Clicks pass through to field layer
5. **Selection highlight is visual-only** - No event handling, no interaction
6. **All layers are position: absolute** within the page container

### 2.3 Layer Count Comparison

| Layer | Original WPF | Current Runtime | Runtime V2 |
|-------|-------------|-----------------|------------|
| Background | Image at Z=0 | BackgroundLayer | BackgroundImage |
| Fields | Controls at Z=1 | RuntimeCanvas (with issues) | FieldLayer (clean) |
| Selection | N/A (no designer) | Selection highlight div | ActiveSelectionHighlight (thin) |
| Designer overlays | N/A | Yellow rectangles (wrong) | Not present in runtime |

### 2.4 Canvas/Viewport Architecture

```
RuntimeViewer
│
├── ViewportContainer (position: relative, overflow: hidden)
│   │
│   └── CameraTransform (transform: scale(Z) translate(X, Y))
│       │
│       ├── PageContainer (display: flex, flex-direction: column, gap: 20px)
│       │   │
│       │   ├── PageWrapper (for page 1)
│       │   │   │
│       │   │   └── PageSurface (position: relative, width: Npx, height: Npx)
│       │   │       │
│       │   │       ├── BackgroundImage (z-index: 0)
│       │   │       │   <img> - pointer-events: none
│       │   │       │
│       │   │       └── FieldLayer (z-index: 1)
│       │   │           │
│       │   │           ├── RuntimeField (id="field-1") - TextField
│       │   │           │   <input> - standard browser behavior
│       │   │           │
│       │   │           ├── RuntimeField (id="field-2") - CheckboxField
│       │   │           │   <input type="checkbox">
│       │   │           │
│       │   │           └── ...
│       │   │
│       │   └── PageWrapper (for page 2) - visibility: hidden if not current
│       │       │
│       │       └── PageSurface (width: Npx, height: Npx)
│       │           ├── BackgroundImage
│       │           └── FieldLayer
│       │               └── ...
│       │
│       └── (no designer overlays in runtime)
│
├── PageNavigation (bottom bar)
│   Previous | Next | Page N of M
│
└── ZoomControls (corner overlay)
    Zoom In | Zoom Out | Fit Page | Fit Width | 1:1
```

### 2.5 Page Visibility Strategy

```
Pages are NEVER unmounted.
Pages are NEVER conditionally rendered.

Active page:   visibility: visible, display: block
Inactive page: visibility: hidden, display: block (keeps layout, preserves state)

Rationale:
- Prevents field unmount → focus loss
- Preserves scroll position
- Preserves IME composition state
- Preserves partial input (no validation reset)
- Zero-cost for hidden pages (no paint, no events)
```

---

## 3. Runtime Component Tree

### 3.1 Complete Component Hierarchy

```
<RuntimeProvider>                          // Page-level provider (one per form page URL)
│
├── <RuntimeStore>                        // Zustand store initialization (no React component)
│
├── <RuntimeViewer>                       // Main viewer component
│   │
│   ├── <ViewportContainer>              // Scrollable container, overflow: hidden
│   │   │
│   │   └── <CameraTransform>            // CSS transform: scale + translate
│   │       │
│   │       └── <PageContainer>          // flex column, page gap
│   │           │
│   │           ├── <RuntimePage key={1} id={1} pageData={...} isActive={true}>
│   │           │   │
│   │           │   ├── <BackgroundImage
│   │           │   │     src={bgUrl}
│   │           │   │     width={pageWidthPx}
│   │           │   │     height={pageHeightPx}
│   │           │   │   />
│   │           │   │
│   │           │   └── <FieldLayer pageId={1}>
│   │           │       │
│   │           │       ├── <RuntimeField field={field1}>
│   │           │       │   └── <TextField />        // or NumberField, DateField, etc.
│   │           │       │
│   │           │       ├── <RuntimeField field={field2}>
│   │           │       │   └── <CheckboxField />
│   │           │       │
│   │           │       └── ...
│   │           │
│   │           └── <RuntimePage key={2} id={2} pageData={...} isActive={false}>
│   │               │   // visibility: hidden, still mounted
│   │               └── ...
│   │
│   ├── <RuntimeToolbar />                // Page nav, zoom controls, save status
│   │
│   └── <RuntimeStatusBar />              // Dirty indicator, field count, errors
```

### 3.2 Component Responsibilities

| Component | Responsibility | Memo Boundary |
|-----------|---------------|---------------|
| `RuntimeProvider` | Initializes Zustand stores. Provides page-level context | N/A (mount once) |
| `RuntimeViewer` | Loads form data. Manages page navigation, zoom state. Orchestrates rendering | Yes - stable form data |
| `ViewportContainer` | Overflow hidden, wheel/scroll handling, resize observer | Yes - stable |
| `CameraTransform` | Applies CSS transform. Handles zoom/pan | Yes - only on zoom change |
| `PageContainer` | Flex layout for pages | Yes - stable |
| `RuntimePage` | Per-page container. Manages visibility. Hosts bg image + field layer | Yes - only on page data change |
| `BackgroundImage` | Renders PNG. Handles load/error states | Yes - only on src change |
| `FieldLayer` | Positions fields absolutely. Renders selection highlight | Yes - only on field array change |
| `RuntimeField` | Field container. Connects field config to specific input component. Handles tab index, required indicator | Yes - only on own field data change |
| `TextField` | Input/textarea with controlled value from store | Yes - only on own value change |
| `CheckboxField` | Checkbox with controlled checked state | Yes - only on own value change |

### 3.3 Why This Hierarchy

**No dedicated SelectionLayer component in runtime.**
- Selection is managed by the browser (focus/blur events)
- Visual highlight is a CSS `:focus-visible` or `:focus-within` on the RuntimeField wrapper
- No React state for selection = no re-renders on focus change
- Focus highlight is a browser-native concern

**No FieldLayer as a separate abstraction?**
- Yes, it exists as a simple `<div>` container for fields
- It does NOT manage field state, selection, or event handling
- Its only job is CSS positioning context

**No PageProvider or FieldProvider.**
- Page context is passed via props (pageId, dimensions)
- Field data is accessed via Zustand selectors directly in RuntimeField
- No prop drilling through 4+ levels
- No context providers that cause cascading re-renders

---

## 4. State Architecture

### 4.1 Store Diagram

```
┌──────────────────────────────────────────────────────────────────┐
│                       Zustand Root Store                         │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     formSlice                                │ │
│  │  - formId: string                                           │ │
│  │  - pages: RuntimeSheet[] (immutable)                        │ │
│  │  - currentPageIndex: number                                 │ │
│  │  - isLoading: boolean                                       │ │
│  │  - loadForm(formId) => void                                 │ │
│  │  - setPage(index) => void                                   │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     valueSlice                               │ │
│  │  - values: Record<fieldId, FieldValue>                      │ │
│  │  - setValue(fieldId, value) => void                         │ │
│  │  - setMultiple(values) => void (batch, undo boundary)       │ │
│  │  - getValue(fieldId) => FieldValue                          │ │
│  │  - resetValues() => void                                    │ │
│  │  - resetField(fieldId) => void                              │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     dirtySlice                               │ │
│  │  - dirty: Record<fieldId, boolean>                          │ │
│  │  - markDirty(fieldId) => void                               │ │
│  │  - markClean(fieldId) => void                               │ │
│  │  - markAllClean() => void                                   │ │
│  │  - isDirty: boolean (computed)                              │ │
│  │  - dirtyCount: number (computed)                            │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     validationSlice                          │ │
│  │  - errors: Record<fieldId, string[]>                        │ │
│  │  - warnings: Record<fieldId, string[]>                      │ │
│  │  - setErrors(fieldId, errors) => void                       │ │
│  │  - clearErrors(fieldId) => void                             │ │
│  │  - setWarning(fieldId, warnings) => void                    │ │
│  │  - validateField(fieldId) => void                           │ │
│  │  - validateAll() => void                                    │ │
│  │  - hasErrors: boolean (computed)                            │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     saveSlice                                │ │
│  │  - saveStatus: 'idle' | 'saving' | 'saved' | 'error'       │ │
│  │  - lastSavedAt: number | null                               │ │
│  │  - saveQueue: Map<fieldId, FieldValue> (debounced)          │ │
│  │  - enqueueSave(fieldId, value) => void                      │ │
│  │  - flushSaveQueue() => Promise<void>                        │ │
│  │  - setSaveStatus(status) => void                            │ │
│  └─────────────────────────────────────────────────────────────┘ │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────────┐ │
│  │                     historySlice                             │ │
│  │  - past: Snapshot[]                                          │ │
│  │  - future: Snapshot[]                                        │ │
│  │  - pushSnapshot() => void                                   │ │
│  │  - undo() => void                                           │ │
│  │  - redo() => void                                           │ │
│  │  - canUndo: boolean (computed)                              │ │
│  │  - canRedo: boolean (computed)                              │ │
│  │  - maxHistory: 50 (configurable)                            │ │
│  └─────────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────┘
```

### 4.2 Local vs Global State

| State | Location | Access Pattern | Why |
|-------|----------|---------------|-----|
| Form metadata (pages, dimensions) | **Global** (formSlice) | Read on mount, stable | Same for all fields |
| Current page index | **Global** (formSlice) | Single source of truth | Navigation depends on it |
| Field values | **Global** (valueSlice) | Individual field access via selector | Must persist across page changes, needed for save |
| Dirty flags | **Global** (dirtySlice) | Individual field access via selector | Must persist, needed for save |
| Validation errors | **Global** (validationSlice) | Individual field access via selector | Must persist, needed for submit |
| Save queue/status | **Global** (saveSlice) | Read by status bar, write by fields | Single save pipeline |
| Undo/redo history | **Global** (historySlice) | Accessed by keyboard shortcuts | Single undo stack |
| **Focus state** | **Local (browser)** | `document.activeElement` | NEVER in React state. Browser manages this. |
| **Hover state** | **Local (CSS)** | `:hover` pseudo-class | NEVER in React state. CSS manages this. |
| **Scroll position** | **Local (browser)** | Native scroll restoration | NEVER in React state. |
| **IME composition** | **Local (browser)** | `compositionstart/compositionend` | NEVER in React state. |
| **Cursor position** | **Local (browser)** | `selectionStart/selectionEnd` | NEVER in React state. |
| **Field animation** | **Local (CSS)** | CSS transitions/animations | NEVER in React state. |
| **Zoom/pan** | **Local component** | ViewportContainer useState | UI-only, not persisted |
| **Tooltip visibility** | **Local component** | useState | UI-only, transient |

### 4.3 State Access Patterns

```
// === GLOBAL STATE ACCESS ===
// Each field component accesses ONLY its own value via selector

// In RuntimeField (per field instance):
const value = useRuntimeStore(s => s.values[fieldId]);
const setValue = useRuntimeStore(s => s.setValue);
const error = useRuntimeStore(s => s.errors[fieldId]?.[0]);

// This creates a REACTIVE BINDING that only re-renders THIS field
// when ITS OWN value or error changes.
// No other field's change triggers this component.

// === NO PROP DRILLING ===
// Values are NOT passed from RuntimePage → FieldLayer → RuntimeField → TextField
// Each field connects directly to the store

// === SAVE TRIGGER ===
// On blur: check dirty, if dirty, enqueue save
// RuntimeField's useEffect cleanup: flush pending save
```

### 4.4 State Initialization

```
loadForm(formId):
  1. Show loading state
  2. Fetch GET /api/form/runtime/{formId}
  3. Set formSlice.pages = response.sheets
  4. Set formSlice.formId = formId
  5. Set valueSlice.values = initializeDefaults(response.sheets)
     (each field's default value, empty string, etc.)
  6. Set dirtySlice = {} (all clean initially)
  7. Set validationSlice = {} (no errors initially)
  8. Set saveSlice.saveStatus = 'idle'
  9. Set historySlice.past = [snapshot of initial values]
  10. Hide loading state
```

### 4.5 Persistence Model

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Persistence Layer                           │
│                                                                     │
│  Zustand persist middleware (optional, local draft save)            │
│  localStorage: submitDraft, restoreDraft                           │
│                                                                     │
│  Primary persistence:                                               │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │  SaveQueue                                                    │   │
│  │  - Accumulates changed values on each keystroke              │   │
│  │  - Debounce: 2 seconds after last change                     │   │
│  │  - Batch: send all queued changes in one request             │   │
│  │  - Retry: 3 attempts with exponential backoff                │   │
│  │  - Offline: store in IndexedDB, flush on reconnect           │   │
│  │                                                               │   │
│  │  POST /api/form/runtime/{formId}/values                      │   │
│  │  Body: { values: Record<fieldId, string> }                   │   │
│  │  Response: { success: boolean, errors?: FieldError[] }       │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  Conflict resolution:                                              │
│  - Last-write-wins (simplest, acceptable for current use case)     │
│  - Future: version vector per field for collaboration              │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 5. Rendering Architecture

### 5.1 Rendered When?

```
RuntimeViewer
  Renders: On mount, on page change (switching which page is visible)
  Does NOT render: On field value change, on field focus, on field blur

RuntimePage
  Renders: When page data changes (rare - only on load)
  Does NOT render: On field value change, on parent re-render (memo'd)

FieldLayer
  Renders: When field array reference changes (rare)
  Does NOT render: On field value change (children access store directly)

RuntimeField
  Renders: When own field config changes (very rare)
  Does NOT render: When other fields change value (selector-based subscription)

TextField / NumberField / CheckboxField / etc.
  Renders: When own value in store changes
  This is THE MINIMUM POSSIBLE RENDER
  Other fields are NOT affected
```

### 5.2 Memoization Strategy

```
RuntimeViewer          → React.memo (form data is stable)
ViewportContainer      → React.memo (stable)
CameraTransform        → React.memo (only on zoom/pan change)
PageContainer          → React.memo (stable)
RuntimePage            → React.memo (only on own page data change)
BackgroundImage        → React.memo (only on src change)
FieldLayer             → React.memo (only on field array change)
RuntimeField           → React.memo (only on own field config + value change)
TextField              → No memo needed (already minimal, receives value + onChange)
                        BUT: use selectors to extract only what's needed

=== CRITICAL ===
RuntimeField must use zustand selector that subscribes ONLY to its own value:

const value = useRuntimeStore(s => s.values[fieldId]);

NOT:

const values = useRuntimeStore(s => s.values);
// This subscribes to ALL values - every field change re-renders every field
```

### 5.3 Context Splitting

**No React.createContext is used for the Runtime V2.**

Rationale:
- Zustand stores are accessed via hooks directly in components
- No Provider component needed at any level
- No context consumer re-renders to manage
- Zustand's selector mechanism is more precise than Context (which re-renders ALL consumers)
- No prop drilling needed (each field connects to store directly)

**Exception**: If a theme/translation system is needed, use a separate Context for that, but keep it outside the runtime rendering pipeline.

### 5.4 Page Rendering Pipeline

```
1. RuntimeViewer mounts
2. Load form data from API
3. Initialize Zustand stores with form data
4. Render ViewportContainer → CameraTransform → PageContainer

5. PageContainer renders ALL RuntimePage components
     (pages.map(page => <RuntimePage key={page.id} pageData={page} isActive={page.id === currentPageId} />))

6. Each RuntimePage:
   a. Renders BackgroundImage (z-index: 0)
   b. Renders FieldLayer (z-index: 1)
   c. FieldLayer maps fields to RuntimeField components
   d. Each RuntimeField connects to store for its value
   e. RuntimeField renders the appropriate field component

7. Visibility: isActive ? 'visible' : 'hidden' (NOT conditional rendering)
   - Hidden pages: display: block, visibility: hidden (maintains layout, children are alive)
   - Active page: display: block, visibility: visible

8. When user navigates to next page:
   - currentPageId changes in store
   - RuntimePage's isActive prop changes
   - No unmount/remount of any component
   - No field loses its state
```

### 5.5 Field Rendering Pipeline

```
Per RuntimeField:

1. Connect to store for:
   - value (selector: values[fieldId])
   - error (selector: errors[fieldId]?.[0])
   - field definition (passed as prop, stable)

2. Render:
   <div
     className="runtime-field"
     style={{
       position: 'absolute',
       left: field.leftPx,
       top: field.topPx,
       width: field.widthPx,
       height: field.heightPx,
     }}
     data-field-id={field.id}
   >
     <FieldComponent
       field={field}
       value={value}
       onChange={handleChange}
       onBlur={handleBlur}
       error={error}
     />
   </div>

3. handleChange:
   - setValue(fieldId, newValue)
   - markDirty(fieldId)
   - enqueueSave(fieldId, newValue) (debounced)

4. handleBlur:
   - validateField(fieldId)
   - flushSaveQueue() (if dirty)
```

### 5.6 Field Type Dispatch

```
// In RuntimeField, not a separate dispatcher:

const FieldComponent = FIELD_COMPONENT_MAP[field.dataType];
if (!FieldComponent) return null;

<FieldComponent ... />

Where:

const FIELD_COMPONENT_MAP: Record<string, React.ComponentType<FieldProps>> = {
  KeyboardText: TextField,
  InputNumeric: NumberField,
  Check: CheckboxField,
  CalendarDate: DateField,
  Select: SelectField,
  MultiSelect: SelectField,
  Signature: SignatureField,
  FixedText: ReadonlyField,
  FreeText: ReadonlyField,
  Calculate: CalculatedField,
  // ... others
};

This is a simple lookup table, NOT a switch-case in render.
No conditional logic in render = more predictable, easier to memoize.
```

### 5.7 No Virtualization Decision

**Decision: No virtualization for fields.**

Rationale:
- WPF renders ALL controls always. Browser DOM handles thousands of elements fine.
- 10,000 fields × a few DOM nodes ≈ 30,000-50,000 DOM nodes. Modern browsers handle this.
- Virtualization causes focus loss (elements are removed/added to DOM)
- Virtualization breaks tab order (non-rendered elements aren't in tab order)
- Virtualization breaks find-in-page
- Virtualization breaks browser autofill
- Virtualization breaks IME composition
- Virtualization breaks scroll-to-field

**Optimization instead:**
- Efficient selectors (never re-render all fields)
- CSS containment (`contain: layout style` on page wrapper)
- `content-visibility: auto` for hidden pages
- No virtualization needed when rendering is efficient

---

## 6. Event Architecture

### 6.1 Guiding Principle

**Events flow like the browser intended.**

- No stopPropagation
- No preventDefault (unless absolutely necessary for specific behavior)
- No event delegation on field layer
- No synthetic event wrapping
- Let the browser handle focus, blur, keyboard, mouse naturally

### 6.2 Event Flow Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Event Flow Patterns                         │
│                                                                     │
│  FOCUS:                                                            │
│  ──────                                                            │
│  User clicks TextField                                             │
│    → Browser fires mousedown on <input>                            │
│    → Browser fires focus on <input>                                │
│    → Browser fires focusin (bubbles) on <input> → container → body │
│    → Browser positions caret                                       │
│    → IME composition can begin                                     │
│    → No React state changes                                        │
│    → No re-renders                                                 │
│    → RuntimeField's onFocusCapture? Only if needed for tracking    │
│                                                                     │
│  KEYBOARD INPUT:                                                   │
│  ──────────────                                                    │
│  User types 'A'                                                    │
│    → Browser fires keydown → keypress → input → keyup              │
│    → input event: <input> value changes                            │
│    → React onChange fires                                          │
│    → handleChange(fieldId, newValue)                               │
│    → setValue(fieldId, newValue)    (Zustand)                      │
│    → markDirty(fieldId)              (Zustand)                     │
│    → enqueueSave(fieldId, value)     (Zustand, debounced)          │
│    → THIS FIELD re-renders (due to value selector)                 │
│    → NO OTHER FIELD re-renders                                     │
│                                                                     │
│  BLUR:                                                             │
│  ────                                                             │
│  User clicks outside field                                         │
│    → Browser fires blur on <input>                                 │
│    → Browser fires focusout (bubbles)                              │
│    → RuntimeField's onBlur fires                                   │
│    → validateField(fieldId)                                        │
│    → flushSaveQueue() (immediate save on blur)                     │
│    → markClean(fieldId) (optional)                                 │
│                                                                     │
│  TAB NAVIGATION:                                                  │
│  ──────────────                                                    │
│  User presses Tab                                                  │
│    → Browser fires keydown (Tab)                                   │
│    → Browser moves focus to next focusable element in tab order    │
│    → blur fires on current field                                   │
│    → focus fires on next field                                     │
│    → No React state changes during tab switch                      │
│    → No re-renders during tab switch                               │
│                                                                     │
│  IME COMPOSITION:                                                 │
│  ────────────────                                                  │
│  User uses IME to input Chinese/Japanese/Korean                   │
│    → compositionstart event                                        │
│    → compositionupdate events (as user selects candidates)         │
│    → compositionend event                                          │
│    → input event with final character                              │
│    → React onChange fires with final character                     │
│    → setValue is called ONCE, with final result                    │
│    → No intermediate value updates during composition              │
│                                                                     │
│  BROWSER AUTOFILL:                                                │
│  ────────────────                                                  │
│  Browser fills saved credentials/address                          │
│    → Browser sets value directly on input                          │
│    → Browser fires input event                                     │
│    → React onChange fires                                          │
│    → setValue(fieldId, filledValue)                                │
│    → Works naturally, no special handling needed                   │
└─────────────────────────────────────────────────────────────────────┘
```

### 6.3 Event Handling Details

| Event | Handler | Action | Re-render? |
|-------|---------|--------|------------|
| **Focus** | None (browser native) | N/A | No |
| **Blur** | `RuntimeField.onBlur` | Validate field, flush save queue | No (unless validation adds errors) |
| **Input** | `field.onChange` | `setValue`, `markDirty`, `enqueueSave` | Yes (only this field via selector) |
| **KeyDown (Tab)** | None (browser native) | Browser handles tab navigation | No |
| **KeyDown (Enter)** | `TextField.onKeyDown` | If single-line, blur. If multi-line, newline. | No |
| **MouseDown** | None (unless intercept needed) | N/A | No |
| **Paste** | None (browser native) | Browser inserts text, fires input event | Same as input above |
| **Cut** | None (browser native) | Browser removes text, fires input event | Same as input above |
| **CompositionStart** | None | N/A | No |
| **CompositionEnd** | None (fires input event) | Same as input above | Same as input above |
| **Scroll** | None (browser native) | Viewport scrolling | No |

### 6.4 Tab Order

```
Tab order follows the DOM order of focusable elements.

Original WPF behavior: fields are tabbed in the order they were created.
Browser behavior: fields are tabbed in DOM order.

Implementation:
- Fields are rendered in the order they come from the API
- Each RuntimeField wrapper has tabIndex={field.tabIndex} if specified
- If no tabIndex is specified, fields are tabbed in DOM order (API order)

No custom tab navigation logic needed.
No keydown handler for Tab needed.
let the browser handle tab order naturally.
```

### 6.5 Accessibility

- All inputs are native HTML elements (screen-reader compatible)
- Labels use `aria-label={field.name}` on each RuntimeField wrapper
- Required fields have `aria-required="true"`
- Error states use `aria-invalid="true"` and `aria-describedby={errorId}`
- Tab order is natural DOM order
- Zoom works with browser zoom (Ctrl + / -)

### 6.6 Touch/Mobile Events

- Native input elements handle touch events
- No custom touch handlers needed
- On mobile: inputs get native mobile keyboard
- `inputMode` is set based on field type:
  - Number: `inputMode="decimal"`
  - Text: `inputMode="text"`
  - Date: `type="date"` (native date picker on mobile)

---

## 7. Runtime vs Designer Separation

### 7.1 Complete Separation

```
Runtime Application              Designer Application
(pages/form/[id])                (pages/designer/[id])
       │                                │
       │                                │
┌──────────────┐               ┌──────────────────┐
│  Runtime V2   │               │  Designer V2      │
│               │               │                   │
│  Zustand      │               │  DesignerController│
│  RuntimeStore │               │  (class-based)    │
│               │               │                   │
│  RuntimeViewer│               │  ZoomableCanvas   │
│  (read-only   │               │  (pan/zoom/select)│
│   form fill)  │               │                   │
│               │               │  Cluster overlays │
│  FieldLayer   │               │  (yellow, Z=1)    │
│  (permanent   │               │                   │
│   controls)   │               │  Selection (Z=2)  │
│               │               │                   │
│  SaveService  │               │  Resize handles   │
│  Validation   │               │  (Z=3)            │
└──────────────┘               └──────────────────┘
        │                                │
        └─────────── SHARED ────────────┘
                    │
            ┌───────┴───────┐
            │  Shared Types  │
            │               │
            │  RuntimeField │
            │  RuntimeSheet │
            │  FieldConfig  │
            │  FieldValue   │
            └───────────────┘
```

### 7.2 What Goes Where

| Component/Model | Belongs To | Reason |
|----------------|------------|--------|
| `RuntimeViewer` | **Runtime only** | Form filling, no selection, no overlays |
| `RuntimePage` | **Runtime only** | Background + fields, no designer chrome |
| `BackgroundImage` | **Shared** (both need it) | Same component, just renders an image |
| `FieldLayer` | **Runtime only** | Only renders fields, no selection |
| `RuntimeField` | **Runtime only** | Input controls connected to store |
| `TextField` etc. | **Shared** (potentially) | Same input components, may need minor differences |
| `RuntimeStore` (Zustand) | **Runtime only** | Designer uses DesignerController instead |
| `PaperlessDesigner` | **Designer only** | Full IDE: toolbars, explorer, config panel |
| `ConfigPanel` | **Designer only** | Field property editing |
| `FieldExplorer` | **Designer only** | Field list with search/sort |
| `ZoomableCanvas` | **Designer only** | Pan/zoom with designer overlays |
| `Selection overlay` | **Designer only** | Yellow highlight in designer |
| `DesignerController` | **Designer only** | FormDefinition, undo/redo, cluster CRUD |
| Save service | **Runtime only** | Designer has publish, not auto-save |
| Validation service | **Shared** | Both need validation (designer for preview) |

### 7.3 Designer Uses Runtime

```
Designer:
  └── Preview Mode:
       └── RuntimeViewer (readOnly: true, showOverlays: false)
           (Renders the form in read-only mode for preview)

The RuntimeViewer accepts a readOnly prop.
In readOnly mode:
  - Fields are not editable (pointer-events: none on inputs, or disabled)
  - Selection highlights are disabled
  - No save/validation logic
  - Background image + fields still render

This is the ONLY interaction between Designer and Runtime.
Runtime NEVER imports designer code.
```

### 7.4 Import Rule

```
Runtime directory:
  ✗ NEVER imports from Designer directory
  ✗ NEVER imports from CanvasModel directory
  ✗ NEVER imports from ConfigPanel or FieldExplorer
  ✓ Only imports from:
    - Shared types
    - Zustand
    - React/Next.js
    - Utility libraries

Designer directory:
  ✓ CAN import from Runtime (for preview mode)
  ✗ NEVER wraps or modifies Runtime components
  ✗ NEVER adds event handlers to Runtime components
  ✓ Uses RuntimeViewer in readOnly mode only

This ensures:
- Runtime NEVER has designer code executing
- Runtime bundle is minimal (no designer code)
- Designer can be split into separate chunk
- Future: Runtime can be deployed as standalone form filler
```

---

## 8. Performance Strategy

### 8.1 At Rest (No User Interaction)

```
Target: 0 CPU usage, 0 React renders

Achieved by:
- All fields are permanent DOM (no create/destroy overhead)
- React.memo on all stable components
- Zustand selectors subscribe per-field (no cross-field subscriptions)
- No polling, no timers, no animation frames
- CSS handles hover/active states without JavaScript
```

### 8.2 During Typing

```
Target: < 16ms per keystroke (60fps)

Achieved by:
- Typing in Field A:
  → Only Field A re-renders
  → RuntimePage, FieldLayer, other fields: NO re-render
  → RuntimeViewer: NO re-render
  → Entire page does NOT re-render
  → No parent component re-renders

  = 1 component re-render per keystroke (the specific field component)
  = < 1ms for typical text field
```

### 8.3 During Page Navigation

```
Target: < 100ms perceived latency

Achieved by:
- All pages are already rendered (no DOM creation)
- Only CSS visibility changes
- No data fetching
- No state initialization
- Zero React re-renders for field components (they don't subscribe to page index)
```

### 8.4 During Form Load

```
Target: < 2s for 100 pages, 5000 fields

Achieved by:
- Initial render creates ALL page containers
- Each page creates ALL field wrappers
- This is fast because:
  - No DOM creation for hidden pages (they're in the initial render)
  - React's reconciliation handles this efficiently
  - CSS visibility takes no layout cost
  - 5000 fields × ~5 DOM nodes = 25000 nodes. Chrome handles 250k+.

  Risk mitigation:
  - If initial load is too slow, use progressive rendering:
    Render first page immediately, defer hidden pages by 50ms each
    <React.unstable_DeferredValue> or requestIdleCallback
```

### 8.5 At Scale (10,000 Fields)

```
Strategy: NO virtualization, optimize rendering

Per field overhead:
  - 1 RuntimeField wrapper div
  - 1 input element
  - 1 label/description element (optional)
  = ~3-5 DOM nodes per field

10,000 fields × 4 nodes = 40,000 DOM nodes
  40,000 nodes is well within browser capabilities
  Typical gmail: ~100,000+ nodes
  Typical Twitter feed: ~50,000+ nodes
  Chrome DevTools can handle 1M+ nodes

Performance concerns:
  1. Initial render time: mitigated by progressive rendering
  2. Memory: 40,000 nodes × ~200 bytes = 8MB (fine)
  3. Re-render cost: still 1 component per keystroke (selector-based)
  4. Layout cost: CSS containment (contain: layout style) prevents cross-field layout

Scaling beyond 10,000:
  At 50,000+ fields, consider:
  - content-visibility: auto on pages (not fields - field virtualization breaks focus)
  - Render only visible pages initially
  - But even 50,000 fields × 4 nodes = 200,000 nodes (still fine)
```

### 8.6 Performance Optimization Checklist

```
✓ React.memo on: RuntimeViewer, ViewportContainer, CameraTransform, PageContainer,
  RuntimePage, BackgroundImage, FieldLayer, RuntimeField
✓ Zustand selector per field (NOT subscribing to entire values object)
✓ CSS contain: layout style on page wrapper
✓ CSS content-visibility: auto on hidden pages
✓ No context providers (Zustand only)
✓ No event delegation on field layer
✓ No virtualization (permanent DOM)
✓ No re-render on focus/blur (browser native)
✓ CSS :hover instead of JS onMouseEnter/Leave
✓ No polling intervals
✓ requestAnimationFrame for scroll-linked animations only
```

---

## 9. Scalability Strategy

### 9.1 Multi-Page Forms

```
Strategy: Render all pages, toggle visibility

Why not render-on-demand:
  - WPF renders all pages (in a ScrollViewer)
  - Field values must persist across page navigation
  - Focus must not be lost
  - Tab order spans pages
  - Browser autofill works across pages

Implementation:
  - All pages are in the DOM from initial render
  - Only current page is visible (visibility: visible)
  - Hidden pages: visibility: hidden, still in layout flow
  - PageContainer uses flex column with gap between pages
  - CameraTransform scales everything uniformly

Memory for 500 pages:
  - If each page has 1MB background image → 500MB (problematic)
  - Solution: only load background image for visible page
  - Field wrappers are lightweight (just div + input)
  - ~200 bytes per field × 10 fields per page × 500 = 1MB (fine)
  - Images are the memory concern, not fields
```

### 9.2 Background Image Strategy

```
Option A: Load all images on init
  - 500 pages × 1MB = 500MB (unacceptable)
  - Rejected

Option B: Load on demand, cache in memory
  - Load current page background immediately
  - Preload next/previous page backgrounds
  - Cache in memory (Map<pageId, ImageData>)
  - Discard images more than 3 pages away
  - Acceptable: max 5 images in memory = 5MB

Option C: Lazy load with placeholder
  - Show loading placeholder for image
  - Load when page becomes active
  - Cache for session
  - Most practical

Decision: Option B + C combined.
  - Current page: loaded (high priority)
  - Adjacent pages: preloaded (low priority)
  - Far pages: not loaded until navigated to
  - Cache limit: 10 images max, LRU eviction
  - Page renders fields even without background (form is usable without background)
```

### 9.3 Handle Large Forms (Architectural)

```
10 sheets, 50 pages, 5000 fields:

  Sheet 1: Page 1-5    (500 fields)
  Sheet 2: Page 6-15   (1000 fields)
  Sheet 3: Page 16-30  (1500 fields)
  Sheet 4: Page 31-50  (2000 fields)

The current page model already handles this:
  - All pages rendered
  - Only current page visible
  - Background images loaded on demand
  - Field selectors prevent mass re-render
  - Tab order spans all pages naturally

No special handling needed for "large forms."
The architecture IS the scaling strategy.
```

### 9.4 Future: Distributed Rendering

```
If forms become truly massive (100+ pages, 50,000+ fields):

Options to consider:
1. Lazy page mount: mount pages only when they become visible,
   keep mounted after first render (solves initial load time)

2. Web Worker for validation: Heavy validation logic offloaded
   (not needed for typical form validation)

3. Virtual scroll for page list: If pages > 100, virtualize page
   containers (not fields within pages)

But: these are FUTURE considerations.
The current architecture handles all realistic paperless form sizes.
```

---

## 10. Future Extensibility

### 10.1 Collaboration

```
How to add without refactoring:

- Add collaborationSlice to Zustand:
  - presence: Record<fieldId, UserId>
  - fieldLocks: Record<fieldId, LockInfo>
  - remoteChanges: Event[] (OT/CRDT operations)

- RuntimeField renders a collaborator indicator if presence[fieldId]
  (small avatar/color dot in corner of field)

- Value changes:
  - onChange: send operation to server (not just value)
  - on remote change: apply operation to local value via OT/CRDT

- The key: Zustand's slice architecture allows adding collaboration
  without touching existing slices. Each field connects to presence
  via a new selector.
```

### 10.2 Comments/Annotations

```
- New layer: AnnotationLayer (z-index: 3)
  - Separate from FieldLayer
  - Renders comment bubbles connected to fields/regions
  - Positioned using the same coordinate system
  - No interference with field events
  - Zustand: commentsSlice

- No changes to existing field components.
- No changes to event handling.
```

### 10.3 Signatures

```
- SignatureField already exists
- Future: integrate with digital signature provider (DocuSign, Adobe Sign)
- Field component connects to signature service
- Value stored as signature data URL or reference
- Same value/dirty/save pipeline as text fields
```

### 10.4 OCR / AI

```
- New service: OcrService / AiService
- Called from ValidationService or SaveService
- Adds AI suggestions as warnings (not errors)
- Uses existing validationSlice for display
- No new components needed
```

### 10.5 Version History

```
- New slice: versionSlice
  - versionId: string
  - versions: VersionMetadata[]
  - loadVersion(versionId) => sets values to that version's snapshot
  - compareVersions(v1, v2) => diff view

- RuntimeViewer accepts optional versionId prop
  - If set, RuntimeViewer shows readonly view of that version
  - Uses same components, just readOnly: true

- No changes to field components.
- No changes to rendering pipeline.
```

### 10.6 Conditional Logic / Formulas

```
- CalculatedField component already handles read-only computed values
- Formula engine: pure function: (allFieldValues) => fieldValue
- Register formula per field in field definition
- On any value change:
  - Recalculate all formulas (or dependency-graph-based)
  - Only affected fields re-render
  - Uses existing setValue pipeline

- Dependency graph: Map<fieldId, Set<fieldId>>
  - Field A depends on B, C → if B or C changes, recalculate A
  - Skip fields that haven't changed
```

### 10.7 Workflow / Permissions / Audit Trail

```
- Permissions: field-level readOnly/enabled/visible flags
  - Already supported by field definition (readOnly, visible, enabled)
  - RuntimeField checks these flags before rendering
  - No architecture changes

- Workflow: currentStep determines which fields are enabled
  - Managed by valueSlice + validationSlice
  - No new components

- Audit trail: log all value changes with timestamp + userId
  - Middleware on Zustand setValue
  - No components affected
```

### 10.8 Extensibility Principle

```
Every future feature should be:
  1. A new Zustand slice (no existing slice modified)
  2. A new component layer (no existing layer modified)
  3. A new service (no existing service modified)
  4. A new field type (no existing field type modified)

This is achieved by:
  - Slice-based Zustand architecture
  - Layer-based rendering hierarchy
  - Strategy pattern for field components
  - Event-driven service communication
```

---

## 11. Migration Roadmap

### 11.1 Current vs Runtime V2 Comparison

| Aspect | Current Runtime | Runtime V2 | Why Change |
|--------|----------------|------------|------------|
| **State management** | `useState` + prop drilling | Zustand store with selectors | Eliminate prop drilling, enable per-field subscriptions, add devtools |
| **Field values storage** | `useRuntimeState` hook (local) | Zustand valueSlice (global) | Must persist across page navigation |
| **Page rendering** | Current page only (conditional) | All pages, visibility toggle | Prevent field unmount, preserve focus/IME |
| **Field lifecycle** | Mount on page show, unmount on hide | Permanent (never unmount) | Prevents focus loss |
| **Selection highlight** | Yellow overlay with React state | CSS `:focus-within` | No re-renders on focus change |
| **Designer/Runtime** | Entangled (same component tree) | Completely separate | Designer code never executes during form fill |
| **Event handling** | stopPropagation on wrapper | Browser-native (no stopPropagation) | Fixes tab order, clipboard, autofill |
| **Background images** | Loaded with page | Lazy loaded, cached, LRU eviction | Reduces memory for multi-page forms |
| **Field coordinates** | pt/px ambiguity | Single px system | Simplify, eliminate conversion |
| **Validation** | Inline in useFieldState | Dedicated validationSlice | Decouple from value management |
| **Save** | None (manual export) | Auto-save with debounce, retry, offline | Production requirement |
| **Undo/redo** | None | Zustand historySlice | Production requirement |
| **Memo boundaries** | None | Strategic React.memo + selectors | Eliminate unnecessary re-renders |
| **Provider hierarchy** | None | Zustand (no providers needed) | Eliminate context re-render issues |
| **Field type dispatch** | switch-case in render | Lookup table constant | Predictable, memoizable |
| **Bundle size** | Includes designer code | Runtime-only bundle (tree-shakeable) | Smaller initial load |

### 11.2 Migration Phases

```
Phase 0: Setup (1-2 days)
  - Install Zustand
  - Define all store slices (types only)
  - Create RuntimeStore with all slices
  - No UI changes
  - Test: store works in isolation

Phase 1: State Separation (2-3 days)
  - Create Zustand store in parallel with existing state
  - Field components read from BOTH old state (prop) and new store
  - Remove useRuntimeState, useFieldState hooks
  - All rendering still uses old component tree
  - Test: form loads, fields display, values work

Phase 2: RuntimePage Extraction (2-3 days)
  - Extract RuntimePage component from RuntimeFormViewer
  - Render all pages (hidden pages: visibility: hidden)
  - Keep RuntimeCanvas/OverlayModel pipeline
  - Test: page navigation works, values persist across pages

Phase 3: RuntimeField Rewrite (3-4 days)
  - Replace RuntimeField with new implementation
  - Remove OverlayModel dependency (use field data directly)
  - Add Zustand selectors (per-field subscription)
  - Add React.memo at correct boundaries
  - Remove stopPropagation
  - Remove selection logic
  - Remove yellow styles
  - Test: form filling works, no focus loss, no unnecessary renders

Phase 4: Background Image Optimization (1-2 days)
  - Replace BackgroundLayer with lazy loading + caching
  - Add preload for adjacent pages
  - Add LRU eviction cache
  - Test: fast page navigation, reasonable memory usage

Phase 5: Save & Validation (2-3 days)
  - Implement auto-save with debounce
  - Implement validation slice
  - Implement dirty tracking
  - Test: values persist, validation works, network errors handled

Phase 6: Undo/Redo (1-2 days)
  - Implement history slice
  - Connect Ctrl+Z / Ctrl+Shift+Z
  - Test: undo/redo works, history is bounded

Phase 7: Designer Separation (3-4 days)
  - Extract PaperlessDesigner designer-specific code
  - Create separate Designer component (if needed)
  - RuntimeViewer becomes standalone form filler
  - Designer uses RuntimeViewer in readOnly mode for preview
  - Test: designer works, runtime works, no cross-imports

Phase 8: Polish & Edge Cases (2-3 days)
  - Tab order verification
  - IME composition testing
  - Browser autofill testing
  - Mobile touch testing
  - Accessibility audit
  - Performance profiling
  - Memory profiling
```

### 11.3 Phase Validation Rules

```
Each phase must satisfy:

1. Existing functionality is preserved (no regressions)
2. The phase is independently testable (can be deployed alone)
3. No breaking changes to API or data flow
4. Backward compatible with existing saved data
5. Each phase either maintains or improves performance

Phase verification:
  - Before phase: run current tests, note baseline
  - After phase: run same tests, compare results
  - Performance: measure render count, memory, frame rate
```

### 11.4 Risk Mitigation

| Risk | Mitigation |
|------|------------|
| Zustand migration breaks existing state | Phase 0: test store in isolation first. Phase 1: run new store alongside old. Gradual cutover |
| All-pages rendering causes memory issues | Monitor memory in Phase 2. Add image lazy loading in Phase 4 before it becomes a problem |
| Removing stopPropagation breaks something | Test tab order, clipboard, autofill, drag-and-drop thoroughly in Phase 3 |
| Designer separation breaks existing workflows | Phase 7 is LAST phase. All functionality is preserved until then |
| Performance regression | Measure at each phase. Zustand selectors + React.memo should IMPROVE performance, not degrade |
| Bundle size increase | Zustand is 1KB. Overall bundle should decrease (remove designer code from runtime) |

---

## 12. Risks and Trade-offs

### 12.1 Trade-off: All Pages in DOM (Memory vs UX)

| Approach | Pros | Cons |
|----------|------|------|
| **All pages rendered (chosen)** | No field unmount, no focus loss, fast navigation, preserved IME | Higher initial memory (mitigated: only images are heavy, fields are lightweight) |
| Render on demand | Lower initial memory | Focus loss on navigation, IME state lost, slower navigation |

**Decision: All pages rendered.**
Fields are lightweight (div + input). Only background images are memory-heavy, handled by lazy loading.

### 12.2 Trade-off: No Virtualization (Performance vs Complexity)

| Approach | Pros | Cons |
|----------|------|------|
| **No virtualization (chosen)** | Simple architecture, no focus issues, no tab order issues, works with browser autofill/find-in-page | Higher DOM node count (handled by browser efficiently) |
| Virtual scroll | Lower DOM node count (for 50k+ fields) | Focus loss on scroll, broken tab order, broken find-in-page, broken autofill, broken IME |

**Decision: No virtualization.**
Browser DOM can handle 50k+ nodes. Virtualization breaks core form behavior.

### 12.3 Trade-off: Zustand over Context (Performance vs Familiarity)

| Approach | Pros | Cons |
|----------|------|------|
| **Zustand (chosen)** | Per-field selectors (no mass re-render), no providers, 1KB, built-in devtools, persist/undo middleware | New dependency (small), different pattern from Context |
| React Context | Built-in, familiar | All consumers re-render on any change, provider nesting, no selector mechanism |
| Redux | Well-known, ecosystem | 12KB, too much boilerplate, dispatching actions for form fill is over-engineered |
| Jotai/Recoil | Atomic, efficient | Overkill for form state, less mature, more API surface |

**Decision: Zustand.**
Best balance of size, performance, and ergonomics for form state management.

### 12.4 Trade-off: CSS Focus over React Focus State

| Approach | Pros | Cons |
|----------|------|------|
| **CSS :focus-within (chosen)** | Zero re-renders, native behavior, simple | Limited styling (can't conditionally render elements based on focus) |
| React focus state (onFocus/onBlur) | Full control over render based on focus | Re-render on every focus/blur, state management overhead |

**Decision: CSS :focus-within.**
For focus highlighting, CSS is sufficient. If conditional rendering based on focus is needed (e.g., show toolbar on focus), use a local `useState` on the field component (not global state, not propagated up).

### 12.5 Trade-off: Single Store vs Multiple Stores

| Approach | Pros | Cons |
|----------|------|------|
| **Single store, multiple slices (chosen)** | Atomic updates across slices, single subscription, simpler devtools, undo across slices | Store file can grow large (mitigated: slices in separate files) |
| Multiple stores (one per concern) | Smaller individual stores | Cross-store coordination, harder undo/redo, multiple subscriptions |
| Atoms (Jotai/Recoil) | Fine-grained reactivity | Overkill, coordination complexity |

**Decision: Single Zustand store with slices.**
Cross-slice coordination (value + dirty + save + history) is simpler in a single store.

### 12.6 Risk: Zustand Selector Performance with 10,000 Fields

```
Risk: Each field subscribes to store.values[fieldId].
  Zustand uses Object.is comparison. Subscribe to s.values[fieldId].
  When setValue(fieldId, newValue) is called:
    - Only subscribers to that specific fieldId re-render
    - Zustand creates a new values object on each setValue
    - But each field subscribes to s.values[fieldId], not s.values
    - So Object.is(freshResult, previousResult) compares the single field's value
    - Only that field re-renders

  Confirmed: this is O(1) per field, not O(N).
  Adding 10000 fields does NOT slow down per-keystroke performance.
  No special handling needed.
```

### 12.7 Risk: Initial Render Time for Large Forms

```
Risk: 500 pages × 10 fields = 5000 fields rendered on initial load.

Mitigation strategies:
  1. First paint: render only active page, defer hidden pages by 50ms each
     (React 18: useDeferredValue or startTransition)
  2. If needed: use content-visibility: auto on hidden pages
     (browser automatically skips layout/paint for hidden content)

  Decision: Start with all pages in initial render.
  Monitor performance. Add deferred rendering if needed.
  This is a performance optimization, NOT an architecture change.
```

### 12.8 Risk: Browser Tab Order with Hidden Pages

```
Risk: Tab navigation goes through hidden pages.

Analysis:
  - Hidden pages have visibility: hidden
  - Elements with visibility: hidden are NOT focusable
  - Tab order skips hidden pages entirely
  - Browser tab order only includes visible, enabled, focusable elements

  Confirmed: Tab order is correct by default.
  No special handling needed.
```

---

## Appendix A: File Structure (Runtime V2)

```
paperless/
├── runtime/                          # Runtime V2 (form filler)
│   ├── index.ts                      # Public API
│   ├── RuntimeProvider.tsx           # Page-level wrapper (minimal)
│   ├── RuntimeViewer.tsx             # Main viewer (orchestrates pages)
│   ├── components/
│   │   ├── ViewportContainer.tsx     # Scrollable container
│   │   ├── CameraTransform.tsx       # CSS transform wrapper
│   │   ├── PageContainer.tsx         # Flex column for pages
│   │   ├── RuntimePage.tsx           # Per-page container
│   │   ├── BackgroundImage.tsx       # Lazy-loaded background PNG
│   │   ├── FieldLayer.tsx            # Field positioning container
│   │   ├── RuntimeField.tsx          # Field wrapper + dispatcher
│   │   ├── fields/
│   │   │   ├── TextField.tsx
│   │   │   ├── NumberField.tsx
│   │   │   ├── CheckboxField.tsx
│   │   │   ├── DateField.tsx
│   │   │   ├── SelectField.tsx
│   │   │   ├── SignatureField.tsx
│   │   │   ├── ReadonlyField.tsx
│   │   │   ├── CalculatedField.tsx
│   │   │   └── index.ts             # FIELD_COMPONENT_MAP
│   │   ├── RuntimeToolbar.tsx       # Page nav, zoom controls
│   │   └── RuntimeStatusBar.tsx      # Save status, dirty indicator
│   ├── store/
│   │   ├── index.ts                  # createRuntimeStore()
│   │   ├── formSlice.ts              # Form metadata, page navigation
│   │   ├── valueSlice.ts             # Field values
│   │   ├── dirtySlice.ts             # Dirty tracking
│   │   ├── validationSlice.ts        # Validation errors/warnings
│   │   ├── saveSlice.ts              # Save queue, status, retry
│   │   └── historySlice.ts           # Undo/redo
│   ├── services/
│   │   ├── SaveService.ts            # Auto-save with debounce + retry
│   │   ├── ValidationService.ts      # Field validation
│   │   ├── FocusService.ts           # Tab order, IME (minimal)
│   │   └── ImageCacheService.ts      # Background image lazy loading + cache
│   └── types/
│       └── index.ts                  # Runtime-specific types
│
├── designer/                          # Designer (form authoring)
│   ├── index.ts                      # Public API
│   ├── PaperlessDesigner.tsx         # Main designer component
│   ├── components/                   # Designer-specific components
│   │   ├── Toolbar/
│   │   ├── FieldExplorer/
│   │   ├── ConfigPanel/
│   │   └── ZoomableCanvas/
│   └── services/
│       └── DesignerController.ts
│
├── shared/                            # Shared between runtime and designer
│   ├── types/
│   │   ├── runtime.ts                # RuntimeField, RuntimeSheet, etc.
│   │   ├── field.ts                  # Field types, FieldConfig
│   │   └── overlay.ts               # Overlay types (used by designer)
│   ├── components/
│   │   └── BackgroundImage.tsx       # Shared background image component
│   └── utils/
│       ├── coordinates.ts            # Coordinate utilities
│       └── validation.ts             # Validation rules
│
└── app/
    ├── page.tsx                       # Runtime entry (form filling)
    └── designer/
        └── page.tsx                   # Designer entry (form authoring)
```

---

## Appendix B: Key Zustand Patterns

```typescript
// === Store Creation ===
// store/index.ts

import { create } from 'zustand';
import { createFormSlice, FormSlice } from './formSlice';
import { createValueSlice, ValueSlice } from './valueSlice';
import { createDirtySlice, DirtySlice } from './dirtySlice';
import { createValidationSlice, ValidationSlice } from './validationSlice';
import { createSaveSlice, SaveSlice } from './saveSlice';
import { createHistorySlice, HistorySlice } from './historySlice';

export type RuntimeStore = FormSlice & ValueSlice & DirtySlice
  & ValidationSlice & SaveSlice & HistorySlice;

export const useRuntimeStore = create<RuntimeStore>()((...a) => ({
  ...createFormSlice(...a),
  ...createValueSlice(...a),
  ...createDirtySlice(...a),
  ...createValidationSlice(...a),
  ...createSaveSlice(...a),
  ...createHistorySlice(...a),
}));

// === Field Component Pattern ===
// RuntimeField.tsx

function RuntimeField({ field }: { field: RuntimeField }) {
  const fieldId = field.id;

  // Per-field selectors - only subscribe to own data
  const value = useRuntimeStore(s => s.values[fieldId]);
  const error = useRuntimeStore(s => s.errors[fieldId]?.[0]);
  const setValue = useRuntimeStore(s => s.setValue);
  const markDirty = useRuntimeStore(s => s.markDirty);
  const validateField = useRuntimeStore(s => s.validateField);
  const enqueueSave = useRuntimeStore(s => s.enqueueSave);

  const handleChange = useCallback((newValue: FieldValue) => {
    setValue(fieldId, newValue);
    markDirty(fieldId);
    enqueueSave(fieldId, newValue);
  }, [fieldId]);

  const handleBlur = useCallback(() => {
    validateField(fieldId);
  }, [fieldId]);

  const FieldComponent = FIELD_COMPONENT_MAP[field.dataType];
  if (!FieldComponent) return null;

  return (
    <div
      className="runtime-field"
      style={{
        position: 'absolute',
        left: field.leftPx,
        top: field.topPx,
        width: field.widthPx,
        height: field.heightPx,
      }}
      data-field-id={fieldId}
    >
      <FieldComponent
        field={field}
        value={value}
        onChange={handleChange}
        onBlur={handleBlur}
        error={error}
      />
    </div>
  );
}

// Memo is safe because props change ONLY when field definition or own value changes
export default React.memo(RuntimeField);

// === Page Navigation ===
// formSlice.ts

createFormSlice: (set, get) => ({
  currentPageId: null,
  pages: [],

  setCurrentPage: (pageId: string) => {
    set({ currentPageId: pageId });
    // NO other state changes
    // NO value clearing
    // NO validation clearing
    // Just update the page ID
  },
});

// === Background Image Loading ===
// BackgroundImage.tsx

function BackgroundImage({ pageId }: { pageId: string }) {
  const src = useRuntimeStore(s => s.bgUrls[pageId]);

  return (
    <img
      src={src}
      alt=""
      role="presentation"
      style={{
        position: 'absolute',
        inset: 0,
        width: '100%',
        height: '100%',
        objectFit: 'none',
        pointerEvents: 'none',
        zIndex: 0,
      }}
    />
  );
}

export default React.memo(BackgroundImage);
```

---

## Appendix C: Event Validation Checklist

```
Feature: Tab navigation
  Without stopPropagation:
    Tab key: browser moves focus to next element in DOM order
    Shift+Tab: browser moves focus to previous element
    Hidden pages: elements are visibility: hidden, not focusable
    Tab skips hidden pages naturally
    ✓ Works

Feature: Browser autofill
  Without stopPropagation:
    Browser fills saved credentials in matching input fields
    Fires input event → React onChange → store update
    ✓ Works

Feature: Clipboard (Ctrl+C / Ctrl+V)
  Without stopPropagation:
    Browser handles clipboard events natively
    ✓ Works

Feature: IME (Chinese/Japanese/Korean input)
  Without stopPropagation:
    compositionstart → compositionupdate → compositionend
    Final input event fires onChange once
    ✓ Works

Feature: Find in page (Ctrl+F)
  Without stopPropagation:
    Browser searches all text in visible inputs
    ✓ Works

Feature: Spell check
  Without stopPropagation:
    Browser spell check works on all inputs
    ✓ Works

Feature: Right-click context menu
  Without stopPropagation:
    Browser context menu shows Cut/Copy/Paste/Spell check
    ✓ Works

Feature: Touch to focus (mobile)
  Without stopPropagation:
    Tap on input → focus → native keyboard opens
    ✓ Works

Feature: radio button group
  Native input type="radio" with shared name attribute
  ✓ Works
```

---

*End of Architecture Specification*

*Next step: Review with team. Validate against requirements. Begin Phase 0 implementation.*
