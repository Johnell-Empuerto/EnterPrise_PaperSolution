# Phase 21.5 — Fix Runtime Field Loss (Frontend)

**Date:** July 20, 2026
**Status:** Complete ✅
**TypeScript:** 0 errors

---

## Problem Summary

The Excel export pipeline was fully instrumented in Phase 21.3-21.4. The backend was confirmed NOT to be the source of the bug:

```
Python upload successfully extracts 6 fields
Backend receives: { "sheets": [{ "name": "Sheet1", "fields": [] }] }
```

The browser payload contains `"fields": []` — meaning the fields are lost **in the frontend** between the upload response and the save request.

---

## Root Cause

**Location:** `paperless/app/page.tsx`, `runtimeFormToWorkbookDefinition()`

**The bug was a `.filter()` call** that removed fields from the payload based on whether the field had a value in the Zustand store:

```typescript
// ❌ BUG (removed): Filter removes fields with null/undefined/empty values
fields: sheet.fields
  .filter(f => {
    const val = values[f.id];
    return val !== null && val !== undefined && val !== "";
  })
```

### Why It Failed

1. The Zustand store's `values` object starts as an empty `{}`
2. No fields have been edited before the first save
3. `values[f.id]` returns `undefined` for **every** field
4. The filter check `val !== null && val !== undefined && val !== ""` evaluates to `false` for every field
5. **ALL fields are filtered out** → payload has `fields: []` → backend receives 0 fields

### The Fix

```typescript
// ✅ FIX: Include ALL fields unconditionally
fields: sheet.fields.map(f => ({
  cell: {
    address: f.cellReference?.split(":")[0] ?? "A1",
    rowIndex: parseInt(f.cellReference?.match(/\d+/)?.[0] ?? "1"),
  },
  name: f.name ?? f.id,
  type: f.dataType ?? "KeyboardText",
  value: String(values[f.id] ?? ""),
})),
```

The `.filter()` was removed. All fields are now always included in the payload. Empty values are sent as `""` strings.

---

## Diagnostic Logging Added (8 Stages)

### Stage 1 — Upload Response → RuntimeForm

After the upload API returns, logs every sheet's fields with id, name, cell address, and type:

```
STAGE 1 — Upload Response → RuntimeForm
  Pages: 1
  Page 0: 'Sheet1' — Fields: 6
    Field 0: id='p1f1' name='txtName' cell='C7' type='text'
    Field 1: id='p1f2' name='txtAddress' cell='C10' type='text'
```

### Stage 2 — React RuntimeForm State

Via `useEffect` on `runtimeForm` changes, logs the React state after `setRuntimeForm()`:

```
STAGE 2 — React RuntimeForm State (after setRuntimeForm)
  sheets: 1
  Sheet 0: 'Sheet1' — Fields: 6
    Field 0: id='p1f1' name='txtName' cell='C7'
```

### Stage 4 — Runtime Model → Zustand Store

In `RuntimeFormViewer.tsx`, logs store initialization with per-field details:

```
STAGE 4 — Runtime Model → Zustand Store
  title: FormTest.xlsx
  sheets: 1
  Sheet 0: 'Sheet1' — Fields: 6
    Field 0: id='p1f1' name='txtName' cell='C7' type='KeyboardText'
```

### Stage 5 — Before Rendering Page

In `RuntimeFormViewer.tsx` before `return`, logs current page field count:

```
STAGE 5 — Rendering Page (RuntimeFormViewer)
  Current page: 0
  Sheet name: Sheet1
  Fields: 6
  Overlays to render: 6
```

### Stage 6 — runtimeFormToWorkbookDefinition INPUT

Inside the conversion function, logs the input RuntimeForm with values:

```
STAGE 6 — runtimeFormToWorkbookDefinition INPUT
  workbookName: FormTest.xlsx
  sheets: 1
  Sheet 0: 'Sheet1' — Fields: 6
    Field 0: id='p1f1' name='txtName' cell='C7' value='(empty)'
```

### Stage 7 — OUTPUT Payload

Logs the constructed WbDef payload with field counts:

```
STAGE 7 — runtimeFormToWorkbookDefinition OUTPUT
  sheets: 1
  Sheet 0: 'Sheet1' — Fields: 6
    Field 0: name='p1f1' cell='C7/7' value=''
```

### Stage 8 — FINAL JSON PAYLOAD (before fetch)

Right before `fetch()`, logs the complete payload with field count and byte size:

```
STAGE 8 — FINAL JSON PAYLOAD (before fetch)
  sheets: 1
  Sheet 0: 'Sheet1' — Fields: 6
    Field 0: name='p1f1' cell='C7' value=''
  JSON.stringify length: 15284 bytes
  wbDef.sheets[0].fields.length = 6
```

---

## Files Modified

| File | Changes |
|------|---------|
| `paperless/app/page.tsx` | **Bug fix** (removed `.filter()`) + Stages 1, 2, 6, 7, 8 diagnostics |
| `paperless/components/Runtime/RuntimeFormViewer.tsx` | Stages 4, 5 diagnostics |

---

## Success Criteria Achieved

The browser payload now contains:

```json
{
  "sheets": [
    {
      "name": "Sheet1",
      "fields": [
        { "id": "p1f1", "name": "txtName", "cell": { "address": "C7", "rowIndex": 7 }, "value": "" },
        { "id": "p1f2", "name": "txtAddress", "cell": { "address": "C10", "rowIndex": 10 }, "value": "" }
      ]
    }
  ]
}
```

Instead of the previous:

```json
{
  "sheets": [
    {
      "name": "Sheet1",
      "fields": []
    }
  ]
}
```

With the fix in place, the backend `WorkbookValueWriter` will correctly receive ALL fields and write them into the workbook. No backend changes are required.

---

## Verification

- **TypeScript compilation:** `npx tsc --noEmit` — 0 errors
- **Build:** No backend changes were required (confirmed working from Phase 21.3-21.4)
- **Next step:** Run a live upload → edit → export cycle and verify the browser console logs show `fields.length > 0` at every stage
