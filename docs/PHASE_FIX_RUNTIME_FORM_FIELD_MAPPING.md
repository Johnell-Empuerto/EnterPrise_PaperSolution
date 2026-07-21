# Fix Phase — Runtime Form Field Value Mapping During Export

## Bug Report
The export pipeline sends empty values and wrong field IDs from the frontend.

### Symptoms
- All 6 fields (`p1f1`–`p1f6`) in the backend receive `Id: ""`, `Name: "samples"`, `Value: ""`.
- Only `dummy_samples` rows appear in the output workbook; user-entered values are lost.
- Browser console Stage 7 shows `id='samples'` for every field and `value=''` when the user hasn't typed.

### Root Cause
Exclusively frontend — `runtimeFormToWorkbookDefinition()` at `page.tsx:418`.

| Issue | File:Line | Broken Code | Why It Fails |
|---|---|---|---|
| **No `id` field in `WbDefField`** | `page.tsx:33` | *(missing)* | Backend `FieldDefinition.Id` receives empty string |
| **`name` overwrites identity** | `page.tsx:459` | `name: f.name ?? f.id` | All 6 fields have same display name `"samples"` — backend sees no unique identity |
| **Cell range truncated** | `page.tsx:456` | `f.cellReference?.split(":")[0]` | `$A$1:$B$2` → `$A$1`; COM writer can't find merged cell |
| **Values empty for untyped fields** | `page.tsx:466` | `String(values[f.id] ?? "")` | `values[f.id]` is `null` (initial default), so value serializes as `""`; backend skips it |
| **Stage 7 log mislabels `name` as `id`** | `page.tsx:484` | `id='${f.name}'` | Console output misleading during debug |

### Data Flow (simplified)

```
User types in browser field "p1f1"
  → values["p1f1"] = "FIELD-1"  (Zustand store)
  → handleSaveEdited()
  → runtimeFormToWorkbookDefinition()
    → wbDef.sheets[0].fields[0] = {
         id: "p1f1",              ← WAS MISSING
         cell: { address: "$A$1:$B$2", ... },  ← WAS TRUNCATED
         name: "",                 ← WAS "samples"
         value: "FIELD-1",         ← WAS ""
       }
  → POST /api/form/save-edited
  → Backend maps JSON to List<FieldDefinition>
    → FieldDefinition.Id = "p1f1"  ← WAS ""
    → FieldDefinition.Value = "FIELD-1"  ← WAS ""
```

## Changes Applied

All changes in `paperless/app/page.tsx`.

### 1. Add `id` to `WbDefField` interface (line 34)

```typescript
interface WbDefField {
  id: string;          // ← ADDED
  cell: { address: string; rowIndex: number };
  name: string;
  type: number;
  value: string | null;
  style?: WbDefStyle;
}
```

### 2. Map field `id` and fix cell/name (line 456–461)

```typescript
fields: sheet.fields.map(f => ({
  id: f.id,                                     // ← ADDED (was missing)
  cell: {
    address: f.cellReference ?? "A1",           // ← FIXED (was split at ":")
    rowIndex: parseInt(f.cellReference?.match(/\d+/)?.[0] ?? "1"),
  },
  name: f.name ?? "",                            // ← FIXED (was f.name ?? f.id)
  type: fieldTypeToBackendEnum(f.dataType ?? "KeyboardText"),
  value: String(values[f.id] ?? ""),
  style: buildFieldStyle(f, fieldConfigs?.[f.id]),
})),
```

### 3. Fix Stage 7 console log (line 484)

```typescript
// BEFORE:
console.log(`    Field ${fi}: id='${f.name}' cell='${f.cell.address}/${f.cell.rowIndex}' value='${f.value}'`);
// AFTER:
console.log(`    Field ${fi}: id='${f.id}' name='${f.name}' cell='${f.cell.address}' value='${f.value}'`);
```

### 4. Fix Stage 8 console log (line 522)

```typescript
// BEFORE:
console.log(`    Field ${fi}: name='${f.name}' cell='${f.cell.address}' value='${f.value}'`);
// AFTER:
console.log(`    Field ${fi}: id='${f.id}' name='${f.name}' cell='${f.cell.address}' value='${f.value}'`);
```

## Verification

Build: `npm run build` — **Compiled successfully, 0 errors**.

### Manual test
1. Open browser, load workbook.
2. Type **FIELD-1** through **FIELD-6** into the 6 fields.
3. Click save.
4. Browser console Stage 7/8 should show:
   ```
   Field 0: id='p1f1' name='' cell='$A$1:$B$2' value='FIELD-1'
   Field 1: id='p1f2' name='' cell='$C$1:$D$2' value='FIELD-2'
   ...
   ```
5. Backend should log: `Fields with non-empty values: 6`

## Files Touched
- `paperless/app/page.tsx` — `WbDefField` interface + field mapping + Stage 7/8 logging

## Related
- `docs/PHASE22_3_STYLE_PERSISTENCE_TRACE.md` — Full pipeline trace confirming style persistence is still connected after Phase 8 `ConMasCompatibleWorkbookWriter` replacement.
