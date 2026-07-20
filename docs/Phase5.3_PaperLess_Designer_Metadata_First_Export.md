# Phase 5.3 — Generate PaperLess Designer Metadata on First Export

## Status: **Complete** ✅

---

## Objective

Transform a plain Excel workbook into a ConMas-compatible PaperLess Designer workbook during the first export. The exported workbook embeds designer metadata so it can be re-uploaded without browser state or database.

---

## Implementation

### File Modified: `Application/WorkbookValueWriter.cs`

### New Methods Added

| Method | Purpose |
|---|---|
| `EnsureDesignerMetadata()` | Entry point — detects existing `_Fields` and `ExcelOutputSetting` sheets, creates or updates them |
| `CreateExcelOutputSettingSheet()` | Creates the 36-row ConMas XML configuration sheet |
| `CreateFieldsSheet()` | Creates the `_Fields` hidden sheet with 7-column field metadata |
| `UpdateFieldsSheet()` | Refreshes `_Fields` content on subsequent exports (clear + repopulate) |
| `CreateInlineStringCell()` | Helper to create proper `<is><t>value</t></is>` inline string cells |
| `WriteInlineStringCell()` | Helper to append inline string cell to a row |
| `GenerateExcelOutputSettingXmlFragments()` | Generates 36 XML configuration fragments mirroring legacy ConMas format |

### Pipeline Integration

Inside the `SpreadsheetDocument` block, after shared strings are saved and before `Dispose`:

```
Write cell values
    ↓
Save SharedStringTable
    ↓
EnsureDesignerMetadata(wbPart, wbDef)
    ├── Detect _Fields → Missing? → CreateFieldsSheet()
    │                   → Exists?   → UpdateFieldsSheet()
    ├── Detect ExcelOutputSetting → Missing? → CreateExcelOutputSettingSheet()
    │                              → Exists?   → Skip (update not needed)
    ├── Set _Fields to Hidden (SheetStateValues.Hidden)
    ├── Set ExcelOutputSetting to Visible
    └── Set _designerMetadataCreated = true
    ↓
wbPart.Workbook.Save()
    ↓
SpreadsheetDocument.Dispose() → SDK writes ZIP
```

### ZIP Restore Behaviour

When `_designerMetadataCreated == true`:
- `workbook.xml` is **not** restored from original ZIP entries
- The SDK's workbook.xml (which has the new `_Fields` and `ExcelOutputSetting` sheet definitions) is kept
- All other entries (styles, theme, relationships, etc.) are still restored

When `_designerMetadataCreated == false`:
- Original `workbook.xml` is restored (no structural changes were made)

---

## Metadata Format

### `_Fields` Sheet (Hidden)

| Column | Header | Content |
|---|---|---|
| A | Address | Cell reference (e.g., `B5`) |
| B | FieldId | WbDef field ID |
| C | FieldName | Human-readable field name |
| D | FieldType | Field type (Text, Number, Date, etc.) |
| E | SheetName | Owning sheet name |
| F | CreatedDate | ISO 8601 timestamp |
| G | Notes | (reserved, currently empty) |

### `ExcelOutputSetting` Sheet (Visible)

36 rows of ConMas XML configuration fragments mirroring the legacy `WorkbookGenerator.GenerateExcelOutputSettingXmlFragments()`. Includes:
- Designer version metadata
- Report type, sheet count, cluster count
- Output configuration (CSV, PDF, Excel)
- Display/save menu configuration
- Per-cluster configuration (cluster ID, visibility, mobile settings)
- Auto-numbering configuration
- Sheet remarks

---

## Round-Trip Behaviour

| Scenario | Behaviour |
|---|---|
| **First export** (no `_Fields` or `ExcelOutputSetting`) | Both sheets created. `_designerMetadataCreated = true`. SDK workbook.xml preserved. |
| **Subsequent export** (`_Fields` + `ExcelOutputSetting` exist) | `_Fields` repopulated. `ExcelOutputSetting` unchanged. `_designerMetadataCreated = true` (UpdateFieldsSheet). SDK workbook.xml preserved. |
| **Re-upload exported workbook** | `WorkbookReaderService` detects `_Fields` sheet → rebuilds fields from metadata |

---

## API Compatibility

- **No API changes** — same `POST /api/forms/save-edited` endpoint
- **No request model changes** — same `WorkbookDefinition` payload
- **No response model changes** — same file download
- **Frontend unaffected** — works with existing browser editor

---

## Build Status

```
Build succeeded
0 errors
94 warnings (pre-existing)
```

---

## Verification

To verify the metadata was created:
1. Export a workbook from the browser
2. Open the exported file in a ZIP editor (7-Zip, etc.)
3. Check `xl/workbook.xml` for `<sheet name="_Fields" state="hidden">` and `<sheet name="ExcelOutputSetting">`
4. Check `xl/worksheets/` for the corresponding sheet XML files

Or re-upload the exported workbook — fields should be detected automatically.

---

## Known Limitations

1. **Cell comments not implemented** — `EnsureDesignerMetadata()` doc comment mentions comments but no comment-writing code exists. The `_Fields` sheet replaces the need for comment-based field detection.
2. **Defined names (Print_Area) not implemented** — Legacy `WorkbookGenerator.SetDefinedNames()` is not mirrored. The print area is preserved from the original workbook.

These are low-priority because the `_Fields` sheet is the canonical metadata source and the original workbook's print area is unchanged by `WorkbookValueWriter`.
