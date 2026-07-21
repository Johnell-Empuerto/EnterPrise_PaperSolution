# Phase — Embedded PaperLess JSON Configuration Persistence

## Objective

Add a PaperLess configuration persistence layer inside exported Excel workbooks. PaperLess-specific field identity, style, and configuration are stored as embedded JSON inside the XLSX, surviving export → re-upload round trips without relying on Excel cell formatting.

## Architecture

```
Excel remains source of truth for:
  ├── Visual workbook appearance
  ├── Cell values
  ├── Print Area / Page Layout
  ├── Merged cells / Borders
  ├── Excel rendering / PNG / PDF output
  └── Native Excel structure

PaperLess JSON becomes source of truth for:
  ├── Stable PaperLess field IDs (p1f1, p1f2, ...)
  ├── Field-to-cell mapping
  ├── Field types
  ├── PaperLess font/style configuration
  ├── PaperLess field configuration (required, maxLength, etc.)
  └── Future PaperLess-specific configuration
```

## Problem Solved

Previously, field IDs and PaperLess styles were regenerated from Excel comments on every upload. If a user changed `p1f1.fontSize` from `14 → 24` and re-uploaded, the style reverted to `14` because Excel doesn't natively preserve PaperLess-specific metadata.

Now the exported workbook contains a `PaperLessConfig` worksheet (VeryHidden) with the JSON. On re-upload, the config is read, parsed, and overlaid on top of Excel-detected fields — restoring the exact last-saved configuration.

## Implementation

### Files Created (3)

| File | Lines | Purpose |
|---|---|---|
| `ExcelAPI/Models/WorkbookDefinition/PaperLessConfig.cs` | 111 | Serialization model (`PaperLessConfig` → `PaperLessSheet` → `PaperLessField`). Reuses existing `CellStyle` from `StyleDefinition.cs` — no duplicate style models. |
| `ExcelAPI/Application/IPaperLessConfigWriter.cs` | 16 | Interface defining `WritePaperLessConfig(WorkbookDefinition, string outputPath)` |
| `ExcelAPI/Application/PaperLessConfigWriter.cs` | 319 | ZIP-level writer that creates/updates the `PaperLessConfig` worksheet with `state="veryHidden"` |

### Files Modified (4)

| File | Change |
|---|---|
| `FormSaveService.cs` | Injected `IPaperLessConfigWriter`, calls `WritePaperLessConfig()` after `WorkbookStyleWriter.ApplyStyles()` |
| `DesignerModelReader.cs` | Added `"PaperLessConfig"` to `ConfigSheetNames`. New Section 12 reads the config sheet, parses JSON, validates schema version. After normal field detection, overlays PaperLess field IDs, styles, and config on matching fields (matched by cell address). Added helper methods: `GetCellInlineText`, `NormalizeCellAddress`, `MapHorizontalAlignment`, `MapVerticalAlignment` |
| `CompatibilityValidator.cs` | Added `"PaperLessConfig"` to `ConfigurationSheetNames` |
| `ServiceRegistration.cs` | Registered `IPaperLessConfigWriter` → `PaperLessConfigWriter` as scoped |

### Files NOT Modified (intentional)

| File | Reason |
|---|---|
| `WorkbookReaderService.cs` | Already has `"PaperLessConfig"` in `ConfigurationSheetNames` (line 40) |
| `FormController.cs` | Both `WorkbookReaderService` and `DesignerModelReader` already called; overlay handled in reader |
| `ConMasCompatibleWorkbookWriter.cs` | Config writer is a separate service (follows `WorkbookStyleWriter` pattern) |
| All frontend files (`page.tsx`, `runtime.ts`, store) | Config is a backend-only persistence mechanism — frontend API contract unchanged |

## Data Flow

### Export

```
Frontend runtimeFormToWorkbookDefinition()
  → POST /api/form/save-edited (WorkbookDefinition JSON)
  → FormSaveService.SaveEditedValuesAsync()
      → ConMasCompatibleWorkbookWriter.WriteValues()
      → WorkbookStyleWriter.ApplyStyles()
      → PaperLessConfigWriter.WritePaperLessConfig()  ← NEW
          → Serialize WorkbookDefinition → PaperLessConfig → minified JSON
          → Open output.xlsx as ZipArchive
          → Delete old PaperLessConfig worksheet if exists
          → Create worksheet with JSON in B1 (inline string)
          → Set state="veryHidden" in workbook.xml
          → Update workbook.xml.rels, [Content_Types].xml
      → Return edited workbook
```

### Re-Upload

```
POST /api/form/upload-excel
  → WorkbookReaderService.Read()           (skips PaperLessConfig — already filtered)
  → DesignerModelReader.Read()             ← MODIFIED
      → Enumerate all sheets
      → Detect _Fields, ExcelOutputSetting, PaperLessConfig as configSheets
      → Detect printable sheets, build fields via existing Excel logic
      → Read PaperLessConfig sheet          ← NEW
          → Parse JSON from cell B1
          → Validate schemaVersion == 1
          → For each configSheet + configField:
              → Match by NormalizedCellAddress (e.g., "$A$1:$B$2" → "A1")
              → Overlay field.Id, Style, Type, Config
      → Return DesignerModel with restored configuration
```

## Logging

### Export
```
[PAPERLESS CONFIG] Writing configuration
[PAPERLESS CONFIG] Fields serialized: 6
[PAPERLESS CONFIG] Removed existing configuration sheet
[PAPERLESS CONFIG] Configuration sheet created (xl/worksheets/paperlessconfig.xml)
[PAPERLESS CONFIG] Configuration persisted successfully
```

### Re-Upload
```
[PAPERLESS CONFIG] Configuration sheet found
[PAPERLESS CONFIG] JSON parsed successfully
[PAPERLESS CONFIG] Schema version: 1
[PAPERLESS CONFIG] Sheets in config: 1
[PAPERLESS CONFIG] Fields in config: 6
[PAPERLESS CONFIG] Field 'p1f1' matched by cell $A$1:$B$2
[PAPERLESS CONFIG]   Restored style: fontSize=24, bold=false, font='Calibri'
[PAPERLESS CONFIG] Fields restored from configuration: 6

— OR if config is missing/corrupted —
[PAPERLESS CONFIG] No PaperLessConfig sheet found — legacy workbook
[PAPERLESS CONFIG] WARNING: Configuration JSON invalid
[PAPERLESS CONFIG] WARNING: Unsupported schema version 2 — ignoring config
```

## JSON Schema

```json
{
  "schemaVersion": 1,
  "paperless": {
    "version": "1.0"
  },
  "sheets": [
    {
      "name": "Sheet1",
      "fields": [
        {
          "id": "p1f1",
          "cell": "$A$1:$B$2",
          "type": "KeyboardText",
          "style": {
            "font": {
              "name": "Calibri",
              "sizePt": 24,
              "bold": false,
              "italic": false,
              "underline": false,
              "colorArgb": null
            },
            "alignment": {
              "horizontal": "left",
              "vertical": "center"
            },
            "fill": null,
            "wrapText": false
          },
          "config": {
            "required": false,
            "maxLength": 0,
            "inputRestriction": "None",
            "lines": 1
          }
        }
      ]
    }
  ]
}
```

- **Runtime values** (FIELD-1, FIELD-2, etc.) are **NOT stored** — values remain ephemeral
- **`style`** reuses `CellStyle` from `StyleDefinition.cs` (FontDefinition, AlignmentDefinition, FillDefinition)
- Minified JSON stored as inline string in cell B1 of `PaperLessConfig` worksheet
- Sheet state: `veryHidden` — invisible to normal Excel users

## Field Matching Strategy

Fields are matched by **normalized cell address**:

| Config Cell | Excel-Detected Cell | Match? |
|---|---|---|
| `$A$1:$B$2` | `A1` | Yes (both normalize to `A1`) |
| `$C$1:$D$2` | `C1` | Yes |
| `A12` | `A12` | Yes |

Sheet name + normalized cell address is the composite key. This is more reliable than matching by generated field ID because the ID may be exactly what we are restoring.

## Overlay Priority

| Property | Excel Detection | PaperLess Config | Winner |
|---|---|---|---|
| Field ID | generated ID (e.g., `field_0_0`) | `p1f1` | Config (stable identity) |
| Font Size | 14 (from Excel) | 24 (user's last edit) | Config |
| Font Name | Calibri | Arial | Config |
| Bold | false | true | Config |
| Required | false | true | Config |
| Cell Address | `A1` | `$A$1:$B$2` | Excel (structure) |
| Print Area | `A1:D12` | *(not stored)* | Excel |
| Merged Cells | detected | *(not stored)* | Excel |

## Build Verification

| Project | Status |
|---|---|
| Backend (`dotnet build`) | **0 errors, 0 new warnings** |
| Frontend (`npm run build`) | **Compiled successfully** |

## Test Cases

| # | Test | Expected |
|---|---|---|
| 1 | Upload original Excel (no config) | Existing field detection works |
| 2 | Change p1f1 fontSize 14→24, export | PaperLessConfig exists, JSON has `sizePt: 24` |
| 3 | Re-upload exported workbook | `p1f1.fontSize = 24` restored (not 14) |
| 4 | Field IDs survive Upload→Export→Upload | `p1f1`–`p1f6` preserved through round-trip |
| 5 | Export same workbook 3 times | Exactly 1 PaperLessConfig sheet (no duplicates) |
| 6 | Legacy workbook upload | Existing pipeline runs normally |
| 7 | Corrupted JSON in config sheet | Warning logged, Excel detection fallback |
| 8 | Config sheet not in rendering | PNG output unchanged, no PaperLessConfig visible |
| 9 | Runtime values still work | Values pipeline unchanged — config is metadata only |
