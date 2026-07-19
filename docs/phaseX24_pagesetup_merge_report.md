# PageSetup Merge Fix Report — Phase X.24

## Problem

After upload → generate, the output Excel loses:
- **Center on page** (horizontally and vertically)
- **Custom margins** (0.70866in left/right, 0.74803in top/bottom)

Page centering reverts to off; margins revert to Excel defaults (0.7in / 0.75in).

## Root Cause

`ReadPageSettings()` reads the **content sheet** (visible sheet), which only has Excel defaults for PageSetup. The authoritative settings live on the hidden **`_Fields`** sheet, but the pipeline never reads from it.

### OOXML Evidence

Extracted from `FormTest - Copy.xlsx`:

| Sheet | `CenterHorizontally` | `CenterVertically` | Left Margin | Right Margin |
|-------|----------------------|--------------------|-------------|--------------|
| `_Fields` (hidden) | `1` (true) | `1` (true) | 0.70866in (51.023pt) | 0.70866in (51.023pt) |
| Sheet1 (content) | `0` (false) | `0` (false) | 0.7in (50.4pt) | 0.7in (50.4pt) |

The content sheet's values are Excel defaults; the `_Fields` sheet holds the workbook's true print settings.

## Pipeline Audit

| Component | Status | Detail |
|-----------|--------|--------|
| `ReadPageSettings()` | Reads from wrong sheet | Called only on visible content sheets, never on `_Fields`. Currently reads `CenterHorizontally`, `CenterVertically`, margins, orientation, zoom, fit-to-pages, paper size — all present in model. |
| `PageSettings` model | Complete | Already has all properties: `CenterHorizontally`, `CenterVertically`, left/right/top/bottom margins, orientation, zoom, fit-to-pages-wide/tall, paper size. |
| `ApplyPageSettings()` | Writes correctly | Current (forensic-logging) version writes centering, margins, and all other PageSetup properties to the output sheet. |
| Upload pipeline | Missing `_Fields` read | After reading content-sheet PageSetup in Step 4a, never reads `_Fields` sheet and never merges. |

## Fix — `MergePageSettings()` Helper

**File:** `ExcelAPI/Services/WorkbookReaderService.cs`

A new method merges `PageSettings` from the content sheet and `_Fields` sheet with the following priority rules:

### Merge Rules

| Property | Content → _Fields Fallback Condition | Rationale |
|----------|--------------------------------------|-----------|
| `CenterHorizontally` | `content \|\| fields` (OR) | If either sheet has it true, preserve it. |
| `CenterVertically` | `content \|\| fields` (OR) | Same as above. |
| Left/Right Margins | `Math.Abs(content - 50.4pt) < 0.5` → use _Fields | Excel default L/R margin = 50.4pt (0.7in). If content matches, it was unset. |
| Top/Bottom Margins | `Math.Abs(content - 54.0pt) < 0.5` → use _Fields | Excel default T/B margin = 54.0pt (0.75in). Same logic. |
| Orientation | `content == "portrait"` → use _Fields | Portrait is Excel default. Landscape is always explicit. |
| Zoom | `content == 100` → use _Fields | 100% is Excel default zoom. |
| FitToPagesWide | `content == 0` → use _Fields | 0 = "not set" in OpenXML. |
| FitToPagesTall | `content == 0` → use _Fields | Same as above. |
| PaperSize | `content == "Letter"` → use _Fields | Letter is the default paper size. |

### Call Site

After Step 4b (cell style capture from `_Fields`), the merge runs for every content sheet:

```csharp
if (fieldsSheet != null && form.Sheets.Count > 0)
{
    var fieldsPs = ReadPageSettings(fieldsSheet);
    foreach (var sheetDef in form.Sheets)
    {
        sheetDef.PageSettings = MergePageSettings(sheetDef.PageSettings, fieldsPs);
    }
}
```

The method also includes structured `_logger.LogInformation` calls that log content, _Fields, and merged values in a single line each, making forensic verification straightforward.

## Scenario Verification

| Scenario | Content Sheet | _Fields Sheet | Expected Result | Merge Produces |
|----------|--------------|---------------|-----------------|----------------|
| **Legacy upload** (initial) | Excel defaults (false, false, 50.4pt, portrait, 100, Letter) | Custom (true, true, 51pt, landscape, 85, Legal) | Use _Fields values | H=true V=true, 51pt margins, landscape, zoom 85, Legal ✓ |
| **Re-upload** (generated file re-imported) | Merged values from prior export (true, true, 51pt, landscape, 85, Legal) | Same merged values | Keep content (already correct) | Content values preserved ✓ |
| **Both custom** (partial edits) | Some values explicit, some default | Covers remaining gaps | Content wins where set, _Fields fills | OR for booleans, non-default for others ✓ |
| **No _Fields sheet** | Any | null | Content values only | Merge skipped entirely ✓ |

## Files Changed

| File | Change |
|------|--------|
| `ExcelAPI/Services/WorkbookReaderService.cs:188-196` | Call site — merge after content loop |
| `ExcelAPI/Services/WorkbookReaderService.cs:1140-1197` | `MergePageSettings()` method |
| `ExcelAPI/Services/WorkbookReaderService.cs:1204-1247` | Helper methods: `MergeMargin()`, `MergeOrientation()`, `MergeIntProperty()`, `MergeStringProperty()` |

No changes were made to `WorkbookGenerator.cs`, `FormDefinition.cs`, or any other pipeline component.

## Build Status

`dotnet build` — **0 errors, 0 new warnings.** All pre-existing warnings are unrelated (nullable, obsolete SKIA API, NuGet advisory).

## Next Steps

1. Re-upload `FormTest - Copy.xlsx` with updated code.
2. Generate a fresh output.
3. Validate in Excel:
   - Page Setup → Center on page (horizontally + vertically) is checked.
   - Margins → Custom (Left=0.71, Right=0.71, Top=0.75, Bottom=0.75).
   - Border on A12 (from Phase X.23 fixes) is visible.
4. If successful, consider adding `HeaderMargin` / `FooterMargin` to the model (identified gap not yet addressed).
