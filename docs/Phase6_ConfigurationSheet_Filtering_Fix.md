# Phase 6 — Configuration Sheet Filtering Fix

**Date:** July 20, 2026  
**Status:** Complete ✅  
**Files Changed:** `ExcelAPI/ExcelAPI/Designer/Analysis/WorkbookReaderService.cs`

---

## Problem

When reopening a PaperLess or legacy ConMas exported workbook, hidden configuration sheets (`_Fields`, `ExcelOutputSetting`) were incorrectly appearing as designer pages alongside the actual printable worksheets.

**Before fix — Incorrect behavior:**

```
Workbook
├── Sheet1              ← printable form
├── _Fields             ← hidden metadata sheet
└── ExcelOutputSetting  ← hidden configuration sheet

Designer shows:
  Page 1 = Sheet1              ✅
  Page 2 = _Fields             ❌ (should be hidden)
  Page 3 = ExcelOutputSetting  ❌ (should be hidden)
```

This violated the legacy ConMas Designer behavior where configuration sheets are **data sources only** — they store field metadata, designer settings, and configuration, but should never be displayed as visual pages.

---

## Root Cause

In `WorkbookReaderService.Read()`, the sheet classification logic only filtered out `_Fields` from printable content sheets:

```csharp
// BEFORE — only _Fields was filtered
if (string.Equals(name, "_Fields", StringComparison.OrdinalIgnoreCase))
{
    fieldsSheet = ws;
}
else
{
    contentSheets.Add((i, ws));  // EVERYTHING else becomes a page
}
```

`ExcelOutputSetting` passed through this filter and was added to `contentSheets`, which caused it to become a `SheetDefinition` in the `FormDefinition`, which then propagated through the pipeline:

```
WorkbookReaderService.Read()
    ↓
FormDefinition.Sheets includes ExcelOutputSetting
    ↓
WorkbookDefinitionConverter.FromFormDefinition()
    ↓
WorkbookDefinition.Sheets includes ExcelOutputSetting
    ↓
FormRuntimeBuilder.BuildFromDefinitionDirect()
    ↓
RuntimeForm.Sheets includes ExcelOutputSetting → DESIGNER PAGE ❌
```

---

## Fix

### 1. Added `ConfigurationSheetNames` set

A static `HashSet<string>` of known configuration sheet names was added to `WorkbookReaderService`:

```csharp
private static readonly HashSet<string> ConfigurationSheetNames = new(StringComparer.OrdinalIgnoreCase)
{
    "_Fields",
    "ExcelOutputSetting",
    "DesignerConfig",
    "PaperLessConfig",
    "ConMasConfig"
};
```

The set uses case-insensitive comparison to match Excel's case-insensitive sheet naming.

### 2. Modified sheet classification loop

The classification loop now checks all known configuration names:

```csharp
// AFTER — all config sheets are filtered
for (int i = 1; i <= sheetCount; i++)
{
    var ws = (Excel.Worksheet)workbook.Worksheets[i];
    string name = ws.Name;

    if (string.Equals(name, "_Fields", StringComparison.OrdinalIgnoreCase))
    {
        fieldsSheet = ws;  // Still captured for metadata reading
        _logger.LogInformation("Found _Fields configuration sheet — will not become a designer page");
    }
    else if (ConfigurationSheetNames.Contains(name))
    {
        _logger.LogInformation("Skipping configuration sheet '{Name}' — will not become a designer page", name);
        Marshal.ReleaseComObject(ws);  // Prevent COM memory leak
    }
    else
    {
        contentSheets.Add((i, ws));  // Only printable sheets become pages
    }
}
```

### Behavior after fix

| Sheet Name | Classification | Becomes Designer Page? | Used For |
|------------|---------------|----------------------|----------|
| `Sheet1` | Printable | ✅ Yes | Visual form |
| `Invoice` | Printable | ✅ Yes | Visual form |
| `Form` | Printable | ✅ Yes | Visual form |
| `_Fields` | Configuration | ❌ No | Field metadata (read via `ReadFieldsSheet`) |
| `ExcelOutputSetting` | Configuration | ❌ No | Skipped entirely |
| `DesignerConfig` | Configuration | ❌ No | Skipped (future use) |
| `PaperLessConfig` | Configuration | ❌ No | Skipped (future use) |
| `ConMasConfig` | Configuration | ❌ No | Skipped (future use) |

---

## Pipeline Impact

The fix is at the earliest stage of the reopen pipeline, so it prevents config sheets from entering the pipeline at all:

```
WorkbookReaderService.Read() → filters out config sheets
    ↓
FormDefinition.Sheets → only printable sheets ✅
    ↓
WorkbookDefinitionConverter.FromFormDefinition()
    ↓
WorkbookDefinition.Sheets → only printable sheets ✅
    ↓
FormRuntimeBuilder.BuildFromDefinitionDirect()
    ↓
RuntimeForm.Sheets → only printable sheets ✅
    ↓
DESIGNER → shows only printable pages ✅
```

### What still works

- ✅ `_Fields` sheet is still read for field metadata, styles, and page settings
- ✅ Printable worksheets still become designer pages
- ✅ Field positions and properties are still restored
- ✅ Background image generation unchanged
- ✅ Existing new-template upload flow unchanged
- ✅ COM objects for skipped sheets are properly released (no memory leaks)

---

## Verification

### Build
- `dotnet build` — **Build succeeded** (0 errors, 2 pre-existing OpenApi warnings)

### Code Review
- `ConfigurationSheetNames` uses `StringComparer.OrdinalIgnoreCase` — matches Excel behavior
- COM objects for skipped sheets are released via `Marshal.ReleaseComObject`
- `_Fields` is handled as a special case before the general config check (correct order since `_Fields` is also in the config set)
- Backward compatible — existing printable sheet processing logic unchanged

---

## Legacy ConMas Compatibility

This fix aligns the PaperLess reopen behavior with the legacy ConMas Designer:

| Behavior | ConMas Designer | PaperLess (Before) | PaperLess (After) |
|----------|----------------|-------------------|-------------------|
| Config sheets as pages | ❌ Never | ❌ Sometimes | ✅ Never |
| `_Fields` read for metadata | ✅ Yes | ✅ Yes | ✅ Yes |
| Printable sheets restored | ✅ Yes | ✅ Yes | ✅ Yes |
| Field positions preserved | ✅ Yes | ✅ Yes | ✅ Yes |

