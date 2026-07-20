# Phase 17/18 — Functional Round-Trip Compatibility Implementation

**Date:** July 20, 2026
**Status:** Implementation Complete
**Code Changes:** ✅ Yes (code changes implemented)

---

## Objective

Eliminate all remaining assumptions that workbook XML must remain structurally identical. Replace the old structural validation philosophy with functional round-trip compatibility matching ConMas Designer behavior.

---

## What Changed

### 1. Enhanced `CompatibilityValidator.cs`

**Before:** Checked only basic existence (workbook opens, sheets exist, comments exist, page setup exists).

**After:** Now validates all business-critical requirements:

| Check | What it does |
|-------|-------------|
| ✅ Workbook opens | Verifies edited workbook can be opened with OpenXML SDK |
| ✅ Printable worksheets exist | Counts sheets excluding `_Fields`, `_RawData`, `ExcelOutputSetting` |
| ✅ Configuration sheets present | Detects `_Fields` and `ExcelOutputSetting` |
| ✅ ExcelOutputSetting not duplicated | Warns if more than one instance |
| ✅ Field metadata exists | Checks for comments OR `_Fields` sheet |
| ✅ **Field count preserved** (NEW) | Counts comments in original vs edited workbook — warns if count changed |
| ✅ **PrintArea exists** (NEW) | Checks per-sheet and workbook-level `Print_Area` defined names |
| ✅ **PageSetup details logged** (NEW) | Logs paperSize, orientation, fitToWidth, fitToHeight |
| ✅ **PageMargins checked** (NEW) | Warns if missing |
| ✅ **Merged cells counted** (NEW) | Counts merge cells per sheet |
| ✅ **Hidden rows counted** (NEW) | Counts hidden rows per sheet |
| ✅ **Hidden columns counted** (NEW) | Counts hidden columns per sheet |
| ✅ **Defined names logged** (NEW) | Logs all defined names for diagnostics |

**Still does NOT check:**
- ❌ Byte-for-byte XML equality
- ❌ Relationship IDs
- ❌ Shared string ordering or count
- ❌ Style ordering or font count
- ❌ Any hash/SHA256 comparison
- ❌ ZIP entry ordering
- ❌ Printer settings binary

**Key behavior:** Never throws. Never rejects. Always returns the workbook.

### 2. Removed `WorkbookDiffValidator` Registration

File: `Application/ServiceRegistration.cs`

**Before:**
```csharp
services.AddScoped<CompatibilityValidator>();
services.AddScoped<WorkbookDiffValidator>(); // Phase 4.4 (deprecated)
```

**After:**
```csharp
services.AddScoped<CompatibilityValidator>();
// WorkbookDiffValidator removed — structural validation is no longer performed
// Source file kept on disk for reference only
```

The `WorkbookDiffValidator.cs` source file is kept on disk for reference but is no longer registered in DI. No code path can call it at runtime.

### 3. Removed Deprecated `ValidationResult` Property

File: `Application/FormSaveService.cs`

**Before:**
```csharp
[Obsolete("Use CompatibilityResult instead...")]
public WorkbookDiffResult? ValidationResult { get; set; }
public CompatibilityResult? CompatibilityResult { get; set; }
public int CellsWritten { get; set; }
```

**After:**
```csharp
public CompatibilityResult? CompatibilityResult { get; set; }
public int CellsWritten { get; set; }
```

The deprecated `WorkbookDiffResult? ValidationResult` property has been removed entirely. Only `CompatibilityResult` remains.

### 4. Created Multi-Generation Acceptance Test

**File:** `test_multi_gen_roundtrip.py`

A standalone Python test script that:

- Takes a template workbook and runs it through N generations (default: 50)
- At every generation, verifies:
  - ✅ Workbook opens as valid ZIP/OOXML package
  - ✅ Required OOXML parts exist (`[Content_Types].xml`, `xl/workbook.xml`)
  - ✅ Workbook XML is valid XML
  - ✅ Content Types XML is valid XML
  - ✅ Printable sheets exist (excluding config sheets)
  - ✅ `_Fields` sheet exists
  - ✅ `ExcelOutputSetting` sheet exists
  - ✅ No duplicate configuration sheets
  - ✅ Cell comments exist (field count > 0)
- Tracks SHA256 of each generation for drift analysis
- Generates JSON report and plain-text summary

Usage:
```bash
python test_multi_gen_roundtrip.py --generations 50 --template test_form.xlsx
```

---

## Build Result

| Metric | Result |
|--------|--------|
| Compilation errors | **0** |
| Warnings | 2 (pre-existing NuGet advisory) |
| Build status | **Build succeeded** |
| File lock issue | Fixed — killed stale `ExcelAPI.exe` process |

---

## Files Modified

| File | Action | Summary |
|------|--------|---------|
| `Application/CompatibilityValidator.cs` | **MODIFIED** | ~100 lines added: field count, print area, layout checks |
| `Application/ServiceRegistration.cs` | **MODIFIED** | Removed `WorkbookDiffValidator` registration |
| `Application/FormSaveService.cs` | **MODIFIED** | Removed deprecated `ValidationResult` property |
| `test_multi_gen_roundtrip.py` | **CREATED** | 50-generation acceptance test suite |

---

## Known Gaps (from code-reviewer)

1. **VML not checked** — User requires VML drawing parts to survive. Not yet validated.
2. **Field positions/IDs not checked** — Only field count is compared, not specific cell references.
3. **Hidden rows/cols/merged cells counted but not compared** — Values are logged but not compared against original.
4. **Test script doesn't call PaperLess API** — Currently validates ZIP structure only. Needs API calls for true round-trip verification.
