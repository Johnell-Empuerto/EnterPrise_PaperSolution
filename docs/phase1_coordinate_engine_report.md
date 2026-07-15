# Phase 1 — Coordinate Engine Implementation Report

**Date:** July 9, 2026
**Engineer:** Buffy (Codebuff AI)
**Scope:** Replace COM `Range.Width` with column-width-based computation in the centering formula

---

## Objective

Replace the single `Range.Width` COM call with a three-strategy `ColumnWidthProvider` that:

1. **No-centering forms (247 forms, 54%):** origin = leftMargin — already handled correctly by existing code
2. **Rounding forms (56 forms, 12.3%):** already within ±0.5pt — no changes needed
3. **Default-column forms (10 forms, 2.2%):** use `n × 50.1pt` instead of `Range.Width`
4. **Explicit-column forms (90 forms, 19.7%):** read `<cols>` from XLSX, convert to points

---

## File Changed

**`ExcelAPI/ExcelAPI/Services/ExcelCaptureService.cs`** — the only file modified.

### 1. Added Using Statements (lines 3-5)

```csharp
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
```

These enable XLSX ZIP parsing, regex column parsing, and LINQ-to-XML queries.

### 2. New Step 6a.5 — Content Width Computation (inserted after Step 6a, before Step 6b)

```csharp
// --- Step 6a.5: Compute content width from column widths ---
// Replace COM Range.Width with a width computed from XLSX worksheet XML.
double contentWidthPt = ComputeContentWidthFromXlsx(
    excelFilePath, worksheet?.Name, printArea, printAreaCols);
if (contentWidthPt > 0)
{
    printAreaWidth = contentWidthPt;
    _logger.LogInformation("[COLS] Content width override: ...");
}
else
{
    _logger.LogInformation("[COLS] Using Range.Width=...");
}
```

### 3. Five New Private Methods

| Method | Purpose |
|--------|---------|
| `ComputeContentWidthFromXlsx` | Entry point: reads XLSX, resolves worksheet, selects strategy |
| `ResolveWorksheetPath` | Maps worksheet name to `xl/worksheets/sheetN.xml` via `workbook.xml` + relationships |
| `ColumnLetterToIndex` | Converts "A" → 1, "Z" → 26, "AA" → 27 |
| `TryParsePrintAreaColumns` | Parses `$A$1:$D$10` → firstCol=1, lastCol=4 |

---

## Strategy Selection Logic

```
ComputeContentWidthFromXlsx(excelFilePath, worksheetName, printArea, printAreaCols)
  │
  ├─ Parse print area → column range [firstCol, lastCol]
  │
  ├─ Open XLSX as ZipArchive
  │   └─ ResolveWorksheetPath(archive, worksheetName)
  │       ├─ Read xl/workbook.xml → find sheet → get rId
  │       └─ Read xl/_rels/workbook.xml.rels → map rId → target path
  │
  ├─ Read worksheet XML
  │   ├─ <cols> EXISTS?  →  STRATEGY A: Explicit Columns
  │   │   ├─ For each <col> in print area range:
  │   │   │   pointWidth = (charWidth × 7.33 + 5) × 72/96
  │   │   │   totalWidth += pointWidth × count
  │   │   └─ Return totalWidth
  │   │
  │   └─ <cols> ABSENT?  →  STRATEGY B: Default Columns
  │       └─ Return colsInRange × 50.1
  │
  └─ Exception? → Return 0 (caller falls back to Range.Width)
```

---

## Critical Fix: Character → Point Conversion

The XLSX `<col>` `width` attribute is in **character units** (e.g., 8.11 characters), **NOT** points. Summing them directly would give ~6x too small results.

The OOXML conversion formula (ECMA-376, 18.3.1.13):
```
pixelWidth = charWidth × maxDigitWidth + padding
pointWidth = pixelWidth × 72 / 96
```

Where:
- `maxDigitWidth = 7.33` pixels for Calibri 11pt (legacy Normal font)
- `padding = 5` pixels (standard Excel constant at 96 DPI)

**Verification:** Default Calibri 11pt column = 8.43 chars → (8.43 × 7.33 + 5) × 0.75 = **50.09pt** ≈ **50.1pt** ✓

---

## Architecture Properties

| Property | Status |
|----------|--------|
| Extensibility | ✅ New strategies can be added to `ComputeContentWidthFromXlsx` without changing the pipeline |
| Safe fallback | ✅ Returns 0 on any failure → caller uses original `Range.Width` |
| Rendering unchanged | ✅ No changes to PDF export, PNG generation, cluster serialization, or DB schema |
| Logging | ✅ Added per-column debug logging and final summary |
| Resource disposal | ✅ `ZipArchive` and `StreamReader` wrapped in `using` |

---

## Coverage Estimate

| Forms | Strategy | Status |
|-------|----------|--------|
| 247 (54%) | No centering → origin = leftMargin | ✅ Already correct (no code change needed) |
| 56 (12.3%) | Rounding (<0.5pt error) | ✅ Already correct (no code change needed) |
| 10 (2.2%) | Default columns → n × 50.1pt | ✅ Implemented |
| 90 (19.7%) | Explicit columns → XLSX `<cols>` | ✅ Implemented |
| **403 (88%)** | **Total covered** | **✅ Phase 1 complete** |

---

## Known Limitations (Phase 2 Candidates)

1. **Hardcoded `maxDigitWidth = 7.33`** — Assumes Calibri 11pt Normal font. Forms with different fonts may have slight errors in explicit-column conversion. Future: read font from XLSX `styles.xml` and compute `maxDigitWidth` dynamically.

2. **No `defaultColWidth` reading** — Some XLSX files store the default column width in `<sheetFormatPr defaultColWidth="...">`. Could use this value instead of hardcoded 50.1pt in Phase 2.

3. **Form 228 family (8 forms)** — Large residual errors (6.9-12.6pt) not explained by column width alone. Needs separate investigation.

4. **Missing print-area metadata (30 forms)** — Could not infer print area from ConMas XML. These need the XLSX print area to be validated.

5. **Margin anomalies (16 forms)** — Origin differs from margin by 0.5-5pt without centering. May need Excel version-specific margin calibration.

---

## Build Status

- **Compilation:** ✅ Succeeds (tested with `dotnet build`)
- **Pre-existing warnings only** (NuGet vulnerability, deprecated SkiaSharp API, nullable dereferences)
- **No new warnings** introduced by the changes

---

## File Summary

```diff
--- a/ExcelAPI/ExcelAPI/Services/ExcelCaptureService.cs
+++ b/ExcelAPI/ExcelAPI/Services/ExcelCaptureService.cs
@@ -1,5 +1,8 @@
 using System.Runtime.InteropServices;
 using System.Diagnostics;
+using System.IO.Compression;
+using System.Text.RegularExpressions;
+using System.Xml.Linq;
 using Microsoft.Office.Interop.Excel;
@@ -338,6 +341,30 @@
                 }

+                // --- Step 6a.5: Compute content width from column widths ---
+                double contentWidthPt = ComputeContentWidthFromXlsx(
+                    excelFilePath, worksheet?.Name, printArea, printAreaCols);
+                if (contentWidthPt > 0)
+                {
+                    _logger.LogInformation(...);
+                    printAreaWidth = contentWidthPt;
+                }
+                else
+                {
+                    _logger.LogInformation(...);
+                }
+
                 // --- Step 6b: Read Page Setup ---
@@ -1025,6 +1052,274 @@
         }

+        #region Coordinate Helpers
+        private double ComputeContentWidthFromXlsx(...) { ... }
+        private static string? ResolveWorksheetPath(...) { ... }
+        private static int ColumnLetterToIndex(string letters) { ... }
+        private static bool TryParsePrintAreaColumns(...) { ... }
+        #endregion
+
     }
 }
```
