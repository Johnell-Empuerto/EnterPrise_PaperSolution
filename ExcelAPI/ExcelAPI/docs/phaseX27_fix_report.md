# Phase X.27 — Print-Position Fix: Three Root Causes & Fixes

## Summary

The generated workbook's print position (margins, centering, print area) was wrong because three distinct bugs corrupted the PageSettings values between upload and generation. All three are fixed in this phase.

---

## Bug 1: FieldsPageSettings safety-net overwrites correct content margins

**File**: `Generators/WorkbookGenerator.cs` (lines 139–163, now removed)

**Root cause**: After `ApplySheetLayout()` correctly set custom margins (51.0236pt left, 53.8583pt top) and centering (true/true) on the content sheet via `ApplyPageSettings(ws, sheetDef.PageSettings)`, a subsequent "safety-net" block applied `FieldsPageSettings` (the _Fields hidden sheet's OOXML PageSetup) to the **same content sheet**. The `_Fields` sheet has default margins (50.4/54pt) and no centering, so it overwrote the correct values with defaults.

```csharp
// OLD — removed in X.27
if (form.FieldsPageSettings != null && sheetIndexMap.TryGetValue(form.Sheets?[0]?.Id ?? "", out int sheet1Idx))
{
    var wsSheet1 = (Excel.Worksheet)workbook.Worksheets[sheet1Idx];
    ApplyPageSettings(wsSheet1, form.FieldsPageSettings);
}
```

**Fix**: Removed the safety-net block entirely. `ApplySheetLayout()` is the sole authority for content sheet PageSetup.

---

## Bug 2: MergePageSettings corrupts round-trip PageSettings JSON

**File**: `Services/WorkbookReaderService.cs` (lines 236–243, now removed)

**Root cause**: The `Read()` method merged `FieldsPageSettings` (defaults) into each content sheet's `PageSettings` via `MergePageSettings()`. While `MergeMargin()` preserved custom margins (tolerance check), `MergeOrientation()` returned `fields` (null/empty) for portrait content, and `MergeIntProperty()` could overwrite `FitToPagesWide`/`FitToPagesTall`. The merged (corrupted) PageSettings was then serialized to JSON and returned to the client. When the client re-posted this JSON to `output-excel`, the content sheet received wrong values.

```csharp
// OLD — removed in X.27
if (form.FieldsPageSettings != null && form.Sheets.Count > 0)
{
    foreach (var sheetDef in form.Sheets)
    {
        sheetDef.PageSettings = MergePageSettings(sheetDef.PageSettings, form.FieldsPageSettings);
    }
}
```

**Fix**: Removed the merge step. `FieldsPageSettings` is still read from OOXML and stored on `FormDefinition` (for client-side display), but it no longer poisons the content sheet's PageSettings.

---

## Bug 3: CalculatePrintArea ignores merge cell extents (missing column D)

**File**: `Generators/WorkbookGenerator.cs` (lines 209–268)

**Root cause**: `CalculatePrintArea()` only iterated over `form.Clusters` — each cluster has a `CellAddress` pointing to the first cell of a merge range (e.g., `"C1"`, `"A3"`). The max column was the highest column among these single-cell references (= 3 = C), so the print area was `$A$1:$C$12`. It missed column D, which only appeared in merge range addresses (`C1:D2`, `A3:D4`, etc.).

```csharp
// OLD — only looked at clusters
foreach (var cluster in form.Clusters) {
    ParseAddressExtents(cluster.CellAddress, ...);
}
return $"$A$1:${endCol}${maxRow}"; // maxCol = 3 → $A$1:$C$12
```

**Fix**: Added a second pass over `SheetDefinition.MergedCells` addresses to extend the extents to cover full merge ranges:

```csharp
// NEW — also considers merge cells
if (form.Sheets != null) {
    foreach (var sheet in form.Sheets) {
        foreach (var mc in sheet.MergedCells) {
            ParseAddressExtents(mc.Address, ...);
        }
    }
}
return $"$A$1:${endCol}${maxRow}"; // maxCol = 4 → $A$1:$D$12
```

Also extracted the address-parsing loop into a shared `ParseAddressExtents()` helper to eliminate duplication.

---

## Verification Results

After all three fixes, the generated workbook's OOXML and COM values match ConMas exactly:

| Property | Before fix | After fix | ConMas | Match? |
|----------|-----------|-----------|--------|--------|
| `pageMargins left/top` | 0.7 / 0.75 | 0.70866 / 0.74803 | 0.70866 / 0.74803 | ✓ |
| `printOptions h=1 v=1` | missing | present | present | ✓ |
| `Print_Area` definedName | `$A$1:$C$12` | `$A$1:$D$12` | `$A$1:$D$12` | ✓ |
| COM CenterHorizontally | False | True | True | ✓ |
| COM LeftMargin (pt) | 70 | 51.0236 | 51.0236 | ✓ |
| COM TopMargin (pt) | 70 | 53.8583 | 53.8583 | ✓ |
| COM FitToPagesWide | 0 | 1 | 1 | ✓ |
| COM FitToPagesTall | 0 | 1 | 1 | ✓ |

## Build Status

`dotnet build` succeeds with 0 errors. Server restarted and tested end-to-end: upload → JSON round-trip → output-excel → OOXML/COM verification all pass.
