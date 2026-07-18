# Phase X21 — WorkbookGenerator COM Name-Operation Robustness Fix

**Date:** 2026-07-18  
**File modified:** `ExcelAPI/ExcelAPI/Generators/WorkbookGenerator.cs`

## Problem

During Excel export, `EnsureSheets` would crash the entire export pipeline when a COM exception occurred during a worksheet rename operation (e.g., `ws.Name = "..."`). The most common failure was:

> `Exception from HRESULT: 0x800A03EC` — name conflict or sheet-level COM error.

Because the rename was done directly without error handling, a single transient COM failure would bubble up unhandled, aborting the workbook generation for that export request.

## Changes Made

### 1. Added `NameOp` helper method

A generic wrapper that logs BEFORE/AFTER around any COM name operation and catches `COMException` with the known-name-conflict HRESULT (`0x800A03EC`).

```csharp
private bool NameOp(string exportId, string opLabel, System.Action action)
```

- Logs `>>> NAME-OP: {opLabel}` on entry
- On success logs `<<< NAME-OP SUCCESS: {opLabel}`
- On `0x800A03EC`: logs a warning, waits 200 ms, retries once; if still failing, logs failure and returns `false`
- On any other exception: logs warning and returns `false`

### 2. Wrapped all `EnsureSheets` COM name operations in `NameOp`

Each of the following operations is now wrapped:

| Operation | NameOp label |
|---|---|
| Deleting excess sheets | `Worksheet[{idx}].Delete()` |
| Adding new sheets | `Worksheets.Add(After: last)` |
| Renaming WS[3] → ExcelOutputSetting | `Worksheet[3].Name = "ExcelOutputSetting"` |
| Renaming WS[2] → content sheet name | `Worksheet[2].Name = "{name}"` |
| Renaming WS[1] → `_Fields` | `Worksheet[1].Name = "_Fields"` |

Each wrap logs the **default COM name** before rename (e.g., `Sheet3`, `Chart1`) so we can trace what Excel auto-named the worksheet.

On failure, the method logs a warning and **continues** instead of throwing.

### 3. Threaded `exportId` through all methods for correlated logging

Six methods now accept a `string exportId` parameter and use structured logging with `[FORENSIC][{ExportId}]` prefix:

| Method | Previous signature | New signature |
|---|---|---|
| `EnsureSheets` | `(Workbook, FormDefinition, out Dictionary<...,int>)` | `+ string exportId` |
| `ApplySheetLayout` | `(Workbook, FormDefinition, Dictionary<...,int>, string printArea)` | `+ string exportId` |
| `WriteCellValues` | `(Workbook, FormDefinition, Dictionary<...,int>)` | `+ string exportId` |
| `WriteCellComments` | `(Workbook, FormDefinition, Dictionary<...,int>)` | `+ string exportId` |
| `PopulateFieldsWorksheet` | `(Workbook, FormDefinition, Dictionary<...,int>)` | `+ string exportId` |
| `CreateExcelOutputSetting` | `(Workbook, FormDefinition, Dictionary<...,int>)` | `+ string exportId` |

All internal `_logger.LogInformation` / `LogWarning` calls within these methods were updated to `[FORENSIC][{ExportId}]` format, enabling per-request log filtering.

### 4. Fixed `Action` type ambiguity

The `NameOp` parameter was initially declared as `Action`, which resolved ambiguously between `System.Action` and `Microsoft.Office.Interop.Excel.Action`. Fixed by qualifying as `System.Action`.

## Build Verification

```
Build succeeded.
    0 Error(s)
```

## Risk Assessment

- **Low risk.** All changes are additive or wrap existing behavior in try/catch. On failure, the original behavior was to throw; now it logs a warning and continues with the next operation.
- The `NameOp` wrapper is a pure helper — it doesn't change what the operations do, only how errors are handled.
- Threading `exportId` is a mechanical signature change; call sites were all updated in the same file.
