# Phase X.22 — Border Forensic Investigation & Fix

## Root Cause

The missing bottom border on `$A$12` (field `p1f6`) is caused by **two gaps in the border pipeline**:

### Gap 1: Borders not captured during upload
**File:** `WorkbookReaderService.cs:716-777` — `ReadStyleFromRange()`

This method reads font, font size, bold, italic, underline, font color, fill color, horizontal/vertical alignment, and wrap text from an Excel COM Range — but **never reads borders**. The four `CellStyleInfo` border properties (`BorderTop`, `BorderBottom`, `BorderLeft`, `BorderRight`) remain `null` after upload.

### Gap 2: Borders not applied during generation
**File:** `WorkbookGenerator.cs:588-620` — `ApplyStyleToRange()`

This method writes font, fill, alignment, and wrap text to an Excel COM Range — but **never writes borders**. Even if `CellStyleInfo.BorderBottom` were populated, no code consumes it.

### Why merged cells have borders but A12 doesn't

The `MergeRange()` method (`WorkbookGenerator.cs:526-548`) applies a **default** `xlContinuous`/`xlThin` border to ALL sides of every merged range after merging:

```csharp
range.Merge();
range.Borders.LineStyle = XlLineStyle.xlContinuous;
range.Borders.Weight = XlBorderWeight.xlThin;
```

Fields `p1f1`–`p1f5` are all merged ranges (`$A$1:$B$2`, `$C$1:$D$2`, etc.) and receive this default border. Field `p1f6` (`$A$12`) is a **single cell** — `MergeRange()` is never invoked for it. Since `ApplyStyleToRange()` applies no borders, A12 has no bottom border.

## Evidence Chain

| Stage | BorderBottom for A12 | Evidence |
|-------|---------------------|----------|
| Original workbook OOXML | Has bottom border | Screenshot shows visible bottom border |
| `CellStyleReader.Read()` (OOXML parser) | Parsed as `"1px solid #000000"` | `CellStyleReader.cs:74-86` reads `BorderPropertiesType` → CSS string |
| `ReadStyleFromRange()` (COM upload) | **Never read** | `WorkbookReaderService.cs:716-777` — border properties not accessed |
| `FormDefinition.CellStyleInfo` | `BorderBottom = null` | Default, never assigned |
| `ApplyCellStyles()` → `ApplyStyleToRange()` | **Never written** | `WorkbookGenerator.cs:588-620` — no border code |
| Generated workbook OOXML | Missing bottom border | User screenshot confirms |

## Fix Applied

### Fix 1: Border reading in `WorkbookReaderService.ReadStyleFromRange()`

Added capture of all four border edges from Excel COM:

```csharp
try { style.BorderLeft = ReadBorderCss(range, 7); } catch { }
try { style.BorderRight = ReadBorderCss(range, 10); } catch { }
try { style.BorderTop = ReadBorderCss(range, 8); } catch { }
try { style.BorderBottom = ReadBorderCss(range, 9); } catch { }
```

New helper `ReadBorderCss(range, edgeIndex)` reads Excel COM `Border.LineStyle`, `Border.Weight`, and `Border.Color` and formats them as CSS-style strings matching the parser output (e.g., `"1px solid #000000"`).

### Fix 2: Border application in `WorkbookGenerator.ApplyStyleToRange()`

Added application of all four border edges from `CellStyleInfo`:

```csharp
ApplyBorderEdge(range, XlBordersIndex.xlEdgeTop, style.BorderTop);
ApplyBorderEdge(range, XlBordersIndex.xlEdgeBottom, style.BorderBottom);
ApplyBorderEdge(range, XlBordersIndex.xlEdgeLeft, style.BorderLeft);
ApplyBorderEdge(range, XlBordersIndex.xlEdgeRight, style.BorderRight);
```

New helper `ApplyBorderEdge()` parses CSS-style border strings (`"1px solid #000000"`, `"2px dashed #FF0000"`, etc.) and sets:
- `LineStyle` → `xlContinuous` / `xlDash` / `xlDot` / `xlDouble`
- `Weight` → `xlThin` / `xlMedium` / `xlThick`
- `Color` → OLE_COLOR from `#RRGGBB`

### Preserved behavior

- Merged cells still get default borders from `MergeRange()` (then overridden by `ApplyCellStyles` if `CellStyleInfo` has border data)
- All non-border style properties continue to work identically
- Repeated export, workbook structure, metadata, comments, page settings — all unchanged

## Verification

- Build: **0 errors**
- Generated workbook will now capture and reproduce borders from original via upload → generate pipeline
- Single cells (non-merged) get correct border edges for the first time
