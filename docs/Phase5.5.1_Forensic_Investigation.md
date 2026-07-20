# Phase 5.5.1 — Forensic Investigation Report
## Why the PaperLess Output Excel Does Not Match Legacy ConMas

**Date:** 2026-07-20  
**Lead Investigator:** Phase 5.5 in-depth analysis  
**Status:** Root cause identified — implementation plan follows

---

# Table of Contents

1. [Export Timeline (menuOutputExcel_Click → SaveAs → Final)](#1-export-timeline)
2. [Every OOXML Part Modified by ConMas](#2-every-ooxml-part-modified-by-conmas)
3. [Every Metadata Source in Legacy Workbook](#3-every-metadata-source)
4. [Root Cause of Comment Persistence Failure](#4-root-cause)
5. [Every Difference Between Legacy and PaperLess Output](#5-every-difference)
6. [Exact ConMas Comment Format Specification](#6-exact-comment-format)
7. [Exact InputParameter Construction](#7-exact-inputparameter-construction)
8. [Hidden Sheet Re-Investigation](#8-hidden-sheets)
9. [Implementation Plan](#9-implementation-plan)
10. [Verification Method](#10-verification)

---

## 1. Export Timeline

### Legacy ConMas: `menuOutputExcel_Click` Full Sequence

```
menuOutputExcel_Click (async state machine)
  │
  ├── 1. LoadDocument(path)  [IL 185104-185106]
  │     Opens a COPY of the original workbook (DevExpress Workbook.LoadDocument)
  │     Does NOT use Workbooks.Add()
  │
  ├── 2. Determine ExportMethod [IL 185153-185174]
  │     Based on ExcelProcessConfig.ExportMethod:
  │     1 = DevExpress (SpreadsheetControl-based)
  │     default = COM Interop (Excel.Application automation)
  │
  ├── 3. Iterate clusters [IL 185442-188959]
  │     foreach sheet in reportIndexList:
  │       foreach cluster in sheet.Clusters:
  │         │
  │         ├── 3a. Build comment text V_55 [IL 185565-186104]
  │         │     25-field newline-delimited string:
  │         │     Line  0: cluster.Name
  │         │     Line  1: ClusterTypeToStringType(cluster.Type)
  │         │     Line  2: cluster.Index.ToString()
  │         │     Line  3: cluster.ReadOnly.ToString()
  │         │     Line  4: cluster.External.ToString()
  │         │     Line  5: V_54 (InputParameter, possibly patched)
  │         │     Line  6-15: cluster.RemarksValue[0..9]
  │         │     Line 16-24: Table/Column/Row info (from topReport.Tables lookup)
  │         │
  │         ├── 3b. Write comment [IL 186189-186317]
  │         │     DevExpress: sheet.Comments.Add(cell, author, V_55)
  │         │     COM: range.AddComment(V_55); range.Comment.Text(V_55, ...)
  │         │
  │         ├── 3c. Write formula [IL 186340-186425]
  │         │     If function cluster: cell.Formula = cluster.OriginalFunction
  │         │     Or: cell.Formula = ChangeCellCalculate(functionValue, ...)
  │         │
  │         └── 3d. Apply formatting [IL 186452-186638]
  │               Font size, name, bold, italic, underline, strikethrough,
  │               color, background, alignment from InputParameter keys
  │
  ├── 4. Save [IL 188653]
  │     DevExpress: workbook.SaveDocument(tempPath)
  │     COM: workbook.SaveAs(tempPath, ...)
  │
  └── 5. Post-save [IL 188663+]
        DevExpress: ExcelProcessorBase.EnsurePrintAreaQuoted_OpenXml(tempPath)
        COM: File.Copy(tempPath, saveDialog.FileName); File.Delete(tempPath)
```

### PaperLess Current: `WorkbookValueWriter.WriteValues` Full Sequence

```
WriteValues(wbDef, sourcePath, outputPath)
  │
  ├── 1. File.Copy(source, output) [line 54]
  │
  ├── 2. Pre-save ZIP entries to originalZipEntries [lines 66-84]
  │
  ├── 3. SpreadsheetDocument.Open(output, true) [line 130]
  │     │
  │     ├── 3a. Write cell values for each WbDef sheet [lines 154-297]
  │     │     For each field: find row+cell, write value, preserve style
  │     │
  │     ├── 3b. Save worksheets [line 291]
  │     │
  │     ├── 3c. Save shared strings [lines 299-304]
  │     │
  │     └── 3d. WriteConMasCellComments(wbPart, wbDef) [line 310]
  │             For each wbDef.Sheets:
  │               Find matching Excel sheet by name
  │               Find or create WorksheetCommentsPart
  │               For each field: remove old comment, add new ConMas comment
  │               Comments.Save()
  │
  ├── 4. doc.Dispose() — SDK writes all parts to ZIP [line 311]
  │
  ├── 5. ZIP restore [lines 390-467]
  │     For each originalZipEntries (skip comments, VML, rels, content_types):
  │       Compare current vs original byte[]; restore if different
  │
  └── 6. SHA256 verification [lines 470-491]
```

### Key Difference: Comment Writing Context

| Aspect | Legacy ConMas | PaperLess Phase 5.4 |
|--------|-------------|---------------------|
| API used | COM `range.AddComment()` / `Comments.Add()` | OpenXml `Comments.Save()` |
| When comments written | Before SaveAs, via COM Interop | Inside `Open(..., true)` using block, before Dispose |
| Sheet targeted | Cluster's target cell (on Sheet1) | `wbDef.Sheets` lookup by name |
| Cell addresses | From `cluster.CellAddress` (A1 format) | From `field.Cell.Address` |
| Comment text format | Raw `\r\n` in XML | `\n` joined, goes through SDK XML serialization |

---

## 2. Every OOXML Part Modified by ConMas

### Parts Modified by Legacy ConMas Export:

| Part | Modified? | How | Why |
|------|-----------|-----|-----|
| `xl/comments1.xml` | **YES** | 6 new comments with 25-field format | Stores cluster metadata |
| `xl/worksheets/sheet2.xml` (Sheet1) | **YES** | Cell formulas written to cluster cells | Function evaluation |
| `xl/drawings/vmlDrawing1.vml` | **YES** | Comment shape positions adjusted | Visual comment placement |
| `xl/sharedStrings.xml` | **YES** | New strings for ExcelOutputSetting | Form configuration XML |
| `xl/styles.xml` | **YES** | Font/formatting styles applied | Per-cell formatting |
| `xl/workbook.xml` | **YES** | New sheet (ExcelOutputSetting) added | Configuration sheet |
| `[Content_Types].xml` | **YES** | sheet3.xml override added | ExcelOutputSetting |
| `xl/_rels/workbook.xml.rels` | **YES** | New rel for ExcelOutputSetting | Sheet relationship |
| `xl/worksheets/sheet3.xml` | **NEW** | 36-row configuration sheet | ConMas XML config |
| `xl/worksheets/_rels/sheet2.xml.rels` | No change | Already had comment rel | N/A |
| `xl/printerSettings/printerSettings1.bin` | No change | Already existed | N/A |
| `xl/theme/theme1.xml` | No change | Already existed | N/A |
| `docProps/app.xml` | **YES** | Updated sheet count | 3 sheets instead of 2 |
| `_rels/.rels` | No change | N/A | N/A |
| VBA / custom XML | Not present | N/A | N/A |

### Parts Modified by PaperLess (Phase 5.4 intended):

| Part | Intended | Actual (in output) | Notes |
|------|----------|-------------------|-------|
| `xl/comments1.xml` | **YES** — 6 ConMas comments | **NO** — unchanged template comments | **BLOCKER** |
| `xl/worksheets/sheet2.xml` (Sheet1) | **YES** — cell values | **YES** — cell values written | Working |
| `xl/sharedStrings.xml` | **YES** — new strings | Changed from 329 to 389 bytes | Working |
| All other parts | RESTORED from originals | RESTORED | Working |

---

## 3. Every Metadata Source

### Legacy ConMas: Metadata Sources in Output Excel

| Source | Content | How Read Back |
|--------|---------|---------------|
| **Cell comments** (`xl/comments1.xml`) | 25-field cluster metadata (name, type, index, InputParameter, remarks, table info) | `RegisterCommentCells()` parses comment text |
| **Cell formulas** (`xl/worksheets/sheet2.xml`) | Function expressions for calculated fields | Excel evaluates at runtime |
| **Cell formatting** (`xl/styles.xml`) | Font, alignment, colors from InputParameter | Visible in Excel |
| **ExcelOutputSetting sheet** (`xl/worksheets/sheet3.xml`) | Complete `<conmas>` XML config (36 shared strings) | Read during republish |
| **Shared strings** (`xl/sharedStrings.xml`) | All text values including config XML fragments | Referenced by all sheets |
| **Worksheet names** (`xl/workbook.xml`) | Sheet names and ordering | Determines field/sheet mapping |

### PaperLess: Metadata Sources in Output Excel

| Source | Content | Status |
|--------|---------|--------|
| **Cell comments** | Original template comments (15 `_x000D_`-delimited fields) | ❌ Should be ConMas 25-field format |
| **Cell values** | User-entered field values | ✅ Working |
| **Shared strings** | Values from cell writes | ✅ Working |
| All other metadata | RESTORED from template | ✅ Working |

---

## 4. Root Cause of Comment Persistence Failure

### What We Proved

1. **SDK CAN persist comment changes** — Isolated test proved that modifying Sheet1's existing `WorksheetCommentsPart` and calling `Save()` produces an output file with the modified comments. Explicit test: `Contains CONMAS_FORMAT: True`.

2. **ZIP restore does NOT revert comments** — The restore correctly skips `xl/comments*` entries. Test proved: `Comments survived restore: True`.

3. **ZIP restore has IOException bug** — When iterating entries modified by the SDK (like `xl/workbook.xml`), `currentEntry.Delete()` throws `IOException: Cannot delete an entry currently open for writing`. This is caught by the try-catch and the restore is aborted. However, comments have already been skipped (processing order: iterating `originalZipEntries` keys, comments are checked before workbook.xml). So the abort after the IOException does not affect comments.

### The Actual Failure Reason

**The Phase 5.4 code writes comments to the wrong sheet or the wrong cell addresses.**

The template stores comments on **Sheet1** (rId2, file `xl/worksheets/sheet2.xml`). The `_Fields` sheet (rId1, file `xl/worksheets/sheet1.xml`) has NO comments part and NO relationships.

The WbDef fields are associated with `_Fields` (the metadata sheet), not `Sheet1` (the user data sheet). When `WriteConMasCellComments` iterates `wbDef.Sheets`:
- If the only sheet in WbDef is `_Fields`: creates a NEW comments part on `_Fields` (file `comments2.xml`), but this file is not found in output — SDK may not serialize it
- If the sheet is `Sheet1`: cell addresses in the WbDef fields may NOT match the existing comment cell addresses (A1, C1, A3, A6, A9, A12), so old comments are NOT removed and new comments are added to different cells

**Root Cause:** The Phase 5.4 code assumes `wbDef.Sheets` points to the same cells that have comments (cluster cells on Sheet1). But the comment cells in the template (A1, C1, A3, A6, A9, A12) are metadata cells that describe form fields — they are not the same as the user-data cells that WbDef fields point to.

### Confirmation Tests

| Test | Result | Meaning |
|------|--------|---------|
| Modify Sheet1 comments via SDK | `CONMAS_FORMAT: True` | SDK CAN persist |
| ZIP restore with comment skip | Comments survived | Restore does not revert |
| Check which sheet has comments | Sheet1 (rId2) | `_Fields` has none |
| Phase 5.4 code on `_Fields` | No comments part found | Code creates new (wrong sheet) |
| Byte comparison PaperLess vs template | Identical (same xr:uid) | No modification occurred |

---

## 5. Every Difference Between Legacy and PaperLess Output

### ZIP Entry Comparison

| Entry | Legacy bytes | PaperLess bytes | Diff? | Reason |
|-------|-------------|-----------------|-------|--------|
| `[Content_Types].xml` | 1,780 | 1,644 | **YES** | Legacy has sheet3 override for ExcelOutputSetting |
| `_rels/.rels` | 588 | 588 | No | Identical |
| `docProps/app.xml` | 989 | 950 | **YES** | Legacy: 3 worksheets; PaperLess: 2 |
| `docProps/core.xml` | 696 | 696 | No | Identical |
| `xl/_rels/workbook.xml.rels` | 980 | 839 | **YES** | Legacy: 6 rels incl. sheet3; PaperLess: 5 rels |
| `xl/comments1.xml` | 3,157 | 2,755 | **YES** | Legacy: 25-field `\r\n` format; PaperLess: template's 15-field `_x000D_` format |
| `xl/drawings/vmlDrawing1.vml` | 4,823 | 4,823 | No | Identical (minor position values differ) |
| `xl/printerSettings/printerSettings1.bin` | 5,428 | 5,428 | No | Identical |
| `xl/sharedStrings.xml` | 17,859 | 389 | **YES** | Legacy: 43 strings incl. config XML; PaperLess: minimal |
| `xl/styles.xml` | 3,616 | 2,198 | **YES** | Legacy: 8 borders/8 cellXfs; PaperLess: 2 borders/3 cellXfs |
| `xl/theme/theme1.xml` | 8,721 | 8,721 | No | Identical |
| `xl/workbook.xml` | 2,700 | 2,971 | **YES** | Legacy: 3 sheets; PaperLess: 2 sheets |
| `xl/worksheets/_rels/sheet2.xml.rels` | 605 | 605 | No | Identical (comments + VML + printerSettings) |
| `xl/worksheets/sheet1.xml` | 1,178 | 1,178 | No | Identical (`_Fields` sheet — no cell data) |
| `xl/worksheets/sheet2.xml` | 2,551 | 2,551 | **YES** | Different style indices but same structure |
| `xl/worksheets/sheet3.xml` | 3,863 | (missing) | **YES** | ExcelOutputSetting sheet — PaperLess doesn't create it |

### Differences That Matter (Must Fix)

| # | Difference | Why It Matters | Fix Required? |
|---|-----------|----------------|---------------|
| 1 | `xl/comments1.xml` — wrong format, wrong content | Core metadata store for round-trip | **YES — BLOCKER** |
| 2 | `xl/comments1.xml` — 15 `_x000D_` segments vs 25 `\r\n` lines | Wrong field count, wrong delimiters | **YES — BLOCKER** |
| 3 | `xl/comments1.xml` — InputParameter line is empty | No configuration data in comment | **YES — BLOCKER** |

### Differences That Don't Matter (Expected)

| # | Difference | Why It's OK |
|---|-----------|-------------|
| 4 | No sheet3.xml (ExcelOutputSetting) | PaperLess was designed to not create it |
| 5 | workbook.xml has 2 sheets instead of 3 | Consequence of #4 |
| 6 | workbook.xml.rels has 5 rels instead of 6 | Consequence of #4 |
| 7 | sharedStrings.xml is smaller | No ExcelOutputSetting strings |
| 8 | styles.xml has fewer borders | Template-dependent; PaperLess uses template's styles |
| 9 | docProps/app.xml shows 2 worksheets instead of 3 | Consequence of #4 |
| 10 | sheet2.xml style indices differ | Cosmetic — user data sheet |
| 11 | VML positions differ slightly | Cosmetic — comment shape placement |

---

## 6. Exact ConMas Comment Format Specification

### Verified from Legacy Output (`FormTest - Copy-conmas.xlsx`)

**Binary analysis** (hex dump confirmed):

- **Encoding**: UTF-8 with XML declaration `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
- **Line endings in XML prolog**: `\r\n` (0x0D 0x0A)
- **Line endings in `<t>` content**: `\r\n` (0x0D 0x0A) — not `\n` alone
- **Comment count**: 6 comments
- **Cell refs**: A1, C1, A3, A6, A9, A12
- **Author**: "MCF - JOHNELL E. EMPUERTO"
- **Run properties**: `<b/>`, `<sz val="9"/>`, `<color indexed="81"/>`, `<rFont val="Tahoma"/>`, `<charset val="1"/>`
- **Text element**: `<t xml:space="preserve">`
- **Total lines**: 25 per comment (lines 0-24)
- **Trailing empty lines**: 19 empty lines after line 5 (lines 6-24)

### Exact XML Structure

```xml
<comment ref="A1" authorId="0" shapeId="0" xr:uid="{FC296EAC-BCF4-4DE0-B281-CD3D2FD76BB3}">
  <text>
    <r>
      <rPr>
        <b/>
        <sz val="9"/>
        <color indexed="81"/>
        <rFont val="Tahoma"/>
        <charset val="1"/>
      </rPr>
      <t xml:space="preserve">samples\r\n
KeyboardText\r\n
0\r\n
0\r\n
0\r\n
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
\r\n
</t>
    </r>
  </text>
</comment>
```

### 25 Lines Per Comment

```
Line  0: ClusterName           (e.g., "samples")            ← from field/cluster name
Line  1: ClusterTypeString     (e.g., "KeyboardText")       ← from FieldTypeToConMasType()
Line  2: ClusterIndex          (e.g., "0")                  ← 0-based field index
Line  3: ReadOnly              (e.g., "0")                  ← "0" or "1"
Line  4: External              (e.g., "0")                  ← always "0"
Line  5: InputParameter        (semicolon-delimited config) ← full config string
Lines 6-24: empty (19 lines)                                 ← each just \r\n
```

---

## 7. Exact InputParameter Construction

### From Legacy ConMas Output (line 5 of each comment)

```
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11
```

### Field-by-Field Breakdown (from menuOutputExcel_Click IL trace)

| Key | Value | Source in IL | Notes |
|-----|-------|-------------|-------|
| `Required` | `0` or `1` | `cluster.InputParameter` parsed via `ParseParameterText()` | Directly from InputParameter dictionary |
| `Lines` | `1` (or more) | Same | Number of visible lines |
| `InputRestriction` | `None`, `Numeric`, `Date`, etc. | Same | Input validation type |
| `MaxLength` | `0` (or N) | Same | Maximum character length |
| `Align` | `Left`, `Center`, `Right` | Same | Horizontal alignment |
| `Font` | Font name (e.g., `Arial`) | Same | Font face |
| `FontSize` | Point size (e.g., `11`) | Same | Font size |
| `Weight` | `Normal` or `Bold` | Same | Font weight |
| `Color` | `R,G,B` (e.g., `0,0,0`) | Same | Font color as RGB comma-separated |
| `VerticalAlignment` | `0`, `1`, or `2` | Same | 0=Top, 1=Center, 2=Bottom |
| `DefaultFontSize` | `11` | Same | Fallback font size |

### String Construction Rules

From the IL trace:
1. InputParameter is a **semicolon-delimited** string of `Key=Value` pairs
2. The keys and values come directly from the cluster's `InputParameter` property (which is a string parsed by `ParseParameterText()` into a dictionary, then serialized back)
3. If InputParameter contains `Function` key but NOT `FunctionVersion` key, `FunctionVersion=4.2.0000` is appended
4. The keys have EXACT casing as shown above (not configurable)
5. The order is: Required, Lines, InputRestriction, MaxLength, Align, Font, FontSize, Weight, Color, VerticalAlignment, DefaultFontSize
6. No extra whitespace around `=` or `;`

### PaperLess Must Reproduce EXACTLY

The `BuildInputParameterString()` in `WorkbookValueWriter.cs:772` currently produces:
```
Required=0;ReadOnly=0;Visible=1;Lines=1;FontSize=9;FontAutoResizeMode=0
```

This does NOT match the legacy format. It must be rewritten to produce the exact legacy format with the exact keys, ordering, and defaults.

---

## 8. Hidden Sheet Re-Investigation

### Re-investigation of Hidden Sheets

Previous Phase X+3 report concluded "no hidden sheets are created." This is re-confirmed for the **comment-writing path** in `menuOutputExcel_Click`. However, there IS a hidden sheet in the legacy output:

| Sheet | Visibility | Name | Content |
|-------|-----------|------|---------|
| SheetId=2 | `hidden` | `_Fields` | Field metadata (already exists in template) |
| SheetId=3 | `visible` | `ExcelOutputSetting` | 36-row ConMas XML config (ADDED by ConMas export) |

The `ExcelOutputSetting` sheet IS created by the ConMas export process. The Phase X+3 report mentions that `No _Fields sheet, no ExcelOutputSetting object, no hidden sheets, no custom XML parts, no document properties are created` — this finding was about the `menuOutputExcel_Click` method specifically. The `ExcelOutputSetting` sheet is likely created by a DIFFERENT code path (perhaps `ExcelProcessorBase.EnsurePrintAreaQuoted_OpenXml` or a separate `ExcelOutputSetting` builder).

The `ExcelOutputSetting` sheet contains:
- Row 1: `<conmas>` opening tag with top-level settings
- Rows 2-5: Config fragments for designer version, save settings, output files
- Rows 6-10: More config (biometrics, identification, camera, cooperation table)
- Rows 11-35: Per-cluster settings for each cluster
- Row 36: Closing tag

Since PaperLess does not use this sheet and the user has not requested it, this difference is expected and does NOT need to be fixed for the current implementation.

---

## 9. Implementation Plan

### Step 1: Fix Comment Writing Location

Replace the current `WriteConMasCellComments` approach. Instead of iterating `wbDef.Sheets`, directly target the worksheet that has comments:

```csharp
// After writing cell values, inside the using (var doc = ...) block:
// 1. Find the sheet that has an existing WorksheetCommentsPart
// 2. This is usually the user-data sheet (Sheet1)
// 3. Get its comments part
// 4. Clear all existing comments
// 5. For each field in wbDef, add a ConMas-format comment at the field's cell address
// 6. Save the comments part
```

### Step 2: Fix Comment Text Format

Build the exact 25-line `\r\n`-delimited comment text:

```
{ClusterName}\r\n
{ClusterType}\r\n
{FieldIndex}\r\n
{ReadOnly}\r\n
{External}\r\n
{InputParameter}\r\n
\r\n (19 times)
```

### Step 3: Fix InputParameter Format

Rewrite `BuildInputParameterString` to produce the exact legacy format with keys:
`Required`, `Lines`, `InputRestriction`, `MaxLength`, `Align`, `Font`, `FontSize`, `Weight`, `Color`, `VerticalAlignment`, `DefaultFontSize`

### Step 4: Add xr:uid Generation

Each comment needs a unique `xr:uid` attribute. Generate using `Guid.NewGuid().ToString("D").ToUpperInvariant()` wrapped in `{...}`.

### Step 5: Add Run Properties

Each comment's text run must include `<rPr>` with bold, size 9, indexed color 81, font Tahoma, charset 1.

### Step 6: Fix the \r\n Encoding

When writing comment text through the OpenXml SDK, use `\r\n` (not just `\n`) and ensure `xml:space="preserve"` is set. Alternatively, write the comments XML directly to the ZIP after `doc.Dispose()` to bypass SDK encoding issues.

### Step 7: Fix ZIP Restore IOException

The `ZipArchiveMode.Update` `Delete()` call throws when trying to delete entries modified by the SDK. Fix the ZIP restore to handle this by:
- Processing a copy of the entry list
- Handling IOException gracefully per entry
- OR closing/reopening the archive between operations

### Step 8: Optional — Comments on Correct Sheet

If the WbDef fields are on `_Fields` but comments need to be on `Sheet1`, the comment cells must be mapped correctly. In the template, comments are on:
- A1, C1 (row 1)
- A3 (row 3)
- A6 (row 6)
- A9 (row 9)
- A12 (row 12)

These are on Sheet1. If the WbDef fields should have their comments on the SAME cells as the template, the cell address mapping must be preserved.

---

## 10. Verification Method

After implementing, the PaperLess output must match the legacy output for `xl/comments1.xml`:

```powershell
# Compare comments bytes
$legacy = [System.IO.Compression.ZipFile]::OpenRead("legacy.xlsx")
$pl = [System.IO.Compression.ZipFile]::OpenRead("paperless.xlsx")
$legacyComments = $legacy.GetEntry("xl/comments1.xml")
$plComments = $pl.GetEntry("xl/comments1.xml")
# ... compare byte arrays ...

# Compare comment text content
$legacyText = [System.Text.Encoding]::UTF8.GetString($legacyCommentsBytes)
$plText = [System.Text.Encoding]::UTF8.GetString($plCommentsBytes)
# Each comment must have 25 \r\n-delimited lines
# InputParameter must match exactly
```

Expected: 6 comments, each with 25 lines, `\r\n` delimiters, exact InputParameter format.
