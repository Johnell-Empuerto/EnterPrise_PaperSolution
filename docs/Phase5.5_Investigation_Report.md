# Phase 5.5 Investigation Report
## Workbook Comparison: Legacy ConMas Output vs PaperLess Output (edited 2)

**Date:** 2026-07-20
**Files Compared:**
- Legacy: `~/Downloads/FormTest - Copy-conmas.xlsx` (16,784 bytes, 16 ZIP entries)
- PaperLess: `~/Downloads/FormTest - Copy_edited (2).xlsx` (14,361 bytes, 15 ZIP entries)
- Original Template: `~/Documents/FormTest - Copy.xlsx` (16,782 bytes, 15 ZIP entries)

---

## 1. Critical Finding: Phase 5.4 Comment Writing DOES NOT Persist

**`xl/comments1.xml` in PaperLess output is byte-for-byte identical to the original template (2,755 bytes).**
- Same `_x000D_` carriage-return XML encoding (from original template)
- Same `xr:uid` GUID values (proves no SDK modification occurred)
- Same 15 `_x000D_`-separated segments (not the ConMas 25-line `\n` format)

**Root Cause:** `WriteConMasCellComments()` at `WorkbookValueWriter.cs:590` creates/modifies `commentsPart.Comments` and calls `.Save()` at line 692, but the modifications are NOT persisted to the output ZIP. The OpenXml SDK's `Comments.Save()` inside a `SpreadsheetDocument.Open(...)` `using` block does not correctly flush changes to the package when `doc.Dispose()` runs.

**Evidence:**
| File | Comments byte length | First `xr:uid` | Line separator |
|------|---------------------|-----------------|----------------|
| Original Template | 2,755 | `{1F59C8FD-...}` | `_x000D_` |
| PaperLess edited 2 | 2,755 | `{1F59C8FD-...}` | `_x000D_` |
| PaperLess edited 1 | 2,755 | `{1F59C8FD-...}` | `_x000D_` |
| PaperLess edited 3 | 2,755 | (same) | `_x000D_` |
| Legacy ConMas | 3,157 | `{FC296EAC-...}` | `\n` |

All three PaperLess outputs have identical comments. The Phase 5.4 code ran but the modifications never made it into the ZIP.

---

## 2. Legacy ConMas Comment Format (Target)

6 comments on the `_Fields` sheet (A1, C1, A3, A6, A9, A12) — each exactly 25 lines with `\n` delimiter:

```
Line  0: ClusterName       (e.g., "samples")
Line  1: ClusterTypeString (e.g., "KeyboardText")
Line  2: FieldIndex        (e.g., "0", "1", "2"...)
Line  3: ReadOnly          ("0" or "1")
Line  4: External          ("0")
Line  5: InputParameter    (semicolon-delimited config string)
Lines 6-24: empty lines (19 lines, each just `\n`)
```

**Key difference in line 5:** Legacy ConMas `InputParameter`:
```
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;
FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11
```

PaperLess's `BuildInputParameterString()` produces a DIFFERENT format:
```
Required=0;ReadOnly=0;Visible=1;Lines=1;FontSize=9;FontAutoResizeMode=0
```

**The InputParameter format must be EXACTLY what ConMas produces** to be round-trip compatible.

---

## 3. Structural Differences (Expected / Not Required for Fix)

| Aspect | Legacy ConMas | PaperLess | Expected? |
|--------|-------------|-----------|-----------|
| **Sheet count** | 3 (`_Fields`, `Sheet1`, `ExcelOutputSetting`) | 2 (`_Fields`, `Sheet1`) | Yes — PaperLess removed `ExcelOutputSetting` (Phase 5.4) |
| **Content_Types** | Has `sheet3.xml` override | No `sheet3.xml` | Yes — mirrors sheet count |
| **workbook.xml.rels** | 6 rels (rId1-6, sheet3 included) | 5 rels (rId1-5, no sheet3) | Yes — mirrors sheet count |
| **workbook.xml** | 3 sheet elements | 2 sheet elements | Yes |
| **docProps/app.xml** | 3 worksheets + 1 named range | 2 worksheets (implicit) | Yes |
| **sharedStrings.xml** | 43 strings incl. ConMas XML config | ~6 strings | Yes — no ExcelOutputSetting |
| **Styles** | 8 borders, 8 cellXfs | 2 borders, 3 cellXfs | Template-dependent |
| **sheet2.xml styles** | Different style indices per cell | Uniform style index 2 | Template-dependent |

All structural differences are CONSEQUENCES of PaperLess not creating `ExcelOutputSetting`/`_Fields`. They do NOT need to be fixed.

---

## 4. VML Drawing Comparison

Legacy and PaperLess VML files are structurally identical:
- Same 6 shapes (`_x0000_s1025` through `_x0000_s1030`)
- Same shape IDs
- Same `visibility:hidden`
- Minor positioning differences (margin-top: `1.2pt` vs `1.8pt`, anchor values vary)
- **No VML fix needed** — the existing VML works

---

## 5. Root Cause Analysis: Why `WriteConMasCellComments` Doesn't Persist

The `using (var doc = SpreadsheetDocument.Open(outputPath, true))` block at `WorkbookValueWriter.cs:130`:

1. SDK loads all parts from the ZIP into memory
2. Phase 5.4 code modifies `commentsPart.Comments` (in-memory object)
3. `commentsPart.Comments.Save()` is called — SDK writes to internal stream
4. `doc.Dispose()` at end of `using` block serializes all parts back to ZIP
5. ZIP restore loop runs — **skips** comments (`isComments = true` → skip restore)

**Why it fails:** The SDK's `Save()` on a pre-existing `WorksheetCommentsPart` does not properly flush to the underlying package stream. When `doc.Dispose()` serializes, it may use a stale cached copy. Also, the ZIP restore correctly skips comments (intending to preserve SDK's version), but the SDK's version is the UNMODIFIED original.

---

## 6. Fix Plan: Phase 5.5 Implementation

### 6.1 Direct ZIP Injection of Comments (Bypass SDK)

Replace the call to `WriteConMasCellComments(wbPart, wbDef)` with a post-dispose ZIP injection approach:

**Step A:** After `doc.Dispose()` (after the `using` block at line 311), re-open the output ZIP and directly write `xl/comments1.xml` with the ConMas format XML.

**Step B:** Update `xl/worksheets/_rels/sheet2.xml.rels` in the ZIP to ensure the comments relationship is intact (it already is — rId3 points to `../comments1.xml`).

**Step C:** Update `[Content_Types].xml` if needed (it already has the comments override).

### 6.2 Exact Comment XML Format

The comments XML must match the ConMas format exactly:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  mc:Ignorable="xr"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <authors>
    <author>MCF - JOHNELL E. EMPUERTO</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0" shapeId="0"
      xr:uid="{FC296EAC-BCF4-4DE0-B281-CD3D2FD76BB3}">
      <text>
        <r>
          <rPr>
            <b/>
            <sz val="9"/>
            <color indexed="81"/>
            <rFont val="Tahoma"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">samples
KeyboardText
0
0
0
Required=0;Lines=1;InputRestriction=None;MaxLength=0;Align=Center;Font=Arial;FontSize=11;Weight=Normal;Color=0,0,0;VerticalAlignment=2;DefaultFontSize=11












</t>
        </r>
      </text>
    </comment>
    <!-- ... 5 more comments ... -->
  </commentList>
</comments>
```

### 6.3 InputParameter String Construction

Must match the EXACT legacy ConMas format from `menuOutputExcel_Click`. Based on Phase X+3 analysis, the format is:

```
Required=<0/1>;Lines=<N>;InputRestriction=<None|Number|...>;MaxLength=<N>;Align=<Left|Center|Right>;Font=<FontName>;FontSize=<Pt>;Weight=<Normal|Bold>;Color=<R,G,B>;VerticalAlignment=<N>;DefaultFontSize=<Pt>
```

The `BuildInputParameterString()` in `WorkbookValueWriter.cs:772` must be rewritten to produce this exact format.

### 6.4 `_Fields` Sheet Style Preservation

The legacy ConMas output preserves the original template's styles for the `_Fields` sheet. The PaperLess output currently flattens style indices (all use style 2). This is because the SDK modifies the `_Fields` sheet when writing cell values, and the ZIP restore restores original styles for all non-modified entries.

**Fix:** The ZIP restore already skips `xl/worksheets/sheet2.xml` (the `_Fields` sheet). If the `_Fields` sheet has cells with values that are written by the WbDef pipeline, those changes persist. The style issue is cosmetic and does not affect functionality.

---

## 7. Implementation Checklist

1. **Direct ZIP injection of `xl/comments1.xml`** — write raw XML with ConMas format after `doc.Dispose()`
2. **Update ZIP skip list** — remove comments from skip list (since we'll inject directly after dispose)
3. **Generate new GUIDs** for `xr:uid` on each comment (or omit — optional, Excel regenerates on open)
4. **Rewrite `BuildInputParameterString`** to match legacy ConMas format exactly
5. **Verify** `xl/worksheets/_rels/sheet2.xml.rels` has comments relationship (already correct)
6. **Verify** `xl/drawings/vmlDrawing1.vml` has shapes matching comment count (already correct from template)
7. **Verify** `[Content_Types].xml` has comments override (already correct)
8. **Build, test** with `FormTest.xlsx` template, compare to legacy ConMas output

---

## 8. Key Risk: Newline Encoding

The ConMas comments use RAW `\n` (0x0A) in the XML `<t>` element. When writing directly to the ZIP via `ZipArchive`, the `\n` characters will be preserved as-is (unlike the OpenXml SDK which encodes `\r` as `_x000D_`).

**Solution:** Use `\n` only (no `\r\n`) in the comment text. The `xml:space="preserve"` attribute ensures whitespace is preserved.

---

## 9. Verification Method

After implementing the fix, compare the PaperLess output to the legacy ConMas output at the ZIP entry level:

```
Compare-Object (Get-ZipEntries legacy.xlsx) (Get-ZipEntries paperless.xlsx)
```

The only expected differences should be:
- Missing `sheet3.xml` (PaperLess doesn't create `ExcelOutputSetting`)
- Different `docProps/app.xml` (different application)
- Different `xr:uid` values (new GUIDs)
- Different `workbook.xml` (fewer sheets)

The comment content, VML content, worksheet rels, and content types should all be structurally equivalent.

---

## 10. Summary

| # | Issue | Severity | Fix Required? |
|---|-------|----------|---------------|
| 1 | Comments not written at all (SDK `Save()` doesn't persist) | **BLOCKER** | Yes — direct ZIP injection |
| 2 | `InputParameter` format doesn't match legacy | **BLOCKER** | Yes — rewrite `BuildInputParameterString` |
| 3 | Missing `ExcelOutputSetting` sheet | None | No — expected removal |
| 4 | Style indices differ on `_Fields` sheet | Low | No — cosmetic |
| 5 | VML positioning slightly different | Low | No — cosmetic |
| 6 | `_Fields` sheet cell values differ | None | Will be resolved when comments work correctly |
