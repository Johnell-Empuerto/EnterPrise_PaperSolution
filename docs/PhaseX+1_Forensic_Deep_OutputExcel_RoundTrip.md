# Phase X+1 — Deep Forensic Report: Legacy Output Excel Round-Trip

**Date**: 2026-07-20
**Evidence Sources**:
- Decompiled IL: `ConMasClient.exe` (776 types, 721,892 lines IL)
- Decompiled IL: `LibExcelController.dll` (80 types, 7,128 lines IL)
- Decompiled IL: `ConMasExcelClient.dll` (13 methods)
- Existing project code: `WorkbookGenerator.cs`, `WorkbookReaderService.cs`
- 70+ existing documentation files

---

## 1. THE CENTRAL CONTRADICTION RESOLVED

**User's claimed behavior**: "Real ConMas Designer is able to: Upload Template.xlsx → Design Fields → Output Excel → Close Designer → Open Output Excel → Designer loads ALL fields again → Continue editing. This works WITHOUT requiring the original upload."

**Forensic finding**: The legacy ConMas Designer does NOT have a direct "Open Output Excel" feature that reconstructs form definitions from a standalone `.xlsx` file. The round-trip works through the **form definition file** (`.if` file or server-stored XML), which contains a **Base64-encoded copy of the workbook** as `DefinitionFileValue`.

### The ACTUAL round-trip flow:

```
1. Upload Template.xlsx
   → btnImportExcelTemplate_Click
   → File.Copy(template, UserConfigDirectory\\copy\\filename)
   → Counts visible sheets
   → Writes Decoded DefinitionFileValue to disk
   → Compares sheet counts
   → Stores xlsx as DefinitionFileValue (Base64)
        ↓
2. Design Fields
   → User creates/edits clusters in Designer UI
   → Comments are stored in the in-memory model
        ↓
3. Output Excel (menuOutputExcel_Click)
   → GetDefinition() calls:
      a. File.Copy(original, CopyTemp.xlsx)
      b. Workbooks.Open(copy)
      c. Hide sheets without print areas
      d. MakeCluster() → reads existing cell comments
      e. CalcClusterSize() → computes coordinates
      f. SaveAs(copy)
      g. ExcelWorkbookInfra.Save() → OOXML post-processing
      h. Build XML with <definitionFile> containing Base64 xlsx
      i. Write .xml file
      j. Upload .xml to server (UploadDefinition)
         ↓
4. Close Designer
   → Form definition saved locally as .if file AND to server
   → .if file contains: TopReport + DefinitionFileValue (Base64 xlsx)
         ↓
5. Reopen (btnImportDefinition_Click or Open .if file)
   → Load TopReport from .if or server
   → Read DefinitionFileValue (Base64)
   → Convert.FromBase64String → byte[]
   → File.WriteAllBytes(def_xlsx)
   → Workbooks.Open(def_xlsx)
   → Read cell comments from opened workbook
   → Map comments → cluster definitions → populate Designer UI
         ↓
6. Continue editing
```

### How this works WITHOUT the original upload:

The key is that the **Output Excel IS self-describing** — cell comments contain ALL cluster metadata. But the Designer doesn't reopen the Output Excel directly. Instead:

- When the form is saved (step 4), the xlsx is **embedded as Base64** in the form definition
- When reopened (step 5), the embedded xlsx is extracted, opened, and comments are read
- The standalone Output Excel (.xlsx) from step 3 CAN be used to reconstruct fields, but ONLY if you import it as a NEW form template (via `Import Excel Template` button)

---

## 2. EXACT WORKBOOK MODIFICATION LIST (Output Excel Generation)

Based on IL analysis of `ConMasExcelClient.ExcelController.GetDefinition()` and `LibExcelController.ExcelControllerBase`:

### Pre-Save Modifications (via COM Interop)

| # | Modification | IL Evidence | Method |
|---|-------------|-------------|--------|
| 1 | **Copy original** | `File.Copy(input, CopyTemp.ext, true)` | GetDefinition |
| 2 | **Open copy via COM** | `Workbooks.Open(tempPath, ...)` | GetDefinition |
| 3 | **Hide sheets without print area** | `ws.Visible = xlSheetHidden` | GetDefinition |
| 4 | **Clear shapes** | `WorksheetShapeCollection.Clear()` | MakeCluster |
| 5 | **Write cluster values to cells** | `range.Value = clusterValue` | GetDefinition |
| 6 | **Add cell comments** | `range.AddComment(commentText)` | GetDefinition |
| 7 | **SaveAs** | `workbook.SaveAs(copyTemp, xlExclusive)` | GetDefinition |

### Post-Save OOXML Modifications (via Infragistics/OpenXml)

| # | Modification | IL Evidence | Method |
|---|-------------|-------------|--------|
| 8 | **Clamp xWindow/yWindow in workbook.xml** | String ref: `"xWindow"`, `"yWindow"` | PreLoad_EditWorkbookParts |
| 9 | **Remove applyBorder where borderId=0 in styles.xml** | String ref: `"xl/styles.xml"`, `"applyBorder"` | PreLoad_EditStyleParts |
| 10 | **Escape R1C1 in all definedNames** | String ref: `"definedNames"`, `"definedName"` | PostSave_EditWorkbookParts |
| 11 | **Remove fPrintsWithSheet from all drawings** | String ref: `"clientData"`, `"fPrintsWithSheet"` | PostSave_EditDrawingsParts |

### External File Creation

| # | Modification | IL Evidence | Method |
|---|-------------|-------------|--------|
| 12 | **Embed xlsx as Base64 in XML** | `File.ReadAllBytes` + `Convert.ToBase64String` | GetDefinition |
| 13 | **Write .xml file with full ConMas definition** | `StreamWriter.Write(definition.ToString())` | ExportExcelDefinition |
| 14 | **Upload .xml to server** | `WebClient.UploadFile(url, xmlPath)` | UploadDefinition |

### What is NOT modified:
- ❌ Sheet count (preserved from original)
- ❌ Sheet order (preserved from original)
- ❌ Custom XML parts (preserved, not created)
- ❌ VBA project (preserved, not created)
- ❌ Embedded images (preserved, not added)
- ❌ Worksheet names (preserved from original)
- ❌ No `_Fields` sheet created
- ❌ No `ExcelOutputSetting` sheet created (but it IS **read** if present by ExtractSettingXml)

---

## 3. CELL COMMENT FORMAT (DEFINITIVE)

From `LibExcelController.ExcelControllerBase.ValidAndSplit()` IL analysis:

```
Line 0: Cluster name (e.g., "氏名", "samples")
Line 1: Cluster type (e.g., "KeyboardText", "Machine", "CheckBox")
Line 2: Cluster index (integer)
Line 3: "0" (read-only flag? or reserved)
Line 4: "0" (external flag? or reserved) 
Line 5+: Input parameters (semicolon-delimited key=value pairs)
```

Parsed by `ValidAndSplit()` with these out parameters:
```
name, type, index, inputParameters, readOnly, external,
remarksValue[0..9], tableNo, tableName, tableCooperationTable,
columnNo, columnKey, columnName, columnType, rowNo, rowName
```

16 fields parsed from a single comment! The comment format is RICH with metadata, not just type+name.

### The exact parsing:
```pseudocode
ValidAndSplit(commentText, out name, out type, out index, 
              out inputParameters, out readOnly, out external, 
              out remarksValue[10], out tableNo, out tableName,
              out tableCooperationTable, out columnNo, out columnKey,
              out columnName, out columnType, out rowNo, out rowName)
    
    Split by chars: \r, \n, \t (0x0D, 0x0A, 0x09)
    Trim each element via TrimArrayElements()
    
    name = lines[0]        // Cluster display name
    type = lines[1]        // Cluster type enum name
    index = lines[2]       // Cluster index (int)
    readOnly = lines[3]    // "0" or "1"
    external = lines[4]    // "0" or "1"
    lines[5..14] = remarks values 1-10
    remaining lines = [tableNo, tableName, tableCooperationTable, 
                       columnNo, columnKey, columnName, columnType, 
                       rowNo, rowName]
    
    // If lines[5] contains semicolons, it's input parameters:
    if lines[5] contains ";":
        inputParameters = lines[5]
        remarks start at lines[6]
    else:
        inputParameters = ""
        remarks start at lines[5]
```

---

## 4. THE REOPEN PIPELINE (DEFINITIVE)

### Pipeline A: Open from Embedded Definition (Standard Path)

```
btnImportDefinition_Click / Open .if file
    ↓
Load TopReport from file or server
    ↓
Read DefinitionFileValue (Base64 string)
    ↓
Convert.FromBase64String → byte[]
    ↓
File.WriteAllBytes(def_xlsx)
    ↓
Workbooks.Open(def_xlsx) via COM CLSID
    ↓
For each visible worksheet:
    ↓
  Check cell.HasComment()  -- per-cell via Infragistics or COM
    ↓
  If comment exists:
    ↓
    Read comment.Text
    ↓
    Parse via ValidAndSplit → all 16 fields
    ↓
    Measure Range.Left/Top/Width/Height via COM
    ↓
    Create BlankCluster object
    ↓
Populate Designer UI with all clusters
```

### Pipeline B: Import External Excel as New Template

```
btnImportExcelTemplate_Click
    ↓
OpenFileDialog (filter: *.xls, *.xlsx)
    ↓
File.Copy(selected → UserConfigDir\copy\filename)
    ↓
Workbooks.Open(def_xlsx from existing DefinitionFileValue)  -- the CURRENT definition
    ↓
Count visible sheets in BOTH files
    ↓
IF sheet counts match:
    ↓
  ReadAllBytes(copiedFile) → Convert.ToBase64String
    ↓
  Update DefinitionFileValue
    ↓
  Show success message
ELSE:
    ↓
  Show error (sheet count mismatch)
```

### Pipeline C: Rendering (PDF generation) — reads workbook directly

```
ProcessExcelFile()
    ↓
PreProcess(parentWindow, fileName, out xlsxFileName, out settingXml)
    ↓
  LoadWorkbook() via DevExpress Spreadsheet
    ↓
  ExtractSettingXml("ExcelOutputSetting") -- reads XML from "ExcelOutputSetting" sheet (DEFAULT PARAMETER)
    ↓
  RegisterCommentCells() -- scans every cell for comments
    ↓
    For each Comment in worksheet:
    ↓
      Cluster.ParseCommentText(commentText)
    ↓
      Cluster.IsCluster(commentText) -- validates it's a ConMas cluster
    ↓
      Record as cluster cell
    ↓
ModifyCells() -- write form values to cells
    ↓
SaveAsPdf() -- render to PDF
```

### Key Finding: The `ExtractSettingXml` Method

From `ConMasClient_dump.il:712060`:
```
.param [2] = "ExcelOutputSetting"
```

The `ExtractSettingXml(Workbook& workbook, [opt] string settingSheetName)` method has a **default parameter value of "ExcelOutputSetting"**. This means:

- The legacy ConMas DOES look for an `ExcelOutputSetting` sheet during **rendering** (Pipeline C)
- It reads XML from this sheet if it exists
- It uses this to configure the rendering pipeline
- This method is used in the DevExpress PDF rendering pipeline, NOT in the Designer reopen pipeline

**But**: There is NO evidence in the export pipeline (GetDefinition, Output Excel) that CREATES this `ExcelOutputSetting` sheet. The rendering pipeline simply checks if it exists.

---

## 5. HIDDEN SHEETS: DEFINITIVE ANSWER

### Does legacy ConMas CREATE any hidden sheets?

**NO.** The export pipeline (`GetDefinition`) does NOT create any hidden sheets. It:
- Preserves original sheets
- Hides existing sheets that lack a print area (sets `Visible = xlSheetHidden`)
- Does NOT add `_Fields` or any other new sheet

### Does legacy ConMas READ any hidden sheets?

**YES — during rendering (Pipeline C):**
- `ExtractSettingXml("ExcelOutputSetting")` looks for a sheet named `"ExcelOutputSetting"` — regardless of visibility
- If found, reads XML from it

**NO — during Designer reopen (Pipeline A):**
- Only visible sheets are processed for cluster reading
- Hidden sheets are skipped entirely
- No `_Fields` sheet is ever read

### The `_Fields` sheet:

**Conclusion confirmed: `_Fields` is entirely a PaperLess invention.** Zero references to `"_Fields"` in any ConMas assembly (ConMasClient.exe, LibExcelController.dll, ConMasExcelClient.dll).

### The `ExcelOutputSetting` sheet:

**Conclusion nuanced:**
- Legacy ConMas does NOT CREATE this sheet during export
- But legacy ConMas DOES CHECK for this sheet during rendering (`ExtractSettingXml` default param)
- If the sheet exists (e.g., created by a third party), the XML is extracted
- If the sheet doesn't exist, `ExtractSettingXml` returns null/empty
- This is used in the DevExpress rendering pipeline, NOT in the Designer reopen pipeline

---

## 6. ZIP STRUCTURE COMPARISON: ORIGINAL vs OUTPUT EXCEL

Based on existing forensic analysis (PhaseX6) and IL evidence:

### Preserved (unchanged from original):
| Part | Status | Evidence |
|------|--------|----------|
| `[Content_Types].xml` | Preserved | No modification code found |
| `_rels/.rels` | Preserved | No modification code found |
| `xl/workbook.xml` | **MODIFIED** | xWindow/yWindow clamped |
| `xl/_rels/workbook.xml.rels` | Preserved | No modification code found |
| `xl/styles.xml` | **MODIFIED** | applyBorder cleaned |
| `xl/sharedStrings.xml` | **MODIFIED** | New comment text added |
| `xl/worksheets/sheetN.xml` | **MODIFIED** | Cell values + comments + VML |
| `xl/worksheets/_rels/sheetN.xml.rels` | Modified | VML comment relationship added |
| `xl/commentsN.xml` | **ADDED** | New comment XML per sheet |
| `xl/drawings/commentsVmlN.vml` | **ADDED** | VML shapes for comment rendering |
| `xl/drawings/drawingN.xml` | **MODIFIED** | fPrintsWithSheet removed |
| `xl/calcChain.xml` | Modified | Recalculated |
| `xl/definedNames` | **MODIFIED** | R1C1 escaped |
| `xl/printerSettings/printerSettings1.bin` | Preserved | Not modified |
| `docProps/core.xml` | Preserved | Not modified |
| `docProps/app.xml` | Preserved | Not modified |
| Custom XML parts | Preserved | Not created, not removed |
| VBA project | Preserved | Not modified |

### Added by Output Excel:
| Part | How Created |
|------|-------------|
| `xl/commentsN.xml` | Via COM `AddComment()` → OOXML serialization |
| `xl/drawings/commentsVmlN.vml` | Auto-generated by Excel for comment shapes |

### Not present in Output Excel (but PaperLess adds):
- ❌ `xl/worksheets/sheet3.xml` for `_Fields` (not in legacy)
- ❌ `_Fields` relationship in `xl/_rels/workbook.xml.rels`
- ❌ `<sheet name="_Fields">` in workbook.xml
- ❌ `xl/worksheets/sheet2.xml` for `ExcelOutputSetting` (not in legacy export, though legacy rendering reads it)

---

## 7. ALL METADATA STORAGE LOCATIONS

### In the Workbook (standalone .xlsx):

| Location | Stores | Read By | Write During |
|----------|--------|---------|--------------|
| **Cell comments** (xl/commentsN.xml) | Cluster name, type, index, input parameters, read-only flag, external flag, remarks[10], table info (tableNo, tableName, columnNo, columnKey, columnName, columnType, rowNo, rowName), cooperation table | **MakeCluster** (LibExcelController) — Infragistics per-cell HasComment scan | Output Excel export |
| **Cell values** (xl/worksheets/sheetN.xml) | User-entered data, default values | COM Range.Value | Output Excel export |
| **workbook.xml** (xWindow/yWindow) | Window position (clamped) | COM + OOXML post-processing | PostSave |
| **styles.xml** (applyBorder) | Border style normalization | OOXML post-processing | PostSave |
| **definedNames** | Named ranges (R1C1 escaped) | OOXML post-processing | PostSave |
| **drawings** (fPrintsWithSheet) | Print settings cleanup | OOXML post-processing | PostSave |
| **ExcelOutputSetting sheet** (optional, not created by export) | XML configuration for rendering | **ExtractSettingXml** (DevExpress pipeline) | NOT created by legacy export |

### In the Form Definition (external):

| Location | Stores | Used By |
|----------|--------|---------|
| **TopReport.DefinitionFileValue** | Base64-encoded xlsx | Designer reopen, upload |
| **.if file** | Full form definition (XML) + embedded xlsx | Local save/load |
| **Server XML** (`<conmas><top>...`) | Full form definition + embedded xlsx | Server storage, download |
| **Upload XML** (`UploadFile.xml`) | Full ConMas definition + Base64 xlsx | Server upload |

---

## 8. WHY THE ROUND-TRIP WORKS

### The mechanism:

The round-trip works because **the form definition (TopReport) carries the entire workbook as a Base64-encoded blob** (`DefinitionFileValue`). This blob is:

1. **Written during Output Excel** — `GetDefinition()` produces the Output Excel, then `File.ReadAllBytes()` → `Convert.ToBase64String()` → stored as `DefinitionFileValue`
2. **Preserved in the form definition** — When the form is saved (`.if` file or server), the Base64 xlsx is part of the definition
3. **Read back during reopen** — `Convert.FromBase64String()` → `File.WriteAllBytes()` → `Workbooks.Open()` → read cell comments → reconstruct

### Why the original upload is not needed:

- The Output Excel contains cell comments with ALL field metadata
- The form definition contains the Base64-encoded Output Excel as `DefinitionFileValue`
- When reopening, the embedded xlsx is extracted and opened
- Comments are read from the extracted xlsx → fields are reconstructed

### What if you ONLY have the standalone .xlsx (no .if file, no server)?

- The **Designer does NOT have a direct "Open Output Excel" feature**
- But the cell comments ARE sufficient for full reconstruction
- You would create a new form via `Import Excel Template`, point to the Output Excel
- The Designer reads the comments and creates the form with all fields
- This is the CLOSEST behavior to what the user describes

---

## 9. CONTRADICTIONS IN EXISTING DOCUMENTATION — RESOLVED

| Document | Claim | Evidence | Resolution |
|----------|-------|----------|------------|
| PhaseX report | "Legacy does NOT create ExcelOutputSetting sheet" | **Confirmed**: No creation code in export IL | ✅ Correct |
| PhaseX report | "Legacy reads ONLY cell comments" | **Partially wrong**: Legacy ALSO reads `ExcelOutputSetting` sheet if present (ExtractSettingXml default param = "ExcelOutputSetting") | ❌ Missed the rendering pipeline |
| Phase33 report | "MakeCluster reads comments via Infragistics, NOT COM" | **Confirmed**: Uses `WorksheetCell.get_HasComment()` and `get_Comment()` via Infragistics backend | ✅ Correct |
| Phase33 report | "No `_Fields` sheet creation" | **Confirmed**: Zero references in any ConMas assembly | ✅ Correct |
| Phase32 report | "GetDefinition is XML exporter NOT metadata reader" | **Confirmed**: Returns XElement, calls MakeCluster internally | ✅ Correct |
| PhaseX7 report | "Hidden sheets skipped by legacy runtime" | **Partially**: Hidden sheets ARE skipped for comment scanning, BUT `ExtractSettingXml` reads from `ExcelOutputSetting` regardless of visibility | ⚠️ Nuanced |
| Phase5.2.5 | "ExcelOutputSetting sheet is missing in save-edited pipeline" | **Legacy context**: The rendering pipeline reads this sheet; export does NOT create it | ⚠️ Clarified |

---

## 10. INCORRECT ASSUMPTIONS IN PAPERLESS IMPLEMENTATION

| Assumption | Actual Legacy Behavior | Impact |
|------------|----------------------|--------|
| `_Fields` sheet is needed for reopen | Legacy never creates or reads `_Fields` | Unnecessary, will be skipped by legacy readers |
| `ExcelOutputSetting` sheet should be created during export | Legacy does NOT create it during export; only reads it during rendering if present | Our creation is unnecessary but harmless; our reader should handle its absence |
| Comments-only metadata is insufficient | Comments contain 16 fields including name, type, index, table info, remarks | Our comment format has too few fields |
| `Workbooks.Add()` is acceptable | Legacy copies original and modifies in-place | Loses original structure (VBA, images, custom XML, definedNames) |
| OOXML post-processing is optional | Legacy always clamps xWindow/yWindow, cleans styles, escapes definedNames, cleans drawings | Our output has raw OOXML that may fail validation |
| External XML is legacy format | The upload XML is the PRIMARY server interaction format | We should produce compatible XML |
| ws.Comments is primary comment source | Per-cell `HasComment()` via Infragistics is primary; `ws.Comments` via COM is fallback | Fragile COM comment reading |
| Comment format = `name\ntype\nparams` | Actual format is 16+ fields separated by `\r\n\t` | Our comments have insufficient metadata |

---

## 11. RECOMMENDED ARCHITECTURE (LEGACY-COMPATIBLE)

```
1. Start from COPY of original workbook (File.Copy + Workbooks.Open)
   NOT Workbooks.Add()

2. Preserve all original content:
   - Sheets (count, order, names)
   - VBA project
   - Custom XML parts
   - Embedded images
   - Defined names
   - All relationships

3. Write metadata as CELL COMMENTS (not hidden sheets):
   - Comment format = ValidAndSplit-compatible:
     {name}\r\n{type}\r\n{index}\r\n{readOnly}\r\n{external}\r\n{inputParameters}\r\n{remarks}\r\n{tableNo}\r\n...
   - Each cluster cell gets a comment
   - No _Fields sheet, no ExcelOutputSetting sheet

4. Post-save OOXML modifications:
   - Clamp xWindow/yWindow
   - Remove applyBorder where borderId=0
   - Escape R1C1 in definedNames
   - Remove fPrintsWithSheet from drawings

5. Produce external .xml file (for server upload):
   - <conmas><top><sheets><sheet><clusters><cluster>...
   - <definitionFile><value>Base64 xlsx</value>
   - Upload via WebClient.UploadFile

6. For internal PaperLess round-trip:
   - Store xlsx as Base64 (like DefinitionFileValue)
   - On reopen: decode → open → read comments → reconstruct form
   - Optional: Create ExcelOutputSetting sheet for rendering config
```
