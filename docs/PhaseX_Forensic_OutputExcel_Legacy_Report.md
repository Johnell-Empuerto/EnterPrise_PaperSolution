# Phase X — Forensic Report: Legacy ConMas Output Excel

**Date**: 2026-07-19
**Type**: Investigation only — no implementation
**Sources**: Decompiled IL (ConMasExcelClient, ConMasClient, LibExcelController, Cimtops.ConMasBiz), existing project code (WorkbookGenerator, WorkbookReaderService), 70+ existing docs

---

## 1. Does ConMas Modify the Original or Generate New?

**Finding: ConMas modifies a COPY of the original workbook. It does NOT generate from scratch.**

Evidence from `ConMasExcelClient.il`, `GetDefinition` method (lines 862–2483):

```
Line 946:  ldstr "CopyTemp"                  // temp filename
Line 948:  Path.Combine(_outputPath, "CopyTemp" + extension)
Line 957:  File.Copy(_inputExcelPath, tempPath, true)  // COPY ORIGINAL
Line 997:  Workbooks.Open(tempPath, ...)               // OPEN THE COPY
```

The full sequence:
1. `File.Copy(inputExcel, outputPath + "CopyTemp" + extension, true)` — **copy original**
2. `Workbooks.Open(copyTempPath)` — **open the copy**
3. Hide sheets without print areas
4. `SaveAs(copyTempPath, xlExclusive)` — **save back to copy**
5. Process clusters, coordinates, cell values
6. `File.Copy(outputDir, ...)` — **copy to final output path**
7. Open second instance of Excel on the output copy for additional processing

**Our current `WorkbookGenerator` does something fundamentally different** — it calls `Workbooks.Add(XlWBATemplate.xlWBATWorksheet)` to create a blank workbook, then builds everything from scratch. This loses all original workbook structure, formatting, metadata, VBA, defined names, and OOXML parts that ConMas preserves by copying.

**Impact**: Our generated workbook has only what we explicitly write. The legacy workbook retains everything from the original template plus modifications.

---

## 2. Exactly Where Does ConMas Store Designer Metadata?

**Finding: Cell comments are the PRIMARY metadata source. The `_Fields` and `ExcelOutputSetting` sheets are PaperLess inventions — they do NOT exist in the legacy ConMas workbook.**

### Evidence from decompiled ConMas

In `ConMasExcelClient.il`, `MakeCluster` method (lines 3320–5366):

```
Line 3616: callvirt instance string ExcelWorksheetBase::GetCommentText()
Line 3617-3624: Trims comment by chars {0x0D, 0x0A, 0x09} (\r\n\t)
Line 3630-3631: Log "get cluster cell info."
Lines 3684-3719: Calls ValidAndSplit to parse comment into cluster fields
Lines 5010-5366: Creates <cluster> element from parsed data
```

The comment text format (from `ValidAndSplit`):
```
Line 0: cluster name
Line 1: cluster type (e.g., "KeyboardText")
Line 2: cluster index
Line 3: "0"
Line 4: "0"
Line 5: input parameters (semicolon-delimited key=value)
```

### What legacy ConMas does NOT use:

| Storage | Used by Legacy? | Evidence |
|---------|----------------|----------|
| **Cell comments** | **YES — PRIMARY** | `GetCommentText()` called on every cell in MakeCluster |
| **`_Fields` hidden sheet** | **NO** | No `_Fields` string found in any decompiled IL. This is a PaperLess invention. |
| **`ExcelOutputSetting` sheet** | **NO** | No `ExcelOutputSetting` sheet creation found in legacy IL. This is a PaperLess invention. |
| **Custom XML parts** | **NO** | Not found in any legacy IL |
| **VBA project** | **NO** | No vbaproject.bin created by legacy Output Excel |
| **workbook.xml (defined names)** | **YES — POST-PROCESS** | `PostSave_EditWorkbookParts` escapes R1C1 in existing defined names |
| **drawings XML** | **YES — POST-PROCESS** | `PostSave_EditDrawingsParts` removes `fPrintsWithSheet` |
| **styles.xml** | **YES — POST-PROCESS** | `PreLoad_EditStyleParts` removes `applyBorder` where borderId=0 |
| **External XML file** | **YES — PRIMARY** | `ExportExcelDefinition` writes `.xml` file with full ConMas definition including `<definitionFile>` containing base64-encoded xlsx |

### Summary of metadata mechanisms in legacy ConMas:

| Mechanism | Read or Write? | Purpose |
|-----------|---------------|---------|
| Cell comments | **Read** → extract cluster definitions | PRIMARY metadata for rebuild |
| workbook.xml | **Write** → clamp xWindow/yWindow | Fix window position after editing |
| styles.xml | **Write** → clean applyBorder | Normalize border styles |
| workbook.xml definedNames | **Write** → escape R1C1 | Normalize named range formulas |
| drawings.xml | **Write** → remove fPrintsWithSheet | Clean print settings from drawings |
| External .xml file | **Write** → full ConMas definition | Upload to server for DB storage |
| Base64 xlsx in XML | **Write** → embed in `<definitionFile>` | Full workbook stored in server |

---

## 3. What Is `ExcelOutputSetting`?

**Finding: `ExcelOutputSetting` is an artifact of our PaperLess implementation, not part of the legacy ConMas workbook format.**

In our code (WorkbookGenerator.cs line 991–1029):
- Creates a visible worksheet named `"ExcelOutputSetting"` (WS[3])
- Writes 36 rows of ConMas XML fragments to column A

In the legacy IL (`ConMasExcelClient.il`, `ConMasClient.il`):
- **Zero references** to creating or writing an `ExcelOutputSetting` worksheet
- The only `"ExcelOutputSetting"` references in `ConMasClient.il` (lines 493921–494197) are in a method called `ChangeExcelOutputSetting` which updates `ExcelOutputValue` on cluster models — this is a **property on a ClusterDefinition object**, not a worksheet

The IL evidence at ConMasClient.il line 493939:
```
ChangeExcelOutputSetting(Nullable<int32> Value)
```
This modifies `BlankCluster.ExcelOutputValue` on the selected cluster. This is a **per-cluster property** stored in the runtime model, not a serialized worksheet.

**Conclusion**: The XML content we write to the `ExcelOutputSetting` sheet IS structurally correct (the `<conmas>` XML matches what legacy ConMas would produce), but the legacy system stores this XML in the uploaded `.xml` file and in the `<definitionFile>` element — not in a visible worksheet. The `ExcelOutputSetting` sheet is a PaperLess storage mechanism that works but is not the legacy approach.

---

## 4. What Is `_Fields`?

**Finding: `_Fields` is entirely a PaperLess invention. The legacy ConMas does not create or use a `_Fields` hidden sheet.**

Evidence from decompiled IL:
- **Zero matches** for `"_Fields"` string in `ConMasExcelClient.il`, `ConMasClient.il`, or `LibExcelController.il`
- The legacy runtime reads metadata exclusively from:
  1. **Cell comments** (primary — `GetCommentText()`)
  2. **workbook structure** (sheets, print area, geometry via COM)
  3. **Server response XML** (from `RequestDefinition` / `GetDefinitionData` API call)

Our `WorkbookReaderService.cs` (line 119–124) treats `_Fields` as the primary metadata source:
```csharp
// Step 2: ReadFieldsSheet(fieldsSheet) — primary metadata source
// Step 3: ReadComments(ws) — fallback metadata
```

But the IL evidence shows legacy ConMas does the **opposite**: comments are primary, and there is no `_Fields`.

### Why this matters for re-import

When our generated workbook is uploaded back to the legacy ConMas server, the server-side re-import pipeline (`DefinitionBiz.Regist`) reads from:
- `HttpRequest.Files[0].InputStream` → parses as XML → `DefinitionBiz.Entry()` which reads from the XML elements (`<sheets>`, `<sheet>`, `<clusters>`, `<cluster>`)
- It does NOT open the workbook as an Excel file. It reads the uploaded XML.

But the **ConMas Designer** (the WPF client) does NOT re-upload a workbook to reconstruct fields. It re-downloads the definition XML from the server via `RequestDefinition(topId)` (which calls `GetDefinitionData`). The XML returned by the server contains all cluster definitions.

### The runtime republish pipeline (legacy ConMas)

When you upload an output Excel back to ConMas:

1. Designer calls `RequestDefinition(topId)` → gets XML with cluster definitions
2. Designer opens the workbook via COM
3. Designer reads cell comments for cluster metadata
4. Designer maps comments → XML definition → populates the designer UI

**Our `_Fields` sheet is only used by our own `WorkbookReaderService`** to reconstruct the `FormDefinition` when doing a round-trip within our PaperLess system. The legacy ConMas never needs it because it reads comments directly.

---

## 5. Legacy Re-Import Pipeline (Exact Order)

When ConMas reopens an exported workbook to reconstruct the designer, it follows this pipeline (reconstructed from `ConMasClient.il` and `ConMasExcelClient.il`):

```
1. Open workbook via COM (Workbooks.Open)
      ↓
2. For each visible worksheet:
      ↓
3.   Get PrintArea (PageSetup.PrintArea)
      ↓
4.   Get UsedRange
      ↓
5.   For each cell in range:
      ↓
6.     Check for CommentText (cell.GetCommentText())
        ↓
7.     Split comment by \r\n\t characters
        ↓
8.     Parse into cluster fields (ValidAndSplit):
          - Line 0: cluster name
          - Line 1: cluster type
          - Line 5+: input parameters (key=value)
        ↓
9.     Measure cell geometry (Range.Left, .Top, .Width, .Height)
        ↓
10.    Create cluster definition from parsed + measured data
        ↓
11. Build sheets → build form → populate designer UI
```

**Key observation**: Step 2 processes only `xlSheetVisible` sheets. The `_Fields` sheet (which we create as `xlSheetHidden`) would be skipped. `xlSheetVeryHidden` would also be skipped.

**The server-side re-import** (`DefinitionBiz.Regist`) is different — it reads from the uploaded XML file, not from the workbook:
```
1. HttpRequest.Files[0].InputStream → XElement.Load
2. ValidateReportTable(xInputReport)
3. Entry(nextRevNo)
4.   InsertTop() → def_top table
5.   InsertSheet() → def_sheet table
6.   BuildCluster() → def_cluster table
7.   Store xlsx BLOB from <definitionFile><value>
```

---

## 6. What Creates the Second Sheet?

**Finding: The legacy ConMas does NOT create a second metadata sheet. The workbook structure remains the same as the original template, with modifications only to cell values, comments, and OOXML internals.**

In our PaperLess `WorkbookGenerator.cs`, `EnsureSheets()` creates exactly 3 sheets:
- WS[1] → `"_Fields"` (hidden)
- WS[2] → content sheet (user data)
- WS[3] → `"ExcelOutputSetting"` (visible)

In the legacy IL (`ConMasExcelClient.il`, `GetDefinition`):
- **No sheet creation** other than the originals from the copied workbook
- Sheets without print areas are **hidden** (set `Visible = xlSheetHidden`), not removed
- No `_Fields` sheet is created
- No `ExcelOutputSetting` sheet is created

**Our call chain** (paperless):
```
FormController.Save()
  → FormSaveService.SaveAsync()
    → WorkbookGenerator.Generate()
      → EnsureSheets()  ← PaperLess invention, 3 sheets
```

**Legacy call chain** (ConMas):
```
menuOutputExcel_Click
  → ExportExcelDefinition()
    → GetDefinition()
      → File.Copy(original) ← start from template
      → Workbooks.Open(copy)
      → Hide unmatched sheets
      → MakeCluster() ← read comments, build XML
      → CalcClusterSize() ← compute coordinates
      → SaveAs
      → ExcelWorkbookInfra.Save() ← OOXML modifications
        → PreLoad_EditWorkbookParts (xWindow/yWindow fix)
        → PreLoad_EditStyleParts (applyBorder cleanup)
        → PostSave_EditWorkbookParts (definedName R1C1 escape)
        → PostSave_EditDrawingsParts (fPrintsWithSheet removal)
      → Build <definitionFile> with base64 xlsx
      → Write .xml file (full ConMas definition)
```

---

## 7. Structural Differences: Original vs ConMas Output Workbook

Based on forensic comparison from `PhaseX6_Legacy_OutputExcel_Complete_Reverse_Engineering.md` and decompiled IL:

| Element | Original | ConMas Output | Our Output |
|---------|----------|---------------|------------|
| **Sheet count** | N | N (same) | **3** (always) |
| **Hidden sheets** | Any | Sheets without print area hidden | `_Fields` (hidden), `ExcelOutputSetting` (visible) |
| **Cell comments** | Any user comments | **Metadata comments added per cluster** | Metadata comments added |
| **Cell values** | User-entered | User-entered (preserved) | User-entered (written from FormDefinition) |
| **workbook.xml** | Original | xWindow/yWindow clamped to valid range | Unchanged from OpenXml SDK output |
| **styles.xml** | Original | `applyBorder` removed where borderId=0 | Unchanged |
| **definedNames** | Original | All R1C1 references escaped | Only `_xlnm.Print_Area` added |
| **drawings XML** | Original | `fPrintsWithSheet` removed from all clientData | Unchanged |
| **Custom XML parts** | Original | Preserved (not modified) | **Stripped** (new workbook has none) |
| **VBA project** | Original | Preserved (not modified) | **Stripped** (new workbook has none) |
| **External .xml file** | N/A | Created: full ConMas definition | Not created (our API returns JSON) |
| **Base64 xlsx** | N/A | Embedded in XML `<definitionFile>` | Not embedded (stored in DB via generator) |
| **Embedded images** | Original | Preserved | **Stripped** (new workbook has none) |
| **Shared strings** | Original | Modified (new comments + values) | Rebuilt from scratch |
| `.rels` structure | Original | Modified (1-2 more relationships) | Rebuilt from scratch |

**Critical structural difference**: Our workbook is a **new OOXML package** with `_Fields` and `ExcelOutputSetting` sheets. The legacy workbook is the **original OOXML package** with only targeted modifications. This means:
- Our workbook **loses custom XML parts, VBA, embedded images, and any non-standard OOXML content**
- Our workbook **adds two new sheets** that legacy ConMas doesn't expect
- Our workbook **sheet order** is: `_Fields`, Content, `ExcelOutputSetting` — legacy has original sheet order

---

## 8. Existing Documentation Summary

The repository has extensive documentation about Output Excel. Here is what each key document covers:

| Document | Phase | Focus | Findings |
|----------|-------|-------|----------|
| `OutputExcelOfForm_ReverseEngineering.md` | Early | Legacy reverse engineering | Correctly identified: comments are metadata, workbook is modified copy NOT new. But our implementation went against this finding. |
| `PhaseX6_Legacy_OutputExcel_Complete_Reverse_Engineering.md` | X.6 | Full format spec | Documents 11 structural differences. Identified that our workbook has `_Fields` and `ExcelOutputSetting` sheets that legacy doesn't. |
| `PhaseX5_Forensic_Comparison_Report.md` | X.5 | Bug hunting | Root cause: our comment format was wrong. Fixed but fundamental mismatch remains. |
| `PhaseX7_Runtime_First_Forensic_Report.md` | X.7 | Upload flow | Legacy runtime skips hidden sheets entirely. `_Fields` is never read by legacy. |
| `Phase31_MetadataReverseEngineeringReport.md` | 31 | Metadata storage | Documented all possible metadata locations. `_Fields` was identified as OUR addition. |
| `Phase32_GetDefinition_ReverseEngineering.md` | 32 | IL analysis | Mapped the `GetDefinition` IL to C#. Proved CopyTemp pattern. |
| `Phase5.2.5_Forensic_Missing_Configuration_Sheet.md` | 5.2.5 | Root cause | `ExcelOutputSetting` sheet exists in our output but `save-edited` pipeline bypasses `WorkbookGenerator` completely. |
| `CalcClusterSize_ReverseEngineering.md` | Calc | Coordinate engine | Detailed pixel-scanning algorithm for coordinate extraction. |
| `PhaseX43_DesignerProducingServerRenderingAssets.md` | X.43 | Upload protocol | How the `<conmas>` XML + `<definitionFile>` base64 xlsx are uploaded to the server. |

### Critical Contradictions in Our Documentation

1. **`OutputExcelOfForm_ReverseEngineering.md`** (section 3) correctly states: "The workbook is a **copy** of the original template, not a new workbook." But our `WorkbookGenerator` uses `Workbooks.Add()`.

2. **`Phase31_MetadataReverseEngineeringReport.md`** correctly states: "Cell comments are the PRIMARY metadata source." But our `WorkbookReaderService` prioritizes `_Fields` over comments.

3. **`PhaseX7_Runtime_First_Forensic_Report.md`** correctly states: "Hidden sheets are SKIPPED by the legacy runtime." But we continue to create and depend on `_Fields`.

---

## 9. Conclusions

### What We Know for Certain

1. **Legacy ConMas never creates a new workbook from scratch.** It copies the original, opens the copy, modifies it in-place, and performs post-save OOXML cleanup.

2. **Cell comments are the sole metadata mechanism in the workbook.** The comment text format is: `name\ntype\nindex\n0\n0\nparams...`. There is no `_Fields` sheet, no `ExcelOutputSetting` sheet, no custom XML parts.

3. **The workbook structure is preserved exactly.** Sheet count, order, VBA, images, custom XML parts, defined names — all remain as-is from the original. ConMas only:
   - Adds cell comments (cluster metadata)
   - Writes cell values (user data)
   - Clamps xWindow/yWindow (workbook.xml)
   - Cleans `applyBorder` (styles.xml)
   - Escapes R1C1 in definedNames (workbook.xml)
   - Removes `fPrintsWithSheet` (drawings.xml)

4. **The `.xml` file (external) is the primary upload artifact**, not the workbook. The workbook is Base64-embedded inside the XML as `<definitionFile><value>`. The server stores the XML, extracts the xlsx BLOB, and serves it back on request.

5. **`_Fields` and `ExcelOutputSetting` are PaperLess-only inventions.** They work within our closed system but are not expected by legacy ConMas.

### What We Don't Know

1. **Does the legacy ConMas runtime re-import use the base64-embedded xlsx or re-download from the server?** The `RequestDefinition` API returns XML with cluster data but the `<definitionFile><value>` is large. The exact client behavior needs more IL analysis.

2. **How does the legacy ConMas Designer handle the re-open of an output Excel?** We know it reads comments. But does it also read the external XML? Or does it only use COM + comments?

3. **Does the ConMas server validate the workbook structure on re-upload?** The `DefinitionBiz.Regist` only validates the XML. Does any server code also open and verify the base64 xlsx?

### Recommendations for Phase X+1 (Implementation)

Based on these findings:

1. **Start from the original workbook.** Replace `Workbooks.Add()` with `File.Copy(original, temp)` + `Workbooks.Open(temp)`. This preserves all OOXML structure.

2. **Remove `_Fields` sheet creation.** It's not needed by legacy. If we need it for internal PaperLess round-trips, make it optional and disabled by default.

3. **Remove `ExcelOutputSetting` sheet creation.** Same reasoning.

4. **Focus on cell comments as the sole workbook metadata mechanism.** Match the legacy format exactly: `name\ntype\nindex\n0\n0\nparams...`

5. **Add OOXML post-processing:**
   - Clamp xWindow/yWindow in workbook.xml
   - Remove `applyBorder` where borderId=0 in styles.xml
   - Escape R1C1 references in all definedNames
   - Remove `fPrintsWithSheet` from all drawings

6. **Produce the external `.xml` file** with the full ConMas definition, including the base64-embedded xlsx. This is what the legacy system actually uploads.

7. **Preserve existing OOXML content:** custom XML parts, VBA, embedded images, relationships.
