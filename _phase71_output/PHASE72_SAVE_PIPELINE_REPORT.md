# Phase 7.2 — Reverse Engineer Save Pipeline Differences

**Date:** July 20, 2026
**Status:** Investigation Only
**DO NOT MODIFY ANY CODE**

---

## Part 1 — Complete PaperLess Save Pipeline (save-edited)

### Entry Point: `POST /api/form/save-edited`

```
Controller              File                               Method
─────────               ────                               ──────
FormController.cs       Controllers/FormController.cs      SaveEdited()
  │
  ├─ Session Resolution
  │   SessionWorkbookStore.cs   Application/SessionWorkbookStore.cs   ResolveWorkbookPath()
  │     └─ Resolves TempWorkbooks/{sessionId}/original.xlsx
  │
  ├─ Core Save Pipeline
  │   FormSaveService.cs        Application/FormSaveService.cs        SaveEditedValuesAsync()
  │     │
  │     ├─ File.Copy(sourcePath → outputPath)
  │     │
  │     ├─ WorkbookValueWriter.WriteValues(wbDef, sourcePath, outputPath)
  │     │   Application/WorkbookValueWriter.cs
  │     │   │
  │     │   ├─ [1] File.Copy(source → output)
  │     │   │     System.IO.File.Copy(line 52)
  │     │   │     → creates output.xlsx as a copy of original
  │     │   │
  │     │   ├─ [2] ZipFile.OpenRead → Save ALL original ZIP entries
  │     │   │     Dictionary<entryName, byte[]> (lines 64-82)
  │     │   │     → captures every ZIP entry BEFORE SDK touches them
  │     │   │
  │     │   ├─ [3] SpreadsheetDocument.Open(output, true)  (line 128)
  │     │   │     DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
  │     │   │     → SDK opens existing workbook read-write
  │     │   │     → SDK MAY mutate workbook.xml, styles.xml, theme, rels, content types
  │     │   │
  │     │   ├─ [4] For each sheet with fields:
  │     │   │     ├─ Resolve sheet name → WorksheetPart
  │     │   │     ├─ Ensure SharedStringTable exists (line 175-180)
  │     │   │     │   └─ wbPart.AddNewPart<SharedStringTablePart>() if missing
  │     │   │     ├─ For each field with value:
  │     │   │     │   ├─ Parse cellRef, rowIndex
  │     │   │     │   ├─ Find or Create Row (line 197-204)
  │     │   │     │   ├─ Find or Create Cell (line 207-227)
  │     │   │     │   ├─ Preserve StyleIndex (line 234)
  │     │   │     │   ├─ Write value (Formula / Number / SharedString) (line 238-268)
  │     │   │     │   └─ row.Append(cell) if new (line 271-274)
  │     │   │     ├─ wsPart.Worksheet.Save() (line 289)
  │     │   │     └─ sst.SharedStringTable.Save() (line 300)
  │     │   │
  │     │   ├─ [5] SpreadsheetDocument.Dispose (SDK flushes to disk)
  │     │   │
  │     │   ├─ [6] ZIP RESTORE: ZipFile.Open(Update mode) (line 394)
  │     │   │     For each entry in originalZipEntries:
  │     │   │       ├─ SKIP (intentionally modified):
  │     │   │       │   - xl/worksheets/sheet*.xml  (cell values changed)
  │     │   │       │   - xl/sharedStrings.xml       (new strings appended)
  │     │   │       │   - xl/comments*                (may be written)
  │     │   │       │   - xl/drawings/vmldrawing*    (VML drawings)
  │     │   │       │   - worksheet _rels/*           (worksheet relationships)
  │     │   │       │   - [Content_Types].xml        (content types)
  │     │   │       └─ RESTORE (all other entries):
  │     │   │           - Compare current bytes to original bytes
  │     │   │           - If different: delete + recreate with original bytes
  │     │   │           - If same: leave unchanged
  │     │   │     → Restores: workbook.xml, styles.xml, theme, workbook rels, etc.
  │     │   │
  │     │   └─ [7] PostProcessZipForConMas(outputPath, wbDef) (line 500-503)
  │     │         Application/WorkbookValueWriter.cs (lines 828-1320)
  │     │         │
  │     │         ├─ GenerateExcelOutputSettingFragments() → 36 XML strings (line 879)
  │     │         ├─ AppendExcelOutputSettingSharedStrings() → 36 <si> in sst (line 1060)
  │     │         ├─ WriteConMasCommentsEntry() → comments1.xml (line 1101)
  │     │         ├─ CreateExcelOutputSettingSheetEntry() → sheet3.xml (line 1153)
  │     │         ├─ ComputeNextWorkbookRelId() → rId{n+1} (line 1190)
  │     │         ├─ UpdateWorkbookXmlForSheet3() → add <sheet> (line 1217)
  │     │         ├─ UpdateWorkbookRelsForSheet3() → add <Relationship> (line 1261)
  │     │         └─ UpdateContentTypesForSheet3() → add <Override> (line 1295)
  │     │
  │     └─ [8] Auto-Validation
  │         WorkbookDiffValidator.Compare(sourcePath, outputPath)
  │         Application/WorkbookDiffValidator.cs (line 295)
  │         │
  │         ├─ SpreadsheetDocument.Open(original, false) (line 299)
  │         ├─ SpreadsheetDocument.Open(edited, false) (line 300)
  │         ├─ ComparePartHashes() (line 374)
  │         │   → SHA256 of styles.xml, theme, worksheets
  │         ├─ CompareWorkbook() (line 453)
  │         │   → sheets count/names/visibility, protection, VBA
  │         ├─ CompareStyles() (line 512)
  │         │   → fonts, fills, borders, cellFormats, cellStyles
  │         ├─ CompareWorksheet() (line 588)
  │         │   → rows, cols, merges, freeze, print area, page setup, cells
  │         ├─ CompareObjectParts() (line 897)
  │         │   → drawings, images, comments, tables
  │         ├─ CompareAdditionalParts() (line 990)
  │         │   → shared strings, workbook.xml, external links, rels
  │         └─ CompareDefinedNames() (line 1126)
  │             → all DefinedName entries
  │
  └─ Response
      File.ReadAllBytesAsync(result.WorkbookPath)
      → Returns FileContentResult (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
```

### Technology Usage per Method

| Method | File | Excel COM? | OpenXML SDK? | ZIP Manip? |
|--------|------|-----------|-------------|-----------|
| SaveEdited() | FormController.cs | No | No | No |
| ResolveWorkbookPath() | SessionWorkbookStore.cs | No | No | No |
| SaveEditedValuesAsync() | FormSaveService.cs | No | No | No |
| WriteValues() | WorkbookValueWriter.cs | No | Yes (SpreadsheetDocument.Open) | Yes (ZipArchive) |
| PostProcessZipForConMas() | WorkbookValueWriter.cs | No | No | Yes (ZipArchive + XDocument) |
| Compare() | WorkbookDiffValidator.cs | No | Yes (SpreadsheetDocument.Open) | No |

---

## Part 2 — Workbook Rebuild Occurrences

### Complete Table of Workbook Creation/Modification Operations

| Operation | File | Line | Creates New? | Copies? | Opens Existing? | Replaces Parts? | Changes rIDs? |
|-----------|------|------|-------------|---------|----------------|----------------|--------------|
| `Workbooks.Add(xlWBATWorksheet)` | WorkbookGenerator.cs | 148 | **YES** — blank workbook | No | No | N/A | N/A (new doc) |
| `workbook.SaveAs(outputPath)` | WorkbookGenerator.cs | 208 | **YES** — COM serialization | No | Opens COM save | Yes (COM rewrites all) | Yes (COM assigns) |
| `File.Copy(source → output)` | WorkbookValueWriter.cs | 52 | No | **YES** (exact copy) | No | No | No |
| `SpreadsheetDocument.Open(output, true)` | WorkbookValueWriter.cs | 128 | No | No | **YES** (read-write) | May mutate via SDK | May change via SDK |
| `SpreadsheetDocument.Open(path, false)` | WorkbookDiffValidator.cs | 299,300 | No | No | **YES** (read-only) | No | No |
| `ZipFile.Open(Update)` restore | WorkbookValueWriter.cs | 394 | No | No | **YES** (ZIP update) | Restores original | Restores original |
| PostProcess zip updates | WorkbookValueWriter.cs | 828+ | No | No | **YES** (ZIP append) | **YES** (adds sheet3 + comments) | **YES** (adds rId{n+1}) |
| `File.Copy(original → session)` | SessionWorkbookStore.cs | 89 | No | **YES** (session copy) | No | No | No |

### Key Finding: `SpreadsheetDocument.Create` and `new WorkbookPart` are NEVER used

**Zero occurrences** of:
- `SpreadsheetDocument.Create(...)` 
- `new WorkbookPart()` / `new WorksheetPart()`
- `AddWorkbookPart()` / `AddWorksheetPart()`
- `.Clone()` on SpreadsheetDocument

The **ONLY** place a brand-new workbook is created from scratch is the **deprecated** `WorkbookGenerator.cs` using COM `Workbooks.Add()`.

The current canonical path (`WorkbookValueWriter.WriteValues()`) correctly copies the original and edits in-place.

---

## Part 3 — Reverse Engineered Legacy ConMas Save Sequence

### Reconstructed from Decompiled IL (`ConMasExcelClient.il`, `ConMasClient.il`, `LibExcelController.il`)

```
Original.xlsx (input template with _Fields, Sheet1, cell comments, clusters)

    │
    ▼
[1] File.Copy(inputExcel, outputPath + "CopyTemp" + extension, true)
    ────────────────────────────────────────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method (line 946-957)
    └─ "CopyTemp" string literal + Path.Combine
    └─ File.Copy exists in IL
    └─ Source: PhaseX_Forensic_OutputExcel_Legacy_Report.md

    │
    ▼
[2] Workbooks.Open(copyTempPath)
    ────────────────────────────────
    Evidence: ConMasExcelClient.il (line 997)
    └─ Workbooks.Open(tempPath) exists in IL
    └─ Proof: ConMas opens the COPY, not the original

    │
    ▼
[3] Hide sheets without print areas
    ──────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method
    └─ Iterates worksheets, checks PageSetup.PrintArea
    └─ Sets ws.Visible = xlSheetHidden if no print area
    └─ Source: Phase6.2_Reverse_Engineer_Legacy_ConMas_Printable_Worksheet_Selection.md

    │
    ▼
[4] MakeCluster() — Read cell comments → build cluster XML
    ─────────────────────────────────────────────────────────
    Evidence: ConMasExcelClient.il, MakeCluster method (lines 3320-5366)
    └─ worksheet.GetCommentText() — reads existing comments
    └─ ValidAndSplit() — parses 25-line \r\n comment format
    └─ Creates <cluster> XML elements (name, type, index, params)
    └─ Cluster XML is used for coordinate computation, NOT workbook modification

    │
    ▼
[5] Write cell values (user-entered data)
    ──────────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method
    └─ Sets Range.Value = ... via COM
    └─ Values come from user input / database

    │
    ▼
[6] workbook.SaveAs(copyTempPath, XlSaveAsAccessMode.xlExclusive)
    ──────────────────────────────────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method
    └─ SaveAs exists in IL
    └─ xlExclusive = 3
    └─ Source: PhaseX_Forensic_OutputExcel_Legacy_Report.md

    │
    ▼
[7] workbook.Close() + excel.Quit()
    ────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method
    └─ Close(false) — don't save changes (already saved)

    │
    ▼
[8] Post-save OOXML cleanup via OpenXML SDK
    ─────────────────────────────────────────
    Evidence: ConMasClient.il, ExcelWorkbookInfra.Save() method
    └─ Each cleanup is an idempotent XML transformation:
    │
    ├─ [8a] PreLoad_EditWorkbookParts
    │     → Clamp xWindow/yWindow to valid range
    │     → Sets window position to [-10800, -10800] if out of range
    │
    ├─ [8b] PreLoad_EditStyleParts  
    │     → Remove applyBorder where borderId=0
    │     → xf[@applyBorder=1 and borderId=0] → remove applyBorder
    │
    ├─ [8c] PostSave_EditWorkbookParts
    │     → Escape R1C1 in definedNames
    │     → DefinedName containing 'C' or 'R' → escape with ''
    │
    └─ [8d] PostSave_EditDrawingsParts
          → Remove fPrintsWithSheet from drawings
          → clientData/fPrintsWithSheet elements removed

    │
    ▼
[9] Build <definitionFile> with base64-encoded xlsx
    ────────────────────────────────────────────────
    Evidence: ConMasClient.il, GetDefinition method
    └─ File.ReadAllBytes(xlsxPath)
    └─ Convert.ToBase64String(bytes)
    └─ Creates <definitionFile><value>base64...</value></definitionFile>

    │
    ▼
[10] Write .xml file (full ConMas definition)
    ──────────────────────────────────────────
    Evidence: ConMasExcelClient.il
    └─ XElement.Save(outputPath + ".xml")
    └─ The .xml file contains complete form definition + embedded xlsx

    │
    ▼
[11] File.Copy(outputPath, finalOutputPath)
    ────────────────────────────────────────
    Evidence: ConMasExcelClient.il, GetDefinition method (line ~2483+)
    └─ Copies the output workbook to the final destination
    └─ Source: PhaseX_Forensic_OutputExcel_Legacy_Report.md

    │
    ▼
Export1.xlsx  +  Export1.xml
```

### Evidence Sources

All ConMas save pipeline knowledge comes from:
1. **Decompiled IL** — `ConMasExcelClient.dll`, `ConMasClient.dll`, `LibExcelController.dll`
2. **Forensic comparison** — original template vs ConMas output (PhaseX reports)
3. **Multi-generation COM test** — Phase 7.1 confirmed COM open/save behavior
4. **Database extraction** — `def_top.def_file` confirmed stored workbook structure

---

## Part 4 — Binary Preservation Analysis: Legacy ConMas

| Part | Preserved? | Modified? | Recreated? | Evidence |
|------|-----------|----------|-----------|----------|
| **workbook.xml** | Mostly | **Minor edits** (xWindow/yWindow clamp, documentId changes) | No | IL shows PreLoad_EditWorkbookParts; Phase 7.1 COM test shows only documentId + windowPos changes |
| **styles.xml** | Mostly | **Minor edit** (applyBorder removal where borderId=0) | No | IL shows PreLoad_EditStyleParts; Phase 7.1 shows byte-identical after COM cycles |
| **sharedStrings.xml** | **YES** | No (new values use new entries) | No | Phase 7.1 COM test: byte-identical across 4 generations |
| **worksheets** | **YES** | Cell values written; rows/cells preserved | No | IL shows Range.Value = ... via COM; Phase 7.1 shows worksheet XML byte-identical |
| **comments** | **YES** | Comments added (not overwritten) | Possibly added | IL shows GetCommentText() reads; comments may be added by designer, not by Output Excel |
| **VML/drawings** | Mostly | **Minor edit** (fPrintsWithSheet removed) | No | IL shows PostSave_EditDrawingsParts; preserves drawing structure |
| **printer settings** | **YES** | No | No | Phase 7.1 COM test: byte-identical; IL shows no printer setting modification |
| **workbook rels** | **YES** | No (unless new sheets needed) | No | Phase 7.1: byte-identical; ConMas doesn't add sheets |
| **worksheet rels** | **YES** | No | No | Phase 7.1: byte-identical |
| **content types** | **YES** | No | No | Phase 7.1: byte-identical |
| **theme** | **YES** | No | No | Phase 7.1: byte-identical |
| **docProps** | Mostly | **Minor** (modified timestamp only) | No | Phase 7.1: only `dcterms:modified` changes; creator/app unchanged |
| **defined names** | Mostly | **Minor** (R1C1 escape if needed) | No | IL shows PostSave_EditWorkbookParts escapes R1C1 references |
| **sheet count** | **YES** | No sheets added/removed | No | IL shows only hiding sheets, never adding/removing |
| **custom XML** | **YES** | No | No | IL shows no modification; preserved from original |
| **VBA** | **YES** | No | No | IL shows no VBA modification; preserved from original |

**Legacy ConMas modifications are MINIMAL and IDEMPOTENT.** Once applied, re-application produces zero net change (except documentId GUID and timestamp).

---

## Part 5 — Binary Preservation Analysis: PaperLess

| Part | Preserved? | Modified? | Recreated? | Evidence |
|------|-----------|----------|-----------|----------|
| **workbook.xml** | **RESTORED** | ZIP restore reverts SDK changes; then PostProcess adds sheet3 entry | No (copy then restore+modify) | WorkbookValueWriter.cs: pre-saves original bytes, restores them after SDK, then adds sheet3 via PostProcess |
| **styles.xml** | **RESTORED** | ZIP restore reverts all SDK mutations | No | WorkbookValueWriter.cs line 394+: compares and restores original bytes |
| **sharedStrings.xml** | **APPENDED** | **YES** — new strings appended for values + 36 config fragments | No (augmented) | WorkbookValueWriter.cs line 175-180: appends to SharedStringTable; PostProcess line 1060: appends 36 config strings |
| **worksheets** | **MODIFIED** | Cell values written to sheet1.xml, sheet2.xml | No | WorkbookValueWriter.cs lines 153-295: writes cell values via SDK |
| **comments** | **OVERWRITTEN** | **YES** — PostProcess replaces comments1.xml | **YES** (replace) | WorkbookValueWriter.cs line 1101: WriteConMasCommentsEntry() replaces entire comments1.xml |
| **VML/drawings** | **NOT RESTORED** | SDK may mutate; NOT in restore list | No | WorkbookValueWriter.cs line 394+: skips restoration (treated as intentionally modified) |
| **printer settings** | **RESTORED** | Restored from original | No | Restored via ZIP restore if changed |
| **workbook rels** | **MODIFIED** | Restored then PostProcess adds sheet3 rel | No (augmented) | PostProcess line 1261: UpdateWorkbookRelsForSheet3 adds new relationship |
| **worksheet rels** | **NOT RESTORED** | SDK may mutate; NOT in restore list | No | Not in restore list |
| **content types** | **MODIFIED** | NOT restored; PostProcess adds sheet3 override | No (augmented) | PostProcess line 1295: adds Override for sheet3 |
| **theme** | **RESTORED** | Restored from original | No | ZIP restore restores original if SDK changed it |
| **docProps** | **RESTORED** | Restored from original | No | ZIP restore reverts SDK changes |
| **defined names** | **RESTORED** | Restored from original | No | ZIP restore reverts SDK changes |
| **sheet count** | **INCREASED** | PostProcess adds sheet3 (ExcelOutputSetting) | **YES** (adds new) | PostProcess line 1153: creates sheet3.xml |
| **custom XML** | **RESTORED** | Restored from original | No | ZIP restore reverts SDK changes |
| **VBA** | **RESTORED** | Restored from original | No | ZIP restore reverts SDK changes |

### Key Differences from ConMas

| Aspect | ConMas | PaperLess |
|--------|--------|-----------|
| **Workbook origin** | Copy of original via File.Copy | Copy of original via File.Copy (same!) |
| **Cell values** | Written via COM Range.Value | Written via SDK cell manipulation |
| **Shared strings** | COM manages internally | SDK appends new entries |
| **Comments** | Preserved (not touched by Output Excel) | **Overwritten entirely** by PostProcess |
| **Additional sheets** | **Never adds sheets** | **Adds ExcelOutputSetting sheet** |
| **Workbook.xml** | Only xWindow/yWindow clamp | SDK may mutate; restored; then sheet3 entry added |
| **Styles** | Only applyBorder cleanup | SDK may mutate; restored to original |
| **Post-processing** | Idempotent XML cleanups | ConMas-format injection + ZIP manipulation |

---

## Part 6 — Side-by-Side Save Pipeline

### Legacy ConMas
```
Original.xlsx + Cell Values
    │
    ▼
File.Copy(original → temp)
    │
    ▼
Workbooks.Open(temp)          ─── COM opens the COPY
    │
    ├─ Hide sheets without print areas
    ├─ Range.Value = user data  ─── Write cell values only
    │
    ▼
SaveAs(output, xlExclusive)   ─── COM serializes everything
    │
    ▼
Close()
    │
    ▼
Post-Process (XML cleanup):
    ├─ PreLoad_EditWorkbookParts   (xWindow clamp)
    ├─ PreLoad_EditStyleParts      (applyBorder cleanup)
    ├─ PostSave_EditWorkbookParts  (R1C1 escape)
    └─ PostSave_EditDrawingsParts  (fPrintsWithSheet removal)
    │
    ▼
File.Copy(temp → final output)
    │
    ▼
Export1.xlsx
```

### PaperLess (Current Canonical Path)
```
Original.xlsx + WorkbookDefinition (JSON)
    │
    ▼
File.Copy(original → output)    
    │
    ▼
ZipFile.OpenRead → Save ALL original entries
    │
    ▼
SpreadsheetDocument.Open(output, true)  ─── SDK opens the COPY
    │
    ├─ For each field with value:
    │     ├─ Find or Create Row
    │     ├─ Find or Create Cell
    │     ├─ Write CellValue (Formula/Number/SharedString)
    │     └─ wsPart.Worksheet.Save()
    │
    ▼
SpreadsheetDocument.Dispose    ─── SDK flushes changes to ZIP
    │
    ▼
ZIP RESTORE: ZipFile.Open(Update)
    ├─ For each entry:
    │     ├─ Skip: worksheets, sharedStrings, comments, VML, rels, contentTypes
    │     └─ Restore: workbook.xml, styles.xml, theme, workbook rels, etc.
    │
    ▼
PostProcessZipForConMas:
    ├─ GenerateExcelOutputSettingFragments()
    ├─ AppendExcelOutputSettingSharedStrings()    ─── Adds 36 strings to sharedStrings.xml
    ├─ WriteConMasCommentsEntry()                 ─── OVERWRITES comments1.xml
    ├─ CreateExcelOutputSettingSheetEntry()       ─── CREATES sheet3.xml
    ├─ ComputeNextWorkbookRelId()
    ├─ UpdateWorkbookXmlForSheet3()               ─── MODIFIES workbook.xml
    ├─ UpdateWorkbookRelsForSheet3()              ─── MODIFIES workbook.xml.rels
    └─ UpdateContentTypesForSheet3()              ─── MODIFIES [Content_Types].xml
    │
    ▼
Auto-Validation:
    Compare(source, output)
    ├─ ComparePartHashes()
    ├─ CompareWorkbook()
    ├─ CompareStyles()
    ├─ CompareWorksheet()
    ├─ CompareObjectParts()
    ├─ CompareAdditionalParts()
    └─ CompareDefinedNames()
    │
    ▼
Export1.xlsx
```

### Side-by-Side Timeline

| Step | Legacy ConMas | PaperLess |
|------|--------------|-----------|
| **1. Input** | Original.xlsx + cell values | Original.xlsx + WorkbookDefinition JSON |
| **2. Workbook creation** | `File.Copy(original → temp)` | `File.Copy(original → output)` |
| **3. Open method** | COM `Workbooks.Open(temp)` | SDK `SpreadsheetDocument.Open(output, true)` |
| **4. Pre-save snapshot** | ❌ None | ✅ ZipFile → saves ALL original bytes |
| **5. Sheet filtering** | Hide sheets without print area | ❌ None |
| **6. Write cells** | COM `Range.Value = ...` | SDK `Cell.CellValue = ...` |
| **7. Styles** | ✅ Preserved (not touched) | ✅ Preserved (ZIP restore || SDK mutates) |
| **8. Shared strings** | ✅ COM manages internally | ⚠️ SDK may append; comment strings appended |
| **9. Comments** | ✅ Preserved (not touched) | ❌ Overwritten by PostProcess |
| **10. Save** | COM `SaveAs(xlExclusive)` | SDK `Dispose()` → ZIP flush |
| **11. SDK revert** | ❌ Not applicable (no SDK) | ✅ ZIP restore reverts SDK mutations |
| **12. Extra sheet** | ❌ Never adds | ❌ Adds ExcelOutputSetting (PostProcess) |
| **13. Post-cleanup** | Idempotent XML transforms | ConMas config injection |
| **14. Validation** | ❌ None | ✅ Auto-validate vs original |
| **15. Output** | `.xlsx` + `.xml` definition | `.xlsx` only |

---

## Part 7 — Structural Divergence Matrix

| Component | Legacy ConMas | PaperLess | Same? |
|-----------|--------------|-----------|-------|
| **Workbook origin** | Copy → File.Copy | Copy → File.Copy | ✅ Same |
| **Open method** | COM Interop | OpenXML SDK | ❌ Different |
| **Sheet count** | Preserved (2→2) | **Increased** (2→3) | ❌ Different |
| **Sheet names** | Preserved (_Fields, Sheet1) | _Fields, Sheet1, **ExcelOutputSetting** | ❌ Different |
| **Sheet order** | Preserved | Preserved then new sheet appended | ❌ Different |
| **Styles** | Preserved (minor: applyBorder) | **Restored from original** (ZIP restore) | ✅ Same net result |
| **Shared strings** | Preserved (COM manages) | **Appended** (new values + 36 config) | ❌ Different |
| **Cell values** | Written via COM | Written via SDK | ✅ Same goal |
| **Comments** | Preserved (not touched) | **Overwritten** | ❌ Different |
| **workbook.xml** | Minor: xWindow clamp, documentId | Restored + sheet3 entry added | ❌ Different |
| **workbook.xml.rels** | Preserved | Restored + sheet3 rel added | ❌ Different |
| **Content types** | Preserved | Restored + sheet3 override added | ❌ Different |
| **Defined names** | Minor: R1C1 escape | **Restored from original** | ⚠️ Same net result |
| **Drawings/VML** | Minor: fPrintsWithSheet removal | Not restored (SDK may mutate) | ❌ Different |
| **Printer settings** | Preserved (not touched) | **Restored from original** | ✅ Same net result |
| **Theme** | Preserved (not touched) | **Restored from original** | ✅ Same net result |
| **docProps** | Timestamp updated | **Restored from original** | ❌ Different |
| **Post-processing** | Idempotent XML cleanups | ConMas config injection | ❌ Different |
| **Validation** | None | Auto-compare vs original | ❌ Different |

### Critical Divergences

The critical divergences that cause PaperLess to fail on second-generation exports:

1. **Sheet count increase**: PaperLess adds `ExcelOutputSetting` sheet. An already-exported workbook already has this sheet → on second export, PaperLess tries to add it again → causes duplicate sheets or corrupted structure.

2. **Shared strings duplication**: PaperLess appends new values AND 36 config strings. On second export, the shared strings from the first export are still present → PaperLess may append duplicates.

3. **Comments overwrite**: PaperLess overwrites comments1.xml entirely. An already-exported workbook has ConMas-format comments → PaperLess overwrites them with the same data → but file bytes may differ due to GUIDs/timestamps.

4. **ZIP restore from original**: PaperLess restores workbook.xml/styles.xml/theme from the ORIGINAL file's bytes. On second export, the "original" IS the already-exported workbook, which has different bytes than the original template → restore actually changes things.

---

## Part 8 — Export Cycle Simulation

### ConMas: Original → Export1 → Export2 → Export3 (from Phase 7.1 COM Test)

| Component | Original→E1 | E1→E2 | E2→E3 | Drift? | Idempotent? |
|-----------|-------------|-------|-------|--------|-------------|
| workbook.xml | documentId changed, windowPos changed | **Only documentId changed** | **Only documentId changed** | **No structural drift** | ✅ After 1st pass, only documentId (unavoidable) |
| styles.xml | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| sharedStrings.xml | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| sheet1.xml | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| sheet2.xml | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| workbook.xml.rels | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| content types | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| docProps | timestamp changed | timestamp changed | timestamp changed | Timestamp only | ✅ (metadata only) |
| printer settings | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |
| theme | ✅ Identical | ✅ Identical | ✅ Identical | None | ✅ |

**ConMas drift per generation: ZERO structural drift.** Only GUID and timestamp change. All idempotent transformations produce zero change after first pass.

### PaperLess (Projected from code analysis): Original → Export1 → Export2

| Component | Original→E1 | E1→E2 | Drift? | Accumulates? |
|-----------|-------------|-------|--------|-------------|
| workbook.xml | Restored + sheet3 added | Restored from Export1 + sheet3 re-added | **YES** | Sheet3 entry may be added twice |
| styles.xml | Restored from original | Restored from Export1 | ⚠️ Depends on SDK mutations | Minimal |
| sharedStrings.xml | Values appended + 36 config added | Values appended again + 36 config added | **YES** | **36+ duplicate strings per generation** |
| sheet1.xml | Values written | Values overwritten | ⚠️ Same values → same XML | Minimal |
| sheet2.xml | Values written | Values overwritten | ⚠️ Same values → same XML | Minimal |
| sheet3.xml | Created (36 rows) | Tries to create again → **DUPLICATE** | **YES** | **Corruption on second generation** |
| comments1.xml | Overwritten with ConMas format | Overwritten again | ⚠️ Same content → may match | May match if GUIDs/timestamps consistent |
| workbook.xml.rels | sheet3 rel added | sheet3 rel added (+1 generation) | **YES** | rId counter grows each generation |
| content types | sheet3 override added | sheet3 override added again | ⚠️ May duplicate | Duplicate Override entry |
| docProps | May be restored from original | Restored from Export1 (different bytes) | **YES** | Byte drift |

**PaperLess drift per generation: SIGNIFICANT.** Multiple components accumulate changes:
- Shared strings grow every generation (new values + 36 config strings each time)
- Sheet3 may be duplicated
- Relationships accumulate rIds
- Content types may get duplicate overrides

---

## Part 9 — Root Cause Ranking

| Rank | Cause | Confidence | Evidence |
|------|-------|-----------|----------|
| **1** | **PaperLess adds ExcelOutputSetting sheet that ConMas never creates** | **Very High** | 3 forensic reports confirm ConMas never adds sheets; WorkoutGenerator.cs adds sheet3; already-exported workbook already has it → conflict |
| **2** | **PaperLess appends shared strings instead of preserving them** | **Very High** | Phase 7.1 COM test confirms ConMas preserves sharedStrings.xml byte-identical; WorkoutValueWriter appends values + 36 config strings every export |
| **3** | **PaperLess overwrites comments instead of preserving them** | **High** | ConMas never modifies comments in Output Excel (reads only); PaperLess PostProcess replaces comments1.xml entirely |
| **4** | **ZIP restore uses original file bytes; second-gen 'original' already differs from template** | **High** | On second generation, the "original" IS the already-exported workbook with modified structure; restoring ITS bytes changes the output |
| **5** | **PostProcess manipulates ZIP directly without checking if sheet3 already exists** | **High** | UpdateWorkbookXmlForSheet3 checks if ExcelOutputSetting exists, but AppendExcelOutputSettingSharedStrings always appends 36 strings |
| **6** | **Validator flags trivial changes (documentId, timestamps) as structure changes** | **Medium** | WorkoutDiffValidator.CompareAdditionalParts compares workbook.xml element-by-element via XPath — documentId change would be flagged |
| **7** | **Two different save paths (COM vs OpenXML) produce different OOXML** | **Medium** | COM path (deprecated) creates from scratch; OpenXML path copies and edits; both produce structurally different ZIP than ConMas |
| **8** | **No pre-save snapshot of ConMas-specific entries before SDK opens** | **Medium** | The ZIP restore captures original entries BEFORE SDK opens, but PostProcess changes are NOT captured for next generation |
| **9** | **PostProcess idempotency not guaranteed across generations** | **High** | WriteConMasCommentsEntry generates GUIDs; AppendExcelOutputSettingSharedStrings always appends; neither checks whether data was already written |
| **10** | **Validator runs on every save, but 'pass' criteria may be too strict** | **Low** | Validator is not the primary cause — it correctly identifies real structural differences. Relaxing it would mask deeper exporter problems |

---

## Part 10 — Final Architecture Recommendation

### Based on Evidence Only

**Option A — Continue rebuilding the workbook**

| Evidence | Supports? | Reason |
|----------|-----------|--------|
| Current approach fails on second generation | ❌ Against | All evidence shows this approach is broken |
| ConMas never rebuilds | ❌ Against | Phase 7.1 proves ConMas preserves everything |
| ZIP restore is fragile | ❌ Against | Restoring original bytes on second gen restores wrong data |

**Option B — Copy original workbook and edit in-place (ConMas approach)**

| Evidence | Supports? | Reason |
|----------|-----------|--------|
| Phase 7.1 COM test proves this preserves structure | ✅ For | 11/13 components byte-identical across 4 generations |
| ConMas IL proves this is the intended design | ✅ For | `File.Copy(original) → Open(copy) → modify → save` |
| Only documentId and timestamp change | ✅ For | These are trivial, unavoidable COM artifacts |
| All post-processing is idempotent | ✅ For | Once applied, re-application produces zero change |
| No extra sheets needed by ConMas | ✅ For | ConMas never adds sheets; `_Fields` is already in the template |
| Comments preserved by ConMas | ✅ For | ConMas reads comments, never overwrites |
| Shared strings preserved by ConMas | ✅ For | COM manages, SDK appends is wrong |

**Option C — Hybrid approach**

| Evidence | Supports? | Reason |
|----------|-----------|--------|
| PostProcess for ConMas config is needed | ⚠️ Partial | ConMas outputs a separate .xml file, not an embedded sheet |
| ZIP restore solves SDK mutation problem | ⚠️ Partial | Works for first generation, fails for second |
| Validator is useful but catches own artifacts | ⚠️ Partial | Should validate against ConMas output, not original template |

### Recommended Architecture

```
Based on ALL forensic evidence gathered across Phases 5-7, the recommended
architecture inherits from Legacy ConMas:

1. Input: Original.xlsx template (created by ConMas Designer, has _Fields + Sheet1)

2. File.Copy(original → output)         ← Already done ✅

3. Workbooks.Open(output) via COM        ← ConMas uses COM for cell value writes
   OR
   SpreadsheetDocument.Open(output) via SDK  ← Current approach, but need to fix
     └─ WRITE cell values only         ← No structural changes
     └─ PRESERVE comments              ← Do NOT touch comments1.xml
     └─ PRESERVE shared strings        ← Do NOT append config strings
     └─ PRESERVE sheet count           ← Do NOT add ExcelOutputSetting
     └─ DO NOT restore ZIP entries     ← No pre-save/restore needed if SDK is read-only

4. workbook.Save() or doc.Dispose()

5. Post-process (IF AND ONLY IF needed):
     ├─ xWindow/yWindow clamp (idempotent)
     ├─ applyBorder cleanup (idempotent)
     ├─ R1C1 escape (idempotent)
     └─ fPrintsWithSheet removal (idempotent)
     → Validation: compare to ConMas output, not to original template
```

### Why This Works (Evidence)

1. **Phase 7.1 Multi-Gen Test**: Excel COM preserves 11/13 components byte-identical across 4 generations. Only documentId and timestamp change — both are trivial and unavoidable.

2. **ConMas IL Analysis**: The decompiled code shows `File.Copy` → `Workbooks.Open` → write values → `SaveAs` → cleanup. No sheet creation, no shared string manipulation, no comment overwriting.

3. **Forensic Comparison**: Existing forensic reports confirm ConMas output has the same sheet count, same styles, same shared strings, same relationships as the original template — only cell values and a few idempotent OOXML changes differ.

4. **Database Evidence**: The `def_top.def_file` stored in the production database is the original template with cell values filled in — not a rebuilt workbook with extra sheets.

### What Must Change

| Current PaperLess Behavior | Target Behavior (ConMas-compatible) |
|---------------------------|-----------------------------------|
| Adds ExcelOutputSetting sheet | **Never add sheets** |
| Overwrites comments.xml | **Preserve existing comments** |
| Appends 36 config strings to shared strings | **Never append config strings to workbook** (config goes in .xml file) |
| ZIP restore from original bytes | **No restore needed** if SDK doesn't mutate |
| Validates against original template | **Validate against ConMas output format** or skip validation |
| PostProcess adds sheet3 + rels + content types | **No PostProcess needed** if config goes to external .xml |

### Summary

**The single architectural difference that allows ConMas to preserve workbook fidelity across unlimited generations is:**

> **ConMas never modifies the workbook structure. It only writes cell values and applies idempotent XML cleanups.**

PaperLess currently modifies the workbook structure (adds sheets, overwrites comments, appends shared strings, manipulates ZIP entries). These structural modifications accumulate across generations, causing second-generation exports to fail.

**The fix is not the validator. The fix is the exporter.**

---

*End of Phase 7.2 Investigation — DO NOT MODIFY ANY CODE*
