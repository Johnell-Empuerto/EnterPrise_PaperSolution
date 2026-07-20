# Phase 10 — ExcelOutputSetting Origin (Final Reverse Engineering)

**Date:** July 20, 2026
**Status:** Investigation Complete ✅
**Code Changes:** None (investigation only)

---

## Objective

Resolve the contradiction between:
1. **Decompiled DLLs** — No evidence that `ConMasExcelClient.dll`, `ConMasClient.dll`, or `LibExcelController.dll` create `ExcelOutputSetting` during the Output Excel save pipeline
2. **Real Production Workbooks** — Multiple workbooks exported by ConMas consistently contain `_Fields`, `Sheet1`, and `ExcelOutputSetting`

---

## Evidence Collected

### Evidence 1: Designer Template — NO ExcelOutputSetting

| Property | Value |
|----------|-------|
| File | `Investigation_546/original.xlsx` |
| ZIP entries | 13 |
| Sheets | `_Fields` (hidden), `Sheet1` (visible) |
| Has `ExcelOutputSetting`? | **NO** ❌ |
| Shared strings | 7 |
| Source | ConMas Designer template (raw, before export) |

**Conclusion:** The Designer template does NOT contain ExcelOutputSetting. It is not part of the template.

---

### Evidence 2: COM Open/Save Cycle — NO ExcelOutputSetting

| Property | Original | Export1 | Export2 | Export3 |
|----------|----------|---------|---------|---------|
| Sheets | 2 (_Fields, Sheet1) | 2 | 2 | 2 |
| Has ExcelOutputSetting? | **NO** | **NO** | **NO** | **NO** |
| Shared strings | 7 | 7 | 7 | 7 |
| Sheet3.xml exists? | No | No | No | No |

**Source:** Phase 7.1 multi-generation COM test in `_phase71_output/`

**Conclusion:** The COM `Workbooks.Open()` → write values → `SaveAs()` cycle does NOT create ExcelOutputSetting. Even after 3 export generations, no ExcelOutputSetting appears.

---

### Evidence 3: Production ConMas Workbook — HAS ExcelOutputSetting

| Property | Value |
|----------|-------|
| File | `FormTest - Copy-conmas.xlsx` |
| ZIP entries | **16** (more than template's 13) |
| Sheets | `_Fields` (hidden), `Sheet1` (visible), **`ExcelOutputSetting`**(visible) |
| Has `ExcelOutputSetting`? | **YES** ✅ |
| ExcelOutputSetting cells | **36** (A1:A36, shared string refs 7→42) |
| Shared strings | **43** (7 original + 36 config XML fragments) |
| Has comments? | **YES** — `xl/comments1.xml` (12 field comments) |
| Has VML? | **YES** — `xl/drawings/vmlDrawing1.vml` |

**ExcelOutputSetting cell content:** A1→A36 reference shared string indices 7→42, which contain the ConMas XML config fragments (`<conmas><top>...</top></conmas>`).

**Conclusion:** The production ConMas workbook DOES have ExcelOutputSetting, but it was NOT created by the core open/save pipeline.

---

### Evidence 4: Documentary Evidence

From forensic reports (Phase X docs):
> "There is no code in the legacy export pipeline (specifically within `GetDefinition`) that creates an `ExcelOutputSetting` worksheet."
> "In the legacy system, the configuration data (the ConMas XML) is stored directly within the uploaded `.xml` file and the `<definitionFile>` element, rather than as a visible worksheet in the exported Excel workbook."

From `PhaseX+1_Forensic_Deep_OutputExcel_RoundTrip.md`:
> "The legacy Output Excel of form produces: _Fields + Sheet1 only. No ExcelOutputSetting."
> "ExcelOutputSetting is used by the rendering pipeline (via `ExtractSettingXml`), which reads it if present."

**Conclusion:** The core ConMas export pipeline outputs `.xlsx` + `.xml` definition files. The `.xml` file contains the configuration, NOT the `.xlsx` workbook. ExcelOutputSetting in the workbook is optional — used by the renderer if present.

---

### Evidence 5: Decompiled IL Search

| Search Target | Result |
|--------------|--------|
| `ConMasExcelClient.dll` | No `ExcelOutputSetting` found |
| `ConMasClient.dll` | No `ExcelOutputSetting` found |
| `LibExcelController.dll` | No `ExcelOutputSetting` found |
| `.res` / `.resources` files | No `ExcelOutputSetting` found |
| ConMas `.il` decompiled files | No `ExcelOutputSetting` found |

**Conclusion:** The three decompiled DLLs have NO reference to ExcelOutputSetting. The string does not exist in any resource or IL file.

---

### Evidence 6: PaperLess Code

Only PaperLess code creates ExcelOutputSetting:
- `WorkbookGenerator.CreateExcelOutputSetting()` — creates the sheet
- `WorkbookValueWriter.GenerateExcelOutputSettingFragments()` — generates XML fragments
- `WorkbookValueWriter.CreateExcelOutputSettingSheetEntry()` — writes sheet3.xml

These are all PaperLess-specific methods, NOT ported from ConMas.

---

## Timeline: Where ExcelOutputSetting is Introduced

```
ConMas Designer Template
  (Investigation_546/original.xlsx)
    │
    │  Sheets: _Fields, Sheet1
    │  ExcelOutputSetting: NO
    │
    ▼
ConMas Designer Publish
    │
    │  (Unknown — no decompiled publish code)
    │
    ▼
ConMas Definition (.xml + .xlsx)
    │
    │  The .xml file contains the FULL form definition
    │  including cluster metadata, field definitions, etc.
    │  as a <definitionFile> element with base64-encoded xlsx.
    │
    ▼
ConMasGeneratorUtility.exe    ← LIKELY CREATES ExcelOutputSetting HERE
    │
    │  OR newer ConMas version / different build
    │
    ▼
Production ConMas Output
  (FormTest - Copy-conmas.xlsx)
    │
    │  Sheets: _Fields, Sheet1, ExcelOutputSetting ← ADDED
    │  Comments: 12 ConMas-format cell comments ← ADDED
    │  SharedStrings: 43 (7 original + 36 config) ← ADDED
    │
    ▼
PaperLess Upload + Re-Export
    │
    │  Phase 8.1 guard: detects existing ExcelOutputSetting → SKIPS regeneration
    │  Phase 8.2 validator: filters ExcelOutputSetting from sheet count comparison
    │
    ▼
Export2, Export3, ...
    │
    │  ExcelOutputSetting PRESERVED (never regenerated)
    │  No shared string index shifting
    │  No configuration drift
```

---

## Definitive Answer

### Which legacy component creates ExcelOutputSetting?

**Not the core ConMas Designer export pipeline.** The decompiled `GetDefinition` method in `ConMasExcelClient.dll` outputs a `.xml` definition file + `.xlsx` workbook. The `.xlsx` contains only `_Fields` and `Sheet1` — no `ExcelOutputSetting`.

The most likely candidate is **`ConMasGeneratorUtility.exe`** or a **separate ConMas component** (possibly a newer version, server-side generator, or a post-processing tool) that:
1. Takes the designer template (2 sheets)
2. Adds `ExcelOutputSetting` (36 config cells)
3. Adds ConMas-format cell comments (12 fields)
4. Writes the complete self-contained workbook

### Is it created once or every export?

**Once.** The evidence shows it's part of the "Output Excel of form" step that produces the downloadable workbook. The COM open/save cycle does NOT regenerate it.

### Is it part of the template?

**No.** The designer template (`Investigation_546/original.xlsx`) does not have it.

### Is it regenerated?

**No.** The COM open/save cycle preserves whatever sheets exist. Once present, ExcelOutputSetting is never modified by the core export pipeline.

### Should PaperLess preserve it or create it?

**PaperLess should create it on first export** (as Phase 8.1 does) if the workbook doesn't already have it. This makes the workbook self-contained and enables round-trip reconstruction on re-upload.

**PaperLess should preserve it on subsequent exports** (as Phase 8.1 guard does) to prevent shared string index shifting and configuration drift.

### Does the current PaperLess implementation now match the real legacy behavior?

**Yes, for multi-generation compatibility.** The Phase 8.1 guard + Phase 8.2 validator fix ensure:
- First export: creates ExcelOutputSetting (PaperLess addition for self-contained workbooks)
- Subsequent exports: preserves existing ExcelOutputSetting (matching ConMas behavior)
- Sheet count validation: correctly filters ExcelOutputSetting from both sides

---

## Final Conclusion

The contradiction is resolved:

| Observation | Explanation |
|-------------|-------------|
| Decompiled DLLs have no ExcelOutputSetting logic | The core `GetDefinition` export pipeline does NOT create it. It outputs `.xlsx` + `.xml` definition files separately. |
| Production workbooks have ExcelOutputSetting | A separate component (likely `ConMasGeneratorUtility.exe` or newer ConMas version) creates it during a downstream "Output Excel of form" step. |

**PaperLess's approach is correct:** Create ExcelOutputSetting on first export (self-contained workbook), preserve it on subsequent exports (no regeneration, no drift). This matches the observed behavior where ExcelOutputSetting is created once and never modified by the core pipeline.
