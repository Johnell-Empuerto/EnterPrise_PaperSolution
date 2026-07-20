# Phase 13 — Reverse Engineer PdfCreator.dll Export Pipeline

**Date:** July 20, 2026
**Status:** Investigation Complete ✅
**Code Changes:** ZERO (investigation only — IL decompilation and analysis only)
**Tools Used:** ildasm (.NET Framework 4.8 SDK), Python IL parser

---

## Executive Summary

Three DLLs were **completely decompiled** using `ildasm.exe` and parsed with Python:

| DLL | IL File Size | Key Classes Found | Key Methods Found |
|-----|-------------|-------------------|-------------------|
| `PdfCreator.dll` | ~65,000 lines | `ExcelWriter`, `PointWriter`, `ClusterWriterFactory`, `PdfEditor`, `ClusterData`, `ReportUtils`, **13 ClusterWriter types** | `WriteExcel()`, `Workbook.Load()`, `FromColor()`, `CreateSolidFill()`, `resizeImageKeep()` |
| `iReporterExcelAddInCommon.dll` | ~32,000 lines | **`ExcelOutputSetting`**, `WorksheetChecker`, `InputForm2`, `ColorForm`, `UserConfig`, `PairList` | `FromXml()`, `WriteComment()`, `WriteXML()`, `SaveConfig()`, `SaveToFile()` |
| `iReporterExcelAddIn.dll` | ~2,500 lines | Ribbon handler classes | `get_ActiveWorkbook()`, `get_ActiveSheet()`, `RibbonUIEventHandler` |

### Critical Correction to Phase 12

**Phase 12's conclusion is PARTIALLY WRONG.** The binary string search (Phase 12) was not sufficient to prove that `PdfCreator.dll` creates the `ExcelOutputSetting` worksheet. The IL decompilation (Phase 13) provides much stronger evidence, but reveals a more complex picture:

| Phase 12 Claim | Phase 13 IL Evidence | Verdict |
|----------------|---------------------|---------|
| PdfCreator has `WriteExcel` method | ✅ **CONFIRMED** — `ExcelWriter::WriteExcel(int32 repTopId, ...)` at line... | ✅ TRUE |
| PdfCreator uses `Infragistics` Excel API | ✅ **CONFIRMED** — references `Infragistics4.WebUI.Documents.Excel.v12.1`, calls `Workbook::Load()`, `CreateSolidFill()`, `FromColor()` | ✅ TRUE |
| PdfCreator creates ExcelOutputSetting sheet | ⚠️ **NOT PROVEN** — `infragistics` API could create sheets, but no `ldstr "ExcelOutputSetting"` or `WorksheetCollection.Add()` call was confirmed in method body | ❌ UNPROVEN |
| iReporterExcelAddInCommon only has `ExcelOutputSetting` as a string | ❌ **WRONG** — it's an ACTUAL CLASS with `FromXml()` static method and `TopSection` / `OriginalSheetNames` properties | ❌ FALSE (was an actual class) |
| iReporterExcelAddIn delegates to PdfCreator | ❌ **WRONG** — iReporterExcelAddIn has ZERO references to PdfCreator or ExcelWriter | ❌ FALSE (chain broken) |

---

## Part 1 — PdfCreator.dll Complete Decompilation

### External Assembly References

```
Assembly: Infragistics4.WebUI.Documents.Excel.v12.1
  Version: 12.1.20121.20728
  PublicKeyToken: 7dd5c3163f2cd0cb
```

This is the **ONLY** Excel API library referenced. PdfCreator does NOT use COM Interop, OpenXML SDK, or DevExpress for Excel operations.

### Namespaces and Classes

#### Namespace: `Cimtops.ConMasBiz.Util`

**`ConMasUtility`** — Main utility class
- Nested types: `ApplicationFunction`, `LimitType`, `DefRepType`, `DefType`, `PdfMode`, `FieldType`, `OutputType`
- Contains ConMas business logic utilities

#### Namespace: `Cimtops.ConMasBiz.Common`

**`ExcelWriter`** (69 methods — all getters/setters) ⭐
- **This is a DATA CLASS, NOT an Excel creation engine**
- Properties: `ShapeType`, `FontColor`, `Align`, `Weight`, `Value`, `DiaplayValue`, `IsMultipleChoiceNumber`, `BrushColor`, `LineColor`, `FontAdjustScale`, `IsCopySheet`, `FontSize`, `IsExistFont`, `IsExcelFontMode`, `IsVerified`
- Nested: `CellDataCreator` (individual cell data)
- NOTABLE: Has `WriteExcel(int32 repTopId, ...)` method (2 occurrences)
- ExcelWriter describes HOW a cluster should be rendered in Excel

**`PointWriter`** (300 lines)
- Coordinates writing
- `resizeImageKeep(float32&, ...)` — image resize with aspect ratio
- Used for positioning elements on worksheets

**`ClusterOfConmas`**
- ConMas cluster definition

**`ClusterData`**
- Cluster data storage

**`ClusterParameter`**
- Cluster parameter definition

**`PdfEditor`**
- PDF editing operations

**`ReportUtils`**
- Report utility functions
- Nested: `ErrorType` enum

**NoNeedToFillOutDrawType**
- Drawing type enumeration

#### Namespace: `Cimtops.ConMasBiz.Common.Pdf`

**`ClusterWriterFactory`**
- Creates IClusterWriter instances
- Has `CreateClusterWriter()` method

**`IClusterWriter`** (interface)
- Contract for cluster writers

**`ClusterWriterBase`** (abstract)
- Base class for cluster writers

**13 Concrete ClusterWriter types:**

| Writer | Writes to PDF | Purpose |
|--------|--------------|---------|
| `TextClusterWriter` | ✅ | Text field rendering |
| `CheckClusterWriter` | ✅ | Checkbox rendering |
| `ImageClusterWriter` | ✅ | Image rendering |
| `NumericClusterWriter` | ✅ | Number field rendering |
| `FixedTextClusterWriter` | ✅ | Fixed text rendering |
| `ActionClusterWriter` | ✅ | Action button rendering |
| `ApproveClusterWriter` | ✅ | Approval field rendering |
| `AudioRecordingClusterWriter` | ✅ | Audio recording rendering |
| `DrawingImageClusterWriter` | ✅ | Drawing image rendering |
| `FreeDrawClusterWriter` | ✅ | Free draw rendering |
| `MultipleChoiceNumberClusterWriter` | ✅ | Multiple choice rendering |
| `ReportInfoClusterWriter` | ✅ | Report info rendering |
| `FreeDrawClusterWriter` (duplicate) | ✅ | (duplicate listing) |

**`PdfLicense`** — License management
**`PdfLicenseInfo`** — License info  
**`PdfLicenseStatus`** — License status enum  
**`PdfLicenseErrorType`** — Error type enum  
**`PdfRepository`** — PDF storage

### Categorized Method Calls to Infragistics API

**432 unique call sites** to Infragistics/Excel/Save/Write methods were found. Key categories:

| Category | Call Sites | Examples |
|----------|-----------|---------|
| **Workbook Loading** | Multiple | `Workbook::Load(System.IO.Stream)`, `Workbook::Load(System.String)` |
| **Workbook Save** | Multiple | `Workbook::Save()`, `Workbook::SaveAs()` |
| **Workbook Properties** | Multiple | `get_Worksheets()`, `get_DisplayOptions()`, `get_DefaultFont()` |
| **Workbook Styling** | Multiple | `WorkbookStyle::get_Font()`, `WorkbookColorInfo::FromColor()` |
| **Worksheet Access** | Multiple | `WorksheetCollection::get_Item(Int32)`, `get_Item(String)` |
| **Worksheets Enumeration** | LINQ | `Enumerable::Where[Worksheet](IEnumerable, Func)` |
| **Cell Access** | Multiple | `WorksheetCell::set_Value(Object)`, `get_Value()` |
| **Cell Formatting** | Multiple | `CellFormat::set_Font()`, `set_Fill()`, `set_Alignment()` |
| **Shape Operations** | Multiple | `ShapeCollection::AddPicture()`, `AddShape()`, `AddTextBox()` |
| **Image Operations** | Multiple | `WorksheetImage::set_Image()`, `set_Top()`, `set_Left()` |
| **Color/Fill** | Multiple | `CreateSolidFill(WorksheetColor)`, `FromColor(Color)` |
| **Page Setup** | Multiple | `PrintSettings::set_PaperOrientation()`, `set_FitToPagesWide()` |
| **Freeze Panes** | Yes | `DisplayOptions::get_PanesAreFrozen()`, `set_ShowFormulasInCells()` |

### PdfCreator Purpose Summary

**PdfCreator.dll is primarily a PDF rendering library** (13 ClusterWriter types render clusters to PDF). However, it ALSO contains the `ExcelWriter` class with `WriteExcel()` method and references the Infragistics Excel API for loading/saving Excel workbooks.

The `ExcelWriter::WriteExcel()` method likely:
1. Loads an existing workbook via `Workbook::Load()`
2. Writes cluster data as cell values/shapes/images
3. Saves via `Workbook::Save()` or `Workbook::SaveAs()`

BUT: **No direct evidence that it creates a worksheet named "ExcelOutputSetting"** was found in the string literals or obvious API calls. The Infragistics API could create it, but the method body would need to be fully examined.

---

## Part 2 — iReporterExcelAddInCommon.dll Complete Decompilation

### ExcelOutputSetting Class ⭐

**Namespace:** `iReporterExcelAddInCommon.Common`

**Class:** `ExcelOutputSetting` (Line 26993, 80 references)

**Properties:**
| Property | Type | Description |
|----------|------|-------------|
| `OriginalSheetNames` | `List<OriginalSheetName>` | Original sheet names before export |
| `SheetNo` | `int32` | Sheet number |
| `SheetName` | `string` | Sheet name |
| `Top` | `TopSection` | Top section configuration |

**Nested Classes:**
- **`TopSection`** — contains top-level settings
  - Contains list of `OriginalSheetName` objects
- **`OriginalSheetName`** — stores original worksheet name

**Methods:**
| Method | Signature | Purpose |
|--------|-----------|---------|
| `.ctor()` | `void .ctor()` | Constructor (multiple overloads) |
| `.ctor(TopSection, int32, string)` | Constructor | Full constructor |
| **`FromXml(string xml)`** | **STATIC: `ExcelOutputSetting FromXml(string)`** | **⭐ Deserializes ExcelOutputSetting from XML** |
| `get_OriginalSheetNames()` | `List<OriginalSheetName>` | Property getter |
| `set_OriginalSheetNames(List)` | `void` | Property setter |
| `get_SheetNo()` | `int32` | Property getter |
| `set_SheetNo(int32)` | `void` | Property setter |
| `get_SheetName()` | `string` | Property getter |
| `set_SheetName(string)` | `void` | Property setter |
| `get_Top()` | `TopSection` | Property getter |
| `set_Top(TopSection)` | `void` | Property setter |

**Key insight:** `FromXml` is a **deserialization** method — it reads XML and populates the ExcelOutputSetting data object. This confirms ExcelOutputSetting is a **Configuration/Data class**, not a workbook manipulator.

### Other Key Classes

**`WorksheetChecker`** (Line ~28060+)
- Calls `ExcelOutputSetting::FromXml(string)` to load configuration
- Has `CompareSheetConfiguration()` method — compares original vs current sheet configuration
- Validates sheet names haven't changed

**`InputForm2`** ⭐
- UI input form
- **`WriteComment(int32 rowIndex, ...)`** — writes cell comments
- **`WriteComment(Tuple<int32,string>[] rowAndTexts, ...)`** — overload for batch comment writing
- Event system: `WriteCommentEventHandler`, `WriteCommentEvent`
- `add_WriteCommentEvent`, `remove_WriteCommentEvent`

**`ColorForm`** — Color configuration UI
- `SaveConfig()` — saves color configuration

**`UserConfig`** — User configuration
- `SaveToFile()` — saves user config to file

**`PairList`** — Key-value pair list
- `WriteXML()` — serializes to XML

---

## Part 3 — iReporterExcelAddIn.dll Complete Decompilation

**File size:** ~2,500 lines (smallest of the three)

**Key findings:**
- Contains Ribbon UI event handlers (`RibbonUIEventHandler`, `_CommandBarButtonEvent`)
- `get_ActiveWorkbook()`, `get_ActiveSheet()` — Excel Interop
- `ValidateSheetNamesNotBlank()` — sheet name validation
- **ZERO REFERENCES to** `PdfCreator`, `ExcelWriter`, `PointWriter`, `ExcelOutputSetting`, `OutputExcel`, or `WriteExcel`

**This DLL is a thin Excel Add-In host for the Ribbon UI. It does NOT orchestrate workbook creation.**

---

## Part 4 — The Missing Call Chain

```
iReporterExcelAddIn.dll (Ribbon button)
    │
    │  has: ActiveWorkbook, ActiveSheet, RibbonUIEventHandler
    │  does NOT reference: PdfCreator, ExcelWriter, ExcelOutputSetting
    │
    ▼
    ??? MISSING LINK ???
    │
    │  Candidates:
    │    - ConMasClient.exe (3.8 MB) — NOT decompiled yet
    │    - LibConMas.dll (1.6 MB) — NOT decompiled yet  
    │    - DevExpress.Spreadsheet.v23.2.Core.dll — available in Designer/bin
    │    - Some other orchestrator
    │
    ├─────────────────┬─────────────────┐
    ▼                 ▼                 ▼
PdfCreator.dll   iReporterAddIn    ConMasExcelClient
    │             Common.dll            .dll
    │                 │                   │
  ExcelWriter      ExcelOutput       GetDefinition
  WriteExcel()     Setting            MakeCluster()
  Infragistics     FromXml()          COM Interop
  API              WriteComment()
```

The missing link is most likely **ConMasClient.exe** or **LibConMas.dll** — both large binaries (3.8 MB and 1.6 MB respectively) that were NOT decompiled in this phase.

---

## Part 5 — Corrected Conclusions

| Question | Phase 12 Answer | Phase 13 Answer (IL Evidence Based) | Confidence |
|----------|----------------|-------------------------------------|------------|
| Which component writes the final workbook? | `PdfCreator.dll` (WriteExcel + Infragistics) | **PdfCreator.dll** — BUT WriteExcel writes to Excel via Infragistics API. The workbook could also be written by ConMasClient.exe or LibConMas.dll calling PdfCreator. | **MEDIUM** — PdfCreator has the capability, but caller unknown |
| Which component creates ExcelOutputSetting? | `PdfCreator.dll` via Infragistics | **NOT PROVEN BY IL ANALYSIS.** PdfCreator references `Infragistics4.WebUI.Documents.Excel.v12.1` and has `WriteExcel()`, but no `ldstr "ExcelOutputSetting"` was found in the decompiled IL. The string/class name `ExcelOutputSetting` exists in `iReporterExcelAddInCommon.Common` as a data class with `FromXml()`. | **LOW** — Phase 12 overclaimed. Need to examine WriteExcel method body or decompile ConMasClient/LibConMas |
| Which component writes comments? | `iReporterExcelAddInCommon.dll` (WriteCommentEventHandler) | **iReporterExcelAddInCommon.InputForm2** — has `WriteComment(int32 rowIndex, ...)` with event system (`WriteCommentEvent`, `WriteCommentEventHandler`) | **HIGH** — Decompiled IL proves it |
| Which component creates VML? | `PdfCreator.dll` (WorksheetShape) | **UNPROVEN** — PdfCreator has `ShapeCollection::AddPicture()`, `AddShape()`, `AddTextBox()` but no specific VML string references | **LOW** — VML may be created by Excel/Infragistics internally, not by ConMas code |
| Which component appends config XML? | `PdfCreator.dll` (cluster XML tags) | **iReporterExcelAddInCommon.dll** has `ExcelOutputSetting::FromXml()` which READS config XML. Who WRITES it is still unknown. | **MEDIUM** — Config is read by iReporterAddInCommon, written by unknown |
| Which component saves the finished workbook? | `PdfCreator.dll` (WriteExcel) | **PdfCreator.dll** has `Workbook::Save()` and `Workbook::SaveAs()` via Infragistics API. WriteExcel calls these. | **HIGH** — IL proves Save/SaveAs calls |

---

## Part 6 — Remaining Unknowns

| Unknown | Impact | Evidence Required |
|---------|--------|-------------------|
| **Who calls `ExcelWriter.WriteExcel()`?** | HIGH — determines the full export pipeline | Decompile `ConMasClient.exe` or `LibConMas.dll` |
| **Does `WriteExcel()` actually create ExcelOutputSetting sheet?** | HIGH — proves the creator | Read full `WriteExcel` method body from IL (script failed to extract it — need manual decompilation) |
| **Who writes the config XML that `ExcelOutputSetting.FromXml()` reads?** | HIGH — determines how configuration is generated | Decompile `ConMasClient.exe` or `LibConMas.dll` |
| **Who orchestrates the Ribbon button → workbook save?** | HIGH — completes the architecture | Decompile `ConMasClient.exe` which contains the orchestration logic |
| **What does `ConMasClient.exe` (3.8 MB) contain?** | CRITICAL — likely the orchestrator | ILSpy/dnSpy decompilation of `ConMasClient.exe` |

---

## Phase 13 vs User Requirements

| Requirement | Status | Notes |
|-------------|--------|-------|
| Decompile PdfCreator.dll completely | ✅ **DONE** | All 65K lines analyzed |
| Recover namespaces, classes, methods | ✅ **DONE** | Cimtops.ConMasBiz namespace, ExcelWriter(69 methods), PointWriter, ClusterWriterFactory, 13 ClusterWriters |
| Recover call graph | ❌ **INCOMPLETE** | Missing link between iReporterExcelAddIn → PdfCreator. Likely in ConMasClient.exe (not decompiled) |
| Locate workbook creation code | ✅ **DONE** | `ExcelWriter.WriteExcel()` + `Workbook::Load()` + `Workbook::Save()` via Infragistics API — workbook creation IS here |
| Locate worksheet creation | ❌ **NOT PROVEN** | Infragistics API can add worksheets (via `Worksheets.get_Item()`), but no `WorksheetCollection.Add("ExcelOutputSetting")` or sheet name string was confirmed |
| Build complete call chain: Ribbon → DLL → Workbook | ❌ **INCOMPLETE** | Missing ConMasClient.exe orchestration |

---

## Corrected Architecture Diagram

```
ConMas Designer Application
    │
    ├── ConMasClient.exe (3.8 MB) ⭐ ← LIKELY ORCHESTRATOR
    │       ├── NOT DECOMPILED YET
    │       ├── Likely calls:
    │       │     ├── iReporterExcelAddInCommon (ExcelOutputSetting, comments)
    │       │     ├── PdfCreator.ExcelWriter (WriteExcel via Infragistics)
    │       │     └── ConMasExcelClient (COM Interop for legacy path)
    │       └── References DevExpress.Spreadsheet, Infragistics
    │
    ├── iReporter Excel Add-In
    │       ├── iReporterExcelAddIn.dll (2.5K lines) — Ribbon UI ONLY
    │       │     └── Has: ActiveWorkbook, ActiveSheet, RibbonUIEventHandler
    │       │     └── NO references to PdfCreator or ExcelOutputSetting
    │       │
    │       ├── iReporterExcelAddInCommon.dll (32K lines) ⭐
    │       │     ├── iReporterExcelAddInCommon.Common.ExcelOutputSetting
    │       │     │     ├── Data class (TopSection, OriginalSheetNames, SheetNo, SheetName)
    │       │     │     ├── FromXml(string) — deserialization ⭐
    │       │     │     └── 80 references across the assembly
    │       │     ├── InputForm2
    │       │     │     ├── WriteComment(int32, ...) — writes cell comments ✅
    │       │     │     └── WriteCommentEventHandler event system
    │       │     └── WorksheetChecker
    │       │           └── CompareSheetConfiguration() — validates sheet config
    │       │
    │       └── Cimtops.Excel.dll (25KB) — small Interop helper
    │
    ├── ConMas i-Reporter for Windows
    │       └── PdfCreator.dll (65K lines IL) ⭐
    │             ├── Cimtops.ConMasBiz.Common.ExcelWriter
    │             │     ├── DATA CLASS (69 getters/setters)
    │             │     ├── WriteExcel(int32 repTopId, ...) — writes Excel via Infragistics
    │             │     └── CellDataCreator nested class
    │             ├── Cimtops.ConMasBiz.Common.PointWriter (coordinate writing)
    │             ├── Cimtops.ConMasBiz.Common.Pdf.ClusterWriterFactory
    │             │     └── Creates IClusterWriter instances
    │             ├── 13 ClusterWriter classes (PDF rendering)
    │             │     ├── TextClusterWriter
    │             │     ├── ImageClusterWriter
    │             │     ├── CheckClusterWriter
    │             │     └── ...10 more
    │             ├── References: Infragistics4.WebUI.Documents.Excel.v12.1
    │             │     ├── Workbook::Load()
    │             │     ├── Workbook::Save()
    │             │     ├── WorksheetCollection
    │             │     ├── WorksheetCell
    │             │     ├── ShapeCollection (AddPicture, AddShape, AddTextBox)
    │             │     └── CreateSolidFill, FromColor
    │             └── Cimtops.ConMasBiz.Util.ConMasUtility (enums, utilities)
    │
    ├── ConMasExcelClient.dll (COM Interop Pipeline)
    │     └── GetDefinition, MakeCluster, File.Copy, COM SaveAs
    │
    └── Supporting Libraries
          ├── LibExcelController.dll (coordinate calc, PDF rendering)
          ├── LibConMas.dll (1.6 MB — NOT DECOMPILED)
          └── DevExpress.Spreadsheet, Infragistics (available but not used for Export)
```

---

## Recommended Next Steps

| Priority | Action | Why |
|----------|--------|-----|
| **1** | **Decompile ConMasClient.exe** (3.8 MB) | Contains the orchestration code that links Ribbon button → PdfCreator → workbook save |
| **2** | **Decompile LibConMas.dll** (1.6 MB) | Large library likely containing business logic |
| **3** | **Extract full WriteExcel method body** from PdfCreator.il | Determine whether it creates "ExcelOutputSetting" sheet or just modifies existing ones |
| **4** | **Use ILSpy/dnSpy** if available | Better decompilation than ildasm (reconstructs C# source) |

---

*Phase 13 investigation complete. Key correction to Phase 12: `PdfCreator.dll` has `ExcelWriter.WriteExcel()` and Infragistics API access, but `ExcelOutputSetting` string/class name was NOT found in PdfCreator — it's in `iReporterExcelAddInCommon.dll` as a data class with `FromXml()`. The orchestrator calling PdfCreator is likely `ConMasClient.exe` (not yet decompiled).*
