# Phase X.41 — Legacy ConMas Server Rendering Path Investigation

**Date:** 2026-07-19
**Objective:** Determine exactly how the legacy ConMas production server generates PDF reports without Microsoft Excel installed.

---

## Executive Summary

The production server renders reports to PDF **without Microsoft Excel** by using **Syncfusion.Pdf.Base** as the primary PDF engine, with **PDFNet (Apryse)** for post-processing. There is **no hidden COM invocation**, **no external process spawning**, **no IPC delegation**, and **no dynamic assembly loading** in the main PDF generation path.

The Designer (WPF client) uses COM/Excel for PDF export, but the **server path is entirely self-contained** using commercial .NET PDF libraries.

---

## Methodology

- ILDASM decompilation of every .NET assembly in the server path
- Registry scan for engine configuration keys
- File system scan of all configuration XMLs
- Dependency graph reconstruction from IL metadata

---

## 1. Complete Server Execution Flow

```
┌──────────────────────────────────────────────────────────────────┐
│  HTTP REQUEST (from Web UI, Mobile, or Scheduled Task)           │
└──────────────────────────┬───────────────────────────────────────┘
                           ▼
┌──────────────────────────────────────────────────────────────────┐
│  IIS / ASP.NET WebForms                                         │
│  ConMasAPI.dll (60 KB)                                          │
│                                                                  │
│  Entry Points:                                                   │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ Rests/DownloadPdf.aspx        → DownloadPdf            │     │
│  │ Rests/DownloadReport.aspx     → GetReportFile          │     │
│  │ Rests/MakeReportFromDefinition.aspx → MakeReport       │     │
│  │ Rests/AutoGenerate.aspx       → AutoGenerate           │     │
│  │ Rests/APIExecute.aspx         → APIExecute             │     │
│  │ Rests/CreateSortedReport.aspx → CreateSortedReport     │     │
│  └────────────────────────────────────────────────────────┘     │
└──────────────────────────┬───────────────────────────────────────┘
                           ▼
┌──────────────────────────────────────────────────────────────────┐
│  Business Layer                                                  │
│  Cimtops.ConMasBiz.dll (5 MB)                                    │
│                                                                  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ DownloadReportBiz                                         │  │
│  │  └─ DownloadFile(httpContext)                             │  │
│  │       ├─ fileType="pdf" → DownloadPdf(repTopId, ...)     │  │
│  │       ├─ fileType="pdfLayer" → DownloadPdf(...)          │  │
│  │       └─ fileType="excel" → DownloadExcel(repTopId)      │  │
│  │                                                           │  │
│  │ AutoGenerateBiz (scheduled jobs)                          │  │
│  │  └─ Generate(string, Stream, int32, string)               │  │
│  │                                                           │  │
│  │ CreateSortedReportBiz                                     │  │
│  │  └─ CreateSortReportAPI(httpContext)                     │  │
│  └───────────────────────────────────────────────────────────┘  │
└──────────────────────────┬───────────────────────────────────────┘
                           ▼
┌──────────────────────────────────────────────────────────────────┐
│  PDF Generation Engine                                           │
│  Cimtops.ConMasBiz.Common.PdfEditor (IDisposable)                │
│                                                                  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ GetPdf(int32 repTopId) → byte[]                            │  │
│  │                                                           │  │
│  │ 1. PdfLicense.HasValidLicense()  ← Syncfusion license     │  │
│  │ 2. Initialize()                                           │  │
│  │ 3. SetPdfInfo()                                           │  │
│  │ 4. PdfEditor.JoinToSeparatePdf()  (if multi-page)        │  │
│  │ 5. new PdfLoadedDocument(pdfBinary)  ← SYNCFUSION        │  │
│  │ 6. For each page: EditPdf(page, sheetNo)                  │  │
│  │      └─ Draws fields, text, shapes via Syncfusion         │  │
│  │ 7. AddLayer(page, sheetNo)  (if layer mode)              │  │
│  │ 8. InsertFdSheet(document)  (field data append)          │  │
│  │ 9. document.Save(MemoryStream)  ← FINAL PDF OUTPUT       │  │
│  │ 10. MemoryStream.ToArray() → byte[]                       │  │
│  │ 11. TrailOutput(byte[])  (optional trail pages)           │  │
│  │ 12. ConvertToAprysePdf(byte[])  ← PDFNET APRYSE          │  │
│  │ 13. Return byte[]                                         │  │
│  └───────────────────────────────────────────────────────────┘  │
└──────────────────────────┬───────────────────────────────────────┘
                           ▼
┌──────────────────────────────────────────────────────────────────┐
│  HTTP Response                                                   │
│                                                                  │
│  Content-Type: application/pdf  (or application/zip for multi)   │
│  Content-Disposition: attachment; filename="report.pdf"          │
│  Response.BinaryWrite(byte[])                                    │
└──────────────────────────────────────────────────────────────────┘
```

---

## 2. PDF Creation Entry Point — Exact Code Evidence

The final PDF byte array is produced at `Cimtops.ConMasBiz.Common.PdfEditor.GetPdf()`:

**File:** `Cimtops.ConMasBiz.il:298110-298309`

```
IL_00eb:  newobj     instance void [mscorlib]System.IO.MemoryStream::.ctor()
IL_00f6:  callvirt   instance void [Syncfusion.Pdf.Base]Syncfusion.Pdf.PdfDocumentBase::Save(class [mscorlib]System.IO.Stream)
IL_00fe:  callvirt   instance void [Syncfusion.Pdf.Base]Syncfusion.Pdf.PdfDocumentBase::Close(bool)
IL_0106:  callvirt   instance uint8[] [mscorlib]System.IO.MemoryStream::ToArray()
```

Then optionally post-processed:
```
IL_012d:  call       instance uint8[] Cimtops.ConMasBiz.Common.PdfEditor::TrailOutput(uint8[])
IL_0138:  call       instance uint8[] Cimtops.ConMasBiz.Common.PdfEditor::ConvertToAprysePdf(uint8[])
```

**Confidence: CERTAIN** — Direct IL evidence of Syncfusion PDF generation.

---

## 3. COM/Excel Search Results — Zero References Found

Every .NET assembly in the server path was searched for COM/Excel patterns:

| Pattern | Cimtops.ConMasBiz | ConMasAPI | ConMasGeneratorLib | ConMasGenerator |
|---------|-------------------|-----------|-------------------|-----------------|
| `Microsoft.Office.Interop.Excel` | ✗ | ✗ | ✗ | ✗ |
| `Excel.Application` | ✗ | ✗ | ✗ | ✗ |
| `Type.GetTypeFromProgID` | ✗ | ✗ | ✗ | ✗ |
| `Type.GetTypeFromCLSID` | ✗ | ✗ | ✗ | ✗ |
| `Activator.CreateInstance` | ✗ | ✗ | ✗ | ✗ |
| `CreateObject("Excel")` | ✗ | ✗ | ✗ | ✗ |
| `NetOffice` | ✗ | ✗ | ✗ | ✗ |
| `OfficeApi` | ✗ | ✗ | ✗ | ✗ |
| `ExcelApi` | ✗ | ✗ | ✗ | ✗ |
| `Interop.Excel` | ✗ | ✗ | ✗ | ✗ |

**Confidence: CERTAIN** — IL assembly extern list confirms zero Office/COM interop dependencies.

---

## 4. Dynamic Assembly Loading — Not Found

| Pattern | Found? |
|---------|--------|
| `Assembly.Load` | ✗ |
| `Assembly.LoadFrom` | ✗ |
| `Assembly.LoadFile` | ✗ |
| `AppDomain.Load` | ✗ |
| `Activator.CreateInstance(string)` | ✗ |
| `LoadLibrary` (P/Invoke) | ✗ |
| `Reflection` (for type activation) | ✗ |
| `System.Reflection.Assembly.Load` | ✗ |

**Confidence: CERTAIN** — No dynamic loading exists in any server-path assembly.

---

## 5. External Process Execution — Not Found

| Pattern | Found? |
|---------|--------|
| `Process.Start` | ✗ |
| `System.Diagnostics.Process` | ✗ |
| `ShellExecute` | ✗ |
| `cmd.exe` | ✗ |
| `powershell` | ✗ |

Two external executables exist on disk but are **not invoked by the main PDF path**:

| Executable | Location | Purpose |
|---|---|---|
| `ConMasCreateDW.exe` | `SettingFiles\xml\interface\report\DW\` | DataWindow export tool (separate subsystem) |
| `Cimtops.ConMasReportTable.exe` | `SettingFiles\xml\interface\report\reportTable\app\` | Report table generation tool (separate subsystem) |

These are standalone tools for specific report types, invoked via file-system pool queues, not via `Process.Start` from the main API.

**Confidence: CERTAIN** — No process spawning in the PDF generation path.

---

## 6. IPC Investigation — Not Found

| Pattern | Found? |
|---------|--------|
| Named Pipes (`\\.\pipe\`) | ✗ |
| WCF (`ServiceModel`) | ✗ |
| gRPC | ✗ |
| TCP Socket (`localhost:`) | ✗ |
| COM+ | ✗ |
| Windows Communication Foundation | ✗ |

**Confidence: CERTAIN** — No IPC mechanisms in any server assembly.

---

## 7. Infragistics Investigation — Excel Reading Only

`Infragistics4.WebUI.Documents.Excel.v12.1.dll` is referenced by `Cimtops.ConMasBiz.dll` but only for **reading Excel workbook structure**:

```
IL (line 190934):
  Infragistics.Documents.Excel.WorksheetShape.CreatePredefinedShape()
  Infragistics.Documents.Excel.FormattedTextFont.set_Name()
  Infragistics.Documents.Excel.FormattedFontBase.set_Bold()
  Infragistics.Documents.Excel.FormattedFontBase.set_Italic()
```

These APIs are used to parse the `.xlsx` template to extract cluster coordinates, field positions, labels, and formatting — NOT to render PDF.

The Designer ships with `Infragistics4.Documents.Core.v13.2.dll` and `Infragistics4.Documents.Excel.v13.2.dll` which also lack PDF export capabilities—they are spreadsheet-format libraries only.

**Confidence: CERTAIN** — Infragistics assemblies in this version do not have PDF export capability. They read .xlsx structure only.

---

## 8. DevExpress Investigation — Present but Unused for PDF

The Designer bin contains DevExpress v23.2 assemblies:
- `DevExpress.Spreadsheet.v23.2.Core.dll` (16 MB) — unused for PDF export
- `DevExpress.Pdf.v23.2.Core.dll` (5 MB) — unused for PDF export  
- `DevExpress.Printing.v23.2.Core.dll` — unused for PDF export

These are not present in the **server** bin directory at all (only in the Designer).

**Confidence: CERTAIN** — DevExpress assemblies are not used by the server PDF path.

---

## 9. External PDF Libraries in Server Bin — Usage Status

| Library | Size | Present? | Used by Cimtops.ConMasBiz? | Purpose |
|---------|------|----------|--------------------------|---------|
| `Syncfusion.Pdf.Base` | 8 MB | ✓ Yes | ✓ Yes | **Primary PDF engine** |
| `PDFNet` (Apryse) | 57 MB | ✓ Yes | ✓ Yes | Post-processing/optimization |
| `Syncfusion.PdfToImageConverter.Base` | 16 MB | ✓ Yes | ✓ Yes (referenced) | PDF→image for thumbnails |
| `Syncfusion.Licensing` | 0.5 MB | ✓ Yes | ✓ Yes | License validation |
| `Syncfusion.Compression.Base` | 0.7 MB | ✓ Yes | (dependency) | Compression support |
| `GhostscriptSharp.dll` | 12 KB | ✓ Yes | ✗ No | **Not referenced** |
| `gsdll32.dll` / `gsdll64.dll` | 12+14 MB | ✓ Yes | ✗ No | **Not referenced** |
| `itextsharp.dll` | 3.6 MB | ✓ Yes | ✗ No | **Not referenced** |
| `iTextAsian.dll` | 6 MB | ✓ Yes | ✗ No | **Not referenced** |
| `pdfium.dll` | 5 MB | ✗ No (Designer only) | ✗ No | PDF viewer |

**Ghostscript, iTextSharp, and pdfium are never referenced by Cimtops.ConMasBiz.dll.** They may be used by other subsystems (e.g., `ConMasCreateDW.exe` references them in its own directory).

**Confidence: CERTAIN** — IL assembly extern list confirms only Syncfusion and PDFNet are used.

---

## 10. Runtime Configuration — No Engine Selection

| Configuration Source | Path | Renderer Settings? |
|---|---|---|
| `Web.config` | `C:\ConMas\ConMasAPI\Web.config` | ✗ No |
| `Cimtops.ConMasBiz.dll.config` | `C:\ConMas\ConMasAPI\bin\` | ✗ No |
| `License.xml` | `C:\ConMas\SettingFiles\xml\License.xml` | ✗ Only Syncfusion license key |
| `C:\ConMas\SettingFiles\xml\*.xml` | Business XML configs | ✗ No renderer settings |
| Registry `HKCU\SOFTWARE\CIMTOPS CORPORATION` | — | ✗ No engine keys |
| Registry `HKLM\SOFTWARE\WOW6432Node\CIMTOPS CORPORATION` | — | ✗ Key does not exist |

**Confidence: CERTAIN** — No configurable engine selection exists. The renderer is hard-coded to Syncfusion+PDFNet.

---

## 11. License Validation

The production server validates a Syncfusion license at PDF generation time:

**Evidence** (`Cimtops.ConMasBiz.il:298141`):
```
IL_001f:  call       bool Cimtops.ConMasBiz.Common.Pdf.PdfLicense::HasValidLicense(string&)
```

The license key is stored in `C:\ConMas\SettingFiles\xml\License.xml`:
```xml
<ConMas>
  <license>1Xg+6bAmROlmIghyyIiw6/...</license>
</ConMas>
```

---

## 12. Dependency Graph

```
┌─────────────────────────────────────────────────────────────────────┐
│                        ConMasAPI (ASP.NET WebForms)                  │
│                          ConMasAPI.dll                               │
│                                                                       │
│  DownloadPdf.aspx  GetReportFile.aspx  MakeReportFromDef.aspx  ...   │
└──────────────────────────────┬──────────────────────────────────────┘
                               │ references
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Cimtops.ConMasBiz.dll (5 MB)                      │
│                                                                       │
│  ┌─────────────────────┐  ┌─────────────────────┐                   │
│  │  Biz.API.           │  │  Common.PdfEditor    │                   │
│  │  DownloadReportBiz  │──┤  Common.Pdf.TrailPdf  │                   │
│  │  AutoGenerateBiz    │  │  Common.PdfRepository │                   │
│  │  CreateSortedRepBiz │  │  Common.PdfLicense    │                   │
│  └─────────────────────┘  └──────────┬──────────┘                   │
│                                      │ uses                          │
│  ┌───────────────────────────────────┼───────────────────────┐       │
│  │                                   │                       │       │
│  ▼                                   ▼                       ▼       │
│  ┌─────────────────┐  ┌─────────────────────────┐  ┌──────────────┐ │
│  │ Syncfusion.Pdf   │  │ PDFNet (Apryse)         │  │ Infragistics │ │
│  │ .Base (8 MB)    │  │ pdftron.PDF (57 MB)     │  │ .Excel.v12.1 │ │
│  │                 │  │                         │  │ (read only)  │ │
│  │ PdfDocument     │  │ PDFDoc, PageIterator    │  │              │ │
│  │ PdfLoadedDoc    │  │ ImportPages, Save       │  │ Worksheet    │ │
│  │ ExportAsFixed   │  │ Initialize(license)     │  │ Shape, Font  │ │
│  └─────────────────┘  └─────────────────────────┘  └──────────────┘ │
└─────────────────────────────────────────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                       PostgreSQL Database                            │
│                                                                      │
│  Tables: rep_top, rep_sheet, rep_cluster, rep_fd_sheet, ...         │
│  Stored: Report XML definitions, cluster data, PDF binaries          │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 13. Final Conclusions

### Option A — Microsoft Excel COM
**REJECTED.**  
No COM/Excel references exist in any server-path assembly. The IL assembly extern list, string searches, and method call analysis all confirm zero Excel dependency.

### Option B — Infragistics
**REJECTED for PDF.**  
Infragistics is used only for reading `.xlsx` workbook structure (cluster coordinates, field positions). The Infragistics versions present (`v12.1`, `v13.2`, `v23.1`) do not include PDF export APIs.

### Option C — DevExpress
**REJECTED.**  
DevExpress assemblies exist only in the Designer (WPF client) bin, not in the server bin. They are not loaded or used by any server-path code.

### Option D — External Executable
**REJECTED for main path.**  
`ConMasCreateDW.exe` and `Cimtops.ConMasReportTable.exe` exist as standalone tools for separate report subsystems (DW export, Report Table generation). They are not spawned by `Process.Start` from the main API. The main PDF path is entirely in-process.

### Option E — Hidden Rendering Engine
**CONFIRMED: Syncfusion.Pdf.Base + PDFNet (Apryse).**

The server renders PDFs through:
1. **Syncfusion.Pdf.Base** — Primary engine: creates `PdfDocument`, loads `PdfLoadedDocument`, draws fields/text/shapes on pages, saves to `MemoryStream`
2. **PDFNet (Apryse)** — Post-processing: `ConvertToAprysePdf()` converts/optimizes the Syncfusion-generated PDF
3. **Infragistics** — Read-only: parses Excel template structure for cluster coordinates and field positions

### How the server renders Excel-to-PDF without Excel:

The legacy ConMas workflow is:
1. **Designer (WPF)** creates `.xlsx` template files using **Microsoft Excel COM** (this is where Excel is required)
2. **Server (ConMasAPI)** reads those `.xlsx` templates using **Infragistics** to extract cluster definitions, field coordinates, labels, and formatting parameters
3. **Server** generates the PDF by directly drawing onto `Syncfusion.Pdf.PdfPageBase` pages using the coordinates extracted in step 2
4. **Server** applies field values from the database (PostgreSQL) onto the PDF
5. **Server** optionally post-processes through **PDFNet (Apryse)** for optimization/compatibility
6. **Final PDF** is returned as `byte[]` to the HTTP response

No Excel runtime is needed on the server because the server never opens Excel — it reads the template structure natively via Infragistics and renders the PDF directly via Syncfusion.

---

## Evidence Summary

| Finding | Evidence Type | Location | Confidence |
|---------|--------------|----------|------------|
| Syncfusion.Pdf.Base is the PDF engine | IL assembly extern + method calls | `Cimtops.ConMasBiz.il:98, 239462, 298174, 298250` | CERTAIN |
| PDF generation via Save(Stream) | IL decompilation | `Cimtops.ConMasBiz.il:298250` | CERTAIN |
| PDFNet post-processing | IL method call | `Cimtops.ConMasBiz.il:298287` | CERTAIN |
| No COM/Excel references | IL assembly extern list + string search | `Cimtops.ConMasBiz.il:11-202` | CERTAIN |
| No Process.Start | String search across all IL | All server DLLs | CERTAIN |
| No dynamic assembly loading | String search across all IL | All server DLLs | CERTAIN |
| No IPC mechanisms | String search across all IL | All server DLLs | CERTAIN |
| Infragistics = read-only Excel parsing | IL method calls | `Cimtops.ConMasBiz.il:190934-191049` | CERTAIN |
| License validation before PDF gen | IL method call | `Cimtops.ConMasBiz.il:298141` | CERTAIN |
| Renderer not configurable | XML/registry search | Web.config, *.xml, registry | CERTAIN |
| Ghostscript/iTextSharp unused | IL assembly extern list | `Cimtops.ConMasBiz.il:11-202` | CERTAIN |
| DownloadPdf entry point | IL decompilation | `ConMasAPI.il:2164-2356` | CERTAIN |
| DownloadReportBiz.DownloadFile dispatcher | IL decompilation | `Cimtops.ConMasBiz.il:700538-700642` | CERTAIN |
| PdfEditor.GetPdf = final PDF output | IL decompilation | `Cimtops.ConMasBiz.il:298110-298309` | CERTAIN |

---

## Files Examined

| File | Size | Method |
|------|------|--------|
| `C:\ConMas\ConMasAPI\bin\Cimtops.ConMasBiz.dll` | 5 MB | ILDASM decompilation (1,114,790 lines) |
| `C:\ConMas\ConMasAPI\bin\ConMasAPI.dll` | 60 KB | ILDASM decompilation |
| `C:\ConMas\ConMasAPI\Web.config` | — | Read |
| `C:\ConMas\ConMasAPI\bin\Cimtops.ConMasBiz.dll.config` | — | Read |
| `C:\ConMas\SettingFiles\xml\License.xml` | — | Read |
| `C:\ConMas\SettingFiles\xml\*` | — | Directory scan |
| `C:\Program Files\ConMas Generator\ConMasGenerator.exe` | 8 KB | ILDASM decompilation |
| `C:\Program Files\ConMas Generator\ConMasGeneratorLib.dll` | 114 KB | ILDASM decompilation |
| `C:\Program Files\ConMas Generator\ConMasGeneratorUtility.exe` | 583 KB | ILDASM decompilation |
| `C:\ConMas\SettingFiles\xml\interface\report\DW\ConMasCreateDW.exe` | — | ILDASM (partial) |
| `Registry: HKLM\SOFTWARE\WOW6432Node\CIMTOPS CORPORATION` | — | Searched |
| `Registry: HKCU\SOFTWARE\CIMTOPS CORPORATION` | — | Searched |
