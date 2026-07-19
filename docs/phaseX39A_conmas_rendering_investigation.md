# Phase X.39A — ConMas Ecosystem Rendering Investigation

## Objective

Investigate how the legacy ConMas ecosystem renders Excel files to PDF, to understand its architecture, dependencies, and engine selection strategy. This informs our Python service's approach to Excel-to-PDF rendering without necessarily requiring Microsoft Office.

---

## 1. Installed ConMas Components

Four installation directories were discovered under `C:\Program Files (x86)\CIMTOPS CORPORATION\`:

| Directory | Primary Executable | Role |
|---|---|---|
| `ConMas i-Reporter for Windows` | `i-Reporter_v1.exe` (5.3 MB) | Main desktop application (WPF/.NET 4.7.2) |
| `ConMas Designer` | `ConMasClient.exe` (3.9 MB) | Template/form designer (.NET 4.7.2) |
| `ConMas Generator` | `ConMasGenerator.exe`, `ConMasJob.exe` | Background batch job scheduler |
| `iReporterExcelAddIn` | `iReporterExcelAddIn.dll` | VSTO Excel add-in for Office integration |

Additionally, two ASP.NET web applications exist under `C:\ConMas\`:

| Directory | Role |
|---|---|
| `ConMasAPI` | REST API for the ConMas ecosystem (server-side rendering) |
| `ConMasWeb` | Web application frontend |

---

## 2. Rendering Library Inventory

### 2.1 ConMas i-Reporter for Windows — Desktop App Path

| Library | Size | Purpose | Office Required? |
|---|---|---|---|
| `Infragistics4.WebUI.Documents.Excel.v12.1.dll` | 2.3 MB | Read/write Excel files | **No** |
| `Syncfusion.Pdf.Base.dll` | 8.1 MB | PDF generation | **No** |
| `Syncfusion.Compression.Base.dll` | 768 KB | PDF compression support | **No** |
| `PDFNet.dll` | 50 MB | PDFTron PDFNet SDK (full PDF suite) | **No** |
| `O2S.Components.PDFRender4NET.WPF.dll` | 4.2 MB | PDF rendering to images (WPF) | **No** |
| `O2S.Components.PDFView4NET.WPF.dll` | 5.2 MB | PDF viewing control (WPF) | **No** |
| `PdfCreator.dll` | 156 KB | PDF creation utility | **No** |

**Rendering pipeline:** `Infragistics Excel` → `Syncfusion PDF` / `PDFNet` → `O2S PDF Render` (for previews)

### 2.2 ConMas Designer\bin — Designer Path

| Library | Size | Purpose | Office Required? |
|---|---|---|---|
| **DevExpress.Spreadsheet.v23.2.Core.dll** | **15 MB** | Full spreadsheet engine (charts, formatting, formulas) | **No** |
| **DevExpress.Pdf.v23.2.Core.dll** | **5 MB** | PDF generation | **No** |
| DevExpress.Docs.v23.2.Core.dll | 4 MB | Document processing core | No |
| DevExpress.Office.v23.2.Core.dll | 3.5 MB | Office document model | No |
| DevExpress.Printing.v23.2.Core.dll | 4.5 MB | Printing infrastructure | No |
| **Infragistics4.Documents.Excel.v23.1.dll** | **4.3 MB** | Excel read/write | **No** |
| Infragistics4.Documents.Excel.v13.2.dll | 2 MB | Older Excel read/write | No |
| Infragistics4.Documents.Core.v23.1.dll | — | Core for v23.1 | No |
| InfragisticsWPF4.Excel.v10.3.dll | 1.5 MB | WPF Excel component | No |
| **DocumentFormat.OpenXml.dll** | **6 MB** | Official OpenXML SDK | **No** |
| PDFBaseLib31.dll | 3.8 MB | PDF base library | No |
| PDFBaseLib32.dll | 3.7 MB | 32-bit PDF base library | No |
| PDFRes12.dll | 4.3 MB | PDF resources | No |
| PdfTk31.dll / PdfTkNet31.dll | 905+692 KB | PDF toolkit | No |
| **pdfium.dll** | **13 MB** | Google PDFium rendering engine | **No** |
| OfficeApi.dll | 773 KB | Office API wrapper | Partial |
| NetOffice.dll | 50 KB | Managed Office interop library | Yes (for COM) |

### 2.3 ConMasExcelConverter — Standalone Converter

| Library | Purpose | Office Required? |
|---|---|---|
| `Infragistics4.Documents.Excel.v13.2.dll` (2 MB) | Excel read/write | **No** |
| `Infragistics4.Documents.Core.v13.2.dll` (1 MB) | Core library | **No** |
| `ExcelApi.dll` (4 MB) | Excel API wrapper | Yes (COM) |
| `NetOffice.dll` (50 KB) | Office interop | Yes (COM) |
| `OfficeApi.dll` (773 KB) | Office API | Yes (COM) |
| `ExcelConverter.xlsm` (43 KB) | VBA macro helper workbook | Yes (Excel) |

### 2.4 ConMasAPI / ConMasWeb — Server-Side Path

| Library | Size | Purpose | Office Required? |
|---|---|---|---|
| **Infragistics4.Documents.Excel.v13.2.dll** | 2 MB | Excel read/write | **No** |
| Infragistics4.WebUI.Documents.Excel.v12.1.dll | 2.3 MB | Old Excel read/write | No |
| **Syncfusion.Pdf.Base.dll** | **8 MB** | PDF generation | **No** |
| **Syncfusion.PdfToImageConverter.Base.dll** | **16 MB** | PDF → image conversion | **No** |
| **PDFNet.dll** | **57 MB** | PDFTron PDFNet (full SDK) | **No** |
| **GhostscriptSharp.dll** | 12 KB | Ghostscript .NET wrapper | **No** |
| **gsdll32.dll** | **12 MB** | Ghostscript 32-bit | **No** |
| **gsdll64.dll** | **14 MB** | Ghostscript 64-bit | **No** |
| **itextsharp.dll** | **3.6 MB** | iTextSharp PDF library | **No** |
| iTextAsian.dll | 6 MB | Asian fonts for iText | No |
| **pdfium.dll** (x86/x64/arm64) | 5 MB each | Google PDFium | **No** |
| Ionic.Zip.dll | 462 KB | ZIP handling | No |

---

## 3. The Key Abstraction: `LibExcelController.dll`

This is the single most important component. It implements a **dual-engine architecture** with abstract base classes and concrete implementations for each engine.

### 3.1 Class Hierarchy

```
ExcelControllerBase          (abstract controller)
├── ExcelControllerInterop   (COM Interop controller)
│
ExcelRangeBase               (abstract cell/range)
├── ExcelRangeCom            (COM Interop implementation)
├── ExcelRangeInfra          (Infragistics implementation)
│
ExcelWorkbookBase            (abstract workbook)
├── ExcelWorkbookCom         (COM Interop implementation)
├── ExcelWorkbookInfra       (Infragistics implementation)
│
ExcelWorksheetBase           (abstract worksheet)
├── ExcelWorksheetCom        (COM Interop implementation)
├── ExcelWorksheetInfra      (Infragistics implementation)
│
ExcelWorksheetsBase          (abstract worksheets collection)
├── ExcelWorksheetsCom       (COM Interop implementation)
├── ExcelWorksheetsInfra     (Infragistics implementation)
│
PrintOptionsBase             (abstract print settings)
DisplayOptionsBase           (abstract display settings)
RowColumnBase                (abstract row/column)
```

### 3.2 Engine Selection Strategy

The controller likely checks for Office availability at runtime:
- **Office COM available** → use `*Com` classes (Excel COM Interop via `GetTypeFromCLSID` / `GetTypeFromProgID` for `EXCEL_PROG_ID`)
- **Office COM unavailable** → fall back to `*Infra` classes (Infragistics Excel — no Office needed)

### 3.3 PDF Export Methods

- `ExportAsFixedFormat` — Standard Excel COM export (uses `xlTypePDF = 0`, `xlTypeXPS = 1`)
- `ExportPdf2007` — Excel 2007-format PDF export
- `ExportPdf2010` — Excel 2010-format PDF export
- `CanExportPdf` — Capability check
- `CreateImageFromPdf` — PDF → image conversion (for preview)
- `FixedFormatExtClassPtr` — External PDF export class pointer

### 3.4 Other Key Capabilities

- **Corrupted file handling**: `CorruptLoad`, `IsZipFile`, `WorkbookLoadOptions`
- **Print area management**: `GetFirstPagePrintArea`, `ExistsPrintAreaOnePage`
- **Shape/image handling**: `MorphShapes`, `ClearShapes`, `DrawingImage`
- **Process management**: `KillExcelProcess`, `SetWindowPos` (`HWND_TOP`, `SWP_SHOWWINDOW`), `GetProcessesByName`
- **Image processing**: `OpenCvSharp` (OpenCV), `BitmapConverter`, `CvtColor`, `ToMat`
- **PDF viewing**: `PdfiumViewer`, `PdfDocument`, `PdfRenderFlags`
- **Cell manipulation**: `CalculateCell`, `MergeArea`, `ClearComments`, `ClearContents`, `FillBorder`

---

## 4. Rendering Pipeline Architecture

```
                    ┌──────────────────────┐
                    │   Excel Source File   │
                    │  (.xlsx / .xls / xlsm)│
                    └──────────┬───────────┘
                               │
                    ┌──────────▼───────────┐
                    │   Engine Selection    │
                    │  (LibExcelController) │
                    │                      │
                    │  Office COM avail?    │
                    │    ├── YES → *Com    │
                    │    └── NO  → *Infra   │
                    └──────────┬───────────┘
                               │
                    ┌──────────▼───────────┐
                    │   PDF Export          │
                    │                      │
                    │  App Tier            │
                    │  ├── Syncfusion PDF  │
                    │  ├── DevExpress PDF  │
                    │  ├── PDFNet          │
                    │  ├── Ghostscript     │
                    │  ├── iTextSharp      │
                    │  └── pdfium          │
                    │                      │
                    │  Server Tier         │
                    │  ├── Syncfusion PDF  │
                    │  ├── PDFNet          │
                    │  ├── Ghostscript     │
                    │  ├── iTextSharp      │
                    │  └── pdfium          │
                    └──────────┬───────────┘
                               │
                    ┌──────────▼───────────┐
                    │   PDF Output File     │
                    └──────────────────────┘
```

---

## 5. Key Findings

### 5.1 Confirmed: No Office Dependency

Every component in the ConMas ecosystem ships with at least one Office-free Excel reader:
- **i-Reporter**: Infragistics Excel v12.1
- **Designer**: DevExpress Spreadsheet v23.2 + Infragistics v23.1 + DocumentFormat.OpenXml
- **API/Web**: Infragistics Excel v12.1 + v13.2
- **Converter**: Infragistics v13.2

### 5.2 Confirmed: Multiple PDF Backends

The server-side API (`ConMasAPI`) ships with **5 different PDF engines**, likely used in a fallback chain:
1. **Syncfusion** (8 MB) — preferred managed .NET PDF engine
2. **PDFNet** (57 MB) — full-featured commercial SDK
3. **Ghostscript** (12+14 MB) — PostScript/PDF interpreter
4. **iTextSharp** (3.6 MB) — open-source PDF library
5. **pdfium** (5+5+5 MB) — Google's PDF rendering (for preview only)

### 5.3 Confirmed: VSTO Add-In Exists

The `iReporterExcelAddIn` provides optional Office COM integration for users who have Excel installed. This gives better fidelity for complex Excel files with macros, ActiveX controls, etc.

### 5.4 The `LibExcelController` Architecture Is Our Template

The dual-engine pattern (`*Com` + `*Infra` base classes) in `LibExcelController.dll` serves as the ideal template for our Python implementation:
- Abstract Excel interface (`ExcelControllerBase`)
- COM engine implementation (via win32com — serialized through the global render queue)
- Infragistics/OpenPyXL engine implementation (no Office needed)
- Runtime engine selection based on availability and file complexity

### 5.5 Job Processing Pipeline

The `ConMasGeneratorLib.dll` reveals:
- **Job scheduler** (Windows Task Scheduler integration)
- **File watcher** (watches directories for new files)
- **REST API integration** (`RestController`, `WebClientEx`)
- **Mail sending** (SMTP with error notifications)
- **SQL CE database** for job persistence (`GeneratorDatabase.sdf`)
- **log4net logging** throughout

---

## 6. Implications for Our Service

1. **Our global render queue (Phase X.38) mirrors the COM *Com classes** — serializing Office COM through a single worker is the correct approach.

2. **We need to implement an *Infra equivalent** — a pure-Python or OpenPyXL-based engine that can read Excel files and produce PDFs without Office. This matches the `Infragistics`/`DevExpress Spreadsheet` approach.

3. **The `LibExcelController` base class pattern** should guide our Python abstraction: a common `ExcelWorkbook` interface with COM and pure-Python implementations, plus runtime engine selection.

4. **For PDF output**, we can adopt the same multi-backend strategy:
   - Primary: Python-native (ReportLab, WeasyPrint, or similar)
   - Fallback: Office COM export (via queue)
   - Future: PDFNet Python bindings or Ghostscript

5. **The `ExcelConverter.xlsm` + `ConMasExcelConverter.exe`** pattern is a hybrid approach — they use a VBA helper workbook when Excel is available, with Infragistics as fallback. We can adopt a similar hybrid strategy.

---

## 7. Next Steps (Phase X.39)

1. Implement Python abstract Excel interface mirroring `LibExcelController`
2. Implement COM engine (already done via render queue) — analogous to `*Com` classes
3. Implement pure-Python engine using OpenPyXL + ReportLab — analogous to `*Infra` classes
4. Implement runtime engine selection with graceful fallback
5. Validate rendering fidelity against known-good PDFs from the legacy system
6. Produce compatibility matrix

---

## Appendix: ConMas Installation Layout

```
C:\Program Files (x86)\CIMTOPS CORPORATION\
├── ConMas i-Reporter for Windows\
│   ├── i-Reporter_v1.exe          (5.3 MB — main app, .NET 4.7.2)
│   ├── i-Reporter.exe              (5.6 KB — launcher stub)
│   ├── Infragistics4.WebUI.Documents.Excel.v12.1.dll
│   ├── Syncfusion.Pdf.Base.dll
│   ├── PDFNet.dll                  (50 MB)
│   ├── O2S.Components.PDFRender4NET.WPF.dll
│   ├── O2S.Components.PDFView4NET.WPF.dll
│   └── zxing.dll                   (barcode reader)
│
├── ConMas Designer\
│   └── bin\
│       ├── ConMasClient.exe        (3.9 MB — designer app)
│       ├── ConMasExcelClient.dll
│       ├── LibConMas.dll
│       ├── LibExcelController.dll  (★ KEY: dual-engine abstraction)
│       ├── DevExpress.Spreadsheet.v23.2.Core.dll (15 MB)
│       ├── DevExpress.Pdf.v23.2.Core.dll
│       ├── Infragistics4.Documents.Excel.v23.1.dll
│       ├── DocumentFormat.OpenXml.dll (6 MB)
│       └── ConMasExcelConverter\   (standalone converter)
│           ├── ConMasExcelConverter.exe
│           ├── ExcelApi.dll         (4 MB — Office COM wrapper)
│           ├── NetOffice.dll
│           ├── ExcelConverter.xlsm  (VBA helper)
│           └── Infragistics4.Documents.Excel.v13.2.dll
│
├── ConMas Generator\
│   ├── ConMasGenerator.exe         (batch job runner)
│   ├── ConMasGeneratorUtility.exe
│   ├── ConMasJob.exe               (job processor)
│   ├── ConMasGeneratorLib.dll      (job scheduling/scheduling)
│   └── ConMasTool.exe
│
└── iReporterExcelAddIn\
    ├── iReporterExcelAddIn.dll     (VSTO add-in)
    ├── Cimtops.Excel.dll
    ├── Cimtops.R2Cluster.dll
    └── Microsoft.Office.Tools.*.dll

C:\ConMas\
├── ConMasAPI\                      (REST API — ASP.NET)
│   ├── Cimtops.ConMasBiz.dll       (5 MB — business logic)
│   ├── Infragistics*.dll
│   ├── Syncfusion*.dll
│   ├── PDFNet.dll                  (57 MB)
│   ├── Ghostscript*.dll
│   ├── itextsharp.dll
│   ├── pdfium.dll
│   └── *.aspx                      (REST endpoints)
│
├── ConMasWeb\                      (web frontend — ASP.NET)
│   └── (same libraries as ConMasAPI)
│
├── ConMasManager\                  (management web app)
├── ConMasGenerator\                (generator data files)
├── gateway\                        (Node.js + Python gateway)
├── SettingFiles\                   (XML configuration)
└── postgreSQL\                     (database)
```
