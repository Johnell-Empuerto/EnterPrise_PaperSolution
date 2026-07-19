# Phase X.40 — COM Dependency Analysis (Legacy ConMas PDF Rendering)

**Date:** 2026-07-19
**Scope:** Determine whether legacy ConMas requires Microsoft Excel for production PDF rendering, or transparently falls back to Infragistics/DevExpress.

---

## Executive Summary

ConMas legacy has **two completely separate PDF rendering paths**:

| Path | Location | PDF Engine | Requires Excel? |
|---|---|---|---|
| **Server-side** | `ConMasAPI` / `Cimtops.ConMasBiz.dll` | Syncfusion.Pdf.Base + PDFNet | **No** |
| **Designer (WPF Client)** | `LibExcelController.dll` | Microsoft Excel COM Interop | **Yes** |

The server-side path already generates PDFs **without any COM/Excel dependency**. The Designer client path is hard-coded to require Excel and has no fallback.

---

## Evidence

### 1. Server-side (`Cimtops.ConMasBiz.dll`) — NO Excel dependency

**ILDASM decompilation** of `C:\ConMas\ConMasAPI\bin\Cimtops.ConMasBiz.dll` (5 MB) reveals its full external assembly dependency list:

```
Assembly external references:
  - Syncfusion.Pdf.Base          ← PDF creation & manipulation
  - PDFNet                       ← PDF page merge & processing
  - Syncfusion.PdfToImageConverter.Base  ← PDF→image conversion
  - Syncfusion.Licensing         ← Syncfusion license validation
  - Infragistics4.WebUI.Documents.Excel.v12.1  ← Read Excel workbooks ONLY
  - [38 other non-COM assemblies]

NO references to:
  - Microsoft.Office.Interop.Excel
  - Excel.Application
  - ExcelController / LibExcelController
  - NetOffice / OfficeApi
```

**PDF generation classes** found in IL:

| Class | Method | Engine |
|---|---|---|
| `TrailPdf` | `CreatePdf(int repTopId)` | `Syncfusion.Pdf.Parsing.PdfLoadedDocument` |
| `PdfEditor` | Various edit methods | `Syncfusion.Pdf.Graphics.PdfFont` |
| `(anonymous merge)` | PDF page merge | `PDFNet.pdftron.PDF.PDFDoc` |

**API endpoints** that use this server-side PDF path:

| Endpoint | Business Logic |
|---|---|
| `MakeReportFromDefinition` | `ReportBiz.MakeReportFromDefinition()` |
| `GetReportFile` | `DownloadReportBiz.DownloadFile()` → returns `byte[]` |

### 2. Generator Service — NO PDF/Excel code at all

| Binary | Purpose | PDF/Excel References? |
|---|---|---|
| `ConMasGenerator.exe` (8 KB) | Windows service host | None — only TaskScheduler wrapper |
| `ConMasGeneratorLib.dll` (114 KB) | Task Scheduler library | None — pure scheduling |
| `ConMasGeneratorUtility.exe` (583 KB) | Config UI (WinForms) | None — settings & views only |

The Generator service **does not generate PDFs**. It schedules tasks that ultimately hit the ConMasAPI server-side endpoints.

### 3. Designer (WPF Client) — REQUIRES Excel

**ILDASM decompilation** of `LibExcelController.dll` (82 KB) at:
`C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\`

```csharp
// Engine selection logic (hard-coded COM, NO fallback)
bool CanExportPdf() {
    Type excelType = Type.GetTypeFromProgID("Excel.Application");
    if (excelType == null)
        throw new ApplicationException("NoExcelInstalled");
    return true;
}

void ExportPdf2010(string workbookPath, string outputPath) {
    var excelApp = new Microsoft.Office.Interop.Excel.Application();
    var workbook = excelApp.Workbooks.Open(workbookPath);
    workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
}
```

**Key observations:**
- `CanExportPdf()` fails fatally (`"NoExcelInstalled"`) when Excel COM is unavailable
- `ExportPdf2007()` / `ExportPdf2010()` both use `Microsoft.Office.Interop.Excel`
- `ExcelControllerInterop::GetDefinition()` creates `Excel.Application` via CLSID `{00024500-0000-0000-C000-000000000046}` AND opens workbook via COM Interop
- Infragistics (`*Infra` classes) is used for **reading** workbook structure only — **not** for PDF export

### 4. Designer DLL Inventory — Non-COM alternatives exist but are unused for PDF export

The Designer ships with these non-COM PDF-capable libraries but **does not use them for PDF export**:

| Library | Size | Used For? |
|---|---|---|
| `DevExpress.Spreadsheet.v23.2.Core.dll` | 16 MB | Spreadsheet editing UI |
| `DevExpress.Pdf.v23.2.Core.dll` | 5 MB | PDF viewing in UI (PdfiumViewer) |
| `Infragistics4.Documents.Excel.v13.2.dll` | 2 MB | Excel reading during definition |
| `Infragistics4.Documents.Excel.v23.1.dll` | 4.4 MB | Excel reading during definition |
| `AHPDFLib10.dll` | 4 MB | PDF manipulation (watermark, merge) |
| `PDFBaseLib32.dll` | 3.8 MB | PDF base operations |

---

## Registry Search

| Path | Findings |
|---|---|
| `HKLM\SOFTWARE\WOW6432Node\CIMTOPS CORPORATION` | **Key does not exist** |
| `HKCU\SOFTWARE\CIMTOPS CORPORATION\ConMas Generator` | Contains only `DesktopFolder` and `ProgramMenuFolder` installer paths — no engine configuration |
| `HKCU\SOFTWARE\CIMTOPS CORPORATION\ConMas i-Reporter for Windows` | Contains only installer paths — no engine configuration |

**Conclusion:** No registry-based engine selection mechanism exists.

---

## Deployment Implications

### Server (ConMasAPI / Linux Container)
- **No Excel required** — PDFs generated via Syncfusion + PDFNet
- Already deployable without Office
- Existing Python render service (`render_service/`) is compatible

### WPF Designer Client (Windows machines)
- **Excel IS required** for PDF export feature
- Without Excel → `"NoExcelInstalled"` exception
- DevExpress and Infragistics libraries present but unused for PDF rendering
- Potential future refactor: replace 50 lines of COM Interop with DevExpress.Spreadsheet `ExportToPdf()`

---

## Summary Diagram

```
┌────────────────────────────────────────────────────────┐
│                   ConMas PDF Rendering                  │
├─────────────────────────┬──────────────────────────────┤
│   Server (ConMasAPI)    │   Designer (WPF Client)       │
├─────────────────────────┼──────────────────────────────┤
│                         │                               │
│  Cimtops.ConMasBiz.dll  │  LibExcelController.dll       │
│       │                 │       │                       │
│       ▼                 │       ▼                       │
│  Syncfusion.Pdf.Base    │  Microsoft.Office.Interop.Excel
│  + PDFNet               │  (ExportAsFixedFormat)        │
│       │                 │       │                       │
│       ▼                 │       ▼                       │
│  ✔ PDF generated       │  ✔ PDF generated              │
│  ✔ NO Excel required   │  ✘ EXCEL REQUIRED             │
│                         │                               │
└─────────────────────────┴──────────────────────────────┘
```

---

## Files Examined

| File | Source |
|---|---|
| `C:\ConMas\ConMasAPI\bin\Cimtops.ConMasBiz.dll` | ILDASM decompiled |
| `C:\ConMas\ConMasAPI\bin\ConMasAPI.dll` | ILDASM decompiled |
| `C:\ConMas\ConMasAPI\bin\*.dll` | Full directory inventory |
| `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Designer\bin\LibExcelController.dll` | ILDASM decompiled (prev. session) |
| `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasGenerator.exe` | ILDASM decompiled |
| `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasGeneratorLib.dll` | ILDASM decompiled |
| `C:\Program Files (x86)\CIMTOPS CORPORATION\ConMas Generator\ConMasGeneratorUtility.exe` | ILDASM decompiled |
| Registry: `HKLM\SOFTWARE\WOW6432Node\CIMTOPS CORPORATION` | Searched |
| Registry: `HKCU\SOFTWARE\CIMTOPS CORPORATION` | Searched |
| Registry: `HKCR\Excel.Application`, `CLSID\{00024500-...}` | Checked (prev. session) |
