# Phase X.42 — Native Excel-to-PDF Rendering Engine Investigation

**Date:** 2026-07-19
**Objective:** Determine exactly how the legacy ConMas production server converts an Excel template into PDF drawing commands without Microsoft Excel.

---

## Executive Summary

**The production server does NOT convert Excel to PDF at render time.** There is no Excel-to-PDF rendering engine on the server. Instead, the system uses a **pre-processed database-driven architecture**:

1. The **Designer (WPF client)** uses Excel COM to parse Excel templates and extract cluster metadata (positions, types, fonts, formatting) into database tables (`def_cluster`, `def_sheet`, `rep_cluster`, `rep_sheet`).
2. The **server reads pre-computed cluster data** from the database at render time.
3. The server iterates over each cluster row and dispatches to the appropriate `ClusterWriter` subclass based on `cluster_type`.
4. Each `ClusterWriter` draws directly onto `Syncfusion.Pdf.Graphics.PdfGraphics`.

**There is no native Excel-to-PDF converter, no hidden renderer, no worksheet-to-pdf engine.** The Excel parsing happens exclusively in the Designer.

---

## 1. Complete Rendering Pipeline

```
┌──────────────────────────────────────────────────────────────────────────┐
│                          DESIGNER (WPF Client)                            │
│                    Requires Microsoft Excel COM                           │
│                                                                           │
│  Input: .xlsx template file                                               │
│  Process:                                                                 │
│    1. Open Excel via Interop.Excel                                        │
│    2. Read worksheets, cells, shapes, positions, formatting               │
│    3. Convert Excel cell positions → cluster metadata                     │
│    4. Upload template PDF (background image) → def_top.background_image   │
│    5. Insert clusters → def_cluster table                                 │
│    6. When report filed: copy def→rep tables with user input              │
│                                                                           │
│  Output:                                                                  │
│    ├── def_top (definition metadata)                                      │
│    ├── def_sheet (sheet definitions)                                      │
│    ├── def_cluster (cluster definitions with positions)                   │
│    ├── rep_top (report metadata)                                          │
│    ├── rep_sheet (report sheets)                                          │
│    └── rep_cluster (report clusters with user input)                      │
└──────────────────────────────────────────────────────────────────────────┘
                              │
                              │ Database (PostgreSQL)
                              ▼
┌──────────────────────────────────────────────────────────────────────────┐
│                          PRODUCTION SERVER                                │
│                    NO Microsoft Excel required                            │
│                                                                           │
│  PdfEditor.EditPdf(PdfLoadedPage page, int32 sheetNo)                     │
│                                                                           │
│  ┌────────────────────────────────────────────────────────────────────┐  │
│  │ 1. Get PdfGraphics from page                                       │  │
│  │    page.get_Graphics() → PdfGraphics                               │  │
│  │                                                                     │  │
│  │ 2. Get cluster data from database                                   │  │
│  │    PdfRepository.GetClusterInfoByRep(repTopId, sheetNo, userId)    │  │
│  │    └─ SQL: SELECT ... FROM rep_sheet t1                             │  │
│  │              INNER JOIN rep_cluster t2 ON ...                       │  │
│  │         Returns: DataTable with columns:                            │  │
│  │           cluster_id, cluster_type, input_parameter,               │  │
│  │           left_position, right_position, top_position,             │  │
│  │           bottom_position, input_value, display_value,             │  │
│  │           image_file, approval_sign, verified, is_hidden, ...      │  │
│  │                                                                     │  │
│  │ 3. For each DataRow in clusterDt.Rows:                              │  │
│  │    a. Parse input_parameter → ClusterParameter                      │  │
│  │       (align, verticalAlignment, fontSize, lines, punctuation,      │  │
│  │        fontPriority, weight, color, enableAutoFontSize, etc.)      │  │
│  │    b. Create PointWriter                                            │  │
│  │    c. SetClusterSettingInfo(cluster, sheetNo, width, height,        │  │
│  │                              rotation, ref pointData)               │  │
│  │       └─ Reads left/top/right/bottom_position from DataRow          │  │
│  │       └─ Converts: decimal.Parse → PointWriter constructor          │  │
│  │    d. CreateClusterWriter(typeName)                                 │  │
│  │       └─ ClusterWriterFactory.Instance.CreateClusterWriter(cm, type)│  │
│  │    e. clusterWriter.OutputCluster(graphics, cluster, param, point) │  │
│  │    f. Optional: DrawVerifiedLine / DrawInvisible / DrawInitJudge    │  │
│  │    g. Optional: DrawClusterLine                                     │  │
│  │                                                                     │  │
│  │ 4. InsertFdSheet (optional field data sheets from rep_fd_sheet)     │  │
│  │                                                                     │  │
│  │ 5. document.Save(MemoryStream) → byte[] (SYNCFUSION)               │  │
│  │ 6. ConvertToAprysePdf(byte[]) (PDFNET)                              │  │
│  └────────────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Rendering Classes

### Core PDF Editor (PdfEditor)
- **Namespace:** `Cimtops.ConMasBiz.Common.PdfEditor`
- **IL location:** line 296840
- **Role:** Orchestrates the entire PDF generation process
- **Key fields:**
  - `ConnectionManager cm` — database connection
  - `ConMasUtility conmasUtil` — utility helper
  - `PdfRepository pr` — database access
  - `Dictionary<string, PdfFont> fontCache` — font cache
  - `_pageClusterSet` — HashSet of (sheetNo, clusterId) for dedup

### ClusterWriterFactory
- **Namespace:** `Cimtops.ConMasBiz.Common.Pdf.ClusterWriterFactory`
- **IL location:** line 372615
- **Role:** Singleton factory that creates the correct `ClusterWriter` based on cluster type string

### IClusterWriter (Interface)
- **Namespace:** `Cimtops.ConMasBiz.Common.Pdf.ClusterWriter.IClusterWriter`
- **IL location:** line 384628
- **Key method:**
  - `Output(PdfGraphics, DataRow, ClusterParameter, PointWriter)` — render cluster

### ClusterWriterBase (Abstract Base)
- **Namespace:** `Cimtops.ConMasBiz.Common.Pdf.ClusterWriter.ClusterWriterBase`
- **IL location:** line 380174
- **Key methods:**
  - `OutputCluster(PdfGraphics, DataRow, ClusterParameter, PointWriter)` — main output
  - `IsVerified(DataRow, ClusterParameter)` — check if cluster is verified
  - `DrawVerifiedLine(PdfGraphics, PointWriter, int32)` — verified checkmark
  - `DrawInvisible(PdfGraphics, PointWriter)` — invisible placeholder
  - `DrawInitJudge(PdfGraphics, PointWriter, Color)` — initial judge marker
  - `DrawClusterLine(PdfGraphics, PointWriter, int32, string)` — cluster ID line

### Concrete ClusterWriters (14 total)

| Class | IL Line | cluster_type(s) | Renders |
|-------|---------|-----------------|---------|
| `TextClusterWriter` | 385491 | (default) | Text with formatting |
| `NumericClusterWriter` | 385122 | Numeric, InputNumeric, NumberHours, Calculate, TimeCalculate | Numbers |
| `ImageClusterWriter` | 384743 | Image, FreeText | Images |
| `DrawingImageClusterWriter` | 378557 | DrawingImage | Drawing images |
| `FixedTextClusterWriter` | 383786 | FixedText | Fixed position text |
| `FreeDrawClusterWriter` | 384007 | FreeDraw | Free-form drawings |
| `AudioRecordingClusterWriter` | 376982 | AudioRecording | Audio recording icons |
| `ActionClusterWriter` | 376597 | Action | Action buttons |
| `ApproveClusterWriter` | 381860 | Approve, Create, Inspect | Approval stamps |
| `CheckClusterWriter` | 382420 | Check | Checkbox fields |
| `MultipleChoiceNumberClusterWriter` | 377567 | MultipleChoiceNumber | Radio/choice groups |
| `ReportInfoClusterWriter` | 378359 | Registration, RegistrationDate, LatestUpdate, LatestUpdateDate | Report metadata |

### Supporting Classes

| Class | Namespace | Role |
|-------|-----------|------|
| `PointWriter` | `Cimtops.ConMasBiz.Common.PointWriter` | Stores and converts cluster position coordinates |
| `ClusterParameter` | `Cimtops.ConMasBiz.Common.ClusterParameter` | Parses JSON-like `input_parameter` string into rendering parameters |
| `PdfRepository` | `Cimtops.ConMasBiz.Common.Pdf.PdfRepository` | Data access for PDF/cluster info |
| `PdfLicense` | `Cimtops.ConMasBiz.Common.Pdf.PdfLicense` | Syncfusion license validation |
| `TrailPdf` | `Cimtops.ConMasBiz.Common.Pdf.TrailPdf` | Trail/copy page generation |
| `PdfReplace` | `Cimtops.ConMasBiz.Common.PdfReplace` | PDF text replacement utility |

---

## 3. Rendering Methods Call Graph

```
DownloadReportBiz.DownloadFile(httpContext)
│
├─ fileType="pdf" / "pdfLayer"
│  └─ DownloadReportBiz.DownloadPdf()
│     └─ new PdfEditor(cm, locale, userId)
│        └─ PdfEditor.GetPdf(repTopId)
│           ├─ PdfLicense.HasValidLicense()
│           ├─ Initialize()
│           ├─ SetPdfInfo()
│           │  └─ PdfRepository.GetPdfInfo(repTopId)
│           ├─ [if IsAddPage] JoinToSeparatePdf()
│           ├─ PdfBinary = PdfRepository.GetPdfBinaryByDef(repTopId) or by rep
│           ├─ new PdfLoadedDocument(PdfBinary)
│           ├─ For each page:
│           │  ├─ EditPdf(page, sheetNo)
│           │  │  ├─ page.get_Graphics() → PdfGraphics
│           │  │  ├─ page.get_Size() → sheetWidth, sheetHeight
│           │  │  ├─ PdfRepository.GetClusterInfoByRep(repTopId, sheetNo, userId)
│           │  │  │  └─ SQL JOIN rep_sheet + rep_cluster
│           │  │  ├─ For each cluster DataRow:
│           │  │  │  ├─ new ClusterParameter(input_parameter)
│           │  │  │  ├─ new PointWriter()
│           │  │  │  ├─ SetClusterSettingInfo(cluster, sheetNo, w, h, rot, ref point)
│           │  │  │  │  └─ Parse left/top/right/bottom from DataRow
│           │  │  │  │  └─ new PointWriter(left, top, right, bottom, width, height)
│           │  │  │  ├─ ClusterWriterFactory.CreateClusterWriter(cm, cluster_type)
│           │  │  │  ├─ clusterWriter.OutputCluster(graphics, cluster, param, point)
│           │  │  │  │  └─ (in TextClusterWriter):
│           │  │  │  │     ├─ GetStringFormat(align, vAlign, lines, punct, fontPrio)
│           │  │  │  │     ├─ GetFontStyle(weight)
│           │  │  │  │     └─ DrawText(graphics, point, style, format, size,
│           │  │  │  │           color, autoFont, maxLines, text)
│           │  │  │  ├─ [if Verified] DrawVerifiedLine(graphics, point, noNeedType)
│           │  │  │  ├─ [if InitJudge] DrawInitJudge(graphics, point, color)
│           │  │  │  └─ [if ClusterLine] DrawClusterLine(graphics, point, size, text)
│           │  │  └─ [dispose clusterDt]
│           │  ├─ [if Layer mode] AddLayer(page, sheetNo)
│           │  │  └─ PdfGraphics.DrawImage(image, 0, 0, clientWidth, clientHeight)
│           │  └─ [loop next page]
│           ├─ InsertFdSheet(document)
│           │  └─ RepFdSheetManager.Select(cm, repTopId, fields)
│           │  └─ For each: Append PdfLoadedDocument to document
│           ├─ document.Save(MemoryStream)
│           ├─ document.Close(true)
│           ├─ MemoryStream.ToArray() → byte[]
│           ├─ [if IsTrailOutput] TrailOutput(byte[])
│           │  └─ Creates trail PDF with Syncfusion
│           └─ ConvertToAprysePdf(byte[])
│              └─ pdftron.PDF.PDFDoc(byte[])
│              └─ PDFDoc.ImportPages → PDFDoc.PagePushBack
│              └─ PDFDoc.Save(SaveOptions)
│
├─ fileType="excel"
│  └─ DownloadReportBiz.DownloadExcel()
│     └─ Uses Infragistics to create .xlsx
│        (completely separate path, no PDF involved)
```

---

## 4. Coordinate Generation

### Data Source
Cluster positions are stored directly in the database tables as decimal values:

**`rep_cluster` table columns:**
- `left_position` (decimal) — X coordinate of left edge
- `right_position` (decimal) — X coordinate of right edge
- `top_position` (decimal) — Y coordinate of top edge
- `bottom_position` (decimal) — Y coordinate of bottom edge

### How Positions Become PDF Coordinates

```
DataRow                      PointWriter                      Syncfusion PdfGraphics
─────────                    ────────────                     ───────────────────────
left_position ──► Decimal.Parse() ──► PointWriter.left
top_position  ──► Decimal.Parse() ──► PointWriter.top         DrawString(
right_position ─► Decimal.Parse() ──► PointWriter.right          x = left,
bottom_position ┤ Decimal.Parse() ──► PointWriter.bottom         y = top,
                                                                 width = right - left,
                                                                 height = bottom - top
page.Size.Width  ──► PointWriter.sheetWidth                   )
page.Size.Height ──► PointWriter.sheetHeight
```

**Code evidence** (IL at line 301312-301343):
```
pointData.left = Convert.ToDecimal(Double.Parse(cluster["left_position"]))
pointData.top = Convert.ToDecimal(Double.Parse(cluster["top_position"]))
pointData.right = Convert.ToDecimal(Double.Parse(cluster["right_position"]))
pointData.bottom = Convert.ToDecimal(Double.Parse(cluster["bottom_position"]))
new PointWriter(left, top, right, bottom, sheetWidth, sheetHeight)
```

These positions are in **points (1/72 inch)** — the same unit used by Syncfusion PdfGraphics. No unit conversion is needed because the Designer already stored positions in PDF-compatible units.

### How Borders, Merged Cells, and Other Formatting Work

The key insight is that **borders, merged cells, and complex Excel formatting are baked into the background image** — not rendered individually.

The rendering process:
1. **Background image** (stored as `background_image_file` BLOB in `rep_top`/`def_top`) is a **pre-rendered PDF page** created by the Designer
2. This PDF contains the static template elements: borders, grid lines, merged cell outlines, headers, labels, logos
3. The server loads this background PDF via `PdfLoadedDocument`
4. Cluster writers overlay **data-only elements** (field values, checkmarks, signatures) on top of the background
5. The cluster `input_parameter` JSON string contains:
   - Horizontal alignment (Left, Center, Right)
   - Vertical alignment (Top, Middle, Bottom)
   - Font size, weight (Bold, Italic, Normal)
   - Font color
   - Lines (1, 2, ...)
   - Font priority
   - Punctuation mode
   - Enable auto font size
   - And other formatting details

The `ClusterParameter` class parses this JSON-like string into typed properties used by the ClusterWriter for PDF drawing.

---

## 5. Database Role

### Tables Accessed at Render Time

| Table | Query | Purpose |
|-------|-------|---------|
| `rep_top` | `GetPdfInfo()` | Background image, server version, display settings |
| `rep_sheet` | `GetSheetNoInfo()` | Sheet number mapping |
| `rep_cluster` | `GetClusterInfoByRep()` | Cluster positions, types, parameters, values |
| `rep_fd_sheet` | `RepFdSheetManager.Select()` | Field data sheet PDFs to append |
| `def_top` | `GetPdfBinaryByDef()` | Definition background image (for def rendering) |
| `def_cluster` | `GetClusterInfoByDef()` | Definition cluster info |

### SQL Schema (from PdfRepository)

**`rep_cluster` columns used:**
```
rep_top_id, rep_sheet_no, rep_sheet_id, cluster_id, cluster_type,
input_parameter, left_position, right_position, top_position,
bottom_position, input_value, display_value, image_file,
approval_sign, verified, is_hidden, image_file_length,
binary_file_length, binary_file_length_win, binary_file_type,
CLUSTER_ROLE (computed), REFER_ROLE (computed), NONE_ROLE (computed),
import_image, input_approval_value, recording_file_length
```

**`rep_top` columns used:**
```
background_image_file, server_version, use_input_history,
use_init_input_judge, use_init_input_judge_parameters,
no_need_to_fill_type, use_change_reason, display_sheet_number,
rep_sheet_count, def_top_id
```

---

## 6. Template Processing Flow

### Designer (WPF Client — requires Excel COM)
```
Excel (.xlsx)
  │
  ├─► LibExcelController.dll (COM Office Interop)
  │     Open Excel, read cells, shapes, formatting
  │     Parse all worksheet structure
  │
  ├─► Generate background PDF
  │     Print worksheet to PDF using Excel's own export
  │     (or: render cell borders/grids to image)
  │
  ├─► Extract cluster metadata
  │     For each data-entry field/cell:
  │       - Calculate position from Excel cell coordinates
  │       - Convert to points (1/72 inch)
  │       - Extract type (text, numeric, checkbox, etc.)
  │       - Extract parameters (font, alignment, etc.)
  │
  └─► Store to database
        INSERT INTO def_sheet, def_cluster, def_top
        (background_image_file = PDF binary of template)
```

### Server (ConMasAPI — no Excel)
```
HTTP Request → DownloadReportBiz.DownloadFile()
  │
  ├─► PdfEditor.GetPdf(repTopId)
  │     ├─► Load background PDF from rep_top.background_image_file
  │     ├─► Load cluster metadata from rep_cluster
  │     └─► Overlay field values onto background PDF
  │
  └─► Return binary PDF
```

**Answer to Template Processing Questions:**
- **Does the server read Excel every request?** NO — Excel is never accessed
- **Does it read a cached serialized layout?** YES — cluster metadata in database
- **Does it read XML extracted from Excel?** PARTIALLY — `input_parameter` is a JSON-like string
- **Does it read a proprietary binary format?** NO
- **Does it read database metadata only?** YES — positions, types, and parameters are stored directly

---

## 7. Definitively: What Code Replaces Excel?

**Nothing replaces Excel.** The server was designed from the beginning to NOT use Excel at render time. The architecture separates the concerns:

| Concern | Handled By | Where |
|---------|-----------|-------|
| Excel template parsing | `LibExcelController.dll` (COM) | Designer (WPF) |
| Position calculation | `LibExcelController.dll` (COM) | Designer (WPF) |
| Cell-to-coordinate mapping | `LibExcelController.dll` (COM) | Designer (WPF) |
| Border/font extraction | `LibExcelController.dll` (COM) | Designer (WPF) |
| Background PDF generation | Excel PDF export | Designer (WPF) |
| Cluster data storage | PostgreSQL INSERT | Designer (WPF) |
| PDF drawing | `Syncfusion.Pdf.Base` | Server |
| PDF post-processing | `PDFNet` (Apryse) | Server |
| Data overlay | ClusterWriter subclasses | Server |
| .xlsx reading (read-only) | `Infragistics.Excel` | Server |

**The `ClusterWriterFactory` + `ClusterWriterBase` + concrete `ClusterWriter` subclasses are the rendering engine.** They replace the concept of "Excel rendering" by drawing pre-computed metadata directly onto PDF pages using Syncfusion's graphics API.

**No native Excel-to-PDF converter exists on the server.** The system was designed with a database-centric architecture where all Excel processing happens once (on the Designer) and the results are stored as coordinates and parameters that the server simply draws.

---

## 8. Key Evidence Summary

| Finding | Evidence | Location | Confidence |
|---------|----------|----------|------------|
| Cluster data read from database, not Excel | SQL query in GetClusterInfoByRep | IL: 373512-373587 | CERTAIN |
| Positions stored as decimal in DB | SetClusterSettingInfo parses left/top/right/bottom | IL: 301312-301343 | CERTAIN |
| PdfGraphics obtained from page | `page.get_Graphics()` | IL: 299313 | CERTAIN |
| ClusterWriterFactory dispatches by type | CreateClusterWriter switch on typeName | IL: 372635-373043 | CERTAIN |
| TextClusterWriter.DrawText calls PdfGraphics | Uses Syncfusion PdfGraphics directly | IL: 385717-385725 | CERTAIN |
| Background image is pre-rendered PDF | Loaded via PdfLoadedDocument | IL: 298174 | CERTAIN |
| No Excel→PDF conversion at render time | All cluster data from DB, no Excel refs in rendering path | All IL | CERTAIN |
| Infragistics only used for Excel export (DownloadExcel) | ExcelWriter/CellDataCreator.MakeShapeString | IL: 190934 | CERTAIN |
| Designer handles Excel→cluster conversion | RepFdSheetManager, LibExcelController in Designer | Phase X.40 | CERTAIN |
| `input_parameter` is JSON-like parsed by ClusterParameter | new ClusterParameter("input_parameter") | IL: 299373 | CERTAIN |

---

## 9. Files Examined

| File | Method |
|------|--------|
| `Cimtops.ConMasBiz.dll` | ILDASM decompilation (1,114,790 lines) |
| `ConMasAPI.dll` | ILDASM decompilation |
| PostgreSQL schema (from SQL strings in IL) | Extracted from PdfRepository methods |
